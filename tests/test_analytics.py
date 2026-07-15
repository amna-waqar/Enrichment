"""Offline unit tests for the analytics + currency + client-helper logic.

These use synthetic Chargebee-shaped records (no network) so the maths can be
verified deterministically. Run with:  python -m pytest -q
"""

from __future__ import annotations

import json

import pandas as pd

from chargebee_analytics import analytics, extract
from chargebee_analytics.currency import currency_exponent, minor_to_major
from chargebee_analytics.client import enum_in, ts_between, merge


# ---------------------------------------------------------------------------
# currency
# ---------------------------------------------------------------------------


def test_currency_exponents():
    assert currency_exponent("USD") == 2
    assert currency_exponent("JPY") == 0
    assert currency_exponent("KWD") == 3
    assert currency_exponent(None) == 2


def test_minor_to_major():
    assert minor_to_major(1000, "USD") == 10.0
    assert minor_to_major(1000, "JPY") == 1000.0   # zero-decimal: no division
    assert minor_to_major(1000, "KWD") == 1.0      # three-decimal
    assert minor_to_major(None, "USD") is None


# ---------------------------------------------------------------------------
# filter builders
# ---------------------------------------------------------------------------


def test_enum_in_json_encodes_list():
    f = enum_in("status", ["active", "non_renewing"])
    assert f == {"status": {"in": json.dumps(["active", "non_renewing"])}}
    # must be valid JSON (not a python repr with single quotes)
    assert json.loads(f["status"]["in"]) == ["active", "non_renewing"]


def test_merge_deep_merges_operators():
    m = merge(ts_between("created_at", 1, 2), enum_in("status", ["active"]))
    assert "created_at" in m and "status" in m


# ---------------------------------------------------------------------------
# fixtures
# ---------------------------------------------------------------------------


def _subscriptions():
    # Two USD monthly active, one annual active, one cancelled, one JPY active.
    records = [
        {"id": "s1", "status": "active", "currency_code": "USD", "mrr": 5000,
         "activated_at": 1704067200,  # 2024-01-01
         "cancelled_at": None},
        {"id": "s2", "status": "active", "currency_code": "USD",
         "plan_amount": 120000, "plan_quantity": 1,
         "billing_period": 1, "billing_period_unit": "year",  # $1200/yr -> $100/mo
         "activated_at": 1706745600,  # 2024-02-01
         "cancelled_at": None},
        {"id": "s3", "status": "non_renewing", "currency_code": "USD", "mrr": 2500,
         "activated_at": 1704067200, "cancelled_at": None},
        {"id": "s4", "status": "cancelled", "currency_code": "USD", "mrr": 0,
         "plan_amount": 3000, "plan_quantity": 1,
         "billing_period": 1, "billing_period_unit": "month",  # $30/mo
         "activated_at": 1704067200,  # 2024-01-01
         "cancelled_at": 1709251200},  # 2024-03-01
        {"id": "s5", "status": "active", "currency_code": "JPY", "mrr": 5000,
         "activated_at": 1706745600, "cancelled_at": None},
    ]
    return extract.normalize(records)


def _invoices():
    records = [
        {"id": "i1", "status": "paid", "currency_code": "USD",
         "amount_paid": 5000, "total": 5000, "paid_at": 1704153600},  # 2024-01-02
        {"id": "i2", "status": "paid", "currency_code": "USD",
         "amount_paid": 10000, "total": 10000, "paid_at": 1706832000},  # 2024-02-02
        {"id": "i3", "status": "paid", "currency_code": "USD",
         "amount_paid": 7500, "total": 7500, "paid_at": 1706918400},  # 2024-02-03
    ]
    return extract.normalize(records)


# ---------------------------------------------------------------------------
# extract normalization
# ---------------------------------------------------------------------------


def test_normalize_converts_timestamps_and_money():
    df = _subscriptions()
    assert pd.api.types.is_datetime64_any_dtype(df["activated_at"])
    # s1 mrr 5000 cents -> $50
    s1 = df[df["id"] == "s1"].iloc[0]
    assert s1["mrr_major"] == 50.0
    # JPY 5000 -> 5000 (no division)
    s5 = df[df["id"] == "s5"].iloc[0]
    assert s5["mrr_major"] == 5000.0


# ---------------------------------------------------------------------------
# MRR
# ---------------------------------------------------------------------------


def test_mrr_summary():
    df = _subscriptions()
    summary = analytics.mrr_summary(df)
    usd = summary[summary["currency_code"] == "USD"].iloc[0]
    # live USD subs: s1 $50 + s2 $100 (annual normalized) + s3 $25 = $175
    assert usd["mrr"] == 175.0
    assert usd["arr"] == 2100.0
    assert usd["subscriptions"] == 3
    # JPY: s5 mrr 5000 -> 5000/mo
    jpy = summary[summary["currency_code"] == "JPY"].iloc[0]
    assert jpy["mrr"] == 5000.0


def test_annual_normalization():
    df = _subscriptions()
    enriched = analytics.enrich_mrr(df)
    s2 = enriched[enriched["id"] == "s2"].iloc[0]
    assert round(s2["computed_mrr_major"], 2) == 100.0  # $1200/yr -> $100/mo


# ---------------------------------------------------------------------------
# revenue
# ---------------------------------------------------------------------------


def test_revenue_by_month():
    inv = _invoices()
    rev = analytics.revenue_by_month(inv)
    jan = rev[rev["month"] == pd.Timestamp("2024-01-01")].iloc[0]
    feb = rev[rev["month"] == pd.Timestamp("2024-02-01")].iloc[0]
    assert jan["revenue"] == 50.0
    assert feb["revenue"] == 175.0   # 100 + 75
    assert feb["invoices"] == 2


# ---------------------------------------------------------------------------
# churn / movement
# ---------------------------------------------------------------------------


def test_logo_churn_by_month():
    df = _subscriptions()
    churn = analytics.logo_churn_by_month(df, start="2024-01-01", end="2024-03-31")
    mar = churn[churn["month"] == pd.Timestamp("2024-03-01")].iloc[0]
    # s4 cancelled 2024-03-01
    assert mar["churned"] == 1
    # active at start of March: s1, s3, s4 activated Jan (s4 not yet cancelled
    # at the instant the month starts), s2 activated Feb, s5 activated Feb = 5
    assert mar["active_start"] == 5


def test_mrr_movement_by_month():
    df = _subscriptions()
    mv = analytics.mrr_movement_by_month(df, start="2024-01-01", end="2024-03-31")
    mar = mv[mv["month"] == pd.Timestamp("2024-03-01")].iloc[0]
    # s4 churned in March, computed mrr $30
    assert mar["churned_mrr"] == 30.0
    assert mar["net_new_mrr"] == -30.0


def test_status_breakdown():
    df = _subscriptions()
    b = analytics.subscription_status_breakdown(df)
    assert b[b["status"] == "active"]["count"].iloc[0] == 3
