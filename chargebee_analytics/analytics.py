"""Revenue & retention analytics computed from Chargebee data.

Chargebee has no general-purpose "metrics" REST endpoint, so MRR / ARR / churn
are derived here from the subscription and invoice objects.

All functions accept the *normalized* DataFrames produced by `extract.py`
(tz-aware datetime columns, `*_major` money columns) and return DataFrames.

Caveat: subscription lists are a point-in-time snapshot. Movement metrics
(new / churned MRR by month) are reconstructed from `activated_at` /
`cancelled_at` plus each subscription's recurring monthly value, which is a
close approximation but not a substitute for a full historical MRR ledger.
"""

from __future__ import annotations

from typing import Iterable

import pandas as pd

from .currency import minor_to_major

# Subscriptions that contribute to live recurring revenue.
LIVE_STATUSES = ("active", "non_renewing")

# Average months per calendar unit, used to normalize any term to "per month".
_UNIT_MONTHS = {
    "day": 1.0 / 30.4375,
    "week": 7.0 / 30.4375,
    "month": 1.0,
    "year": 12.0,
}


# ---------------------------------------------------------------------------
# MRR / ARR
# ---------------------------------------------------------------------------


def _monthly_factor(period: float | None, unit: str | None) -> float:
    """Divisor that converts a per-term amount into a per-month amount."""
    period = float(period) if period else 1.0
    months = _UNIT_MONTHS.get((unit or "month").lower(), 1.0)
    denom = period * months
    return denom if denom > 0 else 1.0


def _row_recurring_mrr(row: pd.Series) -> float:
    """Best monthly recurring value for one subscription row, in major units.

    Preference order:
      1. Chargebee's own `mrr` field (already monthly, minor units) when > 0.
      2. Computed from `plan_amount` * `plan_quantity` (PC 1.0), normalized by
         the billing term.
      3. 0.0 when nothing usable is present.
    """
    currency = row.get("currency_code")

    mrr_raw = row.get("mrr")
    if pd.notna(mrr_raw) and float(mrr_raw) > 0:
        return minor_to_major(int(mrr_raw), currency) or 0.0

    plan_amount = row.get("plan_amount")
    if pd.notna(plan_amount) and float(plan_amount) > 0:
        qty = row.get("plan_quantity")
        qty = float(qty) if pd.notna(qty) else 1.0
        term_amount = minor_to_major(int(plan_amount), currency) or 0.0
        factor = _monthly_factor(row.get("billing_period"), row.get("billing_period_unit"))
        return (term_amount * qty) / factor

    return 0.0


def enrich_mrr(subscriptions: pd.DataFrame) -> pd.DataFrame:
    """Return a copy of the subscriptions DataFrame with `computed_mrr_major`."""
    if subscriptions.empty:
        return subscriptions.copy()
    df = subscriptions.copy()
    df["computed_mrr_major"] = df.apply(_row_recurring_mrr, axis=1)
    return df


def mrr_summary(
    subscriptions: pd.DataFrame,
    statuses: Iterable[str] = LIVE_STATUSES,
) -> pd.DataFrame:
    """Current MRR and ARR per currency, plus active subscription count."""
    empty = pd.DataFrame(columns=["currency_code", "subscriptions", "mrr", "arr"])
    if subscriptions.empty or "status" not in subscriptions:
        return empty
    df = enrich_mrr(subscriptions)
    live = df[df["status"].isin(list(statuses))].copy()
    if live.empty:
        return empty
    if "currency_code" not in live:
        live["currency_code"] = "unknown"
    else:
        live["currency_code"] = live["currency_code"].fillna("unknown")
    grp = (
        live.groupby("currency_code")
        .agg(subscriptions=("computed_mrr_major", "size"),
             mrr=("computed_mrr_major", "sum"))
        .reset_index()
    )
    grp["arr"] = (grp["mrr"] * 12).round(2)
    grp["mrr"] = grp["mrr"].round(2)
    return grp.sort_values("mrr", ascending=False).reset_index(drop=True)


# ---------------------------------------------------------------------------
# status breakdown
# ---------------------------------------------------------------------------


def subscription_status_breakdown(subscriptions: pd.DataFrame) -> pd.DataFrame:
    if subscriptions.empty or "status" not in subscriptions:
        return pd.DataFrame(columns=["status", "count"])
    out = (
        subscriptions.groupby("status")
        .size()
        .reset_index(name="count")
        .sort_values("count", ascending=False)
        .reset_index(drop=True)
    )
    return out


# ---------------------------------------------------------------------------
# revenue from invoices
# ---------------------------------------------------------------------------


def revenue_by_month(
    invoices: pd.DataFrame,
    date_col: str = "paid_at",
    amount_col: str = "amount_paid_major",
) -> pd.DataFrame:
    """Collected revenue per month per currency.

    Uses `paid_at` + `amount_paid` by default (cash basis). Pass
    `date_col="date"` / `amount_col="total_major"` for a billings view.
    """
    if invoices.empty or date_col not in invoices or amount_col not in invoices:
        return pd.DataFrame(columns=["month", "currency_code", "revenue", "invoices"])
    df = invoices.dropna(subset=[date_col]).copy()
    if df.empty:
        return pd.DataFrame(columns=["month", "currency_code", "revenue", "invoices"])
    # Drop tz before period bucketing (to_period discards tz anyway).
    months = df[date_col]
    if isinstance(months.dtype, pd.DatetimeTZDtype):
        months = months.dt.tz_localize(None)
    df["month"] = months.dt.to_period("M").dt.to_timestamp()
    currency = df["currency_code"] if "currency_code" in df else "unknown"
    out = (
        df.assign(currency_code=currency)
        .groupby(["month", "currency_code"])
        .agg(revenue=(amount_col, "sum"), invoices=("id", "count"))
        .reset_index()
        .sort_values(["month", "currency_code"])
        .reset_index(drop=True)
    )
    out["revenue"] = out["revenue"].round(2)
    return out


# ---------------------------------------------------------------------------
# churn & MRR movement
# ---------------------------------------------------------------------------


def _month_floor(series: pd.Series) -> pd.Series:
    return series.dt.to_period("M").dt.to_timestamp()


def _month_range(lo: pd.Timestamp, hi: pd.Timestamp) -> pd.DatetimeIndex:
    """Tz-aware (UTC) month-start range spanning [lo, hi] inclusive."""
    start = pd.Timestamp(year=lo.year, month=lo.month, day=1, tz="UTC")
    end = pd.Timestamp(year=hi.year, month=hi.month, day=1, tz="UTC")
    return pd.date_range(start, end, freq="MS", tz="UTC")


def _drop_tz(ts: pd.Timestamp) -> pd.Timestamp:
    """Return a tz-naive copy of a (possibly tz-aware) timestamp."""
    return ts.tz_convert(None) if ts.tzinfo is not None else ts


def logo_churn_by_month(
    subscriptions: pd.DataFrame,
    start: str | pd.Timestamp | None = None,
    end: str | pd.Timestamp | None = None,
) -> pd.DataFrame:
    """Monthly logo (customer-count) churn.

    For each month:
      active_start = subs activated before the month that were still active at
                     the month's start (not yet cancelled).
      churned      = subs whose `cancelled_at` falls within the month.
      churn_rate   = churned / active_start.
    """
    required = {"activated_at", "cancelled_at"}
    if subscriptions.empty or not required.issubset(subscriptions.columns):
        return pd.DataFrame(columns=["month", "active_start", "new", "churned", "churn_rate"])

    df = subscriptions.copy()
    activated = df["activated_at"]
    cancelled = df["cancelled_at"]

    lo = pd.Timestamp(start, tz="UTC") if start else activated.min()
    hi = pd.Timestamp(end, tz="UTC") if end else pd.Timestamp.now(tz="UTC")
    if pd.isna(lo):
        return pd.DataFrame(columns=["month", "active_start", "new", "churned", "churn_rate"])
    months = _month_range(lo, hi)

    rows = []
    for m_start in months:
        m_end = (m_start + pd.offsets.MonthBegin(1))
        active_start = int(((activated < m_start) &
                            (cancelled.isna() | (cancelled >= m_start))).sum())
        new = int(((activated >= m_start) & (activated < m_end)).sum())
        churned = int(((cancelled >= m_start) & (cancelled < m_end)).sum())
        rate = (churned / active_start) if active_start else 0.0
        rows.append({
            "month": _drop_tz(m_start),
            "active_start": active_start,
            "new": new,
            "churned": churned,
            "churn_rate": round(rate, 4),
        })
    return pd.DataFrame(rows)


def mrr_movement_by_month(
    subscriptions: pd.DataFrame,
    start: str | pd.Timestamp | None = None,
    end: str | pd.Timestamp | None = None,
) -> pd.DataFrame:
    """New vs churned MRR per month (approximate, snapshot-based).

    new_mrr     = Σ recurring MRR of subs activated in the month.
    churned_mrr = Σ recurring MRR of subs cancelled in the month.
    net_new_mrr = new_mrr - churned_mrr.
    """
    required = {"activated_at", "cancelled_at"}
    if subscriptions.empty or not required.issubset(subscriptions.columns):
        return pd.DataFrame(columns=["month", "new_mrr", "churned_mrr", "net_new_mrr"])

    df = enrich_mrr(subscriptions)
    activated = df["activated_at"]
    cancelled = df["cancelled_at"]
    mrr = df["computed_mrr_major"].fillna(0.0)

    lo = pd.Timestamp(start, tz="UTC") if start else activated.min()
    hi = pd.Timestamp(end, tz="UTC") if end else pd.Timestamp.now(tz="UTC")
    if pd.isna(lo):
        return pd.DataFrame(columns=["month", "new_mrr", "churned_mrr", "net_new_mrr"])
    months = _month_range(lo, hi)

    rows = []
    for m_start in months:
        m_end = m_start + pd.offsets.MonthBegin(1)
        in_new = (activated >= m_start) & (activated < m_end)
        in_churn = (cancelled >= m_start) & (cancelled < m_end)
        new_mrr = float(mrr[in_new].sum())
        churned_mrr = float(mrr[in_churn].sum())
        rows.append({
            "month": _drop_tz(m_start),
            "new_mrr": round(new_mrr, 2),
            "churned_mrr": round(churned_mrr, 2),
            "net_new_mrr": round(new_mrr - churned_mrr, 2),
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# top-level report
# ---------------------------------------------------------------------------


def summarize(datasets: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    """Build the standard set of reports from `extract.pull_all` output."""
    subs = datasets.get("subscriptions", pd.DataFrame())
    invoices = datasets.get("invoices", pd.DataFrame())
    return {
        "mrr": mrr_summary(subs),
        "status_breakdown": subscription_status_breakdown(subs),
        "revenue_by_month": revenue_by_month(invoices),
        "logo_churn": logo_churn_by_month(subs),
        "mrr_movement": mrr_movement_by_month(subs),
    }
