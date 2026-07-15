"""High-level extraction: pull Chargebee resources into pandas DataFrames.

Each function returns a normalized DataFrame with unix timestamps converted to
tz-aware datetimes and money amounts converted from minor units to major units
(new `*_major` columns), leaving the raw integer columns intact.
"""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

import pandas as pd

from .client import ChargebeeClient, enum_in, ts_after
from .currency import minor_to_major

# Columns that hold unix-epoch-second timestamps across Chargebee objects.
_TS_COLUMNS = {
    "created_at", "updated_at", "activated_at", "started_at", "cancelled_at",
    "next_billing_at", "current_term_start", "current_term_end", "paused_at",
    "resumed_at", "trial_start", "trial_end", "date", "paid_at", "due_date",
    "generated_at", "expiry_date", "issued_credit_note_date", "occurred_at",
    "start_date", "end_date",
}

# (amount_column, currency_column) pairs to convert to major units.
_MONEY_COLUMNS = {
    "amount", "mrr", "total", "sub_total", "amount_paid", "amount_due",
    "amount_adjusted", "credits_applied", "amount_to_collect", "tax",
    "plan_amount", "plan_unit_price", "unit_price", "discount_amount",
    "round_off_amount", "fractional_correction",
}


def _timestamps_to_datetime(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        if col in _TS_COLUMNS and pd.api.types.is_numeric_dtype(df[col]):
            df[col] = pd.to_datetime(df[col], unit="s", utc=True)
    return df


def _add_major_amounts(df: pd.DataFrame) -> pd.DataFrame:
    currency_col = "currency_code" if "currency_code" in df.columns else None
    for col in list(df.columns):
        if col in _MONEY_COLUMNS:
            if currency_col:
                df[f"{col}_major"] = [
                    minor_to_major(a, c)
                    for a, c in zip(df[col], df[currency_col])
                ]
            else:
                df[f"{col}_major"] = [minor_to_major(a, None) for a in df[col]]
    return df


def normalize(records: list[dict[str, Any]]) -> pd.DataFrame:
    """Turn a list of raw Chargebee dicts into a normalized DataFrame."""
    if not records:
        return pd.DataFrame()
    df = pd.json_normalize(records, sep=".")
    df = _timestamps_to_datetime(df)
    df = _add_major_amounts(df)
    return df


# ---------------------------------------------------------------------------
# per-resource pulls
# ---------------------------------------------------------------------------


def get_customers(client: ChargebeeClient, **kw) -> pd.DataFrame:
    return normalize(client.list_all("Customer", **kw))


def get_subscriptions(
    client: ChargebeeClient,
    statuses: list[str] | None = None,
    **kw,
) -> pd.DataFrame:
    params = kw.pop("params", {})
    if statuses:
        params = {**params, **enum_in("status", statuses)}
    return normalize(client.list_all("Subscription", params=params, **kw))


def get_invoices(
    client: ChargebeeClient,
    statuses: list[str] | None = None,
    **kw,
) -> pd.DataFrame:
    params = kw.pop("params", {})
    if statuses:
        params = {**params, **enum_in("status", statuses)}
    return normalize(client.list_all("Invoice", params=params, **kw))


def get_transactions(client: ChargebeeClient, **kw) -> pd.DataFrame:
    return normalize(client.list_all("Transaction", **kw))


def get_credit_notes(client: ChargebeeClient, **kw) -> pd.DataFrame:
    return normalize(client.list_all("CreditNote", **kw))


def get_items(client: ChargebeeClient, **kw) -> pd.DataFrame:
    """PC 2.0 items."""
    return normalize(client.list_all("Item", **kw))


def get_item_prices(client: ChargebeeClient, **kw) -> pd.DataFrame:
    """PC 2.0 item prices."""
    return normalize(client.list_all("ItemPrice", item_key="item_price", **kw))


def get_plans(client: ChargebeeClient, **kw) -> pd.DataFrame:
    """PC 1.0 plans."""
    return normalize(client.list_all("Plan", **kw))


def get_addons(client: ChargebeeClient, **kw) -> pd.DataFrame:
    """PC 1.0 addons."""
    return normalize(client.list_all("Addon", **kw))


def get_events(client: ChargebeeClient, **kw) -> pd.DataFrame:
    return normalize(client.list_all("Event", **kw))


def get_products(client: ChargebeeClient) -> pd.DataFrame:
    """Return the catalog as a DataFrame regardless of PC version."""
    version = client.product_catalog_version()
    if version == "2.0":
        df = get_items(client)
    else:
        df = get_plans(client)
    if not df.empty:
        df["_catalog_version"] = version
    return df


def pull_all(
    client: ChargebeeClient,
    updated_after: datetime | int | None = None,
) -> dict[str, pd.DataFrame]:
    """Convenience: pull the core resources into a dict of DataFrames.

    `updated_after` (a datetime or epoch seconds) restricts customers,
    subscriptions and invoices to those changed since then — useful for
    incremental syncs.
    """
    inc: dict[str, Any] = {}
    if updated_after is not None:
        epoch = (
            int(updated_after.timestamp())
            if isinstance(updated_after, datetime)
            else int(updated_after)
        )
        inc = {"params": ts_after("updated_at", epoch)}

    version = client.product_catalog_version()
    result = {
        "customers": get_customers(client, **inc),
        "subscriptions": get_subscriptions(client, **inc),
        "invoices": get_invoices(client, **inc),
        "transactions": get_transactions(client),
        "credit_notes": get_credit_notes(client),
        "products": get_products(client),
    }
    result["_catalog_version"] = version  # type: ignore[assignment]
    return result


def now_epoch() -> int:
    return int(datetime.now(timezone.utc).timestamp())
