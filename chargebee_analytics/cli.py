"""Command-line interface for the Chargebee analytics toolkit.

Examples:
    python -m chargebee_analytics.cli check
    python -m chargebee_analytics.cli report
    python -m chargebee_analytics.cli pull --out data --format csv
    python -m chargebee_analytics.cli pull subscriptions invoices --out data
"""

from __future__ import annotations

import argparse
import sys

import pandas as pd

from . import analytics, extract
from .client import ChargebeeClient
from .config import load_config
from .export import write_datasets, write_dataframe

pd.set_option("display.max_columns", 40)
pd.set_option("display.width", 160)

_RESOURCE_PULLERS = {
    "customers": extract.get_customers,
    "subscriptions": extract.get_subscriptions,
    "invoices": extract.get_invoices,
    "transactions": extract.get_transactions,
    "credit_notes": extract.get_credit_notes,
    "products": extract.get_products,
    "items": extract.get_items,
    "item_prices": extract.get_item_prices,
    "plans": extract.get_plans,
    "addons": extract.get_addons,
    "events": extract.get_events,
}


def _client() -> ChargebeeClient:
    return ChargebeeClient(load_config())


def cmd_check(args) -> int:
    client = _client()
    cfg = client.config
    print(f"Site:    {cfg.site}")
    print(f"Domain:  {cfg.resolved_domain}")
    print(f"Key:     {cfg.masked_key()} ({'live' if cfg.is_live else 'test'})")
    status = client.verify_connection()
    print(f"Catalog: PC {client.product_catalog_version()}")
    print(f"Connection OK — reachable, sample_count={status['sample_count']}")
    return 0


def cmd_pull(args) -> int:
    client = _client()
    names = args.resources or [
        "customers", "subscriptions", "invoices", "transactions",
        "credit_notes", "products",
    ]
    datasets: dict[str, pd.DataFrame] = {}
    for name in names:
        puller = _RESOURCE_PULLERS.get(name)
        if not puller:
            print(f"  ! unknown resource: {name}", file=sys.stderr)
            continue
        print(f"  pulling {name} ...", flush=True)
        try:
            datasets[name] = puller(client)
            print(f"    {len(datasets[name])} records")
        except Exception as exc:  # keep going on PC-mismatch etc.
            print(f"    skipped ({exc})", file=sys.stderr)
    if args.out:
        written = write_datasets(datasets, args.out, args.format)
        for path in written:
            print(f"  wrote {path}")
    return 0


def cmd_report(args) -> int:
    client = _client()
    print("Pulling subscriptions and invoices ...", flush=True)
    subs = extract.get_subscriptions(client)
    invoices = extract.get_invoices(client)
    reports = {
        "MRR / ARR by currency": analytics.mrr_summary(subs),
        "Subscription status": analytics.subscription_status_breakdown(subs),
        "Revenue by month": analytics.revenue_by_month(invoices),
        "Logo churn by month": analytics.logo_churn_by_month(subs),
        "MRR movement by month": analytics.mrr_movement_by_month(subs),
    }
    for title, df in reports.items():
        print(f"\n=== {title} ===")
        print(df.to_string(index=False) if not df.empty else "(no data)")
    if args.out:
        for title, df in reports.items():
            slug = title.lower().replace(" / ", "_").replace(" ", "_")
            if not df.empty:
                write_dataframe(df, f"{args.out}/{slug}.{args.format}", args.format)
        print(f"\nReports written to {args.out}/")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="chargebee_analytics",
        description="Pull and analyse data from Chargebee.",
    )
    sub = p.add_subparsers(dest="command", required=True)

    c = sub.add_parser("check", help="verify the API key + site connect")
    c.set_defaults(func=cmd_check)

    pull = sub.add_parser("pull", help="pull raw resources to DataFrames/files")
    pull.add_argument("resources", nargs="*", help="resource names (default: core set)")
    pull.add_argument("--out", help="output directory")
    pull.add_argument("--format", default="csv", choices=["csv", "parquet", "json", "xlsx"])
    pull.set_defaults(func=cmd_pull)

    rep = sub.add_parser("report", help="compute and print MRR/churn/revenue reports")
    rep.add_argument("--out", help="also write each report to this directory")
    rep.add_argument("--format", default="csv", choices=["csv", "parquet", "json", "xlsx"])
    rep.set_defaults(func=cmd_report)

    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return args.func(args)
    except RuntimeError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
