"""End-to-end example: connect, pull, analyse, export.

Run from the repo root after filling in `.env`:

    python scripts/quickstart.py
"""

from __future__ import annotations

from chargebee_analytics import ChargebeeClient, analytics, extract
from chargebee_analytics.export import write_datasets


def main() -> None:
    client = ChargebeeClient()  # reads CHARGEBEE_API_KEY / CHARGEBEE_SITE from .env

    print("Verifying connection ...")
    print(" ", client.verify_connection())
    print("  Product catalog:", client.product_catalog_version())

    print("\nPulling core datasets (this can take a while for large sites) ...")
    datasets = extract.pull_all(client)
    for name, df in datasets.items():
        if hasattr(df, "__len__"):
            print(f"  {name:14s} {len(df):>6} rows")

    print("\nComputing reports ...")
    reports = analytics.summarize(datasets)
    print("\nMRR / ARR by currency:")
    print(reports["mrr"].to_string(index=False))
    print("\nRevenue by month (tail):")
    print(reports["revenue_by_month"].tail(6).to_string(index=False))

    print("\nWriting CSVs to ./data ...")
    written = write_datasets(
        {k: v for k, v in datasets.items() if hasattr(v, "empty")}, "data", "csv"
    )
    write_datasets(reports, "data/reports", "csv")
    print(f"  wrote {len(written)} raw dataset files + reports to ./data")


if __name__ == "__main__":
    main()
