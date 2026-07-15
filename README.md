# Enrichment — Chargebee data pull & analysis

A small Python toolkit to **pull data from your Chargebee account and analyse it
via code**: customers, subscriptions, invoices, transactions, credit notes and
the product catalog — plus computed **MRR / ARR, revenue-by-month, churn and
MRR-movement** reports.

Built on the official [`chargebee`](https://pypi.org/project/chargebee/) v3 SDK
(client-based) with automatic pagination and retry/backoff, and
[`pandas`](https://pandas.pydata.org/) for analysis and export.

---

## 1. Setup

```bash
pip install -r requirements.txt          # chargebee, pandas, python-dotenv
cp .env.example .env                      # then edit .env
```

Fill in `.env`:

```
CHARGEBEE_API_KEY=live_xxxxxxxxxxxxxxxxxxxx   # read-only key is enough
CHARGEBEE_SITE=your-site                       # subdomain before .chargebee.com
# CHARGEBEE_REGION=eu                           # only for non-US clusters
```

> **Why the site name?** Chargebee's API lives at
> `https://{site}.chargebee.com/api/v2/`. The API key authenticates you but does
> **not** encode which site it belongs to, so the site (the subdomain you log in
> at) must be provided separately. `.env` is gitignored — your key never gets
> committed.

Verify it works:

```bash
python -m chargebee_analytics.cli check
```

```
Site:    your-site
Domain:  your-site.chargebee.com
Key:     live_xxx...XIMv (live)
Catalog: PC 2.0
Connection OK — reachable, sample_count=1
```

---

## 2. Command-line usage

```bash
# Verify connectivity + detect Product Catalog version
python -m chargebee_analytics.cli check

# Pull core resources to ./data as CSV (also: parquet, json, xlsx)
python -m chargebee_analytics.cli pull --out data --format csv

# Pull only specific resources
python -m chargebee_analytics.cli pull subscriptions invoices --out data

# Compute & print MRR / churn / revenue reports (optionally export them)
python -m chargebee_analytics.cli report --out data/reports
```

---

## 3. Library usage

```python
from chargebee_analytics import ChargebeeClient, extract, analytics

client = ChargebeeClient()                      # reads .env
print(client.verify_connection())
print("Catalog:", client.product_catalog_version())

# Pull everything into pandas DataFrames
data = extract.pull_all(client)                 # dict of DataFrames
subs = data["subscriptions"]
invoices = data["invoices"]

# Analyse
print(analytics.mrr_summary(subs))              # MRR & ARR per currency
print(analytics.revenue_by_month(invoices))     # collected revenue per month
print(analytics.logo_churn_by_month(subs))      # monthly customer churn
print(analytics.mrr_movement_by_month(subs))    # new vs churned MRR
```

Or run the full example:

```bash
python scripts/quickstart.py
```

### Incremental / filtered pulls

```python
from datetime import datetime, timezone
from chargebee_analytics import extract

# only records changed since a date (good for scheduled syncs)
since = datetime(2025, 1, 1, tzinfo=timezone.utc)
recent = extract.pull_all(client, updated_after=since)

# only active + non-renewing subscriptions
active = extract.get_subscriptions(client, statuses=["active", "non_renewing"])
```

---

## 4. What you get

| Module | Purpose |
|---|---|
| `chargebee_analytics.config` | Loads & validates `.env` credentials (`load_config`). |
| `chargebee_analytics.client` | `ChargebeeClient`: connection, retry, `list_all()` auto-pagination, PC-version detection, filter builders. |
| `chargebee_analytics.extract` | `get_customers/subscriptions/invoices/...` → normalized DataFrames; `pull_all()`. |
| `chargebee_analytics.analytics` | `mrr_summary`, `revenue_by_month`, `logo_churn_by_month`, `mrr_movement_by_month`, `summarize`. |
| `chargebee_analytics.currency` | Minor-unit → major-unit conversion (handles zero-/three-decimal currencies). |
| `chargebee_analytics.export` | `write_dataframe` / `write_datasets` to CSV / Parquet / Excel / JSON. |
| `chargebee_analytics.cli` | `check` / `pull` / `report` commands. |

### Normalization done for you

- **Timestamps** — Chargebee returns unix epoch seconds; these become tz-aware
  UTC datetimes.
- **Money** — amounts are integer minor units (cents). Each money column gets a
  `*_major` companion in real currency units, correctly handling zero-decimal
  (JPY, KRW…) and three-decimal (KWD, BHD…) currencies. Never sum across
  currencies without conversion — reports group by `currency_code`.
- **Product Catalog** — auto-detects 1.0 (plans/addons) vs 2.0
  (items/item_prices); `extract.get_products()` returns the right one.

---

## 5. Notes & limitations

- **Read-only** — this toolkit only lists/reads. It never writes to Chargebee.
- **Rate limits** — the client retries `429`/`5xx` with exponential backoff.
  Very large sites: prefer incremental pulls or Chargebee's Export API.
- **MRR/ARR/churn are computed here** — Chargebee has no general metrics REST
  endpoint. MRR uses each subscription's `mrr` field when present, else derives
  it from plan amount normalized to a monthly figure. Churn/movement are
  reconstructed from a point-in-time snapshot (`activated_at` / `cancelled_at`);
  they're a close approximation, not a full historical MRR ledger.

---

## 6. Tests

Logic is covered by offline unit tests (synthetic data, no network):

```bash
python -m pytest -q
```
