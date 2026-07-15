"""Persist DataFrames to disk (CSV / Parquet / Excel)."""

from __future__ import annotations

from pathlib import Path

import pandas as pd


def write_dataframe(df: pd.DataFrame, path: str | Path, fmt: str | None = None) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fmt = (fmt or path.suffix.lstrip(".") or "csv").lower()
    if fmt == "csv":
        df.to_csv(path, index=False)
    elif fmt in ("parquet", "pq"):
        df.to_parquet(path, index=False)
    elif fmt in ("xlsx", "excel"):
        df.to_excel(path, index=False)
    elif fmt == "json":
        df.to_json(path, orient="records", date_format="iso")
    else:
        raise ValueError(f"Unsupported export format: {fmt}")
    return path


def write_datasets(
    datasets: dict[str, pd.DataFrame],
    out_dir: str | Path,
    fmt: str = "csv",
) -> list[Path]:
    """Write each named DataFrame in a dict to `out_dir/<name>.<fmt>`."""
    out_dir = Path(out_dir)
    written: list[Path] = []
    for name, df in datasets.items():
        if not isinstance(df, pd.DataFrame) or df.empty:
            continue
        written.append(write_dataframe(df, out_dir / f"{name}.{fmt}", fmt))
    return written
