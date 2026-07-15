"""Thin, well-behaved wrapper around the official Chargebee v3 Python SDK.

Responsibilities:
  * Build a configured `Chargebee` client from `ChargebeeConfig`.
  * Enable automatic retry/backoff on 429 + 5xx.
  * Provide a single generic `list_all()` that transparently pages through the
    entire result set for ANY list resource and yields raw JSON dicts.
  * Detect the site's Product Catalog version (1.0 vs 2.0).
"""

from __future__ import annotations

import json
import logging
from typing import Any, Callable, Iterable, Iterator

from chargebee import Chargebee
from chargebee.retry_config import RetryConfig

from .config import ChargebeeConfig, load_config

logger = logging.getLogger("chargebee_analytics")

# Chargebee caps `limit` at 100 per page.
MAX_PAGE_SIZE = 100


class ChargebeeClient:
    """A connected Chargebee client with paging + retry built in."""

    def __init__(self, config: ChargebeeConfig | None = None):
        self.config = config or load_config()
        self._cb = Chargebee(
            api_key=self.config.api_key,
            site=self.config.site,
            chargebee_domain=self.config.domain or None,
        )
        # If a region/domain was resolved, point the client at the right host.
        resolved = self.config.resolved_domain
        if resolved != f"{self.config.site}.chargebee.com":
            try:
                self._cb.env.set_api_endpoint(f"https://{resolved}")
            except Exception:  # pragma: no cover - defensive
                logger.warning("Could not override API endpoint to %s", resolved)

        # Automatic exponential backoff on transient failures.
        try:
            self._cb.update_retry_config(
                RetryConfig(
                    enabled=True,
                    max_retries=5,
                    delay_ms=1000,
                    retry_on=[429, 500, 502, 503, 504],
                )
            )
        except Exception:  # pragma: no cover - older SDKs
            logger.debug("Retry config not applied (SDK too old?)")

    @property
    def sdk(self) -> Chargebee:
        """Escape hatch: the raw SDK client for advanced/unsupported calls."""
        return self._cb

    # -- connectivity ------------------------------------------------------

    def verify_connection(self) -> dict[str, Any]:
        """Make one cheap authenticated call to confirm key + site work.

        Returns a small status dict. Raises on auth/site/network failure so the
        caller sees the real Chargebee error message.
        """
        resp = self._cb.Customer.list({"limit": 1})
        return {
            "ok": True,
            "site": self.config.site,
            "domain": self.config.resolved_domain,
            "live": self.config.is_live,
            "sample_count": len(resp.list),
        }

    # -- generic paging ----------------------------------------------------

    def list_all(
        self,
        resource_name: str,
        params: dict[str, Any] | None = None,
        *,
        item_key: str | None = None,
        page_size: int = MAX_PAGE_SIZE,
        max_records: int | None = None,
    ) -> list[dict[str, Any]]:
        """Fetch EVERY record of a list resource, following `next_offset`.

        Args:
            resource_name: PascalCase SDK resource, e.g. "Subscription".
            params: extra filter/sort params (see `filters` helpers).
            item_key: attribute holding the object on each list entry
                (defaults to the snake_case of `resource_name`).
            page_size: rows per request (max 100).
            max_records: stop after this many records (None = all).

        Returns a list of raw JSON dicts (each object's `raw_data`).
        """
        resource = getattr(self._cb, resource_name)
        key = item_key or _snake(resource_name)
        base = dict(params or {})
        base["limit"] = min(page_size, MAX_PAGE_SIZE)

        out: list[dict[str, Any]] = []
        offset: str | None = None
        page = 0
        while True:
            call_params = dict(base)
            if offset:
                call_params["offset"] = offset
            resp = resource.list(call_params)
            page += 1
            for entry in resp.list:
                obj = getattr(entry, key, None)
                if obj is None:
                    # Fall back to whatever single attribute the entry exposes.
                    obj = _first_model_attr(entry)
                out.append(_to_dict(obj))
                if max_records is not None and len(out) >= max_records:
                    logger.debug("%s: reached max_records=%s", resource_name, max_records)
                    return out
            offset = getattr(resp, "next_offset", None)
            logger.debug("%s: page %d, %d records so far", resource_name, page, len(out))
            if not offset:
                break
        return out

    def iter_all(
        self,
        resource_name: str,
        params: dict[str, Any] | None = None,
        *,
        item_key: str | None = None,
        page_size: int = MAX_PAGE_SIZE,
    ) -> Iterator[dict[str, Any]]:
        """Generator variant of `list_all` for streaming large result sets."""
        resource = getattr(self._cb, resource_name)
        key = item_key or _snake(resource_name)
        base = dict(params or {})
        base["limit"] = min(page_size, MAX_PAGE_SIZE)
        offset: str | None = None
        while True:
            call_params = dict(base)
            if offset:
                call_params["offset"] = offset
            resp = resource.list(call_params)
            for entry in resp.list:
                obj = getattr(entry, key, None) or _first_model_attr(entry)
                yield _to_dict(obj)
            offset = getattr(resp, "next_offset", None)
            if not offset:
                break

    # -- product catalog version ------------------------------------------

    def product_catalog_version(self) -> str:
        """Return "2.0" (items) or "1.0" (plans/addons) for this site."""
        # Preferred: read it from the Configurations endpoint.
        try:
            resp = self._cb.Configuration.list()
            for entry in getattr(resp, "configurations", []) or []:
                raw = _to_dict(entry)
                version = raw.get("product_catalog_version") or raw.get(
                    "product_catalog"
                )
                if version:
                    return "2.0" if "2" in str(version) else "1.0"
        except Exception:
            logger.debug("Configuration endpoint unavailable; probing items instead")

        # Fallback probe: PC 2.0 sites expose /items.
        try:
            self._cb.Item.list({"limit": 1})
            return "2.0"
        except Exception:
            return "1.0"


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _snake(pascal: str) -> str:
    out = []
    for i, ch in enumerate(pascal):
        if ch.isupper() and i > 0:
            out.append("_")
        out.append(ch.lower())
    return "".join(out)


def _to_dict(obj: Any) -> dict[str, Any]:
    """Best-effort conversion of an SDK model object to a plain JSON dict."""
    if obj is None:
        return {}
    if isinstance(obj, dict):
        return obj
    raw = getattr(obj, "raw_data", None)
    if isinstance(raw, dict) and raw:
        return raw
    # Fall back to dataclass fields, dropping empty/None and the raw_data slot.
    data = getattr(obj, "__dict__", None)
    if isinstance(data, dict):
        return {k: v for k, v in data.items() if k != "raw_data" and v is not None}
    return {}


def _first_model_attr(entry: Any) -> Any:
    """Return the first non-empty model attribute on a list entry wrapper."""
    ann = getattr(type(entry), "__annotations__", {})
    for name in ann:
        val = getattr(entry, name, None)
        if val is not None:
            return val
    return entry


# ---------------------------------------------------------------------------
# filter builders — turn friendly kwargs into Chargebee's `field[op]` params
# ---------------------------------------------------------------------------


def ts_after(field: str, epoch_seconds: int) -> dict[str, Any]:
    return {field: {"after": epoch_seconds}}


def ts_between(field: str, start: int, end: int) -> dict[str, Any]:
    # Chargebee expects a JSON array for `between`.
    return {field: {"between": json.dumps([start, end])}}


def enum_in(field: str, values: Iterable[str]) -> dict[str, Any]:
    # List filters must be JSON-encoded, not Python-repr'd.
    return {field: {"in": json.dumps(list(values))}}


def merge(*parts: dict[str, Any]) -> dict[str, Any]:
    """Merge several filter dicts, deep-merging same-field operator maps."""
    result: dict[str, Any] = {}
    for part in parts:
        for k, v in part.items():
            if isinstance(v, dict) and isinstance(result.get(k), dict):
                result[k].update(v)
            else:
                result[k] = v
    return result
