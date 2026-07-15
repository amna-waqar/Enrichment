"""Configuration loading for the Chargebee analytics toolkit.

Reads credentials from environment variables (optionally via a `.env` file)
and validates them before any network call is attempted.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

try:
    from dotenv import load_dotenv
except ImportError:  # dotenv is optional; env vars can be set directly
    load_dotenv = None


# Chargebee's default (US) cluster uses the bare "chargebee.com" domain.
# Other regions insert a region segment: {site}.{region}.chargebee.com
_KNOWN_REGIONS = {"us", "eu", "au", "in"}


@dataclass(frozen=True)
class ChargebeeConfig:
    """Validated connection settings for a Chargebee site."""

    api_key: str
    site: str
    region: str | None = None
    domain: str | None = None

    @property
    def is_live(self) -> bool:
        return self.api_key.startswith("live_")

    @property
    def resolved_domain(self) -> str:
        """The full host passed to the SDK, e.g. `acme.eu.chargebee.com`."""
        if self.domain:
            return self.domain
        if self.region and self.region.lower() not in ("", "us"):
            return f"{self.site}.{self.region.lower()}.chargebee.com"
        return f"{self.site}.chargebee.com"

    def masked_key(self) -> str:
        if len(self.api_key) <= 12:
            return "****"
        return f"{self.api_key[:9]}...{self.api_key[-4:]}"


def load_config(dotenv_path: str | os.PathLike | None = None) -> ChargebeeConfig:
    """Load and validate configuration from the environment.

    Searches for a `.env` file (the given path, then the repo root, then CWD)
    and loads it if `python-dotenv` is installed. Raises a clear error if the
    required variables are missing.
    """
    if load_dotenv is not None:
        candidates = []
        if dotenv_path:
            candidates.append(Path(dotenv_path))
        # repo root = parent of this package directory
        candidates.append(Path(__file__).resolve().parent.parent / ".env")
        candidates.append(Path.cwd() / ".env")
        for path in candidates:
            if path and path.is_file():
                load_dotenv(path, override=False)
                break

    api_key = os.environ.get("CHARGEBEE_API_KEY", "").strip()
    site = os.environ.get("CHARGEBEE_SITE", "").strip()
    region = os.environ.get("CHARGEBEE_REGION", "").strip() or None
    domain = os.environ.get("CHARGEBEE_DOMAIN", "").strip() or None

    missing = [
        name
        for name, value in (("CHARGEBEE_API_KEY", api_key), ("CHARGEBEE_SITE", site))
        if not value
    ]
    if missing:
        raise RuntimeError(
            "Missing required environment variable(s): "
            + ", ".join(missing)
            + ".\nSet them in a `.env` file (see .env.example) or your shell. "
            "The site name is the subdomain before '.chargebee.com' and cannot "
            "be derived from the API key."
        )

    if region and region.lower() not in _KNOWN_REGIONS:
        # Not fatal — Chargebee may add regions — but warn via exception-free note.
        pass

    return ChargebeeConfig(api_key=api_key, site=site, region=region, domain=domain)
