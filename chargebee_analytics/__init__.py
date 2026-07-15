"""chargebee_analytics — pull and analyse Chargebee data via code.

Quick start:

    from chargebee_analytics import ChargebeeClient, extract, analytics

    client = ChargebeeClient()               # reads .env
    print(client.verify_connection())

    subs = extract.get_subscriptions(client)
    print(analytics.mrr_summary(subs))
"""

from __future__ import annotations

from . import analytics, extract, export
from .client import ChargebeeClient
from .config import ChargebeeConfig, load_config

__all__ = [
    "ChargebeeClient",
    "ChargebeeConfig",
    "load_config",
    "extract",
    "analytics",
    "export",
]

__version__ = "0.1.0"
