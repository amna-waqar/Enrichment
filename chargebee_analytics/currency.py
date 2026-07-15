"""Money helpers.

Chargebee returns all amounts as integers in the currency's *minor* unit
(e.g. cents). Most currencies have 2 decimal places, but zero-decimal
currencies (JPY, KRW, ...) and three-decimal currencies (BHD, ...) exist, so
we never hardcode a division by 100.
"""

from __future__ import annotations

from decimal import Decimal

# ISO 4217 currencies whose minor unit == major unit (no division).
ZERO_DECIMAL = {
    "BIF", "CLP", "DJF", "GNF", "JPY", "KMF", "KRW", "MGA", "PYG",
    "RWF", "UGX", "VND", "VUV", "XAF", "XOF", "XPF",
}

# Currencies with three decimal places.
THREE_DECIMAL = {"BHD", "IQD", "JOD", "KWD", "LYD", "OMR", "TND"}


def currency_exponent(currency_code: str | None) -> int:
    """Number of decimal places for a currency code (default 2)."""
    if not currency_code:
        return 2
    code = currency_code.upper()
    if code in ZERO_DECIMAL:
        return 0
    if code in THREE_DECIMAL:
        return 3
    return 2


def minor_to_major(amount: int | float | None, currency_code: str | None) -> float | None:
    """Convert an integer minor-unit amount to a major-unit float.

    Returns None when the input amount is None (Chargebee omits some amounts).
    """
    if amount is None:
        return None
    # Guard against pandas NaN (missing amounts become NaN after json_normalize).
    if isinstance(amount, float) and amount != amount:
        return None
    exponent = currency_exponent(currency_code)
    return float(Decimal(int(amount)) / (Decimal(10) ** exponent))
