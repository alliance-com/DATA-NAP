# -*- coding: utf-8 -*-
"""Montant en lettres EN pour Commercial Invoice."""

_U = [
    "", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE",
    "TEN", "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN", "SIXTEEN",
    "SEVENTEEN", "EIGHTEEN", "NINETEEN",
]
_T = ["", "", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY", "EIGHTY", "NINETY"]


def _under_1000(n):
    n = int(n)
    if n < 20:
        return _U[n]
    if n < 100:
        d, u = divmod(n, 10)
        return _T[d] + ((" " + _U[u]) if u else "")
    c, r = divmod(n, 100)
    return _U[c] + " HUNDRED" + ((" " + _under_1000(r)) if r else "")


def _int_words(n):
    n = int(n)
    if n == 0:
        return "ZERO"
    parts = []
    billions, n = divmod(n, 1_000_000_000)
    millions, n = divmod(n, 1_000_000)
    thousands, n = divmod(n, 1_000)
    if billions:
        parts.append(_under_1000(billions) + " BILLION")
    if millions:
        parts.append(_under_1000(millions) + " MILLION")
    if thousands:
        parts.append(_under_1000(thousands) + " THOUSAND")
    if n:
        parts.append(_under_1000(n))
    return " ".join(parts)


def amount_in_words(amount, currency="EUR"):
    try:
        amount = float(amount)
    except Exception:
        return ""
    neg = amount < 0
    amount = abs(amount)
    whole = int(amount)
    cents = int(round((amount - whole) * 100))
    cur = (currency or "EUR").upper()
    if cur == "EUR":
        prefix, unit, subunit = "EURO", "EURO", "CENTS"
    elif cur == "USD":
        prefix, unit, subunit = "U.S DOLLARS", "U.S DOLLARS", "CENTS"
    else:
        prefix, unit, subunit = cur, cur, "CENTS"

    text = _int_words(whole)
    if cents:
        text = f"{text} AND {_int_words(cents)} {subunit}"
    else:
        text = f"{text}"
    # style NAP samples
    if cur == "EUR":
        out = f"EURO {text} ONLY"
    elif cur == "USD":
        out = f"U.S DOLLARS {text} ONLY"
    else:
        out = f"{prefix} {text} ONLY"
    if neg:
        out = "MINUS " + out
    return " ".join(out.split())
