"""
utils.py — Shared Helpers (Module 2: Tolerance Engine)
------------------------------------------------------
Pure, unit-testable helpers used across the WCS Survey Editor app.

Tolerance thresholds (mm), matched to the shop-floor SOP:
    |Δ| ≤ 75  mm  → 'ok'
    |Δ| ≤ 200 mm → 'warn'
    otherwise    → 'danger'
    survey value missing (None / NaN)  → 'empty'
"""

from __future__ import annotations

import math
from numbers import Real
from typing import Literal, Optional, Union

__version__ = "2.0.0"

# ---------------------------------------------------------------------------
# Configurable thresholds (millimetres)
# ---------------------------------------------------------------------------
OK_MAX_MM: float = 75.0
WARN_MAX_MM: float = 200.0

# Public status vocabulary
Status = Literal["ok", "warn", "danger", "empty"]

# What counts as a "measured" numeric value (accept int/float; reject None/NaN)
Number = Optional[Union[int, float]]


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------
def _is_missing(value: Number) -> bool:
    """True if the value is None or NaN — i.e. no survey reading yet."""
    if value is None:
        return True
    if isinstance(value, Real) and math.isnan(float(value)):
        return True
    return False


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def get_tolerance(order_value: Number, survey_value: Number) -> Status:
    """
    Classify a single dimension (width **or** height) against tolerance bands.

    Args:
        order_value:  Ordered dimension in millimetres. If missing, treated as
                      unmeasured and the result is 'empty'.
        survey_value: Site-measured dimension in millimetres. If missing
                      (None / NaN), the result is 'empty'.

    Returns:
        Status literal:
            'ok'     — |order − survey| ≤ 75 mm         (inclusive)
            'warn'   — 75 mm  < |order − survey| ≤ 200 mm (upper inclusive)
            'danger' — |order − survey| > 200 mm
            'empty'  — either value is missing (None / NaN)

    Boundary rule:
        Exactly 75 mm → 'ok'.
        Exactly 200 mm → 'warn'.
    """
    if _is_missing(order_value) or _is_missing(survey_value):
        return "empty"

    delta = abs(float(order_value) - float(survey_value))  # type: ignore[arg-type]

    if delta < OK_MAX_MM:
        return "ok"
    if delta < WARN_MAX_MM:
        return "warn"
    return "danger"


def row_tolerance(
    order_w: Number,
    order_h: Number,
    survey_w: Number,
    survey_h: Number,
) -> Status:
    """
    Combine width- and height-tolerance results for a single row.

    Rules:
        - If either dimension is 'empty', the whole row is 'empty'.
        - Otherwise, return the worst-case status:
              'danger' > 'warn' > 'ok'.

    Args:
        order_w:  Ordered width  (mm)
        order_h:  Ordered height (mm)
        survey_w: Measured width  (mm) — None/NaN if not yet surveyed
        survey_h: Measured height (mm) — None/NaN if not yet surveyed

    Returns:
        Status literal — one of 'ok', 'warn', 'danger', 'empty'.
    """
    w_status = get_tolerance(order_w, survey_w)
    h_status = get_tolerance(order_h, survey_h)

    if w_status == "empty" or h_status == "empty":
        return "empty"

    severity = {"ok": 0, "warn": 1, "danger": 2}
    worst = max(w_status, h_status, key=lambda s: severity[s])
    return worst  # type: ignore[return-value]


# ---------------------------------------------------------------------------
# Back-compat helper (kept for Module 0 scaffold — thin wrapper on new API)
# ---------------------------------------------------------------------------
def classify_tolerance(
    delta_mm: Number,
    green_mm: float = OK_MAX_MM,
    amber_mm: float = WARN_MAX_MM,
) -> str:
    """
    Legacy helper from the Module-0 scaffold. Prefer `get_tolerance`.

    Maps a *pre-computed* delta into the older green/amber/red/grey vocabulary.
    """
    if _is_missing(delta_mm):
        return "grey"
    d = abs(float(delta_mm))  # type: ignore[arg-type]
    if d < green_mm:
        return "green"
    if d < amber_mm:
        return "amber"
    return "red"
