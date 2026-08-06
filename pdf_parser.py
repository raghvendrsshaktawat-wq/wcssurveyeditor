"""
parser.py — PDF Parsing Module (Module 1, tuned for Fenesta WCS Report layout)
------------------------------------------------------------------------------
Extracts header metadata + line items from Fenesta "WCS Report" PDFs using
PyMuPDF (fitz).

Public API:
    parse_survey_pdf(file_bytes: bytes) -> tuple[dict, list[dict]]

All regex/anchor constants are declared at the top of the file so they can be
tuned against new PDF layouts without touching parsing logic.

Tuned against a representative 29-page order (35 line items, 0001-0035).
"""

from __future__ import annotations

import hashlib
import re
from functools import lru_cache
from typing import Any

import fitz  # PyMuPDF

__version__ = "1.1.0"


# =============================================================================
# CONFIGURABLE REGEX PATTERNS & ANCHORS
# =============================================================================
# Tune these constants first when a new PDF layout shows up.
# Every pattern is compiled once at import time.

# ---- Header / metadata patterns ---------------------------------------------

# Order number: "P" followed by exactly 7 digits (e.g. P1234567).
# (Original spec said "W" + 7 digits — real Fenesta orders use "P".)
ORDER_NUM_RE = re.compile(r"\bP\d{7}\b")

# MSC No. (secondary reference): plain 10-digit integer.
# Prefer the labelled pattern on the Installation Completion Report page:
#     "MSC No. /Job Order: 9400000000 / P1234567"
# Fallback: standalone 10-digit line, but skip phone-number contexts.
MSC_LABELED_RE = re.compile(
    r"MSC\s*No\.?\s*/?\s*(?:Job\s*Order:?)?\s*(\d{10})",
    re.IGNORECASE,
)
MSC_STANDALONE_RE = re.compile(r"^\s*(\d{10})\s*$")
# Skip 10-digit numbers whose PREVIOUS line looks like a phone context.
PHONE_CONTEXT_RE = re.compile(r"\b(?:tel|phone|mobile|contact)\b", re.IGNORECASE)

# Date lines — American slash format on this layout (M/D/YYYY).
DATE_LINE_RE = re.compile(r"^\s*\d{1,2}/\d{1,2}/\d{2,4}\s*$")

# Header label block that repeats on every page and lets us lock onto values.
# Labels appear in this exact order on lines 0–6 of pages 2..N (and 27..33 on p.1).
HEADER_LABELS: tuple[str, ...] = (
    "Order No.",
    "Zone",
    "Quote No.",
    "Date",
    "MSC No.",
    "Print Date",
    "Segment",
)
# Empirically, the 6 values that follow are in this visual (grid) order:
#     [Order No., Date, Zone, Quote No., MSC No., Print Date]
# This mapping tells us which value slot belongs to which field.
HEADER_VALUE_ORDER: tuple[str, ...] = (
    "order_number",     # values[0]  e.g. P1234567
    "date",             # values[1]  e.g. 1/31/2026
    "zone",             # values[2]  e.g. <ZONE NAME>
    "quote_number",     # values[3]  e.g. P1234
    "reference_number", # values[4]  e.g. 9400000000  (== MSC No.)
    "print_date",       # values[5]  e.g. 1/30/2026
)

# Zone (fallback / validation) — an all-caps word/phrase, letters + spaces.
ZONE_LINE_RE = re.compile(r"^\s*([A-Z][A-Z\s]{1,30}[A-Z])\s*$")

# Zones we KNOW are false positives from address blocks (Indian states, country).
ZONE_EXCLUDE = {
    "INDIA", "TELANGANA", "KARNATAKA", "MAHARASHTRA", "TAMIL NADU",
    "KERALA", "GUJARAT", "RAJASTHAN", "PUNJAB", "HARYANA", "UP",
    "WEST BENGAL", "ODISHA", "ANDHRA PRADESH", "MADHYA PRADESH",
    "YES", "NO", "TO", "DFIXED", "DFIX", "SG TGH",
    # Header labels themselves (in case they leak through):
    "ORDER", "ZONE", "QUOTE", "DATE", "SEGMENT", "PROJECT",
}

# Customer name — cleanest source is the inline "Customer Name:" prefix
# (appears on the Installation Completion Report page).
CUSTOMER_INLINE_RE = re.compile(r"^\s*Customer\s*(?:Name)?\s*:\s*(.+?)\s*$")
# Fallback anchor for older layouts.
CUSTOMER_ANCHOR = "Customer"
CUSTOMER_FALSE_POSITIVE = "Customer Name"


# ---- Line-item patterns ------------------------------------------------------

# Sales-line code: 4-digit block starting with "0" (0001..0999 in practice).
# NOTE: config codes like "0204" / "0104" also match this shape but are
# filtered out downstream because the line immediately after them is NOT the
# literal qty "1".
SALES_LINE_RE = re.compile(r"^\s*(0\d{3})\s*$")

# Fixed literal that appears on the line *immediately after* the sales-line.
QTY_LITERAL = "1"

# Order dimension: 3–4 digit number on its own line (mm).
# Layout note: this PDF prints WIDTH first, then HEIGHT (header says "Size (w x h)").
DIMENSION_RE = re.compile(r"^\s*(\d{3,4})\s*$")

# 2-space-indented value lines — these hold Reference / Location / System
# in that order, appearing after the glazing line.
INDENTED_VALUE_RE = re.compile(r"^\s{2,}(\S.*?)\s*$")

# Anchor label that precedes the Reference/Location/System block for a given
# line item. Without locking onto this first, the lookahead can drift onto
# the WRONG item's subfields (or find nothing) for items whose local layout
# doesn't match the "typical" line count — e.g. sub-panels/mullions inside a
# combination window, which sit closer to their neighbour's anchor than to
# their own height value.
SUBFIELDS_ANCHOR_LABEL = "Arch Height(mm)"
SUBFIELDS_ANCHOR_SEARCH_WINDOW = 120  # lines to look forward for the anchor

# Description LOOKBACK skip list — lines to *skip* when searching backward
# for the description string (e.g., "(XiX)/(DFix.DFix)").
DESCRIPTION_SKIP_RE = re.compile(
    r"^(?:"
    r"\d+"                    # any pure-numeric line (config code, dimension)
    r"|DFixed?"              # DFixed / DFix labels
    r"|Viewed\s+from\s+Inside"
    r"|-->|<--|--&gt;|&lt;--"  # arrow markers rendered as text/entities
    r"|Sales\s+Line|Description|Qty|Size.*|Glazing"
    r"|Aperture\s+Size|Production\s+Size"
    r")$",
    re.IGNORECASE,
)

# How far to look back for the description string
DESCRIPTION_LOOKBACK = 8      # lines
# How many indented sub-fields (Reference / Location / System) to grab
SUBFIELDS_COUNT = 3
# How far forward to scan for indented sub-fields
SUBFIELDS_LOOKAHEAD = 40      # lines (plenty of headroom)


# =============================================================================
# CACHING
# =============================================================================

def _digest(file_bytes: bytes) -> str:
    return hashlib.sha1(file_bytes).hexdigest()


@lru_cache(maxsize=32)
def _parse_cached(digest: str, file_bytes: bytes) -> tuple[dict, list[dict]]:
    """Actual parsing entry — cached on the sha1 digest of the bytes."""
    return _parse_impl(file_bytes)


def parse_survey_pdf(file_bytes: bytes) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    """
    Parse a Fenesta WCS Report PDF and return metadata + line items.

    Args:
        file_bytes: Raw bytes of an uploaded PDF.

    Returns:
        (metadata_dict, list_of_row_dicts)

        metadata_dict keys:
            order_number, reference_number, zone, customer_name,
            date, print_date, quote_number   (extras included when available)

        Each row dict keys:
            sales_line, description, system, order_width, order_height,
            reference, location, subfields_missing (bool — True if Ref /
            Location / System could not be confidently located and needs a
            manual check),
            survey_width, survey_height, room, remarks   (empty placeholders)
    """
    if not file_bytes:
        return {}, []
    return _parse_cached(_digest(file_bytes), file_bytes)


# =============================================================================
# IMPLEMENTATION
# =============================================================================

def _extract_lines(file_bytes: bytes) -> list[str]:
    """
    Open PDF from bytes and return a flat list of text lines (page order).

    Defensive: catches encrypted / corrupt PDFs and re-raises as a friendly
    ValueError so main.py can surface it to the surveyor without a stack trace.
    """
    if not file_bytes or len(file_bytes) < 100:
        raise ValueError("PDF appears to be empty or truncated.")

    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
    except Exception as e:
        raise ValueError(f"Could not open PDF (is it corrupt?): {e}") from e

    try:
        if getattr(doc, "needs_pass", False) or getattr(doc, "is_encrypted", False):
            raise ValueError(
                "PDF is password-protected. Please remove the password before uploading."
            )

        lines: list[str] = []
        for page in doc:
            try:
                page_text = page.get_text("text") or ""
            except Exception:
                # Skip unreadable pages rather than aborting the whole file
                page_text = ""
            lines.extend(page_text.splitlines())
        return lines
    finally:
        doc.close()


# -------- Header metadata ----------------------------------------------------

def _extract_metadata(lines: list[str]) -> dict[str, Any]:
    """
    Extract header fields using two strategies:
    1) Locate the repeating header-label block and read the 6 values that
       follow, mapped by HEADER_VALUE_ORDER. This is the most reliable path
       because it's grounded in the PDF's visual grid.
    2) Regex fallbacks for order_number / reference_number / customer_name
       to cover edge cases (e.g. very first page uses a different framing).
    """
    md: dict[str, Any] = {
        "order_number": None,
        "reference_number": None,
        "zone": None,
        "customer_name": None,
        "date": None,
        "print_date": None,
        "quote_number": None,
    }

    # ---- Strategy 1: locked-on header-label block --------------------------
    header_values = _find_header_values_block(lines)
    if header_values:
        for slot, value in zip(HEADER_VALUE_ORDER, header_values):
            md[slot] = value

    # ---- Strategy 2: regex fallbacks ---------------------------------------
    # Order number — first "P" + 7 digits anywhere in the doc.
    if not md["order_number"]:
        for line in lines:
            m = ORDER_NUM_RE.search(line)
            if m:
                md["order_number"] = m.group(0)
                break

    # Reference (MSC No.) — labeled pattern first, then standalone 10-digit
    # (skipping phone contexts).
    if not md["reference_number"]:
        md["reference_number"] = _find_msc_number(lines)

    # Zone — validate what we already have; if wrong or missing, hunt manually.
    md["zone"] = _validate_or_find_zone(lines, md.get("zone"))

    # Customer name — prefer inline "Customer Name: <value>" pattern.
    md["customer_name"] = _find_customer_name(lines)

    return md


def _find_header_values_block(lines: list[str]) -> list[str] | None:
    """
    Locate the 7-line label block (Order No. .. Segment) and return the
    next 6 non-empty lines as values. Returns None if the block isn't found.
    """
    n = len(lines)
    labels = HEADER_LABELS
    L = len(labels)

    for i in range(n - L):
        # All 7 label lines must match exactly (stripped).
        if all(lines[i + k].strip() == labels[k] for k in range(L)):
            values: list[str] = []
            j = i + L
            # Collect up to 6 non-empty lines that immediately follow.
            while j < n and len(values) < 6:
                candidate = lines[j].strip()
                if candidate:
                    values.append(candidate)
                j += 1
            if len(values) >= 5:  # tolerate a missing Print Date
                # Pad to length 6 so the zip with HEADER_VALUE_ORDER is safe
                while len(values) < 6:
                    values.append(None)  # type: ignore[arg-type]
                return values
    return None


def _find_msc_number(lines: list[str]) -> str | None:
    """Find the MSC No. — labeled pattern first, then plain 10-digit line."""
    # 1) Labeled pattern (very reliable — appears on Installation report page).
    for line in lines:
        m = MSC_LABELED_RE.search(line)
        if m:
            return m.group(1)

    # 2) Standalone 10-digit line whose *previous* line isn't a phone context.
    for idx, line in enumerate(lines):
        m = MSC_STANDALONE_RE.match(line)
        if not m:
            continue
        prev = lines[idx - 1].strip() if idx > 0 else ""
        if PHONE_CONTEXT_RE.search(prev):
            continue  # skip customer phone numbers
        return m.group(1)

    return None


def _validate_or_find_zone(lines: list[str], candidate: str | None) -> str | None:
    """
    If the header-block gave us a zone that isn't a false positive, keep it.
    Otherwise, scan for an all-caps line that isn't in ZONE_EXCLUDE.
    """
    def _ok(z: str | None) -> bool:
        if not z:
            return False
        z_upper = z.strip().upper()
        return (
            bool(ZONE_LINE_RE.match(z_upper))
            and z_upper not in ZONE_EXCLUDE
            and not any(ch.isdigit() for ch in z_upper)
        )

    if _ok(candidate):
        return candidate.strip().upper()  # type: ignore[union-attr]

    for line in lines:
        s = line.strip()
        if _ok(s):
            return s.upper()
    return None


def _find_customer_name(lines: list[str]) -> str | None:
    """
    Prefer the inline "Customer Name: <value>" pattern (Installation Report
    page). Fallback: the line right after a "Customer" anchor, excluding the
    known false-positive header label "Customer Name".
    """
    # 1) Inline "Customer Name: ..." pattern (most reliable).
    for line in lines:
        m = CUSTOMER_INLINE_RE.match(line)
        if m:
            value = m.group(1).strip()
            if value and value.upper() != CUSTOMER_FALSE_POSITIVE.upper():
                return value

    # 2) Legacy anchor: line after "Customer" (skip the label "Customer Name").
    for idx, line in enumerate(lines):
        if CUSTOMER_ANCHOR in line and idx + 1 < len(lines):
            candidate = lines[idx + 1].strip()
            if candidate and candidate != CUSTOMER_FALSE_POSITIVE:
                return candidate

    return None


# -------- Line items ---------------------------------------------------------

def _extract_rows(lines: list[str]) -> list[dict[str, Any]]:
    """
    Walk the flattened line list and extract each order item.

    Fenesta WCS Report line-item shape (per sales-line):
        <description>             e.g. "(XiX)/(DFix.DFix)"   ← lookback target
        <config code>             e.g. "0204" / "0104"        (skipped)
        <sales_line>              e.g. "0001"
        1                         literal qty
        <order_width>             3-4 digits
        <order_height>            3-4 digits
        <glazing/description>     e.g. "SG TGH 5MM ..."      ← used as description
        ... field labels ...
          <reference>             e.g. "  W1"                 ← indented value
          <location>              e.g. "  Bedroom abv 29F"    ← indented value
          <system>                e.g. "  SY46 Premium ..."   ← indented value
    """
    rows: list[dict[str, Any]] = []
    n = len(lines)
    i = 0

    while i < n:
        sales_match = SALES_LINE_RE.match(lines[i].strip())
        if not sales_match:
            i += 1
            continue

        # Next non-empty line MUST be the literal qty "1" — this filters out
        # the config codes ("0204" / "0104") that also match SALES_LINE_RE.
        qty_idx = _next_non_empty_index(lines, i + 1)
        if qty_idx is None or lines[qty_idx].strip() != QTY_LITERAL:
            i += 1
            continue

        # Then: width, height (PDF header says "Size (w x h)")
        w_idx = _next_non_empty_index(lines, qty_idx + 1)
        h_idx = _next_non_empty_index(lines, (w_idx + 1) if w_idx is not None else n)

        order_height = _match_dim(lines[w_idx]) if w_idx is not None else None
        order_width = _match_dim(lines[h_idx]) if h_idx is not None else None

        if order_width is None or order_height is None:
            i += 1
            continue

        # Description via LOOKBACK from the sales-line (skip numeric/boilerplate).
        description = _lookback_description(lines, i)

        # Reference / Location / System — locked onto the "Arch Height(mm)"
        # anchor first (mirrors the original v4.2 parser), THEN collected as
        # the indented lines that follow it. Starting the scan right after
        # the order-height value (the old approach) drifts onto the wrong
        # item — or finds nothing — for items whose local block doesn't
        # follow the "typical" line count (e.g. mullion/sub-panel items
        # inside a combination window).
        search_start = h_idx + 1 if h_idx is not None else i + 1
        anchor_idx = _find_subfields_anchor(lines, search_start)

        subfields_missing = False
        if anchor_idx is not None:
            subfields = _collect_indented_values(
                lines,
                start=anchor_idx + 1,
                count=SUBFIELDS_COUNT,
                max_scan=SUBFIELDS_LOOKAHEAD,
            )
        else:
            # Anchor not found within the search window — fall back to the
            # old unanchored scan so we still attempt *something*, but flag
            # the row so the UI can surface it instead of silently showing
            # blank cells.
            subfields = _collect_indented_values(
                lines,
                start=search_start,
                count=SUBFIELDS_COUNT,
                max_scan=SUBFIELDS_LOOKAHEAD,
            )
            subfields_missing = True

        reference = subfields[0] if len(subfields) > 0 else ""
        location  = subfields[1] if len(subfields) > 1 else ""
        system    = subfields[2] if len(subfields) > 2 else ""

        # Even with the anchor found, treat a fully-empty result as a flagged
        # gap rather than a silent blank — surveyors need to know this item's
        # Ref/Location/System couldn't be confirmed from the PDF text.
        if not (reference or location or system):
            subfields_missing = True

        rows.append({
            "sales_line":    sales_match.group(1),
            "description":   description,   # e.g. "(XiX)/(DFix.DFix)"
            "system":        system,        # e.g. "SY46 Premium Slim Slider Combination (uPVC)"
            "order_width":   order_width,
            "order_height":  order_height,
            "reference":     reference,     # e.g. "W1"
            "location":      location,      # e.g. "Bedroom abv 29F"
            "subfields_missing": subfields_missing,  # True => needs manual check
            # Empty placeholders — filled in by the survey UI later
            "survey_width":  None,
            "survey_height": None,
            "room":          "",
            "remarks":       "",
        })

        # Advance past what we consumed to avoid re-matching this item.
        i = (h_idx + 1) if h_idx is not None else (i + 1)

    return rows


def _match_dim(raw: str) -> int | None:
    m = DIMENSION_RE.match(raw)
    return int(m.group(1)) if m else None


def _next_non_empty_index(lines: list[str], start: int) -> int | None:
    for j in range(start, len(lines)):
        if lines[j].strip():
            return j
    return None


def _lookback_description(lines: list[str], sales_idx: int) -> str:
    """
    Scan backwards from the sales-line for up to DESCRIPTION_LOOKBACK non-empty
    lines, skipping numeric lines and known boilerplate. Return the first
    surviving candidate (typically the "(XiX)/(DFix.DFix)" config string).
    """
    seen = 0
    for j in range(sales_idx - 1, -1, -1):
        candidate = lines[j].strip()
        if not candidate:
            continue
        seen += 1
        if seen > DESCRIPTION_LOOKBACK:
            break
        if DESCRIPTION_SKIP_RE.match(candidate):
            continue
        # Skip anything that itself looks like a sales-line code
        if SALES_LINE_RE.match(candidate):
            continue
        return candidate
    return ""


def _find_subfields_anchor(lines: list[str], start: int) -> int | None:
    """
    Find the "Arch Height(mm)" label that anchors this item's Reference /
    Location / System block, searching forward from `start` within
    SUBFIELDS_ANCHOR_SEARCH_WINDOW lines.

    Returns the index of the anchor line, or None if not found in range.
    """
    end = min(start + SUBFIELDS_ANCHOR_SEARCH_WINDOW, len(lines))
    for j in range(start, end):
        if SUBFIELDS_ANCHOR_LABEL in lines[j]:
            return j
    return None


def _collect_indented_values(
    lines: list[str],
    start: int,
    count: int,
    max_scan: int,
) -> list[str]:
    """
    Collect up to `count` 2-space-indented value lines starting at `start`,
    scanning at most `max_scan` lines forward. Returns the values with the
    leading whitespace stripped.
    """
    values: list[str] = []
    end = min(start + max_scan, len(lines))
    for j in range(start, end):
        if len(values) >= count:
            break
        m = INDENTED_VALUE_RE.match(lines[j])
        if m:
            values.append(m.group(1).strip())
    return values


# -------- Top-level implementation entry -------------------------------------

def _parse_impl(file_bytes: bytes) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    lines = _extract_lines(file_bytes)
    metadata = _extract_metadata(lines)
    rows = _extract_rows(lines)
    return metadata, rows
