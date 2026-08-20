"""
overlay.py — PDF Overlay / Annotation Module (Module 3, v3.2 — geometry + style parity with app.py)
------------------------------------------------------------------------------
Stamps site-survey values on top of the original Fenesta WCS Report PDF using
PyMuPDF (fitz).

Public API:
    overlay_survey_data(pdf_bytes, rows, surveyor_name="") -> bytes

⚠️  Why the X-coordinates are hardcoded:
    Fenesta's data-entry cells are drawn as long horizontal underlines,
    not as closed rectangles. get_drawings() returns each stroke as a
    separate path, so there's no reliable "cell shape" to pick up. The
    visual columns are fixed in the template — measured once, they stay
    stable across every page of every order. If the template changes,
    re-measure the CELL_X_* constants.

🩹  v3.1 fix (2026-07-23):
    The "Aperture Size" row in the Fenesta template is only ~6-8 pt tall
    between two closely-spaced horizontal rules. PyMuPDF's insert_textbox()
    silently fails when the rect is shorter than one line of text at the
    minimum font, so the previous version rendered the coloured box but
    NO TEXT INSIDE IT.

    Fix:
      1) Enforce MIN_MAIN_CELL_HEIGHT so we always have room to draw.
      2) Use insert_text() (point-based, always renders) for the main
         cell instead of insert_textbox() (rect-based, height-sensitive).
      3) Auto-fit font size by measuring text width first.

🩹  v3.2 fix (2026-07-26):
    Two regressions vs. the original monolithic app.py were traced back to
    THIS module and fixed:

    (a) Border placement was off. `_find_cell_bounds` capped the search for
        the surrounding horizontal ruling lines to a narrow window (6pt
        above / 30pt below the anchor) and, when nothing was found in that
        window, fell back to `anchor_top - 1.0` — almost no gap above the
        label. app.py searched the WHOLE page for ruling lines (no distance
        cap, just direction) and fell back to `hit.y0 - 13` / `hit.y1 + 3`
        when none were found. This version now matches app.py exactly:
        unbounded search, same fallback offsets. This is very likely the
        actual cause of "the border placement was better in app.py."

    (b) Color/remarks styling didn't match app.py's look:
          - app.py: WHITE fill, border AND text both colored per tolerance
            status (green/amber/red/blue) — high-contrast, severity is
            visible at a glance in both the box outline and the text color.
          - This module (pre-v3.2): pastel colored FILL, border colored,
            but text was always near-black — status was only visible in the
            box outline, not the text.
          - app.py's remarks row sits DIRECTLY below the main cell with the
            SAME height (row_h) and no "Remarks:" label prefix — just the
            remark text, colored the same as the status.
          - This module (pre-v3.2) used a fixed 12pt remarks row with a
            small gap, white fill, and a "Remarks: " prefix.
        v3.2 reverts to app.py's scheme for both, since that's the style
        that was preferred. The safer point-based/auto-fit text rendering
        from v3.1 is KEPT (it's a correctness fix, not a style choice) —
        it just now renders in app.py's colors and geometry.
"""

from __future__ import annotations

import io
from typing import Any, Iterable, Optional

import fitz  # PyMuPDF

from utils import row_tolerance

__version__ = "3.2.0"


# =============================================================================
# CONFIGURABLE CONSTANTS (tune these against your source template)
# =============================================================================

# ---- Colors (RGB, 0..1) — keyed to tolerance status -------------------------
# Matches app.py: white fill, border AND text both colored per status.
STATUS_COLORS: dict[str, tuple[float, float, float]] = {
    "ok":     (0.0,  0.6,  0.2),    # green
    "warn":   (0.91, 0.13, 0.18),   # red — amber retired (single-indication)
    "danger": (0.91, 0.13, 0.18),   # red   — Fenesta Red #E8212E
    "empty":  (0.0,  0.36, 0.67),   # blue  — Fenesta Blue #005BAC
}
WHITE: tuple[float, float, float] = (1.0, 1.0, 1.0)


# ---- Anchor labels used to LOCATE cells --------------------------------------
CELL_ANCHOR_LABEL = "Aperture Size"     # appears once per item on every page
SURVEYOR_NAME_ANCHOR = "Name"           # 2nd match on page 1 = Surveyor slot

# ---- Hardcoded X-coordinates for the survey-value cell ----------------------
# Measured from Fenesta WCS Report template (A4 portrait, ~595pt wide).
# Reverted to app.py's measured value (459.65) — the wider 555.0 used
# previously extended past the data-entry cell into unrelated content.
CELL_X_LEFT:  float = 78.25     # left edge of the data-entry cell
CELL_X_RIGHT: float = 459.65    # right edge of the data-entry cell
CELL_INSET:   float =   0.5     # matches app.py's +0.5/-0.5 border inset

# ---- Vertical geometry constants -------------------------------------------
# app.py searched the WHOLE page for ruling lines near the anchor (direction
# only, no distance cap) and used fixed fallback offsets when none were
# found. Replicated exactly here — see v3.2 fix note above.
ANCHOR_RULE_BUFFER_PT: float =  2.0   # matches app.py's "+2 / -2" buffer
FALLBACK_TOP_OFFSET_PT: float = 13.0  # top = anchor_top - 13   (app.py)
FALLBACK_BOT_OFFSET_PT: float =  3.0  # bot = anchor_bottom + 3 (app.py)
MIN_MAIN_CELL_HEIGHT:  float = 13.0   # 🩹 v3.1: enforce so text always renders

# ---- Text sizing ------------------------------------------------------------
# app.py: fs = min(9.0, row_h * 0.62) — same formula used here as the
# starting point; the v3.1 width-based auto-shrink still applies underneath
# as a safety net for unusually long remarks/room strings.
FONT_SIZE_ROW_HEIGHT_RATIO: float = 0.62
MAX_FONT_SIZE: float = 9.0
MIN_FONT_SIZE: float = 5.5
FONT_NAME:      str  = "helv"     # Helvetica — built-in, no embedding
FONT_NAME_BOLD: str  = "hebo"     # Helvetica-Bold

# ---- Surveyor name stamp geometry ------------------------------------------
SURVEYOR_BOX_WIDTH:  float = 140.0
SURVEYOR_BOX_HEIGHT: float =  14.0
SURVEYOR_X_OFFSET:   float =  35.0    # push right of the "Name" label
SURVEYOR_Y_OFFSET:   float =  -2.0    # tiny nudge up to sit on the underline
SURVEYOR_TEXT_COLOR: tuple[float, float, float] = (0.06, 0.09, 0.16)  # near-black


# =============================================================================
# PUBLIC API
# =============================================================================

def overlay_survey_data(
    pdf_bytes: bytes,
    rows: list[dict[str, Any]],
    surveyor_name: str = "",
    sketches: dict[str, bytes] | None = None,
) -> bytes:
    """
    Stamp measured survey values on top of the original order PDF.

    Args:
        pdf_bytes:     Original order PDF (as bytes).
        rows:          List of row dicts from parse_survey_pdf(). Each row is
                       stamped on the page that contains its sales-line.
        surveyor_name: Optional. Stamped on page 1 near the "Surveyor Name"
                       slot (the 2nd "Name" occurrence).

        sketches:      Optional {sales_line: png_bytes}. Each sketch is stamped
                       INLINE on its own item page, into the blank column to
                       the right of the "Production Size" row. Any sketch whose
                       anchor cannot be found falls back to an appended page,
                       so a sketch is never silently lost.

    Returns:
        bytes: The annotated PDF, ready for st.download_button().
    """
    # ---- Defensive guards --------------------------------------------------
    if not pdf_bytes or len(pdf_bytes) < 100:
        raise ValueError("overlay_survey_data: pdf_bytes is empty or truncated.")
    if rows is None:
        rows = []
    if not isinstance(rows, list):
        raise TypeError("overlay_survey_data: rows must be a list of dicts.")
    surveyor_name = (surveyor_name or "").strip()

    facets_by_sales_line: dict[str, list[dict[str, Any]]] = {}
    for r in rows:
        if not (isinstance(r, dict) and r.get("sales_line")):
            continue
        code = str(r.get("sales_line", "")).strip()
        facets_by_sales_line.setdefault(code, []).append(r)

    def _fno_sort(x: dict[str, Any]) -> float:
        try:
            return float(x.get("facet_no") or 0)
        except (TypeError, ValueError):
            return 0.0

    for code, group in facets_by_sales_line.items():
        group.sort(key=_fno_sort)

    by_sales_line: dict[str, dict[str, Any]] = {
        code: group[0] for code, group in facets_by_sales_line.items()
    }

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        raise ValueError(f"Could not open PDF for overlay: {e}") from e

    try:
        if getattr(doc, "needs_pass", False) or getattr(doc, "is_encrypted", False):
            raise ValueError(
                "PDF is password-protected. Remove the password before overlaying."
            )

        # ---- (1) Surveyor name stamp on page 1 -----------------------------
        if surveyor_name and len(doc) > 0:
            _stamp_surveyor_name(doc[0], surveyor_name)

        # ---- (2) Per-row overlays ------------------------------------------
        for page_index in range(len(doc)):
            page = doc[page_index]
            _annotate_page(page, by_sales_line, facets_by_sales_line)

        # ---- (3) Site sketches, stamped inline on each item page ------------
        if sketches:
            placed = _stamp_sketches_inline(doc, sketches, by_sales_line)
            leftover = {
                sl: png for sl, png in sketches.items() if sl not in placed
            }
            if leftover:
                _append_sketch_pages(doc, leftover, by_sales_line)

        # ---- (4) Serialize to bytes ----------------------------------------
        buf = io.BytesIO()
        doc.save(buf, garbage=4, deflate=True)
        return buf.getvalue()
    finally:
        try:
            doc.close()
        except Exception:
            pass


# =============================================================================
# INTERNALS
# =============================================================================

def _stamp_surveyor_name(page: fitz.Page, name: str) -> None:
    """Whites out the surveyor "Name" slot on page 1 and inserts the name."""
    hits = page.search_for(SURVEYOR_NAME_ANCHOR)
    if not hits:
        return
    anchor_rect = hits[1] if len(hits) >= 2 else hits[0]

    box = fitz.Rect(
        anchor_rect.x1 + SURVEYOR_X_OFFSET,
        anchor_rect.y0 + SURVEYOR_Y_OFFSET,
        anchor_rect.x1 + SURVEYOR_X_OFFSET + SURVEYOR_BOX_WIDTH,
        anchor_rect.y0 + SURVEYOR_Y_OFFSET + SURVEYOR_BOX_HEIGHT,
    )
    page.draw_rect(box, color=WHITE, fill=WHITE, overlay=True)
    _insert_single_line(
        page, box, text=name.strip(),
        fontname=FONT_NAME_BOLD, color=SURVEYOR_TEXT_COLOR, align="left",
    )


def _annotate_page(
    page: fitz.Page,
    by_sales_line: dict[str, dict[str, Any]],
    facets_by_sales_line: dict[str, list[dict[str, Any]]] | None = None,
) -> None:
    """Find each "Aperture Size" anchor and stamp the matching sales-line."""
    anchor_hits = page.search_for(CELL_ANCHOR_LABEL)
    if not anchor_hits:
        return
    sales_line_rects = _find_sales_line_rects_on_page(page, by_sales_line.keys())
    h_lines = _find_horizontal_rules(page)

    for anchor_rect in anchor_hits:
        row = _pick_row_for_anchor(anchor_rect, sales_line_rects, by_sales_line)
        if row is None:
            continue
        code = str(row.get("sales_line") or "").strip()
        facets = (facets_by_sales_line or {}).get(code) or [row]
        _draw_row_overlay(page, anchor_rect, facets, h_lines)


def _find_sales_line_rects_on_page(
    page: fitz.Page,
    known_sales_lines: Iterable[str],
) -> list[tuple[str, fitz.Rect]]:
    """
    Locate each sales-line code on the page as a WHOLE WORD.

    CRITICAL FIX (20-Aug-2026): the previous implementation used
    page.search_for(code), which does SUBSTRING matching. A 4-digit
    sales-line code such as "0011" is a substring of the 10-digit MSC number
    printed in every page header (e.g. 9001172578 contains "0011"), so a
    sketch for line 0011 matched the header on page 1 and was stamped on the
    wrong line. Matching whole word tokens instead — the sales line is its own
    space-delimited token, the MSC number is a different token — eliminates
    that false positive entirely. One pass over the page's words also replaces
    N search_for() calls, so it is faster too.
    """
    wanted = {str(c).strip() for c in known_sales_lines if str(c).strip()}
    if not wanted:
        return []
    found: list[tuple[str, fitz.Rect]] = []
    for w in page.get_text("words"):
        token = str(w[4]).strip()
        if token in wanted:
            found.append((token, fitz.Rect(w[0], w[1], w[2], w[3])))
    return found


def _pick_row_for_anchor(
    anchor_rect: fitz.Rect,
    sales_line_rects: list[tuple[str, fitz.Rect]],
    by_sales_line: dict[str, dict[str, Any]],
) -> Optional[dict[str, Any]]:
    """Pick the closest sales-line ABOVE this "Aperture Size" anchor."""
    candidates = [
        (code, rect) for code, rect in sales_line_rects
        if rect.y1 <= anchor_rect.y0
    ]
    if not candidates:
        return None
    code, _ = min(candidates, key=lambda cr: anchor_rect.y0 - cr[1].y1)
    return by_sales_line.get(code)


def _draw_row_overlay(
    page: fitz.Page,
    anchor_rect: fitz.Rect,
    facets: list[dict[str, Any]],
    h_lines: list[fitz.Rect],
) -> None:
    """Draw the coloured cell + text for one survey row against its anchor.

    Styling matches app.py: WHITE fill, border AND text both colored per
    tolerance status. The remarks row (if any) sits directly below the main
    cell with the SAME height, same styling, no label prefix — matching
    app.py's "Production Size row" behaviour exactly.
    """
    row = facets[0]
    is_multi = len(facets) > 1

    cell_top, cell_bot = _find_cell_bounds(anchor_rect, h_lines)

    if (cell_bot - cell_top) < MIN_MAIN_CELL_HEIGHT:
        cell_bot = cell_top + MIN_MAIN_CELL_HEIGHT

    row_h = cell_bot - cell_top

    if is_multi:
        status = _aggregate_status(facets)
    else:
        status = row_tolerance(
            row.get("order_width"), row.get("order_height"),
            row.get("survey_width"), row.get("survey_height"),
        )
    color = STATUS_COLORS.get(status, STATUS_COLORS["empty"])

    # ---- Main "Aperture Size" cell: white fill, colored border + text -----
    main_cell = fitz.Rect(
        CELL_X_LEFT + CELL_INSET,
        cell_top + CELL_INSET,
        CELL_X_RIGHT - CELL_INSET,
        cell_bot - CELL_INSET,
    )
    page.draw_rect(main_cell, color=color, fill=WHITE, width=1.5, overlay=True)

    text = _format_facet_text(facets) if is_multi else _format_survey_text(row)
    _insert_single_line(
        page, main_cell, text=text, row_h=row_h,
        fontname=FONT_NAME, color=color, align="left",
    )

    # ---- Remarks row: directly below, SAME height, no gap, no label ------
    # (mirrors app.py's "Production Size row" — plain remark text, same
    # color as the status, same left-aligned layout as the main cell)
    # Boxes stack downward from the main cell: remarks first (if any), then a
    # raw-measurements box (three-measurement openings only). next_top tracks
    # where the next box goes so they never overlap.
    next_top = cell_bot

    if is_multi:
        remarks = ""
        for _f in facets:
            _rr = str(_f.get("remarks", "") or "").strip()
            if _rr:
                remarks = _rr
                break
    else:
        remarks = str(row.get("remarks", "") or "").strip()
    if remarks:
        rem_cell = fitz.Rect(
            CELL_X_LEFT + CELL_INSET,
            next_top + CELL_INSET,
            CELL_X_RIGHT - CELL_INSET,
            next_top + row_h - CELL_INSET,
        )
        page.draw_rect(rem_cell, color=color, fill=WHITE, width=1.5, overlay=True)
        _insert_single_line(
            page, rem_cell, text=remarks, row_h=row_h,
            fontname=FONT_NAME, color=color, align="left",
        )
        next_top += row_h

    # ---- Raw measurements box (three-measurement mode only) --------------
    meas_text = _format_measurements(facets)
    if meas_text:
        meas_cell = fitz.Rect(
            CELL_X_LEFT + CELL_INSET,
            next_top + CELL_INSET,
            CELL_X_RIGHT - CELL_INSET,
            next_top + row_h - CELL_INSET,
        )
        page.draw_rect(meas_cell, color=color, fill=WHITE, width=1.5, overlay=True)
        _insert_single_line(
            page, meas_cell, text=meas_text, row_h=row_h,
            fontname=FONT_NAME, color=color, align="left",
        )
        next_top += row_h


def _find_horizontal_rules(page: fitz.Page) -> list[fitz.Rect]:
    """
    Collect every thin, wide horizontal ruling line on the page — the same
    filter app.py used (height < 3pt, width > 100pt) — once per page.
    """
    rules: list[fitz.Rect] = []
    for d in page.get_drawings():
        rect: fitz.Rect = d.get("rect")
        if rect is None:
            continue
        if abs(rect.y1 - rect.y0) < 3 and (rect.x1 - rect.x0) > 100:
            rules.append(rect)
    return rules


def _find_cell_bounds(
    anchor_rect: fitz.Rect,
    h_lines: list[fitz.Rect],
) -> tuple[float, float]:
    """
    Locate the horizontal cell-boundary lines around the anchor.

    🩹 v3.2: matches app.py exactly — search the WHOLE set of ruling lines on
    the page (no distance cap, direction only), and use the same fallback
    offsets (-13 above / +3 below) when no ruling line is found. The
    previous version capped the search window and used a near-zero fallback
    gap, which misplaced the border whenever the true ruling line fell
    outside that window.
    """
    above = [r for r in h_lines if r.y1 <= anchor_rect.y0 + ANCHOR_RULE_BUFFER_PT]
    below = [r for r in h_lines if r.y0 >= anchor_rect.y1 - ANCHOR_RULE_BUFFER_PT]

    # app.py picks the CLOSEST line above (max y0) and CLOSEST line below
    # (min y0); fall back to the fixed offsets when no ruling line is found.
    top = max((r.y0 for r in above), default=anchor_rect.y0 - FALLBACK_TOP_OFFSET_PT)
    bot = min((r.y0 for r in below), default=anchor_rect.y1 + FALLBACK_BOT_OFFSET_PT)

    return top, bot


# ---- Text helpers -----------------------------------------------------------

def _format_measurements(facets: list[dict[str, Any]]) -> str:
    """
    One-line raw-measurement string for three-measurement openings, e.g.
        "W: 1204/1200/1207   H: 1502/1500/1499"
    Returns "" when no measurement columns are populated (direct-entry mode),
    so the extra box is only drawn when there is something to show.
    For a multi-facet opening, the first facet that carries measurements wins.
    """
    def _num(v: Any) -> Optional[int]:
        try:
            f = float(v)
        except (TypeError, ValueError):
            return None
        if f != f:                     # NaN
            return None
        return int(round(f))

    for f in facets:
        ws = [_num(f.get(c)) for c in ("meas_w1", "meas_w2", "meas_w3")]
        hs = [_num(f.get(c)) for c in ("meas_h1", "meas_h2", "meas_h3")]
        if any(v is not None for v in ws + hs):
            w_txt = "/".join(str(v) if v is not None else "-" for v in ws)
            h_txt = "/".join(str(v) if v is not None else "-" for v in hs)
            return f"W: {w_txt}   H: {h_txt}"
    return ""


def _format_survey_text(row: dict[str, Any]) -> str:
    """
    Build the '{room} : {surveyed_W} x {surveyed_H}' string.

    Matches app.py exactly: if room is blank, show the size alone (no
    placeholder room text); if only one dimension is measured, show '--'
    for the missing one; if neither is measured, show 'Not surveyed'.
    """
    room = str(row.get("room", "") or "").strip()

    sw = row.get("survey_width")
    sh = row.get("survey_height")

    def _has_value(v: Any) -> bool:
        if v is None:
            return False
        try:
            import math as _math
            return not _math.isnan(float(v))
        except (TypeError, ValueError):
            return False

    def _fmt(v: Any) -> str:
        return str(int(round(float(v))))

    if _has_value(sw) and _has_value(sh):
        size_txt = f"{_fmt(sw)} x {_fmt(sh)}"
    elif _has_value(sw):
        size_txt = f"{_fmt(sw)} x --"
    elif _has_value(sh):
        size_txt = f"-- x {_fmt(sh)}"
    else:
        size_txt = "Not surveyed"

    return f"{room} : {size_txt}" if room else size_txt


def _aggregate_status(facets: list[dict[str, Any]]) -> str:
    order = {"danger": 3, "warn": 2, "ok": 1, "empty": 0}
    worst = "empty"
    for f in facets:
        s = row_tolerance(
            f.get("order_width"), f.get("order_height"),
            f.get("survey_width"), f.get("survey_height"),
        )
        if order.get(s, 0) > order.get(worst, 0):
            worst = s
    return worst


def _format_facet_text(facets: list[dict[str, Any]]) -> str:
    def _has_value(v: Any) -> bool:
        if v is None:
            return False
        try:
            import math as _math
            return not _math.isnan(float(v))
        except (TypeError, ValueError):
            return False

    def _fmt(v: Any) -> str:
        return str(int(round(float(v))))

    def _pair(sw: Any, sh: Any) -> str:
        if _has_value(sw) and _has_value(sh):
            return f"{_fmt(sw)} x {_fmt(sh)}"
        if _has_value(sw):
            return f"{_fmt(sw)} x --"
        if _has_value(sh):
            return f"-- x {_fmt(sh)}"
        return "-- x --"

    def _fno(v: Any, default: int) -> int:
        try:
            return int(float(v))
        except (TypeError, ValueError):
            return default

    first = facets[0]
    label = (str(first.get("room", "") or "").strip()
             or str(first.get("location", "") or "").strip())
    parts = []
    for idx, f in enumerate(facets, start=1):
        no = _fno(f.get("facet_no"), idx)
        parts.append(f"Facet {no}: {_pair(f.get('survey_width'), f.get('survey_height'))}")
    body = ", ".join(parts)
    return f"{label} - {body}" if label else body


def _insert_single_line(
    page: fitz.Page,
    rect: fitz.Rect,
    text: str,
    row_h: Optional[float] = None,
    fontname: str = FONT_NAME,
    color: tuple[float, float, float] = STATUS_COLORS["empty"],
    align: str = "left",  # "left" | "center" | "right"
) -> None:
    """
    Guaranteed-render single-line text (🩹 v3.1 fix kept), sized using
    app.py's formula: fs = min(MAX_FONT_SIZE, row_h * 0.62).

    Uses fitz.get_text_length() to measure the string and shrinks further if
    needed so it never overflows the cell width, then places it with
    insert_text() at an explicit (x, y) baseline. Unlike insert_textbox(),
    insert_text() has no rect-height check, so text ALWAYS renders — even
    on the ~6-8pt tall cells this template uses.
    """
    if not text:
        return

    ref_height = row_h if row_h is not None else rect.height
    fontsize = max(
        MIN_FONT_SIZE,
        min(MAX_FONT_SIZE, ref_height * FONT_SIZE_ROW_HEIGHT_RATIO),
    )

    # Measure text width; shrink font until it fits the rect width (minus pad).
    usable_width = max(rect.width - 4.0, 8.0)
    text_width = fitz.get_text_length(text, fontname=fontname, fontsize=fontsize)
    if text_width > usable_width:
        shrunk = fontsize * (usable_width / text_width)
        fontsize = max(4.0, shrunk)   # hard floor to always render *something*
        text_width = fitz.get_text_length(text, fontname=fontname, fontsize=fontsize)

    # Horizontal alignment
    if align == "left":
        x = rect.x0 + 4.0   # matches app.py's CELL_PAD
    elif align == "right":
        x = rect.x1 - 4.0 - text_width
    else:  # center
        x = rect.x0 + (rect.width - text_width) / 2.0

    # Vertical baseline — app.py: top_y + row_h * 0.73
    y = rect.y0 + ref_height * 0.73

    page.insert_text(
        fitz.Point(x, y),
        text,
        fontname=fontname,
        fontsize=fontsize,
        color=color,
        overlay=True,
    )


# =============================================================================
# MODULE 7 — SITE SKETCHES
# =============================================================================
# Sketches are stamped INLINE on the item page, into the blank column to the
# right of the spec list, level with / below the "Production Size" row. That
# column is empty on every item block in the Fenesta WCS template (it sits
# under the elevation drawing), so nothing printed is covered.
#
# ⚠️  CALIBRATION WARNING (05-Aug-2026)
#     The X constants below were derived PROPORTIONALLY FROM A SCREENSHOT,
#     not measured from a real PDF. They are almost certainly close but may
#     be a few points out. Run `python calibrate_sketch_box.py <your.pdf>`
#     (ships alongside this module) — it renders the proposed box onto page
#     images so you can see exactly where the sketch will land, then nudge
#     SKETCH_BOX_X_LEFT / SKETCH_BOX_RIGHT_MARGIN / SKETCH_BOX_MAX_HEIGHT.
#
# The Y position needs no calibration: it is derived at runtime from the
# "Production Size" text anchor, which PyMuPDF finds on every item page.
# (The parser's DESCRIPTION_SKIP_RE already matches "Production Size",
# confirming the string is in the text layer and not part of an image.)

SKETCH_ANCHOR_LABEL: str = "Production Size"

# Left edge of the blank column. CELL_X_RIGHT (459.65) is the right edge of
# the data-entry cells, so the free space starts just past it.
SKETCH_BOX_X_LEFT: float = 462.0
SKETCH_BOX_RIGHT_MARGIN: float = 15.0   # fallback gap from the page's right edge
SKETCH_BOX_INNER_PAD: float = 3.0       # keep clear of the block's printed border

SKETCH_BOX_TOP_GAP: float = 8.0         # below the "Production Size" row
SKETCH_BOX_MAX_HEIGHT: float = 215.0    # how far the blank column runs
SKETCH_BOX_MIN_HEIGHT: float = 55.0     # below this, don't bother — append instead
SKETCH_NEXT_ITEM_SAFETY: float = 14.0   # keep clear of the next item block

# The sketch may spill slightly outside the nominal box when that buys
# legibility — the surrounding area is blank anyway. 1.0 = strict fit.
SKETCH_OVERSIZE: float = 1.22

SKETCH_BORDER_COLOR: tuple[float, float, float] = (0.62, 0.69, 0.78)
SKETCH_DRAW_BORDER: bool = True


def _stamp_sketches_inline(
    doc: "fitz.Document",
    sketches: dict[str, bytes],
    by_sales_line: dict[str, dict[str, Any]],
) -> set[str]:
    """
    Stamp each sketch onto its own item page. Returns the set of sales lines
    successfully placed, so the caller can append the rest as fallback pages.
    """
    placed: set[str] = set()

    for page_index in range(len(doc)):
        page = doc[page_index]

        anchors = page.search_for(SKETCH_ANCHOR_LABEL)
        if not anchors:
            continue
        anchors = sorted(anchors, key=lambda r: r.y0)

        sales_line_rects = _find_sales_line_rects_on_page(page, sketches.keys())
        if not sales_line_rects:
            continue

        for i, anchor in enumerate(anchors):
            code = _sales_line_for_anchor(anchor, sales_line_rects)
            if code is None or code in placed:
                continue

            png = sketches.get(code)
            if not _is_renderable_image(png):
                continue

            # Vertical room: stop before the next item block on this page.
            next_limit = page.rect.height - SKETCH_BOX_RIGHT_MARGIN
            if i + 1 < len(anchors):
                next_limit = min(
                    next_limit, anchors[i + 1].y0 - SKETCH_NEXT_ITEM_SAFETY
                )

            box = _sketch_box(page, anchor, next_limit)
            if box is None:
                continue

            try:
                _place_sketch(page, box, png)
                placed.add(code)
            except Exception:
                # Never let one bad sketch damage the page.
                continue

    return placed


def _sales_line_for_anchor(
    anchor_rect: "fitz.Rect",
    sales_line_rects: list[tuple[str, "fitz.Rect"]],
) -> Optional[str]:
    """Closest sales-line code ABOVE this 'Production Size' anchor."""
    candidates = [
        (code, rect) for code, rect in sales_line_rects
        if rect.y1 <= anchor_rect.y0
    ]
    if not candidates:
        return None
    code, _ = min(candidates, key=lambda cr: anchor_rect.y0 - cr[1].y1)
    return code


def _block_right_edge(page: "fitz.Page") -> float:
    """
    Right edge of the printed item block.

    Measured from the page's own ruling lines rather than assumed, so the
    sketch can never overshoot the template border (seen in testing when a
    fixed page-margin was used). Falls back to the page margin if no wide
    rule is found.
    """
    best = 0.0
    for d in page.get_drawings():
        rect = d.get("rect")
        if rect is None:
            continue
        if (rect.x1 - rect.x0) > 200 and abs(rect.y1 - rect.y0) < 400:
            best = max(best, rect.x1)
    fallback = page.rect.width - SKETCH_BOX_RIGHT_MARGIN
    if best <= SKETCH_BOX_X_LEFT + 40 or best > page.rect.width:
        return fallback
    return min(best, fallback)


def _sketch_box(
    page: "fitz.Page",
    anchor_rect: "fitz.Rect",
    bottom_limit: float,
) -> Optional["fitz.Rect"]:
    """The blank column beside/below the Production Size row."""
    x0 = SKETCH_BOX_X_LEFT
    x1 = _block_right_edge(page) - SKETCH_BOX_INNER_PAD
    if x1 - x0 < 40:                      # unexpected page geometry
        return None

    y0 = anchor_rect.y1 + SKETCH_BOX_TOP_GAP
    y1 = min(y0 + SKETCH_BOX_MAX_HEIGHT, bottom_limit)
    if y1 - y0 < SKETCH_BOX_MIN_HEIGHT:
        return None

    return fitz.Rect(x0, y0, x1, y1)


def _place_sketch(page: "fitz.Page", box: "fitz.Rect", png: bytes) -> None:
    """Aspect-fit the PNG into the box (with a little overspill allowed)."""
    target = _fit_rect(box, png, oversize=SKETCH_OVERSIZE)

    # Clamp back inside the item block, whatever the oversize did. Using the
    # measured block edge (not the page edge) keeps the sketch from spilling
    # over the template's printed border.
    right = _block_right_edge(page) - SKETCH_BOX_INNER_PAD
    page_area = fitz.Rect(
        SKETCH_BOX_X_LEFT - 4, SKETCH_BOX_RIGHT_MARGIN,
        right, page.rect.height - SKETCH_BOX_RIGHT_MARGIN,
    )
    target = target & page_area
    if target.is_empty or target.width < 10 or target.height < 10:
        target = box

    page.insert_image(target, stream=png, keep_proportion=True, overlay=True)
    if SKETCH_DRAW_BORDER:
        page.draw_rect(target, color=SKETCH_BORDER_COLOR, width=0.6, overlay=True)


def _is_renderable_image(data: Any) -> bool:
    """True only if PyMuPDF can actually decode these bytes as an image."""
    if not data or not isinstance(data, (bytes, bytearray)) or len(data) < 32:
        return False
    try:
        pix = fitz.Pixmap(bytes(data))
        return pix.width > 0 and pix.height > 0
    except Exception:
        return False


def _image_size(png: bytes) -> tuple[float, float]:
    try:
        pix = fitz.Pixmap(png)
        if pix.width > 0 and pix.height > 0:
            return float(pix.width), float(pix.height)
    except Exception:
        pass
    return 780.0, 500.0


def _fit_rect(
    box: "fitz.Rect",
    png: bytes,
    oversize: float = 1.0,
    top_align: bool = True,
) -> "fitz.Rect":
    """Largest rect inside box preserving the PNG's aspect ratio, centred."""
    iw, ih = _image_size(png)
    scale = min(box.width / iw, box.height / ih) * max(1.0, oversize)
    w, h = iw * scale, ih * scale
    cx = box.x0 + (box.width - w) / 2.0
    cy = box.y0 if top_align else box.y0 + (box.height - h) / 2.0
    return fitz.Rect(cx, cy, cx + w, cy + h)


# -----------------------------------------------------------------------------
# FALLBACK — appended pages
# -----------------------------------------------------------------------------
# Used only when the "Production Size" anchor can't be found for a sketch
# (unexpected template variant). Better an extra page than a lost drawing.

SKETCH_PAGE_W: float = 595.0
SKETCH_PAGE_H: float = 842.0
SKETCH_MARGIN: float = 36.0
SKETCH_HEADER_H: float = 52.0
SKETCH_TITLE_COLOR: tuple[float, float, float] = (0.04, 0.24, 0.57)


def _append_sketch_pages(
    doc: "fitz.Document",
    sketches: dict[str, bytes],
    by_sales_line: dict[str, dict[str, Any]],
) -> None:
    for sales_line in sorted(sketches.keys()):
        png = sketches.get(sales_line)
        # Validate BEFORE creating a page — otherwise a corrupt PNG leaves an
        # empty page behind in the output.
        if not _is_renderable_image(png):
            continue

        row = by_sales_line.get(str(sales_line), {})
        before = len(doc)
        try:
            _draw_sketch_page(doc, str(sales_line), png, row)
        except Exception:
            while len(doc) > before:
                doc.delete_page(len(doc) - 1)
            continue


def _draw_sketch_page(
    doc: "fitz.Document",
    sales_line: str,
    png: bytes,
    row: dict[str, Any],
) -> None:
    page = doc.new_page(width=SKETCH_PAGE_W, height=SKETCH_PAGE_H)
    x0, x1 = SKETCH_MARGIN, SKETCH_PAGE_W - SKETCH_MARGIN

    page.insert_text(
        fitz.Point(x0, SKETCH_MARGIN + 12),
        f"SITE SKETCH  —  Sales Line {sales_line}",
        fontname=FONT_NAME_BOLD, fontsize=13, color=SKETCH_TITLE_COLOR,
    )

    bits: list[str] = []
    for label, key in (("Ref", "reference"), ("Location", "location"), ("Room", "room")):
        val = str(row.get(key) or "").strip()
        if val:
            bits.append(f"{label}: {val}")
    ow, oh = row.get("order_width"), row.get("order_height")
    if ow is not None and oh is not None:
        bits.append(f"Order: {_fmt_mm(ow)} x {_fmt_mm(oh)} mm")
    if bits:
        page.insert_textbox(
            fitz.Rect(x0, SKETCH_MARGIN + 18, x1, SKETCH_MARGIN + 46),
            "   ·   ".join(bits),
            fontname=FONT_NAME, fontsize=8.5,
            color=SURVEYOR_TEXT_COLOR, align=0,
        )

    box = fitz.Rect(
        x0, SKETCH_MARGIN + SKETCH_HEADER_H,
        x1, SKETCH_PAGE_H - SKETCH_MARGIN,
    )
    target = _fit_rect(box, png)
    page.draw_rect(target, color=SKETCH_BORDER_COLOR, width=0.7, overlay=True)
    page.insert_image(target, stream=png, keep_proportion=True, overlay=True)


def _fmt_mm(value: Any) -> str:
    try:
        return f"{int(round(float(value)))}"
    except (TypeError, ValueError):
        return "—"
