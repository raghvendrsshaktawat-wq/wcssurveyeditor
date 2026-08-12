"""
sketch.py — Site Sketch Pad (Module 7)
------------------------------------------------------------------------------
Per-sales-line freehand sketching with text boxes, in a modal pop-up.

Flow:
    click "✏️ 0007" in the sketch rail  →  st.dialog opens  →  draw freehand /
    drop text boxes  →  "💾 Save sketch"  →  PNG kept in session_state
    →  dialog closes  →  the sales-line cell in the grid turns green
    →  PNG is stamped INLINE on the item page of the annotated PDF, beside
       the "Production Size" row, and embedded in Excel.

Why a hand-rolled component instead of streamlit-drawable-canvas / aggrid:
    • No extra pip dependency, no npm build, no licence question.
    • ZERO Streamlit reruns while drawing — every tool switch, colour change,
      stroke and keystroke happens inside the iframe. Streamlit is contacted
      exactly once, when the user hits Save (or Cancel). That is what kills
      the flicker / "canvas resets when I change tool" problem.

STORAGE NOTE (05-Aug-2026):
    Sketches are held in st.session_state ONLY — nothing is written to disk
    by default. On Streamlit Community Cloud the filesystem is ephemeral
    anyway, and keeping order data off the host is a deliberate choice.
    Set WCS_SKETCH_DIR to a real path if you DO want local copies (useful
    when running on a Fenesta machine with a OneDrive-synced folder).

    Consequence: refreshing the browser starts a new Streamlit session and
    the sketches are gone. Generate the annotated PDF before closing the tab.

Public API:
    render_sketch_dialog_if_open()              -> call once per script run
    open_sketch(file_key, sales_line, subtitle) -> queue the pop-up
    has_sketch(file_key, sales_line)            -> bool
    drawn_set(file_key, sales_lines)            -> set[str]  (grid formatting)
    get_png(file_key, sales_line)               -> bytes | None
    collect_for_file(file_key, sales_lines)     -> {sales_line: png_bytes}
    total_bytes(file_key, sales_lines)          -> int
    delete_sketch(file_key, sales_line)
"""

from __future__ import annotations

import base64
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import streamlit as st
import streamlit.components.v1 as components

__version__ = "1.3.0"

# =============================================================================
# Optional on-disk copies (OFF by default — see STORAGE NOTE above)
# =============================================================================
_SKETCH_DIR_ENV = os.environ.get("WCS_SKETCH_DIR", "").strip()
SKETCH_ROOT: Path | None = Path(_SKETCH_DIR_ENV) if _SKETCH_DIR_ENV else None

# Session-state keys
_STORE_KEY = "_wcs_sketches"      # {(file_key, sales_line): {...}}
_OPEN_KEY = "_wcs_sketch_open"    # (file_key, sales_line, subtitle, order_no)
_SEQ_KEY = "_wcs_sketch_seq"      # {canvas_key: last processed seq}

# =============================================================================
# Component registration
# =============================================================================
_BUILD_DIR = Path(__file__).parent / "sketch_canvas"

_canvas = components.declare_component("wcs_sketch_canvas", path=str(_BUILD_DIR))


def component_available() -> bool:
    """True if the sketch_canvas/index.html asset shipped alongside this file."""
    return (_BUILD_DIR / "index.html").exists()


# =============================================================================
# Store helpers
# =============================================================================
def _store() -> dict[tuple[str, str], dict[str, Any]]:
    if _STORE_KEY not in st.session_state:
        st.session_state[_STORE_KEY] = {}
    return st.session_state[_STORE_KEY]


def _slug(value: Any, fallback: str = "x") -> str:
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", str(value or "").strip()).strip("_")
    return s[:60] or fallback


def has_sketch(file_key: str, sales_line: str) -> bool:
    return (str(file_key), str(sales_line)) in _store()


def drawn_set(file_key: str, sales_lines: Iterable[str]) -> set[str]:
    """
    Set of sales lines (for THIS file) that already have a sketch.

    Used by main._build_row_styles to conditionally format the Sales Line
    column. Computed once per run and passed down, so the grid never does a
    session_state lookup per cell.
    """
    fk = str(file_key)
    return {str(sl) for sl in sales_lines if (fk, str(sl)) in _store()}


def get_entry(file_key: str, sales_line: str) -> dict[str, Any] | None:
    return _store().get((str(file_key), str(sales_line)))


def get_png(file_key: str, sales_line: str) -> bytes | None:
    entry = get_entry(file_key, sales_line)
    return entry["png"] if entry else None


def collect_for_file(file_key: str, sales_lines: Iterable[str]) -> dict[str, bytes]:
    """{sales_line: png_bytes} for every line of this file that has a sketch."""
    out: dict[str, bytes] = {}
    for sl in sales_lines:
        png = get_png(file_key, str(sl))
        if png:
            out[str(sl)] = png
    return out


def count_for_file(file_key: str, sales_lines: Iterable[str]) -> int:
    return sum(1 for sl in sales_lines if has_sketch(file_key, str(sl)))


def total_bytes(file_key: str, sales_lines: Iterable[str]) -> int:
    """Total PNG weight held in memory for this file — surfaced in the UI."""
    return sum(len(p) for p in collect_for_file(file_key, sales_lines).values())


def delete_sketch(file_key: str, sales_line: str) -> None:
    entry = _store().pop((str(file_key), str(sales_line)), None)
    if entry and entry.get("path"):
        try:
            Path(entry["path"]).unlink(missing_ok=True)
        except OSError:
            pass


def _save(
    file_key: str,
    sales_line: str,
    png_bytes: bytes,
    items: list[Any],
    order_no: str,
    meta: dict[str, Any] | None = None,
) -> None:
    """Register the sketch in session_state (and on disk if WCS_SKETCH_DIR is set)."""
    path = ""
    if SKETCH_ROOT is not None:
        try:
            folder = SKETCH_ROOT / _slug(order_no, "order")
            folder.mkdir(parents=True, exist_ok=True)
            path = str(folder / f"{_slug(sales_line, 'line')}.png")
            with open(path, "wb") as fh:
                fh.write(png_bytes)
        except OSError:
            path = ""      # read-only / bad path — memory copy still works

    entry = {
        "png": png_bytes,
        "items": items,
        "path": path,
        "saved_at": datetime.now().strftime("%d-%b-%Y %H:%M"),
    }
    entry.update(meta or {})
    _store()[(str(file_key), str(sales_line))] = entry


# =============================================================================
# Pop-up plumbing
# =============================================================================
def open_sketch(
    file_key: str,
    sales_line: str,
    subtitle: str = "",
    order_no: str = "",
) -> None:
    """Queue the sketch pop-up for this sales line and rerun."""
    st.session_state[_OPEN_KEY] = (
        str(file_key), str(sales_line), subtitle, str(order_no or ""),
    )
    st.rerun()


def _close() -> None:
    st.session_state.pop(_OPEN_KEY, None)
    st.rerun()


def _decode_png(data_url: str) -> bytes:
    """'data:image/png;base64,AAA...' -> raw bytes."""
    if not data_url:
        return b""
    if "," in data_url:
        data_url = data_url.split(",", 1)[1]
    try:
        return base64.b64decode(data_url)
    except Exception:
        return b""


def _dialog(title: str):
    """st.dialog on modern Streamlit, st.experimental_dialog on older builds."""
    deco = getattr(st, "dialog", None) or getattr(st, "experimental_dialog", None)
    if deco is None:
        return None
    try:
        return deco(title, width="large")
    except TypeError:            # older signature without width=
        return deco(title)


def render_sketch_dialog_if_open() -> None:
    """Call this ONCE per script run (early in main())."""
    target = st.session_state.get(_OPEN_KEY)
    if not target:
        return

    file_key, sales_line, subtitle, order_no = target

    # Mobile: expand the sketch dialog to fill the screen so the canvas is
    # usable on a phone (reported too small). Media-query gated at 768px, so
    # desktop/laptop are completely unaffected.
    st.markdown(
        """<style>
        @media (max-width: 768px) {
          div[data-testid="stDialog"] > div,
          div[data-testid="stDialog"] div[role="dialog"] {
            width: 100vw !important; max-width: 100vw !important;
            height: 100dvh !important; max-height: 100dvh !important;
            margin: 0 !important; border-radius: 0 !important;
          }
        }
        </style>""",
        unsafe_allow_html=True,
    )

    deco = _dialog(f"✏️  Site Sketch — Sales Line {sales_line}")

    def _body() -> None:
        if subtitle:
            st.caption(subtitle)

        if not component_available():
            st.error(
                "Sketch pad asset missing. Make sure the folder "
                "`sketch_canvas/` (with `index.html`) sits next to `sketch.py`."
            )
            if st.button("Close", use_container_width=True):
                _close()
            return

        entry = get_entry(file_key, sales_line)
        canvas_key = f"wcs_canvas_{_slug(file_key)}_{_slug(sales_line)}"

        value = _canvas(
            key=canvas_key,
            initial_items=(entry or {}).get("items") or [],
            grid=True,
            default=None,
        )

        if entry:
            kb = len(entry["png"]) / 1024
            st.caption(
                f"Existing sketch loaded — saved {entry['saved_at']} · {kb:.1f} KB"
            )

        # ---- handle the one message the canvas sends ------------------------
        # The canvas stamps every message with a monotonic Date.now(). Streamlit
        # re-delivers the LAST component value on every rerun, so without this
        # guard the dialog would re-save (and re-close) itself in a loop.
        # Monotonic '>' — not '!=' — so a replayed/stale value is always ignored.
        seq_map = st.session_state.setdefault(_SEQ_KEY, {})
        if isinstance(value, dict):
            try:
                seq = int(value.get("seq") or 0)
            except (TypeError, ValueError):
                seq = 0

            if seq > int(seq_map.get(canvas_key, 0)):
                seq_map[canvas_key] = seq

                if value.get("action") == "cancel":
                    _close()
                    return

                png = _decode_png(value.get("png", ""))
                if int(value.get("count") or 0) == 0:
                    st.warning("Canvas is empty — draw something before saving.")
                elif not png:
                    st.error("Nothing came back from the canvas — try Save again.")
                else:
                    _save(
                        file_key, sales_line, png,
                        value.get("items") or [], order_no,
                        meta={
                            "px_w": value.get("px_w"),
                            "px_h": value.get("px_h"),
                            "scale": value.get("scale"),
                        },
                    )
                    _close()
                    return

        # ---- footer --------------------------------------------------------
        c1, c2 = st.columns([1, 1])
        with c1:
            if entry and st.button(
                "🗑 Delete saved sketch",
                key=f"del_{canvas_key}", use_container_width=True,
            ):
                delete_sketch(file_key, sales_line)
                _close()
                return
        with c2:
            if st.button("Close without saving", key=f"cls_{canvas_key}",
                         use_container_width=True):
                _close()
                return

    if deco is None:
        # Streamlit too old for modals — degrade to an inline panel, never crash.
        st.warning("Your Streamlit build has no modal support — showing inline.")
        _body()
    else:
        deco(_body)()
