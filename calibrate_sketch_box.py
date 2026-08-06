"""
calibrate_sketch_box.py — verify WHERE the site sketch will land
------------------------------------------------------------------------------
The sketch is stamped inline on each item page, into the blank column beside
the "Production Size" row. The vertical position is found at runtime from the
text anchor, so it needs no tuning. The HORIZONTAL constants in overlay.py
were estimated from a screenshot, NOT measured from a real PDF — so run this
once against a real order before trusting them.

WHAT IT DOES
    • finds every "Production Size" anchor in the PDF
    • computes the exact box overlay.py would use
    • draws that box in RED, with a magenta dashed line at the measured
      right edge of the item block
    • writes preview PNGs you can just look at
    • prints the numbers, so you know what to change if it's off

USAGE (from the app folder)
    python calibrate_sketch_box.py "C:\\path\\to\\order.pdf"
    python calibrate_sketch_box.py order.pdf --pages 1,2 --dpi 130

IF THE BOX IS WRONG, edit these in overlay.py:
    SKETCH_BOX_X_LEFT        move the left edge  (bigger = further right)
    SKETCH_BOX_MAX_HEIGHT    how far down the column runs
    SKETCH_BOX_TOP_GAP       gap under the Production Size row
    SKETCH_BOX_INNER_PAD     clearance from the block's printed border
    SKETCH_OVERSIZE          1.0 = strict fit, >1 lets the sketch spill a little
Then re-run this script until the red box sits where you want the sketch.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("PyMuPDF missing.  Run:  python -m pip install pymupdf")

try:
    import overlay
except ImportError:
    sys.exit("Run this from the same folder as overlay.py.")


RED = (0.91, 0.13, 0.18)
MAGENTA = (0.8, 0.0, 0.8)
BLUE = (0.0, 0.36, 0.67)


def calibrate(pdf_path: Path, pages: list[int] | None, dpi: int) -> int:
    doc = fitz.open(str(pdf_path))
    out_dir = pdf_path.parent / "sketch_box_preview"
    out_dir.mkdir(exist_ok=True)

    print(f"\nFile      : {pdf_path.name}")
    print(f"Pages     : {len(doc)}")
    print(f"Anchor    : {overlay.SKETCH_ANCHOR_LABEL!r}")
    print(f"X left    : {overlay.SKETCH_BOX_X_LEFT}")
    print(f"Max height: {overlay.SKETCH_BOX_MAX_HEIGHT}")
    print(f"Oversize  : {overlay.SKETCH_OVERSIZE}")
    print("-" * 78)

    total_anchors = 0
    previews = 0

    for pno in range(len(doc)):
        if pages and (pno + 1) not in pages:
            continue

        page = doc[pno]
        anchors = sorted(page.search_for(overlay.SKETCH_ANCHOR_LABEL),
                         key=lambda r: r.y0)
        if not anchors:
            continue

        total_anchors += len(anchors)
        right = overlay._block_right_edge(page)
        print(f"\nPage {pno + 1}:  {len(anchors)} anchor(s)   "
              f"page width {page.rect.width:.0f} pt   "
              f"measured block right edge {right:.1f} pt")

        # magenta = where we think the block border is
        page.draw_line(fitz.Point(right, 0), fitz.Point(right, page.rect.height),
                       color=MAGENTA, width=0.9, dashes="[3 3] 0")

        for i, anchor in enumerate(anchors):
            limit = page.rect.height - overlay.SKETCH_BOX_RIGHT_MARGIN
            if i + 1 < len(anchors):
                limit = min(limit,
                            anchors[i + 1].y0 - overlay.SKETCH_NEXT_ITEM_SAFETY)

            box = overlay._sketch_box(page, anchor, limit)

            # blue = the anchor text itself
            page.draw_rect(anchor, color=BLUE, width=0.7)

            if box is None:
                print(f"   anchor {i + 1}: y={anchor.y0:.0f}  -> NO BOX "
                      f"(not enough room; sketch would fall back to its own page)")
                continue

            page.draw_rect(box, color=RED, width=1.4)
            page.insert_text(
                fitz.Point(box.x0 + 2, box.y0 - 3),
                f"SKETCH BOX {box.width:.0f} x {box.height:.0f} pt",
                fontsize=6, color=RED,
            )
            print(f"   anchor {i + 1}: y={anchor.y0:.0f}  -> box "
                  f"x {box.x0:.0f}..{box.x1:.0f}   y {box.y0:.0f}..{box.y1:.0f}   "
                  f"({box.width:.0f} x {box.height:.0f} pt)")

        png = out_dir / f"page_{pno + 1:02d}.png"
        page.get_pixmap(dpi=dpi).save(str(png))
        previews += 1

    doc.close()

    print("\n" + "=" * 78)
    if total_anchors == 0:
        print("NO ANCHORS FOUND.")
        print(f"  '{overlay.SKETCH_ANCHOR_LABEL}' is not in this PDF's text layer.")
        print("  Either the label differs on this template, or the page is a")
        print("  scanned image. Every sketch would fall back to its own appended")
        print("  page — still usable, just not inline.")
        return 1

    print(f"{total_anchors} anchor(s) found · {previews} preview image(s) written to:")
    print(f"   {out_dir}")
    print("\nOpen them and check the RED box:")
    print("  • sitting in the blank column, clear of printed text  -> good")
    print("  • overlapping the spec list  -> raise SKETCH_BOX_X_LEFT")
    print("  • past the magenta line      -> raise SKETCH_BOX_INNER_PAD")
    print("  • too short/tall             -> adjust SKETCH_BOX_MAX_HEIGHT")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="Preview the inline sketch box.")
    ap.add_argument("pdf", help="a real Fenesta WCS order PDF")
    ap.add_argument("--pages", default="", help="e.g. 1,2,5 (default: all)")
    ap.add_argument("--dpi", type=int, default=120)
    args = ap.parse_args()

    pdf_path = Path(args.pdf).expanduser()
    if not pdf_path.exists():
        print(f"Not found: {pdf_path}")
        return 1

    pages = None
    if args.pages.strip():
        try:
            pages = [int(x) for x in args.pages.split(",") if x.strip()]
        except ValueError:
            print("--pages must look like: 1,2,5")
            return 1

    return calibrate(pdf_path, pages, args.dpi)


if __name__ == "__main__":
    raise SystemExit(main())
