"""
Extractor for embedded images in an Excel sheet - used to pull images from
sheet "1. Product Description" into Section 2 (Product Hierarchy) of the
generated document, per spec.

Images are saved to disk and their paths returned in top-to-bottom sheet
order (sorted by each image's anchor row), so they can be inserted into
the target document section in a sensible reading order.

UNVALIDATED: no sheet has been seen with real embedded images yet. This
currently places all images together under the Section 2 intro (not tied
to a specific product/attribute row). If images need to appear next to a
SPECIFIC row, share a sample workbook with images and their intended
target rows so the anchor-row matching can be made precise (the same way
stacked_tables.py matches by row range).
"""
from __future__ import annotations
import os
import openpyxl
from extractors.sheet_utils import find_sheet


def extract_hierarchy_images(
    workbook_path: str,
    sheet_name: str = "1. Product Description",
    output_dir: str = "output/_extracted_images",
) -> list[str]:
    wb = openpyxl.load_workbook(workbook_path)
    resolved = find_sheet(wb, sheet_name)
    ws = wb[resolved]

    images = list(getattr(ws, "_images", []))
    if not images:
        return []

    def _anchor_row(img) -> int:
        try:
            return img.anchor._from.row
        except AttributeError:
            return 0

    images.sort(key=_anchor_row)

    os.makedirs(output_dir, exist_ok=True)
    paths: list[str] = []
    for i, img in enumerate(images):
        ext = (getattr(img, "format", None) or "png").lower()
        out_path = os.path.join(output_dir, f"hierarchy_image_{i + 1}.{ext}")
        try:
            img.image.save(out_path)  # PIL-backed path
        except Exception:
            with open(out_path, "wb") as f:
                f.write(img._data())  # raw-bytes fallback
        paths.append(out_path)

    return paths

    images = list(getattr(ws, "_images", []))
    if not images:
        return []

    def _anchor_row(img) -> int:
        try:
            return img.anchor._from.row
        except AttributeError:
            return 0

    images.sort(key=_anchor_row)

    os.makedirs(output_dir, exist_ok=True)
    paths: list[str] = []
    for i, img in enumerate(images):
        ext = (getattr(img, "format", None) or "png").lower()
        out_path = os.path.join(output_dir, f"hierarchy_image_{i + 1}.{ext}")
        try:
            img.image.save(out_path)  # PIL-backed path
        except Exception:
            with open(out_path, "wb") as f:
                f.write(img._data())  # raw-bytes fallback
        paths.append(out_path)

    return paths
