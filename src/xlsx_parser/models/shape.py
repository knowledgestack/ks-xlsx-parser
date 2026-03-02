"""
Shape and image DTOs for embedded objects in worksheets.

Covers images, text boxes, and other drawing objects extracted from
the OOXML drawing parts.
"""

from __future__ import annotations

from pydantic import Field

from .common import CellRange, StableModel, compute_hash


class ShapeAnchor(StableModel):
    """Position anchor for a shape/image on the worksheet."""

    from_col: int
    from_row: int
    to_col: int | None = None
    to_row: int | None = None
    anchor_type: str = "twoCellAnchor"  # or "oneCellAnchor", "absoluteAnchor"


class ShapeDTO(StableModel):
    """
    An embedded shape, image, or text box in a worksheet.

    Shapes include images (PNG, JPEG, etc.), text boxes, rectangles,
    and other drawing objects. Each is anchored to cell coordinates
    for positioning.
    """

    model_config = {"frozen": False, "extra": "forbid"}

    # Identity
    shape_index: int
    sheet_name: str
    shape_id: str = Field(default="", description="Deterministic ID")

    # Type
    shape_type: str  # "image", "textBox", "rectangle", "line", "connector", etc.

    # Content
    alt_text: str | None = None
    text_content: str | None = None  # For text boxes
    image_ref: str | None = None  # Path within the OOXML package (e.g., "/xl/media/image1.png")
    image_content_type: str | None = None  # MIME type

    # Position
    anchor: ShapeAnchor | None = None

    # Dimensions (in EMUs - English Metric Units)
    width_emu: int | None = None
    height_emu: int | None = None

    # Hash
    content_hash: str = Field(default="")

    def finalize(self, workbook_hash: str) -> None:
        """Compute stable ID and content hash."""
        self.shape_id = compute_hash(
            workbook_hash, self.sheet_name, str(self.shape_index)
        )
        self.content_hash = compute_hash(
            self.sheet_name,
            self.shape_type,
            self.alt_text or "",
            self.text_content or "",
            self.image_ref or "",
        )
