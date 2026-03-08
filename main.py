"""
Image ↔ Excel Pixel Art Converter

Converts an image into a Microsoft Excel spreadsheet where each cell
represents a single pixel (coloured via cell fill), and vice versa.

Usage
-----
Configure the Settings dataclass at the bottom of this file and run:

    python image_excel_converter.py

Or import and call the functions directly:

    from image_excel_converter import image_to_excel, excel_to_image, Settings
"""

from __future__ import annotations

import logging
import sys
import time
from dataclasses import dataclass, field
from enum import Enum, auto
from pathlib import Path
from typing import Final

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image, UnidentifiedImageError

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Square-ish cell geometry that looks good at typical zoom levels.
CELL_WIDTH: Final[float] = 3.0
CELL_HEIGHT: Final[float] = 18.0

# openpyxl encodes colour as AARRGGBB; the first two hex chars are the alpha
# channel, which we always skip when converting back to RGB.
_OPENPYXL_ALPHA_PREFIX_LEN: Final[int] = 2

# Fallback colour when a cell carries no fill information.
_WHITE: Final[tuple[int, int, int]] = (255, 255, 255)


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

class OperationMode(Enum):
    TO_EXCEL = auto()
    TO_IMAGE = auto()
    BOTH = auto()


@dataclass(frozen=True)
class Settings:
    """All user-facing knobs in one place."""

    mode: OperationMode = OperationMode.TO_IMAGE

    input_image_path: Path = Path("headshot-ascii-square.png")
    excel_path: Path = Path("pixels.xlsx")
    output_image_path: Path = Path("z.png")

    # Pixel-art resolution.
    grid_width: int = 890
    grid_height: int = 890

    # Multiplier applied when reconstructing the image from Excel.
    upscale_factor: int = 1

    def __post_init__(self) -> None:
        if self.grid_width < 1 or self.grid_height < 1:
            raise ValueError("grid_width and grid_height must be positive integers.")
        if self.upscale_factor < 1:
            raise ValueError("upscale_factor must be a positive integer.")


# ---------------------------------------------------------------------------
# Core logic
# ---------------------------------------------------------------------------

def image_to_excel(
    image_path: Path,
    excel_path: Path,
    grid_width: int = 64,
    grid_height: int = 64,
) -> None:
    """
    Resize *image_path* to (*grid_width* × *grid_height*) and write each pixel
    as a coloured cell fill to a new Excel workbook saved at *excel_path*.

    Parameters
    ----------
    image_path:
        Source image (any format supported by Pillow).
    excel_path:
        Destination ``.xlsx`` file.
    grid_width:
        Number of pixel-columns in the output grid.
    grid_height:
        Number of pixel-rows in the output grid.

    Raises
    ------
    FileNotFoundError
        If *image_path* does not exist.
    UnidentifiedImageError
        If *image_path* cannot be decoded as an image.
    """
    log.info("Image → Excel  |  source=%s  grid=%dx%d", image_path, grid_width, grid_height)

    if not image_path.exists():
        raise FileNotFoundError(f"Input image not found: {image_path}")

    with Image.open(image_path) as raw:
        pixel_art = raw.resize((grid_width, grid_height), Image.Resampling.NEAREST).convert("RGB")

    wb = Workbook()
    ws = wb.active
    ws.title = "Pixel Art"

    # Read all pixel data in one shot — far faster than per-pixel getpixel() calls.
    pixels: list[tuple[int, int, int]] = list(pixel_art.getdata())

    # Cache PatternFill objects by hex colour; openpyxl is slow to create styles.
    fill_cache: dict[str, PatternFill] = {}

    log.info("Writing %d pixels to workbook…", len(pixels))
    for idx, (r, g, b) in enumerate(pixels):
        hex_colour = f"{r:02x}{g:02x}{b:02x}"
        if hex_colour not in fill_cache:
            fill_cache[hex_colour] = PatternFill(
                start_color=hex_colour,
                end_color=hex_colour,
                fill_type="solid",
            )
        row = idx // grid_width + 1   # 1-indexed
        col = idx % grid_width + 1    # 1-indexed
        ws.cell(row=row, column=col).fill = fill_cache[hex_colour]

    log.info("Adjusting cell dimensions…")
    for col_idx in range(1, grid_width + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = CELL_WIDTH
    for row_idx in range(1, grid_height + 1):
        ws.row_dimensions[row_idx].height = CELL_HEIGHT

    wb.save(excel_path)
    log.info("Saved workbook → %s  (%d unique colours)", excel_path, len(fill_cache))


def excel_to_image(
    excel_path: Path,
    output_image_path: Path,
    upscale_factor: int = 8,
) -> None:
    """
    Reconstruct an image from a pixel-art Excel workbook produced by
    :func:`image_to_excel`.

    Cells without an explicit fill are treated as white.

    Parameters
    ----------
    excel_path:
        Source ``.xlsx`` file.
    output_image_path:
        Destination image file (format inferred from extension by Pillow).
    upscale_factor:
        Integer scale applied to the final image via nearest-neighbour
        resampling, so individual pixels remain sharp.

    Raises
    ------
    FileNotFoundError
        If *excel_path* does not exist.
    """
    log.info("Excel → Image  |  source=%s  scale=×%d", excel_path, upscale_factor)

    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    # read_only=True skips loading write-time metadata and is meaningfully
    # faster for large sheets.
    wb = load_workbook(filename=excel_path, read_only=True, data_only=True)
    ws = wb.active

    grid_width: int = ws.max_column
    grid_height: int = ws.max_row
    log.info("Sheet dimensions: %dx%d cells", grid_width, grid_height)

    img = Image.new("RGB", (grid_width, grid_height), color=_WHITE)
    pixel_access = img.load()

    for row_idx, row in enumerate(ws.iter_rows()):
        for col_idx, cell in enumerate(row):
            colour = _parse_cell_colour(cell)
            pixel_access[col_idx, row_idx] = colour

    if upscale_factor > 1:
        img = img.resize(
            (grid_width * upscale_factor, grid_height * upscale_factor),
            Image.Resampling.NEAREST,
        )

    img.save(output_image_path)
    log.info("Saved image → %s", output_image_path)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _parse_cell_colour(cell) -> tuple[int, int, int]:
    """
    Extract the RGB fill colour from an openpyxl cell.

    openpyxl expresses colours as 8-character AARRGGBB hex strings.
    Returns *_WHITE* when no valid fill colour is present.
    """
    try:
        rgb_hex = cell.fill.start_color.rgb
        if not rgb_hex or len(rgb_hex) < 8:
            return _WHITE
        hex_body = rgb_hex[_OPENPYXL_ALPHA_PREFIX_LEN:]  # Strip alpha prefix
        return (
            int(hex_body[0:2], 16),
            int(hex_body[2:4], 16),
            int(hex_body[4:6], 16),
        )
    except (AttributeError, ValueError, TypeError):
        return _WHITE


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def run(settings: Settings) -> None:
    """Execute the conversion(s) described by *settings*."""
    start = time.perf_counter()

    try:
        if settings.mode in (OperationMode.TO_EXCEL, OperationMode.BOTH):
            image_to_excel(
                image_path=settings.input_image_path,
                excel_path=settings.excel_path,
                grid_width=settings.grid_width,
                grid_height=settings.grid_height,
            )

        if settings.mode in (OperationMode.TO_IMAGE, OperationMode.BOTH):
            excel_to_image(
                excel_path=settings.excel_path,
                output_image_path=settings.output_image_path,
                upscale_factor=settings.upscale_factor,
            )

    except (FileNotFoundError, UnidentifiedImageError) as exc:
        log.error("%s", exc)
        sys.exit(1)
    except Exception:
        log.exception("Unexpected error during conversion.")
        sys.exit(1)

    elapsed = time.perf_counter() - start
    log.info("Done in %.2f s", elapsed)


if __name__ == "__main__":
    # -----------------------------------------------------------------------
    # ✏️  Edit this block to configure your run.
    # -----------------------------------------------------------------------
    settings = Settings(
        mode=OperationMode.BOTH,
        input_image_path=Path("image-1920x1080.png"),
        excel_path=Path("pixels.xlsx"),
        output_image_path=Path("z.png"),
        grid_width=1920,
        grid_height=1080,
        upscale_factor=5,
    )
    # -----------------------------------------------------------------------
    run(settings)
