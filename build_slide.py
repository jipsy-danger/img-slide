#!/usr/bin/env python3
"""
IS — Image to PowerPoint Slide Converter

Converts every image in an input directory into one PowerPoint slide.
Each input image becomes exactly one slide, preserving its aspect ratio
and centering it on a 16:9 presentation canvas.

Local usage:
    python build_slide.py ./images -o IS_Images_to_PowerPoint.pptx

Browser usage:
    The IS web app sets IMG_SLIDE_INPUT_DIR and IMG_SLIDE_OUTPUT and
    executes this same file through Pyodide/WebAssembly.
"""

from __future__ import annotations

import argparse
import io
import os
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Emu


SLIDE_W = Emu(12192000)
SLIDE_H = Emu(6858000)

IMAGE_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".webp",
    ".bmp",
    ".gif",
    ".tif",
    ".tiff",
}


def image_files(input_dir: Path) -> list[Path]:
    return sorted(
        (
            path
            for path in input_dir.iterdir()
            if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS
        ),
        key=lambda path: path.name.lower(),
    )


def add_image_slide(prs: Presentation, image_path: Path) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    background = slide.background.fill
    background.solid()
    background.fore_color.rgb = RGBColor(255, 255, 255)

    with Image.open(image_path) as image:
        image.load()
        width_px, height_px = image.size

        if width_px <= 0 or height_px <= 0:
            raise ValueError(f"Invalid image dimensions: {image_path.name}")

        # Normalize every supported input format to PNG bytes so WEBP/BMP/GIF/TIFF
        # files are still accepted by the final PowerPoint.
        rgba = image.convert("RGBA")
        image_bytes = io.BytesIO()
        rgba.save(image_bytes, format="PNG")
        image_bytes.seek(0)

    slide_ratio = int(SLIDE_W) / int(SLIDE_H)
    image_ratio = width_px / height_px

    if image_ratio >= slide_ratio:
        picture_width = int(SLIDE_W)
        picture_height = int(round(picture_width / image_ratio))
        left = 0
        top = int((int(SLIDE_H) - picture_height) / 2)
    else:
        picture_height = int(SLIDE_H)
        picture_width = int(round(picture_height * image_ratio))
        top = 0
        left = int((int(SLIDE_W) - picture_width) / 2)

    slide.shapes.add_picture(
        image_bytes,
        Emu(left),
        Emu(top),
        width=Emu(picture_width),
        height=Emu(picture_height),
    )


def convert_images_to_pptx(input_dir: Path, output_file: Path) -> Path:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder does not exist: {input_dir}")

    if not input_dir.is_dir():
        raise NotADirectoryError(f"Input path is not a folder: {input_dir}")

    images = image_files(input_dir)
    if not images:
        raise ValueError(
            "No supported images were found. Add PNG, JPG, JPEG, WEBP, BMP, GIF, "
            "TIF or TIFF files."
        )

    output_file.parent.mkdir(parents=True, exist_ok=True)

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    for image_path in images:
        add_image_slide(prs, image_path)

    prs.save(str(output_file))
    return output_file


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert images into a PowerPoint presentation."
    )
    parser.add_argument(
        "input_dir",
        nargs="?",
        default=os.environ.get("IMG_SLIDE_INPUT_DIR", "/workspace/input"),
        help="Folder containing source images.",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=os.environ.get(
            "IMG_SLIDE_OUTPUT",
            "/mnt/user-data/outputs/IS_Images_to_PowerPoint.pptx",
        ),
        help="Output .pptx path.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_dir = Path(args.input_dir)
    output = convert_images_to_pptx(input_dir, Path(args.output))
    count = len(image_files(input_dir))
    print("IS BUILD COMPLETE")
    print(f"IMAGES: {count}")
    print(f"SLIDES: {count}")
    print(f"OUTPUT: {output}")


if __name__ == "__main__":
    main()
