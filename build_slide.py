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
import os
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu


# 16:9 PowerPoint canvas matching the original project.
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
    files = [
        p for p in input_dir.iterdir()
        if p.is_file() and p.suffix.lower() in IMAGE_EXTENSIONS
    ]
    return sorted(files, key=lambda p: p.name.lower())


def add_image_slide(prs: Presentation, image_path: Path) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Clean white presentation background.
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = __import__("pptx.dml.color", fromlist=["RGBColor"]).RGBColor(
        255, 255, 255
    )

    with Image.open(image_path) as im:
        width_px, height_px = im.size

    if width_px <= 0 or height_px <= 0:
        raise ValueError(f"Invalid image dimensions: {image_path.name}")

    # Fit the complete image inside the slide without distortion.
    slide_ratio = SLIDE_W / SLIDE_H
    image_ratio = width_px / height_px

    if image_ratio >= slide_ratio:
        pic_w = int(SLIDE_W)
        pic_h = int(round(pic_w / image_ratio))
        left = 0
        top = int((SLIDE_H - Emu(pic_h)) / 2)
    else:
        pic_h = int(SLIDE_H)
        pic_w = int(round(pic_h * image_ratio))
        top = 0
        left = int((SLIDE_W - Emu(pic_w)) / 2)

    slide.shapes.add_picture(
        str(image_path),
        Emu(left),
        Emu(top),
        width=Emu(pic_w),
        height=Emu(pic_h),
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

    # Presentation() starts empty; add exactly one slide for each image.
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
    output = convert_images_to_pptx(Path(args.input_dir), Path(args.output))
    count = len(image_files(Path(args.input_dir)))
    print(f"IS BUILD COMPLETE")
    print(f"IMAGES: {count}")
    print(f"SLIDES: {count}")
    print(f"OUTPUT: {output}")


if __name__ == "__main__":
    main()
