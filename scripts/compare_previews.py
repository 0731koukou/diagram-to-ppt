from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import numpy as np
from PIL import Image, ImageChops, ImageDraw, ImageFilter, ImageFont


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def foreground_bbox(image: Image.Image) -> list[int] | None:
    gray = np.asarray(image.convert("L"))
    mask = Image.fromarray(np.where(gray < 245, 255, 0).astype(np.uint8))
    box = mask.getbbox()
    return list(box) if box else None


def compare(source: Image.Image, candidate: Image.Image) -> dict[str, object]:
    source_rgb = source.convert("RGB")
    candidate_rgb = candidate.convert("RGB").resize(source_rgb.size, Image.Resampling.LANCZOS)
    a = np.asarray(source_rgb, dtype=np.float32)
    b = np.asarray(candidate_rgb, dtype=np.float32)
    delta = np.abs(a - b)
    gray_a = np.asarray(source_rgb.convert("L"))
    gray_b = np.asarray(candidate_rgb.convert("L"))
    foreground = (gray_a < 245) | (gray_b < 245)

    edge_a = np.asarray(source_rgb.convert("L").filter(ImageFilter.FIND_EDGES)) > 35
    edge_b = np.asarray(candidate_rgb.convert("L").filter(ImageFilter.FIND_EDGES)) > 35
    edge_union = np.logical_or(edge_a, edge_b).sum()
    edge_intersection = np.logical_and(edge_a, edge_b).sum()

    return {
        "source_size": list(source_rgb.size),
        "candidate_original_size": list(candidate.size),
        "pixel_similarity": round(1.0 - float(delta.mean() / 255.0), 6),
        "foreground_similarity": round(
            1.0 - float(delta[foreground].mean() / 255.0) if foreground.any() else 1.0,
            6,
        ),
        "edge_iou": round(float(edge_intersection / edge_union) if edge_union else 1.0, 6),
        "source_foreground_bbox": foreground_bbox(source_rgb),
        "candidate_foreground_bbox": foreground_bbox(candidate_rgb),
        "manual_review_required": True,
        "note": "Similarity metrics are advisory; white backgrounds can inflate pixel similarity.",
    }


def make_contact_sheet(source: Image.Image, candidate: Image.Image, output: Path) -> None:
    source_rgb = source.convert("RGB")
    candidate_rgb = candidate.convert("RGB").resize(source_rgb.size, Image.Resampling.LANCZOS)
    diff = ImageChops.difference(source_rgb, candidate_rgb).point(lambda value: min(255, value * 3))
    width, height = source_rgb.size
    label_height = 34
    sheet = Image.new("RGB", (width * 3, height + label_height), "white")
    sheet.paste(source_rgb, (0, label_height))
    sheet.paste(candidate_rgb, (width, label_height))
    sheet.paste(diff, (width * 2, label_height))
    draw = ImageDraw.Draw(sheet)
    font = ImageFont.load_default(size=18)
    for index, label in enumerate(("SOURCE", "PPT RENDER", "DIFF x3")):
        draw.text((index * width + 12, 7), label, fill="#111827", font=font)
    output.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(output)


def main() -> int:
    parser = argparse.ArgumentParser(description="Create advisory visual metrics and a comparison sheet for a diagram redraw.")
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--candidate", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--contact-sheet", type=Path)
    args = parser.parse_args()

    source = Image.open(args.source)
    candidate = Image.open(args.candidate)
    result = compare(source, candidate)
    result["source"] = str(args.source)
    result["candidate"] = str(args.candidate)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.contact_sheet:
        make_contact_sheet(source, candidate, args.contact_sheet)
        result["contact_sheet"] = str(args.contact_sheet)
        args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
