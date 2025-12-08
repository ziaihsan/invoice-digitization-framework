#!/usr/bin/env python3
"""Replace the working annotations/images set with a fresh random sample."""

from __future__ import annotations

import argparse
import random
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple


ROOT_DIR = Path(__file__).resolve().parent
DEFAULT_ANNOTATIONS = ROOT_DIR / "annotations"
DEFAULT_IMAGES = ROOT_DIR / "images"
DEFAULT_TRASH = ROOT_DIR / "trash"
DEFAULT_SOURCE_IMAGES = ROOT_DIR / "Full_invoices_dataset" / "images"
DEFAULT_SOURCE_ANNOTATIONS = (
    ROOT_DIR / "Full_invoices_dataset" / "Annotations" / "Original_Format"
)
DEFAULT_CONVERTER_SCRIPT = ROOT_DIR / "convert_json_to_csv.py"


JsonImagePair = Tuple[Path, Path]
TEMPLATE1_PREFIX = "Template1_Instance"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Move existing annotations/images into trash, then copy a random sample "
            "of JPG/JSON pairs from the invoices dataset back into place."
        )
    )
    parser.add_argument(
        "--annotations-dir",
        type=Path,
        default=DEFAULT_ANNOTATIONS,
        help="Destination folder for JSON annotations.",
    )
    parser.add_argument(
        "--images-dir",
        type=Path,
        default=DEFAULT_IMAGES,
        help="Destination folder for JPG images.",
    )
    parser.add_argument(
        "--trash-dir",
        type=Path,
        default=DEFAULT_TRASH,
        help="Folder where existing files will be moved.",
    )
    parser.add_argument(
        "--source-annotations-dir",
        type=Path,
        default=DEFAULT_SOURCE_ANNOTATIONS,
        help="Folder that holds the source JSON annotations.",
    )
    parser.add_argument(
        "--source-images-dir",
        type=Path,
        default=DEFAULT_SOURCE_IMAGES,
        help="Folder that holds the source JPG images.",
    )
    parser.add_argument(
        "--pairs",
        type=int,
        default=5,
        help="Number of JPG/JSON pairs to copy over (default: 50).",
    )
    parser.add_argument(
        "--seed",
        type=int,
        help="Optional random seed for reproducible sampling.",
    )
    parser.add_argument(
        "--converter-script",
        type=Path,
        default=DEFAULT_CONVERTER_SCRIPT,
        help="Path to the JSON-to-CSV converter script.",
    )
    return parser.parse_args()


def ensure_directory(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def unique_destination(dest_dir: Path, filename: str) -> Path:
    target = dest_dir / filename
    if not target.exists():
        return target

    stem = Path(filename).stem
    suffix = Path(filename).suffix
    counter = 1
    while True:
        candidate = dest_dir / f"{stem}_{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def move_matching_files(src_dir: Path, suffix: str, dest_dir: Path) -> int:
    suffix = suffix.lower()
    moved = 0
    for path in sorted(src_dir.iterdir()):
        if path.is_file() and path.suffix.lower() == suffix:
            target = unique_destination(dest_dir, path.name)
            shutil.move(str(path), target)
            moved += 1
    return moved


def iter_files_with_suffix(root: Path, suffixes: Iterable[str]) -> Iterable[Path]:
    normalized = {suffix.lower() for suffix in suffixes}
    for path in root.rglob("*"):
        if path.is_file() and path.suffix.lower() in normalized:
            yield path


def require_directory(path: Path, description: str) -> Path:
    if not path.is_dir():
        raise FileNotFoundError(f"{description} does not exist: {path}")
    return path


def build_pair_candidates(
    source_annotations: Path, source_images: Path
) -> List[JsonImagePair]:
    def is_template1(name: str) -> bool:
        return name.startswith(TEMPLATE1_PREFIX)

    json_files: Dict[str, Path] = {}
    for path in iter_files_with_suffix(source_annotations, [".json"]):
        json_files.setdefault(path.stem, path)

    image_files: Dict[str, Path] = {}
    for path in iter_files_with_suffix(source_images, [".jpg", ".jpeg"]):
        image_files.setdefault(path.stem, path)

    shared_keys = sorted(
        key for key in json_files.keys() & image_files.keys() if is_template1(key)
    )
    return [(json_files[key], image_files[key]) for key in shared_keys]


def copy_pairs(
    pairs: Sequence[JsonImagePair],
    annotations_dir: Path,
    images_dir: Path,
) -> None:
    for json_path, image_path in pairs:
        shutil.copy2(json_path, annotations_dir / json_path.name)
        shutil.copy2(image_path, images_dir / image_path.name)


def run_converter(script_path: Path, annotations_dir: Path) -> None:
    if not script_path.is_file():
        raise FileNotFoundError(f"Converter script not found: {script_path}")
    command = [sys.executable, str(script_path), str(annotations_dir)]
    print(f"Running converter: {' '.join(command)}")
    subprocess.run(command, check=True)


def main() -> None:
    args = parse_args()

    if args.pairs <= 0:
        raise ValueError("The number of pairs to copy must be greater than zero.")

    if args.seed is not None:
        random.seed(args.seed)

    annotations_dir = ensure_directory(args.annotations_dir)
    images_dir = ensure_directory(args.images_dir)
    trash_dir = ensure_directory(args.trash_dir)
    source_annotations = require_directory(
        args.source_annotations_dir, "Source annotations folder"
    )
    source_images = require_directory(args.source_images_dir, "Source images folder")
    converter_script = args.converter_script

    moved_json = move_matching_files(annotations_dir, ".json", trash_dir)
    moved_jpg = move_matching_files(images_dir, ".jpg", trash_dir)
    print(f"Moved {moved_json} JSON files and {moved_jpg} JPG files to {trash_dir}.")

    candidates = build_pair_candidates(source_annotations, source_images)
    if len(candidates) < args.pairs:
        raise RuntimeError(
            "Requested {req} pairs, but only found {found} in the provided sources "
            "(annotations: {ann}, images: {img}).".format(
                req=args.pairs,
                found=len(candidates),
                ann=source_annotations,
                img=source_images,
            )
        )

    selected_pairs = random.sample(candidates, args.pairs)
    copy_pairs(selected_pairs, annotations_dir, images_dir)
    print(f"Copied {len(selected_pairs)} random pairs into {annotations_dir} and {images_dir}.")

    run_converter(converter_script, annotations_dir)


if __name__ == "__main__":
    main()
