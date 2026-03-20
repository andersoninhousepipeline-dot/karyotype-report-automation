#!/usr/bin/env python3
"""
Image Discovery Diagnostic Tool
================================
Tests the image auto-discovery function on Windows/Linux

Usage:
    python test_image_discovery.py <image_folder> <sample_number>

Example:
    python test_image_discovery.py "C:\\Images" 260154818
    python test_image_discovery.py /home/user/images 260154818
"""

import sys
import os
import re
from pathlib import Path


def find_images_for_sample(sample_no: str, search_dir: str) -> list:
    """
    Find images matching sample number in search_dir.
    Handles Windows/Linux path differences and case-insensitive matching.

    Matches:
      - Exact: 260154818.jpg
      - Numbered: 260161295 1.jpg, 260161295 2.jpg
    """
    if not sample_no or not search_dir or not os.path.isdir(search_dir):
        return []

    sample_no = str(sample_no).strip()
    if not sample_no:
        return []

    results = []

    # Use Path for better cross-platform compatibility
    search_path = Path(search_dir)

    # Search for common image extensions (case-insensitive on Windows, explicit on Linux)
    extensions = ['.jpg', '.jpeg', '.png', '.JPG', '.JPEG', '.PNG']

    try:
        # List all files in directory
        for file_path in search_path.iterdir():
            if not file_path.is_file():
                continue

            # Check if extension matches (case-insensitive comparison)
            if file_path.suffix.lower() not in [ext.lower() for ext in extensions]:
                continue

            # Extract filename without extension
            fname = file_path.stem

            # Match patterns:
            # 1. Exact match: filename == sample_no
            # 2. Numbered: sample_no followed by space and digit(s)
            if fname == sample_no or re.match(rf"^{re.escape(sample_no)}\s+\d+$", fname):
                results.append(str(file_path.absolute()))

        # Sort results for consistent ordering
        results.sort(key=lambda x: Path(x).name)

    except (OSError, PermissionError) as e:
        # Handle permission errors or invalid paths gracefully
        print(f"Error: Could not access directory {search_dir}: {e}")
        return []

    return results


def main():
    print("=" * 70)
    print("Image Discovery Diagnostic Tool")
    print("=" * 70)
    print()

    if len(sys.argv) < 3:
        print("Usage: python test_image_discovery.py <image_folder> <sample_number>")
        print()
        print("Examples:")
        print('  Windows: python test_image_discovery.py "C:\\Images" 260154818')
        print('  Linux:   python test_image_discovery.py /home/user/images 260154818')
        sys.exit(1)

    search_dir = sys.argv[1]
    sample_no = sys.argv[2]

    print(f"Platform:      {sys.platform}")
    print(f"Search Dir:    {search_dir}")
    print(f"Sample Number: {sample_no}")
    print()

    # Normalize path
    search_dir = os.path.normpath(search_dir)
    print(f"Normalized:    {search_dir}")
    print()

    # Check if directory exists
    if not os.path.exists(search_dir):
        print(f"ERROR: Directory does not exist: {search_dir}")
        sys.exit(1)

    if not os.path.isdir(search_dir):
        print(f"ERROR: Path is not a directory: {search_dir}")
        sys.exit(1)

    print("Directory exists and is accessible.")
    print()

    # List all files in directory
    print("All files in directory:")
    print("-" * 70)
    try:
        search_path = Path(search_dir)
        all_files = []
        for item in search_path.iterdir():
            if item.is_file():
                all_files.append(item.name)

        if not all_files:
            print("  (No files found)")
        else:
            for fname in sorted(all_files):
                file_path = search_path / fname
                print(f"  {fname}")
                print(f"    - Stem: '{file_path.stem}'")
                print(f"    - Suffix: '{file_path.suffix}'")
                print(f"    - Suffix lower: '{file_path.suffix.lower()}'")

    except Exception as e:
        print(f"  ERROR: {e}")

    print()
    print("=" * 70)
    print("Running Image Discovery...")
    print("=" * 70)

    # Run the image discovery
    results = find_images_for_sample(sample_no, search_dir)

    if results:
        print(f"✓ Found {len(results)} matching image(s):")
        print()
        for img_path in results:
            print(f"  • {Path(img_path).name}")
            print(f"    Full path: {img_path}")
        print()
        print("SUCCESS: Image discovery is working!")
    else:
        print("✗ No images found matching the sample number.")
        print()
        print("Troubleshooting:")
        print("  1. Check that image files exist with filenames like:")
        print(f"     - {sample_no}.jpg")
        print(f"     - {sample_no} 1.jpg")
        print(f"     - {sample_no} 2.jpg")
        print("  2. Check file extensions (.jpg, .jpeg, .png)")
        print("  3. Check that filenames match exactly (no extra spaces/characters)")
        print("  4. Try listing files manually in the directory")

    print()


if __name__ == "__main__":
    main()
