"""
Test Script: Image Border Detection
====================================
This script tests if images have built-in borders using the _image_has_border method.

Usage:
    python test_border_detection.py <image_folder>

Example:
    python test_border_detection.py "C:\Path\To\Images"
"""

import os
import sys
from pathlib import Path

def test_border_detection(image_folder):
    print("=" * 70)
    print("Testing Image Border Detection")
    print("=" * 70)

    # Test 1: Check if module can be imported
    try:
        from karyotype_template import KaryotypeReportGenerator
        print("✓ Successfully imported KaryotypeReportGenerator")
    except ImportError as e:
        print(f"✗ Failed to import: {e}")
        return False

    # Test 2: Check if _image_has_border method exists
    if hasattr(KaryotypeReportGenerator, '_image_has_border'):
        print("✓ Method _image_has_border() exists (correction applied)")
    else:
        print("✗ Method _image_has_border() NOT FOUND")
        print("  → This means the Windows installation has OLD code")
        print("  → Border detection will not work")
        return False

    # Test 3: Check if PIL/Pillow is installed
    try:
        from PIL import Image
        print("✓ PIL/Pillow is installed")
    except ImportError:
        print("✗ PIL/Pillow NOT installed")
        print("  → Run: pip install Pillow")
        return False

    # Test 4: Find image files
    if not os.path.isdir(image_folder):
        print(f"✗ Image folder not found: {image_folder}")
        return False

    print(f"\nScanning folder: {image_folder}")

    image_extensions = ['.jpg', '.jpeg', '.png', '.JPG', '.JPEG', '.PNG']
    image_files = []

    for ext in image_extensions:
        image_files.extend(Path(image_folder).glob(f"*{ext}"))

    if not image_files:
        print(f"✗ No image files found in {image_folder}")
        print(f"  Searched for: {', '.join(image_extensions)}")
        return False

    print(f"✓ Found {len(image_files)} image(s)\n")

    # Test 5: Check border detection for each image
    print("-" * 70)
    print(f"{'Image File':<40} {'Has Border':<15} {'Action'}")
    print("-" * 70)

    for img_path in sorted(image_files)[:10]:  # Test first 10 images
        try:
            has_border = KaryotypeReportGenerator._image_has_border(str(img_path))
            action = "No border added" if has_border else "Will add border"
            status = "Yes (built-in)" if has_border else "No"

            print(f"{img_path.name:<40} {status:<15} {action}")

        except Exception as e:
            print(f"{img_path.name:<40} ERROR: {str(e)[:30]}")

    print("-" * 70)
    print("\n✓ Border detection test completed")
    print("\nInterpretation:")
    print("  • 'Yes (built-in)' = Image has dark frame, no extra border will be added")
    print("  • 'No' = Image has no frame, a 1pt black border will be added")

    return True


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python test_border_detection.py <image_folder>")
        print("\nExample:")
        print('  python test_border_detection.py "C:\\Path\\To\\Images"')
        print('  python test_border_detection.py /home/user/images')
        sys.exit(1)

    image_folder = sys.argv[1]
    success = test_border_detection(image_folder)
    sys.exit(0 if success else 1)
