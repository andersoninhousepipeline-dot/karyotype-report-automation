"""
Test Script: Generate Normal Report (2-page layout)
===================================================
This script tests if the karyotype_template.py has the latest corrections.

Expected Results:
- 2-page PDF (because COMMENTS field is empty)
- ISCN text color: GREEN (#00B050)
- No comments section rendered
- Method _iscn_color() should exist
"""

import os
import sys

def test_normal_report():
    print("=" * 70)
    print("Testing Normal Report Generation")
    print("=" * 70)

    # Test 1: Check if module can be imported
    try:
        from karyotype_template import KaryotypeReportGenerator
        print("✓ Successfully imported KaryotypeReportGenerator")
    except ImportError as e:
        print(f"✗ Failed to import: {e}")
        return False

    # Test 2: Check if _iscn_color method exists
    if hasattr(KaryotypeReportGenerator, '_iscn_color'):
        print("✓ Method _iscn_color() exists (correction applied)")
    else:
        print("✗ Method _iscn_color() NOT FOUND (code not updated!)")
        print("  → This means the Windows installation has old code")
        return False

    # Test 3: Check if _image_has_border method exists
    if hasattr(KaryotypeReportGenerator, '_image_has_border'):
        print("✓ Method _image_has_border() exists (correction applied)")
    else:
        print("✗ Method _image_has_border() NOT FOUND (code not updated!)")
        return False

    # Test 4: Generate a normal male report
    print("\nGenerating test report...")

    data = {
        "NAME": "Test Patient Normal Male",
        "GENDER": "Male",
        "AGE": "25",
        "SPECIMEN": "Peripheral blood",
        "PIN": "TEST001",
        "SAMPLE NUMBER": "TEST123456",
        "SAMPLE COLLECTION DATE": "15-03-2025",
        "SAMPLE RECEIPT DATE": "16-03-2025",
        "REPORT DATE": "20-03-2025",
        "REFERRING CLINICIAN": "Dr. Test Clinician",
        "HOSPITAL/CLINIC": "Test Hospital",
        "TEST INDICATION": "To rule out gross chromosomal abnormality",
        "RESULT": "46,XY",
        "METAPHASE ANALYSED": "25",
        "ESTIMATED BAND RESOLUTION": "475",
        "AUTOSOME": "Normal",
        "SEX CHROMOSOME": "Normal",
        "INTERPRETATION": "Karyotype shows an apparently normal male.",
        "COMMENTS": "",  # EMPTY → should result in 2-page layout
        "RECOMMENDATIONS": "• Genetic counseling is recommended to discuss the implications of the result.\n• Additional genetic testing may be warranted based on the specific phenotypic indication."
    }

    output_dir = "test_output"
    os.makedirs(output_dir, exist_ok=True)

    try:
        gen = KaryotypeReportGenerator(data, [], output_dir)
        pdf_path = gen.generate()

        print(f"✓ PDF generated successfully: {pdf_path}")

        # Test 5: Check layout type
        if gen.three_page:
            print("✗ Layout: 3-page (WRONG! Expected 2-page for normal report)")
            print("  → Comments field is empty, should be 2-page layout")
            return False
        else:
            print("✓ Layout: 2-page (correct for normal report)")

        # Test 6: Check ISCN color
        color = gen._iscn_color()
        print(f"\nISCN Color Analysis:")
        print(f"  Autosome: {data['AUTOSOME']}")
        print(f"  Sex Chromosome: {data['SEX CHROMOSOME']}")
        print(f"  Computed Color: {color}")

        # Check if it's GREEN
        from reportlab.lib.colors import HexColor
        GREEN = HexColor('#00B050')

        if color == GREEN:
            print("✓ ISCN color is GREEN (correct for Normal/Normal)")
        else:
            print(f"✗ ISCN color is NOT GREEN (got {color})")
            print("  → Color correction not working!")
            return False

        print("\n" + "=" * 70)
        print("✓ ALL TESTS PASSED")
        print("=" * 70)
        print(f"\nGenerated PDF: {os.path.abspath(pdf_path)}")
        print("\nPlease open the PDF and verify:")
        print("  1. ISCN text is GREEN")
        print("  2. Only 2 pages (no page 3)")
        print("  3. No 'Comments' section visible")
        print("  4. Alignment looks correct")

        return True

    except Exception as e:
        import traceback
        print(f"✗ Error during generation: {e}")
        print(traceback.format_exc())
        return False


if __name__ == "__main__":
    success = test_normal_report()
    sys.exit(0 if success else 1)
