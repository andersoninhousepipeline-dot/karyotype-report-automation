"""
Test Script: Generate Variant Report (3-page layout)
====================================================
This script tests if the karyotype_template.py has the latest corrections.

Expected Results:
- 3-page PDF (because COMMENTS field has content)
- ISCN text color: BLACK (because AUTOSOME is "Variant Observed")
- Comments section visible on page 2
"""

import os
import sys

def test_variant_report():
    print("=" * 70)
    print("Testing Variant Report Generation")
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
        return False

    # Test 3: Generate a variant report
    print("\nGenerating test report...")

    data = {
        "NAME": "Test Patient Variant",
        "GENDER": "Female",
        "AGE": "28",
        "SPECIMEN": "Peripheral blood",
        "PIN": "TEST003",
        "SAMPLE NUMBER": "TEST345678",
        "SAMPLE COLLECTION DATE": "15-03-2025",
        "SAMPLE RECEIPT DATE": "16-03-2025",
        "REPORT DATE": "20-03-2025",
        "REFERRING CLINICIAN": "Dr. Test Clinician",
        "HOSPITAL/CLINIC": "Test Hospital",
        "TEST INDICATION": "To rule out gross chromosomal abnormality",
        "RESULT": "46,XX,del(5)(q13q33)",
        "METAPHASE ANALYSED": "25",
        "ESTIMATED BAND RESOLUTION": "475",
        "AUTOSOME": "Variant Observed",
        "SEX CHROMOSOME": "Normal",
        "INTERPRETATION": "Karyotype shows a female with a deletion in chromosome 5.",
        "COMMENTS": "Chromosomal variants may or may not have clinical significance. Clinical correlation is advised.",
        "RECOMMENDATIONS": "Genetic counseling advised."
    }

    output_dir = "test_output"
    os.makedirs(output_dir, exist_ok=True)

    try:
        gen = KaryotypeReportGenerator(data, [], output_dir)
        pdf_path = gen.generate()

        print(f"✓ PDF generated successfully: {pdf_path}")

        # Test 4: Check layout type
        if gen.three_page:
            print("✓ Layout: 3-page (correct for variant report with comments)")
        else:
            print("✗ Layout: 2-page (WRONG! Expected 3-page)")
            return False

        # Test 5: Check ISCN color
        color = gen._iscn_color()
        print(f"\nISCN Color Analysis:")
        print(f"  Autosome: {data['AUTOSOME']}")
        print(f"  Sex Chromosome: {data['SEX CHROMOSOME']}")
        print(f"  Computed Color: {color}")

        # Check if it's BLACK
        from reportlab.lib.colors import black
        BLACK = black

        if color == BLACK:
            print("✓ ISCN color is BLACK (correct for Variant)")
        else:
            print(f"✗ ISCN color is NOT BLACK (got {color})")
            print("  → Color correction not working!")
            return False

        print("\n" + "=" * 70)
        print("✓ ALL TESTS PASSED")
        print("=" * 70)
        print(f"\nGenerated PDF: {os.path.abspath(pdf_path)}")
        print("\nPlease open the PDF and verify:")
        print("  1. ISCN text is BLACK")
        print("  2. 3 pages total")
        print("  3. 'Comments' section visible on page 2")

        return True

    except Exception as e:
        import traceback
        print(f"✗ Error during generation: {e}")
        print(traceback.format_exc())
        return False


if __name__ == "__main__":
    success = test_variant_report()
    sys.exit(0 if success else 1)
