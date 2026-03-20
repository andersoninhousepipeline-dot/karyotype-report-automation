"""
Test Script: Generate Abnormal Report (3-page layout)
======================================================
This script tests if the karyotype_template.py has the latest corrections.

Expected Results:
- 3-page PDF (because COMMENTS field has content)
- ISCN text color: RED (because AUTOSOME is Abnormal)
- Comments section visible on page 2
- Proper spacing/alignment
"""

import os
import sys

def test_abnormal_report():
    print("=" * 70)
    print("Testing Abnormal Report Generation (Trisomy 21)")
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

    # Test 3: Generate an abnormal report (Trisomy 21)
    print("\nGenerating test report...")

    data = {
        "NAME": "Test Patient Trisomy",
        "GENDER": "Male",
        "AGE": "3",
        "SPECIMEN": "Peripheral blood",
        "PIN": "TEST002",
        "SAMPLE NUMBER": "TEST789012",
        "SAMPLE COLLECTION DATE": "15-03-2025",
        "SAMPLE RECEIPT DATE": "16-03-2025",
        "REPORT DATE": "20-03-2025",
        "REFERRING CLINICIAN": "Dr. Test Clinician",
        "HOSPITAL/CLINIC": "Test Hospital",
        "TEST INDICATION": "To rule out gross chromosomal abnormality",
        "RESULT": "47,XY,+21",
        "METAPHASE ANALYSED": "25",
        "ESTIMATED BAND RESOLUTION": "475",
        "AUTOSOME": "Abnormal",
        "SEX CHROMOSOME": "Normal",
        "INTERPRETATION": "The constitutional karyotype shows a male with three copies of chromosome 21, indicating Down (Trisomy 21) syndrome.",
        "COMMENTS": "Trisomy 21 is a genetic syndrome associated with impairment of cognitive ability and physical growth as well as a particular set of facial characteristics.",
        "RECOMMENDATIONS": "Advised genetic counseling for the parents."
    }

    output_dir = "test_output"
    os.makedirs(output_dir, exist_ok=True)

    try:
        gen = KaryotypeReportGenerator(data, [], output_dir)
        pdf_path = gen.generate()

        print(f"✓ PDF generated successfully: {pdf_path}")

        # Test 4: Check layout type
        if gen.three_page:
            print("✓ Layout: 3-page (correct for abnormal report with comments)")
        else:
            print("✗ Layout: 2-page (WRONG! Expected 3-page)")
            print("  → Comments field has content, should be 3-page layout")
            return False

        # Test 5: Check ISCN color
        color = gen._iscn_color()
        print(f"\nISCN Color Analysis:")
        print(f"  Autosome: {data['AUTOSOME']}")
        print(f"  Sex Chromosome: {data['SEX CHROMOSOME']}")
        print(f"  Computed Color: {color}")

        # Check if it's RED
        from reportlab.lib.colors import Color
        RED = Color(1, 0, 0)

        if color == RED:
            print("✓ ISCN color is RED (correct for Abnormal)")
        else:
            print(f"✗ ISCN color is NOT RED (got {color})")
            print("  → Color correction not working!")
            return False

        print("\n" + "=" * 70)
        print("✓ ALL TESTS PASSED")
        print("=" * 70)
        print(f"\nGenerated PDF: {os.path.abspath(pdf_path)}")
        print("\nPlease open the PDF and verify:")
        print("  1. ISCN text is RED")
        print("  2. 3 pages total")
        print("  3. 'Comments' section visible on page 2")
        print("  4. Alignment and spacing look correct")

        return True

    except Exception as e:
        import traceback
        print(f"✗ Error during generation: {e}")
        print(traceback.format_exc())
        return False


if __name__ == "__main__":
    success = test_abnormal_report()
    sys.exit(0 if success else 1)
