#!/usr/bin/env python3
"""
Monkey-Patch Fix for python-docx and python-pptx Int() Conversion Bug

PROBLEM:
Both python-docx and python-pptx call int() directly on XML attribute values,
which fails when Office files contain decimal font sizes/measurements.

ERROR:
ValueError: invalid literal for int() with base 10: '19.5'

ROOT CAUSE:
- OOXML spec expects integer values for font sizes (in half-points)
- Some Office files (especially from non-MS tools) contain decimals
- Libraries call int(str_value) without defensive float() conversion first

AFFECTED FILES:
- python-docx: docx/oxml/simpletypes.py line 69, 318
- python-pptx: pptx/oxml/simpletypes.py line 94, 232, 305, 339, 448, 458, 513, 641, 671, 699

SOLUTION:
Monkey-patch BaseIntType.convert_from_xml() to handle decimals via float() first.

USAGE:
    from fix_ooxml_int_conversion import apply_ooxml_patches
    apply_ooxml_patches()  # Call BEFORE importing Document/Presentation

    from docx import Document
    from pptx import Presentation
"""


def safe_int_from_xml(str_value):
    """
    Safely convert XML string to int, handling decimal values.

    Args:
        str_value: String value from XML attribute

    Returns:
        Integer value, with decimals rounded

    Examples:
        safe_int_from_xml("24") -> 24
        safe_int_from_xml("19.5") -> 20
        safe_int_from_xml("22.079999923706055") -> 22
    """
    try:
        # Try direct int conversion first (fast path for valid values)
        return int(str_value)
    except ValueError:
        # Handle decimal values by converting via float first
        try:
            return int(round(float(str_value)))
        except (ValueError, TypeError):
            # If all else fails, return safe default (preserves document structure)
            # Better to have font size 0 than to crash
            return 0


def patch_python_docx():
    """
    Patch python-docx to handle decimal font sizes and measurements.

    Patches:
        - BaseIntType.convert_from_xml() for all integer attribute parsing
    """
    try:
        from docx.oxml import simpletypes

        # Save original method for potential debugging
        if not hasattr(simpletypes.BaseIntType, '_original_convert_from_xml'):
            simpletypes.BaseIntType._original_convert_from_xml = simpletypes.BaseIntType.convert_from_xml

            # Create patched version
            @classmethod
            def safe_convert_from_xml(cls, str_value):
                """Patched version that handles decimal values."""
                return safe_int_from_xml(str_value)

            # Apply patch
            simpletypes.BaseIntType.convert_from_xml = safe_convert_from_xml

            # Verify patch works
            try:
                test_result = simpletypes.BaseIntType.convert_from_xml('19.5')
                if test_result != 20:
                    import logging
                    logging.error(f"PATCH VERIFICATION FAILED: Expected 20, got {test_result}")
                    return False
            except Exception as e:
                import logging
                logging.error(f"PATCH VERIFICATION FAILED: {e}")
                return False

        return True
    except ImportError:
        # python-docx not installed - skip patch
        return False


def patch_python_pptx():
    """
    Patch python-pptx to handle decimal values in presentations.

    Patches:
        - BaseIntType.convert_from_xml() for all integer attribute parsing
    """
    try:
        from pptx.oxml import simpletypes

        # Save original method for potential debugging
        if not hasattr(simpletypes.BaseIntType, '_original_convert_from_xml'):
            simpletypes.BaseIntType._original_convert_from_xml = simpletypes.BaseIntType.convert_from_xml

            # Create patched version
            @classmethod
            def safe_convert_from_xml(cls, str_value):
                """Patched version that handles decimal values."""
                return safe_int_from_xml(str_value)

            # Apply patch
            simpletypes.BaseIntType.convert_from_xml = safe_convert_from_xml

        return True
    except ImportError:
        # python-pptx not installed - skip patch
        return False


def apply_ooxml_patches():
    """
    Apply all OOXML int() conversion patches.

    Call this ONCE at the start of your program, BEFORE importing Document/Presentation.

    Returns:
        Tuple of (docx_patched, pptx_patched) booleans
    """
    import logging
    logger = logging.getLogger(__name__)

    docx_patched = patch_python_docx()
    pptx_patched = patch_python_pptx()

    # Log results to BOTH stdout and logger
    patches_applied = []
    if docx_patched:
        patches_applied.append("python-docx")
    if pptx_patched:
        patches_applied.append("python-pptx")

    if patches_applied:
        msg = f"✓ Applied OOXML int() conversion patches to: {', '.join(patches_applied)}"
        print(msg)
        logger.info(msg)

    return (docx_patched, pptx_patched)


# Auto-apply if run as main (for testing)
if __name__ == "__main__":
    print("Testing OOXML int() conversion patches...")
    print()

    # Test safe_int_from_xml
    test_cases = [
        ("24", 24),
        ("19.5", 20),
        ("22.079999923706055", 22),
        ("0", 0),
        ("100.4", 100),
        ("100.6", 101),
    ]

    print("Testing safe_int_from_xml():")
    for input_val, expected in test_cases:
        result = safe_int_from_xml(input_val)
        status = "✓" if result == expected else "✗"
        print(f"  {status} safe_int_from_xml('{input_val}') = {result} (expected {expected})")

    print()

    # Apply patches
    docx_ok, pptx_ok = apply_ooxml_patches()

    print()
    print(f"python-docx patch: {'✓ Applied' if docx_ok else '✗ Not available'}")
    print(f"python-pptx patch: {'✓ Applied' if pptx_ok else '✗ Not available'}")
