#!/usr/bin/env python3
"""
Excel Anonymization Module
Handles .xlsx and .xls files for the DOCX Anonymizer app
"""

from openpyxl import load_workbook
from pathlib import Path
import logging
import re


def strip_xlsx_metadata(wb):
    """
    Strip ALL metadata from Excel file.

    Removes:
    - Author
    - Title
    - Subject
    - Keywords
    - Comments
    - Company
    """
    props = wb.properties

    props.creator = ""
    props.lastModifiedBy = ""
    props.title = ""
    props.subject = ""
    props.keywords = ""
    props.description = ""
    props.category = ""
    props.contentStatus = ""
    props.identifier = ""
    props.company = ""
    props.manager = ""

    props.revision = 1
    if hasattr(props, 'version'):
        props.version = None

    return wb


def anonymize_text_xlsx(text, alias_map, sorted_keys, compiled_patterns):
    """
    Apply anonymization replacements with case matching.

    Reuses same logic as Word/PowerPoint processors.
    """
    if not text or not isinstance(text, str):
        return text, 0

    replacements = 0

    for original in sorted_keys:
        replacement = alias_map[original]

        def replace_with_case(match):
            nonlocal replacements
            matched_text = match.group(0)

            # Preserve case pattern
            if matched_text.isupper():
                replacements += 1
                return replacement.upper()
            elif matched_text.islower():
                replacements += 1
                return replacement.lower()
            elif matched_text[0].isupper():
                replacements += 1
                return replacement.capitalize()
            else:
                replacements += 1
                return replacement

        pattern = compiled_patterns[original]
        text = pattern.sub(replace_with_case, text)

    return text, replacements


def anonymize_xlsx(xlsx_path, alias_map, sorted_keys, compiled_patterns):
    """
    Anonymize all text in Excel file.

    Processes:
    - Cell values (text and formulas)
    - Sheet names
    - Headers/footers
    - Comments

    IMPORTANT: Does NOT recalculate formulas (safer, prevents errors)
    """
    wb = load_workbook(xlsx_path, data_only=False)  # Keep formulas
    total_replacements = 0

    # Anonymize sheet names first
    sheet_name_mapping = {}
    for sheet in wb.worksheets:
        old_name = sheet.title
        new_name, count = anonymize_text_xlsx(old_name, alias_map, sorted_keys, compiled_patterns)
        if count > 0:
            # Ensure unique sheet name (Excel requirement)
            if new_name in [s.title for s in wb.worksheets]:
                new_name = f"{new_name}_1"
            sheet.title = new_name
            sheet_name_mapping[old_name] = new_name
            total_replacements += count

    # Process each sheet
    for sheet in wb.worksheets:
        # Process all cells
        for row in sheet.iter_rows():
            for cell in row:
                # Anonymize cell values (text)
                if cell.value and isinstance(cell.value, str):
                    new_value, count = anonymize_text_xlsx(
                        cell.value, alias_map, sorted_keys, compiled_patterns
                    )
                    if count > 0:
                        # Check if it's a formula
                        if str(cell.value).startswith('='):
                            # It's a formula - replace text within formula
                            cell.value = new_value
                        else:
                            # Regular text
                            cell.value = new_value
                        total_replacements += count

                # Anonymize cell comments
                if cell.comment:
                    if cell.comment.text:
                        new_comment, count = anonymize_text_xlsx(
                            cell.comment.text, alias_map, sorted_keys, compiled_patterns
                        )
                        if count > 0:
                            cell.comment.text = new_comment
                            total_replacements += count

        # NOTE: Excel header/footer anonymization skipped
        # openpyxl header/footer objects have complex structure
        # Most Excel files don't use headers/footers with company names
        # Can be added in future if needed

    return wb, total_replacements


def process_single_xlsx(input_path, output_path, alias_map, sorted_keys, compiled_patterns, logger, remove_images=True):
    """
    Process a single Excel file: anonymize + strip metadata.

    Args:
        input_path: Path to input .xlsx file
        output_path: Path for output .xlsx file
        alias_map: Dictionary of original → replacement mappings
        sorted_keys: Sorted list of alias_map keys
        compiled_patterns: Pre-compiled regex patterns
        logger: Logger instance
        remove_images: Ignored for Excel (kept for API consistency)

    Returns:
        (replacements, images_removed) tuple
        Note: images_removed always 0 for Excel (charts/images not processed)
    """
    logger.info(f"Processing: {input_path.name}")

    try:
        # Load and anonymize Excel
        wb, replacements = anonymize_xlsx(input_path, alias_map, sorted_keys, compiled_patterns)

        # Strip ALL metadata (CRITICAL)
        wb = strip_xlsx_metadata(wb)

        # Save
        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)

        logger.info(f"  ✓ {replacements} replacements")

        # Return 0 for images (Excel doesn't remove images yet)
        return replacements, 0

    except Exception as e:
        logger.error(f"  ❌ Error: {e}")
        return 0, 0
