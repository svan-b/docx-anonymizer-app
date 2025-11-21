#!/usr/bin/env python3
"""
Hyperlink Metadata Removal Utilities
=====================================
Removes hyperlink metadata (clickable links) while preserving text.

IMPORTANT: This should be called AFTER anonymization, so that:
1. Text in hyperlinks is anonymized first
2. Then hyperlink metadata is removed
3. Result: Anonymized text without clickable links

Usage:
    from hyperlink_utils import remove_hyperlinks_docx, remove_hyperlinks_xlsx, remove_hyperlinks_pptx

    # After anonymization
    hyperlinks_removed = remove_hyperlinks_docx(doc)
"""


def remove_hyperlinks_docx(doc):
    """
    Remove all hyperlink metadata from Word document while preserving display text.

    This should be called AFTER anonymization.

    Process:
    1. Removes hyperlink relationships from document.xml.rels
    2. Removes <w:hyperlink> XML elements but preserves text runs
    3. Processes main body, tables (CRITICAL for 10-K), headers, and footers

    Args:
        doc: python-docx Document object (already anonymized)

    Returns:
        int: Count of hyperlinks removed
    """
    removed_count = 0

    # Step 1: Remove hyperlink relationships
    if hasattr(doc, 'part') and hasattr(doc.part, 'rels'):
        rels_to_remove = []

        for rel_id, rel in doc.part.rels.items():
            if hasattr(rel, 'reltype') and 'hyperlink' in rel.reltype.lower():
                rels_to_remove.append(rel_id)
                removed_count += 1

        # Remove the relationships
        for rel_id in rels_to_remove:
            if rel_id in doc.part.rels:
                del doc.part.rels[rel_id]

    # Step 2: Remove hyperlink elements from paragraphs (preserves text)
    def remove_hyperlink_elements(paragraphs):
        """Helper to remove hyperlink elements from a list of paragraphs"""
        for paragraph in paragraphs:
            p_elem = paragraph._element
            hyperlink_elems = p_elem.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hyperlink')

            for hl_elem in hyperlink_elems:
                parent = hl_elem.getparent()
                if parent is not None:
                    # Get index of hyperlink element
                    index = list(parent).index(hl_elem)

                    # Move all runs out of hyperlink element
                    for run_elem in list(hl_elem):
                        parent.insert(index, run_elem)
                        index += 1

                    # Remove the now-empty hyperlink element
                    parent.remove(hl_elem)

    # Process main document body
    remove_hyperlink_elements(doc.paragraphs)

    # Process tables (CRITICAL for 10-K documents)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                remove_hyperlink_elements(cell.paragraphs)

    # Process headers and footers
    for section in doc.sections:
        # Headers
        remove_hyperlink_elements([section.header] if hasattr(section.header, 'paragraphs') else [])
        if hasattr(section, 'first_page_header'):
            remove_hyperlink_elements([section.first_page_header] if hasattr(section.first_page_header, 'paragraphs') else [])
        if hasattr(section, 'even_page_header'):
            remove_hyperlink_elements([section.even_page_header] if hasattr(section.even_page_header, 'paragraphs') else [])

        # Footers
        remove_hyperlink_elements([section.footer] if hasattr(section.footer, 'paragraphs') else [])
        if hasattr(section, 'first_page_footer'):
            remove_hyperlink_elements([section.first_page_footer] if hasattr(section.first_page_footer, 'paragraphs') else [])
        if hasattr(section, 'even_page_footer'):
            remove_hyperlink_elements([section.even_page_footer] if hasattr(section.even_page_footer, 'paragraphs') else [])

    return removed_count


def remove_hyperlinks_xlsx(wb):
    """
    Remove all hyperlink metadata from Excel workbook while preserving cell values.

    This should be called AFTER anonymization.

    Process:
    1. Iterates through all sheets
    2. Sets cell.hyperlink = None (removes link, keeps value)

    Args:
        wb: openpyxl Workbook object (already anonymized)

    Returns:
        int: Count of hyperlinks removed
    """
    removed_count = 0

    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]

        for row in sheet.iter_rows():
            for cell in row:
                if cell.hyperlink:
                    # Remove hyperlink but keep the cell value
                    cell.hyperlink = None
                    removed_count += 1

    return removed_count


def remove_hyperlinks_pptx(prs):
    """
    Remove all hyperlink metadata from PowerPoint presentation while preserving text.

    This should be called AFTER anonymization.

    Process:
    1. Removes hyperlinks from text runs
    2. Removes hyperlinks from shape click actions

    Args:
        prs: python-pptx Presentation object (already anonymized)

    Returns:
        int: Count of hyperlinks removed
    """
    removed_count = 0

    for slide in prs.slides:
        for shape in slide.shapes:
            # Remove hyperlinks from text runs
            if hasattr(shape, 'text_frame'):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if hasattr(run, 'hyperlink') and run.hyperlink.address:
                            run.hyperlink.address = None
                            removed_count += 1

            # Remove hyperlinks from shape click actions
            if hasattr(shape, 'click_action') and hasattr(shape.click_action, 'hyperlink'):
                if shape.click_action.hyperlink and hasattr(shape.click_action.hyperlink, 'address'):
                    if shape.click_action.hyperlink.address:
                        shape.click_action.hyperlink.address = None
                        removed_count += 1

    return removed_count
