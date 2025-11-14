#!/usr/bin/env python3
"""
PowerPoint Anonymization Module
Handles .pptx and .ppt files for the DOCX Anonymizer app
"""

from pptx import Presentation
from pathlib import Path
import logging
import re


def strip_pptx_metadata(prs):
    """
    Strip ALL metadata from PowerPoint file.

    Similar to Word metadata stripping - removes:
    - Author
    - Title
    - Subject
    - Keywords
    - Comments
    - Company
    """
    props = prs.core_properties

    props.author = ""
    props.last_modified_by = ""
    props.title = ""
    props.subject = ""
    props.keywords = ""
    props.comments = ""
    props.category = ""
    props.content_status = ""
    props.identifier = ""

    # Clear company (important for SEC filings)
    if hasattr(props, 'company'):
        props.company = ""

    props.revision = 1
    if hasattr(props, 'version'):
        props.version = None

    return prs


def remove_all_images_pptx(prs):
    """
    Remove ALL images from PowerPoint slides.

    Returns count of removed images.
    """
    removed_count = 0

    for slide in prs.slides:
        # Find all picture shapes
        shapes_to_remove = []
        for shape in slide.shapes:
            # Check if shape is a picture
            if shape.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
                shapes_to_remove.append(shape)

        # Remove pictures (must be done after iteration)
        for shape in shapes_to_remove:
            sp = shape.element
            sp.getparent().remove(sp)
            removed_count += 1

    return removed_count


def anonymize_text_pptx(text, alias_map, sorted_keys, compiled_patterns):
    """
    Apply anonymization replacements with case matching.

    Reuses same logic as Word processor.
    """
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


def anonymize_pptx(pptx_path, alias_map, sorted_keys, compiled_patterns):
    """
    Anonymize all text in PowerPoint file.

    Processes:
    - Slide text frames
    - Tables
    - Notes
    - Shapes
    """
    prs = Presentation(pptx_path)
    total_replacements = 0

    for slide in prs.slides:
        # Process all shapes with text
        for shape in slide.shapes:
            # Text frames (title, text boxes, etc.)
            if hasattr(shape, 'text_frame'):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.text:
                            new_text, count = anonymize_text_pptx(
                                run.text, alias_map, sorted_keys, compiled_patterns
                            )
                            if count > 0:
                                run.text = new_text
                                total_replacements += count

            # Tables
            if shape.shape_type == 19:  # MSO_SHAPE_TYPE.TABLE
                table = shape.table
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if run.text:
                                    new_text, count = anonymize_text_pptx(
                                        run.text, alias_map, sorted_keys, compiled_patterns
                                    )
                                    if count > 0:
                                        run.text = new_text
                                        total_replacements += count

        # Process notes (speaker notes)
        if slide.has_notes_slide:
            notes_slide = slide.notes_slide
            if hasattr(notes_slide, 'notes_text_frame'):
                for paragraph in notes_slide.notes_text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.text:
                            new_text, count = anonymize_text_pptx(
                                run.text, alias_map, sorted_keys, compiled_patterns
                            )
                            if count > 0:
                                run.text = new_text
                                total_replacements += count

    return prs, total_replacements


def process_single_pptx(input_path, output_path, alias_map, sorted_keys, compiled_patterns, logger, remove_images=True):
    """
    Process a single PowerPoint file: anonymize + strip metadata + optional image removal.

    Args:
        input_path: Path to input .pptx file
        output_path: Path for output .pptx file
        alias_map: Dictionary of original → replacement mappings
        sorted_keys: Sorted list of alias_map keys
        compiled_patterns: Pre-compiled regex patterns
        logger: Logger instance
        remove_images: If True, removes all images from presentation

    Returns:
        (replacements, images_removed) tuple
    """
    logger.info(f"Processing: {input_path.name}")

    try:
        # Load and anonymize PowerPoint
        prs, replacements = anonymize_pptx(input_path, alias_map, sorted_keys, compiled_patterns)

        # Remove all images (if requested)
        images_removed = 0
        if remove_images:
            images_removed = remove_all_images_pptx(prs)

        # Strip ALL metadata (CRITICAL)
        prs = strip_pptx_metadata(prs)

        # Save
        output_path.parent.mkdir(parents=True, exist_ok=True)
        prs.save(output_path)

        logger.info(f"  ✓ {replacements} replacements, {images_removed} images removed")

        return replacements, images_removed

    except Exception as e:
        logger.error(f"  ❌ Error: {e}")
        return 0, 0
