"""
Utility functions for anonymization
"""

from .anonymizer_utils import anonymize_text, merge_details
from .hyperlink_utils import remove_hyperlinks_docx
from .fix_ooxml_int_conversion import apply_ooxml_patches

__all__ = ['anonymize_text', 'merge_details', 'remove_hyperlinks_docx', 'apply_ooxml_patches']
