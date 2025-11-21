"""
Document processors for DOCX, Excel, and PowerPoint files
"""

from .docx_processor import (
    process_single_docx,
    load_aliases_from_excel,
    categorize_and_sort_aliases,
    precompile_patterns
)
from .excel_processor import process_single_xlsx, process_single_xls
from .pptx_processor import process_single_pptx

__all__ = [
    'process_single_docx',
    'load_aliases_from_excel',
    'categorize_and_sort_aliases',
    'precompile_patterns',
    'process_single_xlsx',
    'process_single_xls',
    'process_single_pptx'
]
