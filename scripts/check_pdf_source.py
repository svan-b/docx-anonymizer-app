#!/usr/bin/env python3
"""
Check if PDFs were converted with Adobe tools (not third-party converters).

Usage:
    python check_pdf_source.py <folder_path>
    python check_pdf_source.py single_file.pdf

Requirements:
    pip install pypdf
"""
from pathlib import Path
import sys
from pypdf import PdfReader

# Tools approved for PDF conversion
APPROVED_PRODUCERS = [
    'Adobe PDF',
    'Adobe Acrobat',
    'Adobe InDesign',
    'Adobe Illustrator',
    'Adobe Photoshop'
]

# Optional: Add your own tools if needed
CUSTOM_APPROVED = [
    # 'LibreOffice',  # Uncomment if LibreOffice is acceptable
    # 'Microsoft',    # Uncomment if MS Print to PDF is acceptable
]

ALL_APPROVED = APPROVED_PRODUCERS + CUSTOM_APPROVED


def check_pdf_source(pdf_path):
    """
    Check if PDF was created with approved tools.

    Returns:
        dict: {
            'file': filename,
            'producer': PDF producer string,
            'creator': PDF creator string,
            'is_approved': True if Adobe tool,
            'path': full path
        }
    """
    try:
        reader = PdfReader(pdf_path)
        metadata = reader.metadata

        if not metadata:
            return {
                'file': pdf_path.name,
                'producer': 'No metadata found',
                'creator': 'No metadata found',
                'is_approved': False,
                'path': pdf_path
            }

        producer = metadata.get('/Producer', 'Not specified')
        creator = metadata.get('/Creator', 'Not specified')

        # Check if producer or creator contains any approved signature
        is_approved = any(
            approved.lower() in str(producer).lower() or
            approved.lower() in str(creator).lower()
            for approved in ALL_APPROVED
        )

        return {
            'file': pdf_path.name,
            'producer': producer,
            'creator': creator,
            'is_approved': is_approved,
            'path': pdf_path
        }

    except Exception as e:
        return {
            'file': pdf_path.name,
            'producer': f'Error reading PDF: {e}',
            'creator': 'N/A',
            'is_approved': False,
            'path': pdf_path
        }


def format_metadata_string(value):
    """Format metadata string for display (truncate if too long)"""
    value_str = str(value)
    if len(value_str) > 60:
        return value_str[:57] + '...'
    return value_str


def main():
    if len(sys.argv) < 2:
        print("Usage: python check_pdf_source.py <folder_or_file>")
        print("\nExample:")
        print("  python check_pdf_source.py /path/to/pdfs")
        print("  python check_pdf_source.py document.pdf")
        sys.exit(1)

    path = Path(sys.argv[1])

    # Handle single file or directory
    if path.is_file():
        if path.suffix.lower() != '.pdf':
            print(f"Error: {path} is not a PDF file")
            sys.exit(1)
        pdfs = [path]
    elif path.is_dir():
        pdfs = sorted(path.glob('**/*.pdf'))
    else:
        print(f"Error: {path} does not exist")
        sys.exit(1)

    if not pdfs:
        print(f"No PDFs found in {path}")
        return

    print(f"\n{'='*80}")
    print(f"PDF CONVERTER VALIDATION REPORT")
    print(f"{'='*80}")
    print(f"Checking {len(pdfs)} PDF(s)...\n")

    # Process all PDFs
    approved = []
    violations = []

    for pdf_path in pdfs:
        result = check_pdf_source(pdf_path)

        if result['is_approved']:
            approved.append(result)
        else:
            violations.append(result)

    # Display results grouped by status
    if approved:
        print(f"✅ APPROVED ({len(approved)} files)")
        print(f"{'-'*80}")
        for r in approved:
            print(f"  ✓ {r['file']}")
            print(f"      Producer: {format_metadata_string(r['producer'])}")
            print(f"      Creator:  {format_metadata_string(r['creator'])}")
        print()

    if violations:
        print(f"❌ NOT APPROVED ({len(violations)} files)")
        print(f"{'-'*80}")
        for r in violations:
            print(f"  ✗ {r['file']}")
            print(f"      Producer: {format_metadata_string(r['producer'])}")
            print(f"      Creator:  {format_metadata_string(r['creator'])}")
            print(f"      Path:     {r['path']}")
        print()

    # Summary
    print(f"{'='*80}")
    print(f"SUMMARY")
    print(f"{'='*80}")
    print(f"Total PDFs:    {len(pdfs)}")
    print(f"✅ Approved:   {len(approved)}")
    print(f"❌ Violations: {len(violations)}")

    if violations:
        print(f"\n⚠️  WARNING: {len(violations)} PDF(s) were NOT converted with approved tools!")
        print(f"\nApproved tools: {', '.join(APPROVED_PRODUCERS)}")
        print(f"\nViolations found:")
        for r in violations:
            print(f"  • {r['file']}")

        # Exit with error code if violations found
        sys.exit(1)
    else:
        print(f"\n✓ All PDFs were converted with approved tools!")
        sys.exit(0)


def check_single_pdf_detailed(pdf_path):
    """
    Detailed check of a single PDF (for debugging).
    Call this directly in Python for more info.
    """
    print(f"\nDetailed PDF Analysis: {pdf_path}")
    print(f"{'='*80}")

    try:
        reader = PdfReader(pdf_path)
        metadata = reader.metadata

        if not metadata:
            print("❌ No metadata found in PDF")
            return

        print("\nAll Metadata Fields:")
        for key, value in metadata.items():
            print(f"  {key}: {value}")

        print(f"\nPDF Version: {reader.pdf_header}")
        print(f"Number of Pages: {len(reader.pages)}")
        print(f"Encrypted: {reader.is_encrypted}")

        # Check approval
        result = check_pdf_source(Path(pdf_path))
        print(f"\nApproval Status: {'✅ APPROVED' if result['is_approved'] else '❌ NOT APPROVED'}")

    except Exception as e:
        print(f"❌ Error: {e}")


if __name__ == '__main__':
    main()
