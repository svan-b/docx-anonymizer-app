# DOCX Anonymizer + PDF Converter

A web application built with Streamlit that anonymizes Word documents (.docx/.doc) based on custom mappings and converts them to PDF.

## Features

- **Batch Processing**: Upload multiple DOCX/DOC files at once
- **Excel-based Mapping**: Define before/after anonymization rules in Excel
- **Image Removal**: Optional removal of all images from documents
- **Header/Footer Clearing**: Clear headers and footers (useful for presentations with logos)
- **PDF Conversion**: Automatic conversion to PDF using LibreOffice
- **ZIP Download**: Download all processed files as ZIP archives

## How to Use

1. **Upload Word Files**: Drag and drop your DOCX or DOC files
2. **Upload Requirements Excel**: Provide an Excel file with anonymization mappings
   - Column 1: Text to replace (Before)
   - Column 2: Replacement text (After)
3. **Configure Options**:
   - Remove all images (checked by default)
   - Clear headers/footers (for presentations with logos)
4. **Process**: Click "Process Files" to anonymize and convert
5. **Download**: Get your anonymized DOCX and PDF files as ZIP archives

## Excel Format

Your requirements Excel file should have:
- **Column 1 (Before)**: Original text to find
- **Column 2 (After)**: Replacement text

Example:
```
Before              | After
--------------------|--------------------
Dutch Bros Inc.     | Project Barista Inc.
Founder Name        | Executive A
Secret Product      | Product X
```

## Technical Details

- Built with Python and Streamlit
- Uses `python-docx` for Word document processing
- LibreOffice for DOC to DOCX and DOCX to PDF conversion
- Supports case-sensitive and case-insensitive replacements
- Handles complex documents with tables, headers, footers, and text boxes

## Local Development

```bash
# Install dependencies
pip install -r requirements.txt

# Install LibreOffice (Ubuntu/WSL)
sudo apt-get install libreoffice

# Run the app
streamlit run app.py
```

## Deployment

This app is deployed on Streamlit Cloud with LibreOffice support via `packages.txt`.

## Privacy

All processing happens server-side. Files are temporarily stored during processing and automatically deleted when the session ends.
