# DOCX Anonymizer + PDF Converter

A professional web application built with Streamlit for anonymizing Word documents (.docx/.doc) with custom mappings and PDF conversion.

🌐 **Live Application**: [https://docx-anonymizer-app.streamlit.app](https://docx-anonymizer-app.streamlit.app)

## Features

- **Batch Processing**: Upload multiple DOCX/DOC files at once
- **Excel-based Mapping**: Define before/after anonymization rules in Excel
- **Image Removal**: Optional removal of all images from documents
- **Header/Footer Clearing**: Clear headers and footers (useful for presentations with logos)
- **PDF Conversion**: Automatic conversion to PDF using LibreOffice
- **ZIP Download**: Download all processed files as ZIP archives
- **xAI Branded UI**: Clean, modern interface with soft aesthetic

## How to Use

1. **Upload Word Files**: Drag and drop your DOCX or DOC files
2. **Upload Requirements Excel**: Provide an Excel file with anonymization mappings
   - Column 1: Text to replace (Before)
   - Column 2: Replacement text (After)
3. **Configure Options**:
   - Remove all images (checked by default)
   - Clear headers/footers (for presentations with logos)
4. **Process**: Click "Execute Anonymization" to process files
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

### Live Production Instance

This application is deployed on **Streamlit Cloud** with automatic deployments from the `main` branch.

- **Platform**: Streamlit Cloud (Community Tier)
- **Auto-Deploy**: Enabled on every push to `main`
- **Uptime**: 24/7 availability
- **Repository**: Public (required for Streamlit Community hosting)

### Deployment Configuration

- **System Dependencies**: LibreOffice (via `packages.txt`)
- **Python Dependencies**: All packages in `requirements.txt`
- **Configuration**: Custom theme in `.streamlit/config.toml`
- **Max Upload Size**: 200MB
- **Processing**: Server-side with automatic cleanup

### Making Changes

1. Update files locally in `/mnt/c/Users/vanbo/Development/Projects/xAI/anonymous/vdr-processor-docx/ui/`
2. Run sync script: `/tmp/sync_to_streamlit.sh`
3. Changes automatically deploy within 1-2 minutes

### Monitoring

- Streamlit Cloud provides automatic health checks
- Build logs available in Streamlit Cloud dashboard
- App automatically restarts on deployment or errors

## Privacy & Security

- **Server-side Processing**: All operations happen on secure Streamlit servers
- **Temporary Storage**: Files stored only during active session
- **Automatic Cleanup**: All files deleted when session ends
- **No Data Retention**: No documents are saved or logged
- **HTTPS**: All traffic encrypted in transit
