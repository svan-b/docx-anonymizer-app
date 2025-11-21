# DOCX Anonymizer - xAI

Professional document anonymization tool for financial data rooms with web interface and batch processing capabilities.

🌐 **Live Application**: [https://docx-anonymizer-app.streamlit.app](https://docx-anonymizer-app.streamlit.app)

## Features

- **Multi-format support**: DOCX, DOC, XLSX, XLS, PPTX, PPT
- **Smart anonymization**: 128+ pattern rules with longest-first matching
- **Metadata stripping**: Removes all author, company, title metadata
- **Image removal**: Optional removal of logos and embedded images
- **Hyperlink cleaning**: Removes hyperlink metadata while preserving text
- **PDF generation**: Optional PDF conversion after anonymization
- **Batch processing**: Command-line tool for processing entire folder structures
- **Detailed tracking**: Excel reports showing all replacements made
- **Web interface**: User-friendly Streamlit app for single-file processing

## Quick Start

### Web Interface (Streamlit)

```bash
streamlit run streamlit_app.py
```

Visit `http://localhost:8501` and upload files through the web interface.

### Batch Processing

```bash
cd batch
python3 batch_anonymize.py \
  --input "input/Project Name" \
  --output "output" \
  --tracker "tracker/Anon Tracker.xlsx" \
  --auto-no-images \
  --remove-hyperlinks \
  --timestamp-output
```

## Project Structure

```
/
├── streamlit_app.py              # Streamlit Cloud entrypoint (wrapper)
├── src/                          # Core application code
│   ├── streamlit_app.py          # Main Streamlit application
│   ├── processors/               # Document processors
│   │   ├── docx_processor.py     # Word document processing
│   │   ├── excel_processor.py    # Excel processing
│   │   └── pptx_processor.py     # PowerPoint processing
│   └── utils/                    # Utility modules
│       ├── anonymizer_utils.py   # Core anonymization logic
│       ├── hyperlink_utils.py    # Hyperlink removal
│       └── fix_ooxml_int_conversion.py
│
├── batch/                        # Batch processing system
│   ├── batch_anonymize.py        # Main batch script
│   ├── input/                    # Input documents
│   ├── output/                   # Anonymized outputs
│   ├── tracker/                  # Excel tracking files
│   ├── commands/                 # Saved command files
│   └── logs/                     # Processing logs
│
├── scripts/                      # Utility scripts
│   └── check_pdf_source.py       # Verify PDF source (Adobe vs 3rd-party)
│
├── tools/                        # Standalone tools
│   └── google_apps_script/       # Google Drive PDF validator
│
└── docs/                         # Documentation
    ├── DEPLOYMENT_GUIDE.md       # Deployment instructions
    ├── V2.1_ROADMAP.md           # Future features
    └── [technical reports]       # Bug fixes and feature docs
```

## Installation

```bash
# Clone repository
git clone <repository-url>
cd docx-anonymizer-app

# Install dependencies
pip install -r requirements.txt

# Install LibreOffice (for PDF conversion)
# Ubuntu/WSL:
sudo apt-get install libreoffice

# macOS:
brew install --cask libreoffice
```

## How to Use

### Web Interface

1. **Upload Requirements Excel**: Provide Excel file with anonymization mappings
   - Column 1: Text to replace (Before)
   - Column 2: Replacement text (After)
2. **Upload Documents**: Drag and drop DOCX, XLSX, or PPTX files
3. **Configure Options**:
   - Remove all images (checked by default)
   - Remove hyperlink metadata
   - Generate PDFs
4. **Process**: Click "Execute Anonymization"
5. **Download**: Get anonymized files as ZIP archive

### Batch Processing

See `batch/commands/` for real-world examples:

```bash
# Process entire folder structure with auto-prompting
python3 batch_anonymize.py \
  --input "input/Project Sunridge VDR" \
  --output "output" \
  --tracker "tracker/Anon Tracker - Sunridge.xlsx" \
  --auto-no-images \
  --remove-hyperlinks \
  --timestamp-output

# Process with specific folders for image removal
python3 batch_anonymize.py \
  --input "input/Project Alpaca" \
  --output "output" \
  --tracker "tracker/Anon Tracker - Alpaca.xlsx" \
  --auto-no-images \
  --remove-hyperlinks \
  --timestamp-output \
  --remove-images-from-folders "1. Financials\1.6 Audited Financials"
```

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
NASDAQ: BROS        | NASDAQ: XXXX
```

## Technology Stack

- **Python 3.9+**
- **Streamlit** - Web interface
- **python-docx** - Word document processing
- **python-pptx** - PowerPoint processing
- **openpyxl** - Excel processing
- **LibreOffice** - PDF conversion

## Version History

### v2.0.1 - Performance Optimization (Nov 20, 2025)
- 2.5x speedup on large legal documents (235s → 96s)
- Optimized header/footer processing for multi-section documents
- Set-based hyperlink detection (O(1) lookup)

### v2.0 - Hyperlink Removal (Nov 20, 2025)
- Remove hyperlink metadata while preserving display text
- Fixes clickable blue links in anonymized documents

### v1.9 - UX Enhancement (Nov 20, 2025)
- Clear file uploads on "New Batch" button click
- Improved user experience for multi-batch workflows

## Deployment

### Live Production Instance

This application is deployed on **Streamlit Cloud** with automatic deployments from the `main` branch.

- **Platform**: Streamlit Cloud (Community Tier)
- **Auto-Deploy**: Enabled on every push to `main`
- **Uptime**: 24/7 availability
- **Max Upload Size**: 200MB

### Making Changes

1. Commit and push to `main` branch
2. Streamlit Cloud automatically redeploys
3. Changes live within 1-2 minutes

## Privacy & Security

- **Server-side Processing**: All operations happen on secure Streamlit servers
- **Temporary Storage**: Files stored only during active session
- **Automatic Cleanup**: All files deleted when session ends
- **No Data Retention**: No documents are saved or logged
- **HTTPS**: All traffic encrypted in transit

## License

Proprietary - xAI Internal Use Only

## Support

For issues or questions, contact the development team or check `docs/` for detailed documentation.
