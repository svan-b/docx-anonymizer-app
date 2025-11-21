# PDF Converter Validation - Setup Guide

## Problem
Users are converting PDFs using third-party tools (pdf2go, iLovePDF, etc.) instead of Adobe Acrobat, which may compromise quality or compliance requirements.

## Solution
Automatically validate PDF metadata to ensure files were converted with Adobe tools.

---

## How PDF Metadata Works

Every PDF contains metadata showing which tool created it:

### Adobe Acrobat Signatures ✅
```
Producer: Adobe PDF Library 15.0
Creator: Adobe Acrobat Pro DC
```

### Third-Party Tools ❌
```
Producer: Skia/PDF m116 (pdf2go)
Producer: iLovePDF (ilovepdf.com)
Producer: Microsoft Print to PDF
Creator: Smallpdf.com
```

---

## Option 1: Google Apps Script (Recommended)

Automatically monitors a Google Drive folder and flags non-Adobe PDFs.

### Setup Instructions

1. **Open Google Drive**
   - Navigate to the folder you want to monitor
   - Copy the folder ID from the URL
     ```
     https://drive.google.com/drive/folders/1ABC123XYZ456
                                              ^^^^^^^^^^^^ (this part)
     ```

2. **Create Apps Script**
   - In Google Drive: Extensions > Apps Script
   - Create new project: "PDF Validator"
   - Copy code from `validate_pdf_source.gs`
   - Paste into script editor

3. **Configure Settings** (top of script)
   ```javascript
   const FOLDER_ID = '1ABC123XYZ456';  // Your folder ID
   const NOTIFICATION_EMAIL = 'your-email@company.com';
   const VIOLATION_ACTION = 'MOVE_TO_QUARANTINE';  // or 'NOTIFY_ONLY'
   ```

4. **Enable Drive API**
   - In Apps Script editor: Services (+ icon)
   - Search "Drive API" > Add

5. **First Run**
   - Select `setup` function from dropdown
   - Click Run ▶️
   - Authorize permissions when prompted

6. **Set Up Trigger** (Automatic Monitoring)
   - Click Triggers (⏰ icon on left)
   - Add Trigger:
     - Function: `validateNewPDFs`
     - Event source: Time-driven
     - Type: Minute timer
     - Interval: Every 5 minutes

7. **Test It**
   - Upload a test PDF to the monitored folder
   - Wait 5 minutes for trigger
   - Check Executions tab for results

### What It Does

1. **Scans** folder every 5 minutes for new PDFs
2. **Checks** PDF metadata for Adobe signatures
3. **Actions** when non-Adobe PDF detected:
   - `MOVE_TO_QUARANTINE`: Moves to `_NON_ADOBE_PDFs_QUARANTINE` subfolder
   - `NOTIFY_ONLY`: Sends email but leaves file in place
   - `DELETE`: Moves to trash
4. **Notifies** via email with detailed report

### Email Report Example
```
⚠️ Non-Adobe PDFs Detected (3 files)

File Name           | Producer              | Creator
--------------------|-----------------------|------------------
report_v2.pdf       | Skia/PDF m116         | pdf2go.com
invoice.pdf         | iLovePDF             | ilovepdf.com
summary.pdf         | Smallpdf             | smallpdf.com

Expected Producer: Adobe PDF, Adobe Acrobat
```

---

## Option 2: Local Python Script

For batch validation of existing PDFs on your computer.

### Requirements
```bash
pip install pypdf
```

### Script: `check_pdf_source.py`

```python
#!/usr/bin/env python3
"""
Check if PDFs were converted with Adobe tools
"""
from pathlib import Path
import sys
from pypdf import PdfReader

APPROVED_PRODUCERS = [
    'Adobe PDF',
    'Adobe Acrobat',
    'Adobe InDesign',
    'Adobe Illustrator'
]

def check_pdf_source(pdf_path):
    """Check if PDF was created with Adobe tools"""
    try:
        reader = PdfReader(pdf_path)
        metadata = reader.metadata

        producer = metadata.get('/Producer', 'Unknown')
        creator = metadata.get('/Creator', 'Unknown')

        is_adobe = any(approved.lower() in str(producer).lower()
                      for approved in APPROVED_PRODUCERS)

        return {
            'file': pdf_path.name,
            'producer': producer,
            'creator': creator,
            'is_adobe': is_adobe
        }

    except Exception as e:
        return {
            'file': pdf_path.name,
            'producer': f'Error: {e}',
            'creator': 'N/A',
            'is_adobe': False
        }

def main():
    if len(sys.argv) < 2:
        print("Usage: python check_pdf_source.py <folder_path>")
        sys.exit(1)

    folder = Path(sys.argv[1])
    pdfs = list(folder.glob('**/*.pdf'))

    if not pdfs:
        print(f"No PDFs found in {folder}")
        return

    print(f"Checking {len(pdfs)} PDFs...\n")
    print("="*80)

    violations = []
    for pdf_path in pdfs:
        result = check_pdf_source(pdf_path)

        if not result['is_adobe']:
            violations.append(result)
            print(f"❌ {result['file']}")
            print(f"   Producer: {result['producer']}")
            print(f"   Creator: {result['creator']}")
        else:
            print(f"✓ {result['file']}")

    print("="*80)
    print(f"\nResults: {len(violations)} non-Adobe PDFs found")

    if violations:
        print("\n⚠️ Files NOT converted with Adobe:")
        for v in violations:
            print(f"  - {v['file']}")

if __name__ == '__main__':
    main()
```

### Usage
```bash
# Check all PDFs in a folder
python check_pdf_source.py /path/to/pdfs

# Example output:
✓ document1.pdf
❌ document2.pdf
   Producer: Skia/PDF m116
   Creator: pdf2go.com
✓ document3.pdf
```

---

## Common PDF Producer Signatures

### ✅ APPROVED (Adobe)
- `Adobe PDF Library 10.0` - 17.0
- `Adobe Acrobat Pro DC`
- `Adobe Acrobat DC`
- `Adobe InDesign CC`
- `Adobe Illustrator CC`
- `Adobe Photoshop PDF Engine`

### ❌ NOT APPROVED (Third-Party)
- `Skia/PDF m116` → pdf2go
- `iLovePDF` → ilovepdf.com
- `Smallpdf` → smallpdf.com
- `Microsoft Print to PDF` → Windows built-in
- `macOS Version 10.15.7 Quartz PDFContext` → Mac Preview
- `LibreOffice 7.x` → LibreOffice Writer

---

## Testing

### Quick Test in Google Apps Script
1. Open script editor
2. Select `testValidation` function
3. Click Run ▶️
4. Check Execution log for results

### Upload Test Files
1. Convert a document using pdf2go.com
2. Upload to monitored folder
3. Wait 5 minutes
4. Verify it gets quarantined/flagged

---

## Troubleshooting

### "Cannot read property 'get' of undefined"
- Drive API not enabled
- Solution: Add Drive API in Services

### "Insufficient permissions"
- Need to authorize script
- Solution: Run `setup()` and approve permissions

### "No violations detected" but should be
- Metadata extraction failed
- Check Execution log for errors
- Try local Python script to verify metadata

### Email not sending
- Check `NOTIFICATION_EMAIL` is set correctly
- Verify Apps Script has Gmail permissions
- Check spam folder

---

## Best Practices

1. **Quarantine First**: Use `MOVE_TO_QUARANTINE` initially to avoid data loss
2. **Review Regularly**: Check quarantine folder weekly
3. **Whitelist Exceptions**: Add legitimate tools to `APPROVED_PRODUCERS` if needed
4. **Audit Trail**: Script adds comments to flagged files automatically
5. **Notify Users**: Send team email explaining Adobe requirement

---

## Integration with Your Workflow

Since your app uses LibreOffice for PDF conversion (`app.py:692`), you might want to:

1. **Add LibreOffice signature** to approved list (if acceptable):
   ```javascript
   const APPROVED_PRODUCERS = [
     'Adobe PDF',
     'Adobe Acrobat',
     'LibreOffice 7'  // Add this
   ];
   ```

2. **Separate validation** for user-uploaded vs. app-generated PDFs:
   - Monitor different folders
   - Use different approval rules

3. **Add to Streamlit App**: Show PDF metadata before download
   ```python
   # In app.py after PDF conversion
   from pypdf import PdfReader
   reader = PdfReader(pdf_output_path)
   st.info(f"PDF Producer: {reader.metadata.get('/Producer', 'Unknown')}")
   ```

---

## Support

For issues or questions:
- Check Execution log in Apps Script editor
- Test with local Python script first
- Verify PDF metadata manually using Adobe Acrobat (File > Properties)
