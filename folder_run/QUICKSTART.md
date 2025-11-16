# Quick Start Guide - Folder Batch Anonymization

## Setup (One-Time)

### 1. Install Dependencies

```bash
cd folder_run
pip install -r requirements.txt
```

### 2. Verify LibreOffice is Installed

```bash
libreoffice --version
```

If not installed:
- **Windows:** `winget install LibreOffice.LibreOffice`
- **Linux/WSL:** `sudo apt-get install libreoffice`

## Your First Run - Project Nautilus

### Step 1: Preview (Dry Run)

**See what will be processed without making changes:**

```bash
cd folder_run

python3 batch_anonymize.py \
  --input "./input/Project Nautilus" \
  --output "./output" \
  --tracker "./tracker/Anon Tracker - Nautilus.xlsx" \
  --dry-run
```

**You'll see:**
- ✅ 139 processable files discovered
- ✅ 367 anonymization mappings loaded
- ⚠️ 19 non-processable files flagged (14 PDFs, 1 PNG - will be skipped)
- File type breakdown per folder
- Estimated processing time

### Step 2: Process with Prompting

**Run live processing with per-folder image removal decisions:**

```bash
python3 batch_anonymize.py \
  --input "./input/Project Nautilus" \
  --output "./output" \
  --tracker "./tracker/Anon Tracker - Nautilus.xlsx"
```

**You'll be prompted for each folder:**
```
────────────────────────────────────────────────────────
Folder: 1. Financials
Subdirectories: 13
Processable Files: 58 total
Types: DOCX: 37 | XLSX: 20 | PPTX: 1

⚠ WARNING: This folder contains non-processable files:
  • PDF: 3 file(s) (will be SKIPPED)
  Total non-processable: 3

Sample processable files:
  • 2023_10K.docx
  • Q1_2024_Financials.xlsx
  • Investor_Presentation.pptx

Est. time: ~4.8 minutes
────────────────────────────────────────────────────────

Remove images from this folder? [yes / no / auto-yes / skip / quit]:
```

**Your choices:**
- `y` → Remove images from this folder (for 10Ks/10Qs with logos)
- `n` → Preserve images (for industry reports with charts/graphs)
- `a` → Auto-yes for ALL remaining folders
- `s` → Skip this entire folder
- `q` → Stop processing now

**Recommended approach for Nautilus:**
- **10K/10Q folders** → Type `y` (remove logos)
- **Industry Reports** → Type `n` (preserve charts)
- **Legal documents** → Type `y` (remove letterhead)
- **Investor presentations** → Type `n` (preserve slides)

### Step 3: Monitor Progress

**You'll see live updates:**
```
Progress: [████████████████████░░░░░░░░] 75.5%
Files: 105/139 | Success: 102 | Failed: 1 | Skipped: 2
Stats: Replacements: 3,847 | Images Removed: 127
Time: 8m 34s
Current: 3. Operational and Commercial/3.2 Real Estate/Lease_Agreement_NYC.docx
```

### Step 4: Review Results

**After completion, check:**

1. **Terminal Summary:**
```
╔════════════════════════════════════════════════════════════╗
║              BATCH PROCESSING SUMMARY                      ║
╚════════════════════════════════════════════════════════════╝

Files Processed:     139
  ✓ Succeeded:       135
  ✗ Failed:          2
  ⊘ Skipped:         2

Anonymization:
  Replacements:       4,127
  Images Removed:     148

PDF Conversion:
  ✓ Succeeded:       132
  ✗ Failed:          3

Performance:
  Success Rate:       97.1%
  Total Processing Time:         12m 45s
```

2. **Output Files:**
   - `output/Project Nautilus/` - Anonymized files (original formats)
   - `pdf_output/Project Nautilus/` - PDF versions

3. **Excel Report:**
   - `reports/batch_report_20251115_143022.xlsx`
   - 4 sheets: File Details, Folder Summary, Run Summary, Error Log

4. **Processing Log:**
   - `logs/batch_run_20251115_143022.log`
   - Detailed file-by-file logs for troubleshooting

## Faster Options

### Skip PDF Conversion (40% faster)

```bash
python3 batch_anonymize.py \
  --input "./input/Project Nautilus" \
  --output "./output" \
  --tracker "./tracker/Anon Tracker - Nautilus.xlsx" \
  --no-pdf
```

### Auto-Remove All Images (No prompting)

```bash
python3 batch_anonymize.py \
  --input "./input/Project Nautilus" \
  --output "./output" \
  --tracker "./tracker/Anon Tracker - Nautilus.xlsx" \
  --auto-yes-images
```

### Fastest (No PDFs + Auto-remove images)

```bash
python3 batch_anonymize.py \
  --input "./input/Project Nautilus" \
  --output "./output" \
  --tracker "./tracker/Anon Tracker - Nautilus.xlsx" \
  --no-pdf \
  --auto-yes-images
```

**Estimated time:** ~8 minutes (vs. 15-20 minutes with PDFs)

## Process Single Folder

**Test on one folder first:**

```bash
python3 batch_anonymize.py \
  --input "./input/Project Nautilus/1. Financials/1.1 Company Overview" \
  --output "./output/test" \
  --tracker "./tracker/Anon Tracker - Nautilus.xlsx"
```

## What Gets Processed

### ✅ Fully Processed
- **Word:** `.docx`, `.doc` (legacy converted first)
  - All text, tables, headers, footers, textboxes, comments
  - Images (removed if you choose `y`)
  - Hyperlinks (text and URLs)
  - **All metadata stripped** (author, company, etc.)

- **Excel:** `.xlsx`, `.xls` (legacy converted first)
  - **ALL sheets** (not just first sheet)
  - All cells, formulas, comments
  - Sheet names
  - **All metadata stripped**
  - Note: Images/charts NOT removed (complex objects)

- **PowerPoint:** `.pptx`, `.ppt` (legacy converted first)
  - All slides, text frames, tables
  - Speaker notes
  - Images (removed if you choose `y`)
  - **All metadata stripped**

### ⚠️ Skipped (with warnings)
- `.pdf` - Cannot be anonymized (output format only)
- `.png`, `.jpg`, `.jpeg` - Image files (reference materials)
- Tracker files - Automatically excluded

## Safety Features

✅ **Input files NEVER modified** - Separate output directory
✅ **Folder structure preserved** - Exact hierarchy maintained
✅ **Error isolation** - One file failure doesn't stop batch
✅ **Metadata always stripped** - Even if anonymization fails
✅ **Detailed logging** - Every operation logged
✅ **Timeout protection** - 5-minute timeout per file

## Troubleshooting

**No files found:**
- Check input path is correct
- Ensure files have supported extensions (.docx, .xlsx, .pptx)

**Failed to load tracker:**
- Verify tracker file path
- Check tracker has "Before" and "After" columns

**LibreOffice errors:**
- PDF conversion may fail for very large files (>10 MB)
- Legacy format conversion requires LibreOffice
- Check that LibreOffice is installed and in PATH

**Processing very slow:**
- Use `--no-pdf` to skip PDF conversion (40% faster)
- Large files (>5 MB) take 20-60 seconds each
- Close other applications to free resources

## Need Help?

- See `README.md` for full documentation
- Check `logs/` directory for detailed error messages
- Review `reports/` Excel file for statistics
- Dry-run mode (`--dry-run`) previews without processing

## Expected Performance - Nautilus Project

Based on dry-run analysis:
- **139 processable files**
- **7 top-level folders**
- **19 non-processable files** (will be skipped)

**Estimated times:**
- With PDF conversion: **~15-20 minutes**
- Without PDF conversion: **~8-10 minutes**
- Actual time depends on your choices (image removal takes ~2 seconds per file)

**Expected success rate:** 95-97% (based on proven backend performance)

---

## Ready to Process!

You're all set! Your input files are **safe** (never modified), the proven backend has processed thousands of files successfully, and comprehensive reporting will show you exactly what happened.

**Start with a dry-run to preview, then process with confidence!** 🚀
