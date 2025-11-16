# Folder Batch Anonymization System

## Overview

Production-grade batch processing system for anonymizing entire folder structures. Extends the proven Streamlit app backend to process folders locally with per-folder image removal prompting, real-time progress monitoring, and comprehensive reporting.

## Features

✅ **Safe File Processing** - Input files NEVER modified (separate output directory)
✅ **Folder Structure Preservation** - Exact folder hierarchy maintained in output
✅ **Per-Folder Prompting** - Choose image removal per folder (10Ks vs. industry reports)
✅ **Multi-Format Support** - DOCX, XLSX, PPTX, DOC, XLS, PPT
✅ **Real-Time Progress** - Live terminal display with stats
✅ **Comprehensive Reporting** - Detailed Excel reports with per-file stats
✅ **Error Isolation** - One file failure doesn't stop the batch
✅ **Dry-Run Mode** - Preview before processing
✅ **Non-Processable File Detection** - Flags PDFs, images, and other unsupported formats

## Directory Structure

```
folder_run/
├── batch_anonymize.py          # Main batch processing script
├── README.md                   # This file
├── requirements.txt            # Python dependencies
│
├── input/                      # YOUR SOURCE FILES (NEVER MODIFIED)
│   └── Project Nautilus/       # Example project
│       ├── Anon Tracker - Nautilus.xlsx  (can be here or in tracker/)
│       ├── 1. Financials/
│       ├── 2. Legal and Corporate/
│       └── ...
│
├── output/                     # ANONYMIZED FILES (original formats)
│   └── Project Nautilus/       # Mirrors input structure
│
├── pdf_output/                 # PDF CONVERSIONS (optional)
│   └── Project Nautilus/       # Mirrors input structure
│
├── tracker/                    # ANONYMIZATION MAPPINGS
│   └── Anon Tracker - Nautilus.xlsx
│
├── logs/                       # PROCESSING LOGS
│   └── batch_run_20251115_143022.log
│
└── reports/                    # EXCEL REPORTS
    └── batch_report_20251115_143022.xlsx
```

## Prerequisites

### System Requirements

- Python 3.7+
- LibreOffice (for PDF conversion and legacy format conversion)

### Install LibreOffice

**Windows:**
```bash
winget install LibreOffice.LibreOffice
```

**Linux (WSL/Ubuntu):**
```bash
sudo apt-get update
sudo apt-get install libreoffice
```

**macOS:**
```bash
brew install libreoffice
```

### Python Dependencies

```bash
pip install -r requirements.txt
```

Required packages:
- `python-docx` - Word document processing
- `python-pptx` - PowerPoint processing
- `openpyxl` - Excel processing and report generation

## Quick Start

### 1. Dry Run (Preview Only)

Preview what will be processed without making any changes:

```bash
python batch_anonymize.py \
  --input "./input/Project Nautilus" \
  --output "./output" \
  --tracker "./tracker/Anon Tracker - Nautilus.xlsx" \
  --dry-run
```

**Output:**
- Shows all folders and file counts
- Displays file type breakdowns
- **Flags folders with PDFs and non-processable files**
- No files are processed

### 2. Live Processing with Prompting

Process files with per-folder image removal prompting:

```bash
python batch_anonymize.py \
  --input "./input/Project Nautilus" \
  --output "./output" \
  --tracker "./tracker/Anon Tracker - Nautilus.xlsx"
```

**Per-Folder Prompt Options:**
- `y` (yes) - Remove images from this folder (use for 10Ks, 10Qs with logos)
- `n` (no) - Preserve images in this folder (use for industry reports with charts)
- `a` (auto-yes) - Auto-remove images for ALL remaining folders
- `s` (skip) - Skip this entire folder
- `q` (quit) - Stop processing immediately

### 3. Auto-Remove Images (No Prompting)

Automatically remove images from all folders:

```bash
python batch_anonymize.py \
  --input "./input/Project Nautilus" \
  --output "./output" \
  --tracker "./tracker/Anon Tracker - Nautilus.xlsx" \
  --auto-yes-images
```

### 4. Auto-Preserve Images (No Prompting)

Automatically preserve images in all folders:

```bash
python batch_anonymize.py \
  --input "./input/Project Nautilus" \
  --output "./output" \
  --tracker "./tracker/Anon Tracker - Nautilus.xlsx" \
  --auto-no-images
```

### 5. Skip PDF Conversion (Faster)

Process files without generating PDFs:

```bash
python batch_anonymize.py \
  --input "./input/Project Nautilus" \
  --output "./output" \
  --tracker "./tracker/Anon Tracker - Nautilus.xlsx" \
  --no-pdf
```

## Command-Line Options

| Option | Description | Required | Default |
|--------|-------------|----------|---------|
| `--input` | Input directory containing files to process | ✅ Yes | - |
| `--output` | Output directory for anonymized files | ✅ Yes | - |
| `--tracker` | Path to Excel anonymization tracker file | ✅ Yes | - |
| `--pdf-output` | Custom PDF output directory | ❌ No | `./pdf_output` |
| `--no-pdf` | Skip PDF conversion (faster) | ❌ No | `False` |
| `--dry-run` | Preview without processing | ❌ No | `False` |
| `--auto-yes-images` | Auto-remove images (no prompting) | ❌ No | `False` |
| `--auto-no-images` | Auto-preserve images (no prompting) | ❌ No | `False` |

## Tracker File Format

The Excel tracker must have these columns (flexible naming):

| Column | Alternatives | Description |
|--------|-------------|-------------|
| Before | Original | Text to find (e.g., "Shake Shack Inc.") |
| After | Replacement | Replacement text (e.g., "Bronco Corp.") |

**Special Features:**
- **Blank "After" values** - Deletes text entirely
- **Auto name reversal** - "John Doe" also matches "Doe, John"
- **Suffix stripping** - "Shake Shack Inc." also matches "Shake Shack"
- **Case preservation** - "MATADOR" → "BRONCO", "Matador" → "Bronco"

## Processing Behavior

### Supported File Types

**Processed:**
- `.docx`, `.pptx`, `.xlsx` (modern Office formats)
- `.doc`, `.ppt`, `.xls` (legacy formats - auto-converted first)

**Skipped (with warnings):**
- `.pdf` - Cannot be anonymized (output format only)
- `.png`, `.jpg`, `.jpeg` - Image files (reference materials)
- Any other file types

### Folder Processing Logic

1. **Discovery:** Scans folder tree recursively
2. **Organization:** Groups files by top-level folder
3. **Prompting:** Before each folder, shows:
   - Folder path and subdirectory count
   - Processable file count and types
   - **⚠ WARNING if PDFs or non-processable files present**
   - Sample filenames (first 5)
   - Estimated processing time
4. **Processing:** Processes all files in folder tree
5. **Reporting:** Aggregates stats by folder

### File Processing Steps

For each file:

1. **Convert legacy formats** (`.doc` → `.docx` via LibreOffice)
2. **Load document** (using proven backend modules)
3. **Anonymize text** (paragraph-level, case-preserving)
4. **Remove images** (if enabled for this folder)
5. **Strip metadata** (author, company, etc.)
6. **Save anonymized file** (preserves folder structure)
7. **Convert to PDF** (optional, via LibreOffice)
8. **Log results** (replacements, images removed, errors)

### Safety Mechanisms

✅ **Input files NEVER modified** - All output goes to separate directories
✅ **Folder structure preserved** - Exact hierarchy maintained
✅ **Error isolation** - File failures don't stop batch
✅ **Detailed logging** - Every operation logged to file
✅ **Timeout protection** - 5-minute timeout per LibreOffice operation
✅ **Metadata always stripped** - Even if anonymization fails

## Output Files

### 1. Anonymized Files (Original Formats)

**Location:** `output/` directory
**Format:** Same as input (.docx, .xlsx, .pptx)
**Structure:** Mirrors input folder structure exactly

### 2. PDF Files (Optional)

**Location:** `pdf_output/` directory
**Format:** All files converted to .pdf
**Structure:** Mirrors input folder structure exactly

### 3. Processing Log

**Location:** `logs/batch_run_YYYYMMDD_HHMMSS.log`
**Contents:**
- Timestamp for each operation
- File-by-file processing details
- Replacement counts, image removal counts
- Error messages with full stack traces
- Summary statistics

**Example:**
```
2025-11-15 14:30:22 | INFO     | Processing DOCX: 1. Financials/1.6 Audited Financials/10k/2023_10K.docx
2025-11-15 14:30:25 | INFO     | Completed: 47 replacements, 3 images removed in 2.8s
```

### 4. Excel Report

**Location:** `reports/batch_report_YYYYMMDD_HHMMSS.xlsx`
**Sheets:**

#### Sheet 1: File Details
Per-file statistics with columns:
- File Path, Folder, Filename, Type
- Status (success/failed/skipped)
- Replacements, Images Removed
- Processing Time (seconds)
- Error (if any)

#### Sheet 2: Folder Summary
Aggregated stats per folder:
- Total Files, Succeeded, Failed, Skipped
- Total Replacements, Images Removed
- Success Rate (%)

#### Sheet 3: Run Summary
Overall batch statistics:
- Total files processed
- Success/failure counts
- Total replacements and images removed
- PDF conversion results
- Processing start/end times
- Total processing time

#### Sheet 4: Error Log (if errors occurred)
Detailed error information:
- Timestamp, File Path, Error Message

## Progress Monitoring

### Terminal Display (Live Updates)

```
Progress: [████████████████████░░░░░░░░] 75.5%
Files: 115/152 | Success: 110 | Failed: 2 | Skipped: 3
Stats: Replacements: 2,847 | Images Removed: 89
Time: 9m 34s
Current: 1. Financials/1.11 Industry Reports/Global Coffee Market Analysis.docx
```

**Updated after each file completion**

### Log File (Detailed)

All operations logged to `logs/batch_run_YYYYMMDD_HHMMSS.log` with:
- Timestamps for every operation
- File-level success/failure
- Detailed error messages
- Summary statistics

## Examples

### Example 1: Process Nautilus Project with Prompting

```bash
cd folder_run

python batch_anonymize.py \
  --input "./input/Project Nautilus" \
  --output "./output" \
  --tracker "./tracker/Anon Tracker - Nautilus.xlsx"
```

**Interactive prompts for each folder:**
- 10K/10Q folders → Choose `y` (remove logos)
- Industry Reports → Choose `n` (preserve charts)
- Legal documents → Choose `y` (remove letterhead)

### Example 2: Dry Run First, Then Process

```bash
# Preview what will be processed
python batch_anonymize.py \
  --input "./input/Project Nautilus" \
  --output "./output" \
  --tracker "./tracker/Anon Tracker - Nautilus.xlsx" \
  --dry-run

# After reviewing, process with auto-remove images
python batch_anonymize.py \
  --input "./input/Project Nautilus" \
  --output "./output" \
  --tracker "./tracker/Anon Tracker - Nautilus.xlsx" \
  --auto-yes-images
```

### Example 3: Fast Processing (No PDFs)

```bash
# Skip PDF conversion for faster processing
python batch_anonymize.py \
  --input "./input/Project Nautilus" \
  --output "./output" \
  --tracker "./tracker/Anon Tracker - Nautilus.xlsx" \
  --auto-yes-images \
  --no-pdf
```

### Example 4: Process Single Folder

```bash
# Process just the Financials folder
python batch_anonymize.py \
  --input "./input/Project Nautilus/1. Financials" \
  --output "./output/1. Financials" \
  --tracker "./tracker/Anon Tracker - Nautilus.xlsx"
```

## Troubleshooting

### LibreOffice Not Found

**Error:** `FileNotFoundError: 'libreoffice' not found`

**Solution:** Install LibreOffice (see Prerequisites)

### Permission Denied

**Error:** `PermissionError: [Errno 13] Permission denied`

**Solution:**
- Ensure output directories are writable
- Close any files open in Excel/Word
- Run terminal as administrator (Windows)

### File Conversion Timeout

**Warning:** `PDF conversion timeout for large_file.docx`

**Solution:**
- Large files (>10 MB) may timeout
- Timeout is set to 5 minutes per file
- Processing continues for remaining files
- Check log file for details

### Empty Output Directory

**Issue:** No files in output directory after processing

**Check:**
1. Input directory contains processable files (.docx, .xlsx, .pptx)
2. Tracker file loaded successfully (check terminal output)
3. No errors in log file (`logs/batch_run_*.log`)
4. Not all folders were skipped during prompting

### Non-Processable Files Warning

**Warning:** `⚠ WARNING: This folder contains non-processable files`

**Explanation:**
- PDFs, PNGs, JPGs cannot be anonymized
- These files will be **SKIPPED** during processing
- Only DOCX, XLSX, PPTX files are processed
- Warning is informational - processing continues

## Performance

### Expected Processing Times

**Per File:**
- Small files (10-50 KB): 2-3 seconds
- Medium files (50-500 KB): 3-7 seconds
- Large files (500 KB - 5 MB): 10-20 seconds
- Very large files (>5 MB): 20-60 seconds

**For 152-File Project:**
- Sequential processing: ~13 minutes
- With PDF conversion: ~15-20 minutes
- Without PDF conversion: ~8-10 minutes

### Optimization Tips

1. **Skip PDF conversion** if not needed (`--no-pdf`) - saves ~40% time
2. **Use auto-modes** to avoid prompting delays (`--auto-yes-images`)
3. **Process subfolders separately** for parallelization across terminals
4. **Close other applications** to free up system resources

## FAQ

**Q: Will my original files be modified?**
A: **NO.** Input files are NEVER modified. All output goes to separate directories.

**Q: What happens if processing fails mid-batch?**
A: Successfully processed files remain in output. Check the log and Excel report for details. Re-run to process remaining files (they'll be overwritten if needed).

**Q: Can I resume a failed batch?**
A: Currently no auto-resume. You can manually process remaining folders by pointing `--input` to specific subdirectories.

**Q: How do I handle folders with mixed content (some need image removal, some don't)?**
A: Process the folder tree twice with different input paths, or manually separate files before processing.

**Q: Why are PDFs skipped?**
A: PDFs are complex binary formats that require specialized tools. This system processes editable Office documents. PDFs in input are likely reference materials.

**Q: Can I customize the timeout for LibreOffice?**
A: Currently hardcoded to 5 minutes. Edit `batch_anonymize.py` line ~410 to change timeout value.

**Q: How do I verify anonymization quality?**
A: Review the Excel report for replacement counts. Spot-check output files. Search for company names that should have been replaced.

## Support

**Issues:** Report bugs or request features via project issues
**Logs:** Always check `logs/batch_run_*.log` for detailed error messages
**Reports:** Excel report provides comprehensive statistics for analysis

## Version History

**v1.0 (2025-11-15)**
- Initial release
- Multi-format support (DOCX, XLSX, PPTX)
- Per-folder image removal prompting
- Real-time progress monitoring
- Comprehensive Excel reporting
- Non-processable file detection and warnings
- Dry-run mode
- Safety-first file handling

## Credits

Built on the proven, battle-tested backend from the Streamlit Document Anonymization App. All core processing modules reused without modification for maximum reliability.
