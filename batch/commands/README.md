# Batch Processing Command Files

This directory contains saved command-line arguments for batch processing different projects.

## Usage

```bash
cd ..  # Go to batch/ directory
bash commands/alpaca.txt
# OR
python3 $(cat commands/alpaca.txt)
```

## Command File Format

Each `.txt` file contains a complete `batch_anonymize.py` command with all arguments.

### Example: Basic Processing

```bash
python3 batch_anonymize.py \
  --input "input/Project Name" \
  --output "output" \
  --tracker "tracker/Anon Tracker.xlsx" \
  --auto-no-images \
  --remove-hyperlinks \
  --timestamp-output
```

### Example: With Specific Folder Image Removal

```bash
python3 batch_anonymize.py \
  --input "input/Project Alpaca" \
  --output "output" \
  --tracker "tracker/Anon Tracker - Alpaca.xlsx" \
  --auto-no-images \
  --remove-hyperlinks \
  --timestamp-output \
  --remove-images-from-folders "1. Financials\1.6 Audited Financials,1. Financials\1.5 Compliance Certificates"
```

## Available Options

- `--input` - Input directory path (required)
- `--output` - Output directory path (required)
- `--tracker` - Excel tracker file path (required)
- `--auto-no-images` - Automatically prompt for image removal per folder
- `--remove-hyperlinks` - Remove hyperlink metadata
- `--timestamp-output` - Add timestamp to output folder names
- `--no-pdf` - Skip PDF generation
- `--remove-images-from-folders` - Comma-separated list of specific folders for image removal
- `--parallel-workers N` - Use N parallel workers (default: 1 = sequential)

## Saved Commands

- `alpaca.txt` - Project Alpaca processing
- `enclave.txt` - Project Enclave processing
- `vitals.txt` - Project Vitals processing
- `sunridge.txt` - Project Sunridge processing

## Notes

- Always use `--auto-no-images` for interactive folder-by-folder prompting
- Use `--remove-images-from-folders` to specify exact folders (comma-separated paths)
- Parallel workers (`--parallel-workers 6`) can speed up large batches but may cause issues with some documents
- Output folders are timestamped by default when using `--timestamp-output`
