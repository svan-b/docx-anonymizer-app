#!/usr/bin/env python3
"""
DOCX Anonymizer + PDF Converter - Streamlit App
Drag and drop DOCX files + Excel requirements to anonymize and convert to PDF
"""
import streamlit as st
import sys
import os
from pathlib import Path
import tempfile
import shutil
import subprocess
import zipfile
from datetime import datetime

# Import our anonymization functions (now in same directory)
from process_adobe_word_files import (
    load_aliases_from_excel,
    categorize_and_sort_aliases,
    process_single_docx
)
import logging

# Configure page
st.set_page_config(
    page_title="DOCX Anonymizer + PDF Converter",
    page_icon="📄",
    layout="wide"
)

# Initialize session state for results persistence
if 'processing_complete' not in st.session_state:
    st.session_state.processing_complete = False
if 'results' not in st.session_state:
    st.session_state.results = []
if 'total_files' not in st.session_state:
    st.session_state.total_files = 0
if 'total_replacements' not in st.session_state:
    st.session_state.total_replacements = 0
if 'total_images' not in st.session_state:
    st.session_state.total_images = 0
if 'docx_zip_data' not in st.session_state:
    st.session_state.docx_zip_data = None
if 'pdf_zip_data' not in st.session_state:
    st.session_state.pdf_zip_data = None
if 'timestamp' not in st.session_state:
    st.session_state.timestamp = None

# Title and description
st.title("📄 DOCX Anonymizer + PDF Converter")

# File uploads
col1, col2 = st.columns([2, 1])

with col1:
    st.subheader("1. Upload Word Files")
    docx_files = st.file_uploader(
        "Drag and drop DOCX or DOC files here",
        type=['docx', 'doc'],
        accept_multiple_files=True,
        key="docx_upload"
    )

with col2:
    st.subheader("2. Upload Requirements Excel")
    excel_file = st.file_uploader(
        "Before/After mappings (.xlsx)",
        type=['xlsx'],
        key="excel_upload"
    )

st.divider()

# Processing options
st.subheader("3. Processing Options")
col1, col2 = st.columns(2)
with col1:
    remove_images = st.checkbox("Remove all images from document", value=True, key="remove_images")
with col2:
    clear_headers_footers = st.checkbox("Clear headers/footers (for presentations with logos)", value=False, key="clear_headers_footers")

st.divider()

# Process button
if st.button("🚀 Process Files", type="primary", disabled=(not docx_files or not excel_file)):
    # Reset session state for new processing
    st.session_state.processing_complete = False
    st.session_state.results = []
    st.session_state.docx_zip_data = None
    st.session_state.pdf_zip_data = None

    # Validate LibreOffice is available
    with st.spinner("Checking LibreOffice installation..."):
        try:
            result = subprocess.run(['soffice', '--version'], capture_output=True, timeout=5)
            if result.returncode != 0:
                st.error("❌ LibreOffice not found. Please install LibreOffice to use this tool.")
                st.info("Ubuntu/WSL: `sudo apt-get install libreoffice`")
                st.stop()
        except FileNotFoundError:
            st.error("❌ LibreOffice not found. Please install LibreOffice to use this tool.")
            st.info("Ubuntu/WSL: `sudo apt-get install libreoffice`")
            st.stop()
        except Exception as e:
            st.error(f"❌ Error checking LibreOffice: {e}")
            st.stop()

    # Defensive check - ensure files are still present
    if not docx_files or not excel_file:
        st.error("❌ Files are missing. Please upload both DOCX files and Excel requirements.")
        st.stop()

    # Create temporary directories
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        input_dir = temp_path / "input"
        docx_output_dir = temp_path / "docx_output"
        pdf_output_dir = temp_path / "pdf_output"

        input_dir.mkdir()
        docx_output_dir.mkdir()
        pdf_output_dir.mkdir()

        # Save Excel requirements
        excel_path = temp_path / "requirements.xlsx"
        with open(excel_path, 'wb') as f:
            f.write(excel_file.getbuffer())

        # Save Word files and convert .doc to .docx if needed
        files_to_process = []
        for docx_file in docx_files:
            file_path = input_dir / docx_file.name
            with open(file_path, 'wb') as f:
                f.write(docx_file.getbuffer())

            # Convert .doc to .docx using LibreOffice
            if docx_file.name.lower().endswith('.doc') and not docx_file.name.lower().endswith('.docx'):
                st.info(f"Converting {docx_file.name} from DOC to DOCX...")
                try:
                    cmd = [
                        'soffice',
                        '--headless',
                        '--norestore',
                        '--nologo',
                        '--nofirststartwizard',
                        '--convert-to', 'docx',
                        '--outdir', str(input_dir),
                        str(file_path)
                    ]
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

                    # Check for converted file
                    converted_path = file_path.with_suffix('.docx')
                    if converted_path.exists():
                        files_to_process.append((docx_file.name, converted_path))
                    else:
                        st.error(f"Failed to convert {docx_file.name}")
                        continue
                except Exception as e:
                    st.error(f"Error converting {docx_file.name}: {e}")
                    continue
            else:
                files_to_process.append((docx_file.name, file_path))

        # Load mappings
        with st.spinner("Loading anonymization mappings..."):
            try:
                alias_map = load_aliases_from_excel(excel_path)
                sorted_keys = categorize_and_sort_aliases(alias_map)
                st.success(f"✓ Loaded {len(alias_map)} mappings")
            except Exception as e:
                st.error(f"Error loading mappings: {e}")
                st.stop()

        st.divider()
        st.subheader("📝 Processing Files")

        # Process each file
        total_replacements = 0
        total_images = 0
        results = []

        progress_bar = st.progress(0)
        status_text = st.empty()

        logger = logging.getLogger(__name__)

        for i, (original_name, input_path) in enumerate(files_to_process):
            status_text.text(f"Processing {i+1}/{len(files_to_process)}: {original_name}")

            # Output name based on original file
            docx_output_path = docx_output_dir / Path(original_name).with_suffix('.docx').name

            # Anonymize DOCX
            with st.expander(f"📄 {original_name}", expanded=True):
                try:
                    replacements, images = process_single_docx(
                        input_path, docx_output_path, alias_map, sorted_keys, logger,
                        remove_images=remove_images,
                        clear_headers_footers_flag=clear_headers_footers
                    )

                    total_replacements += replacements
                    total_images += images

                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Replacements", replacements)
                    with col2:
                        st.metric("Images Removed", images)
                    with col3:
                        st.metric("Status", "✓ DOCX Done")

                    # Convert to PDF
                    pdf_output_path = pdf_output_dir / Path(original_name).with_suffix('.pdf').name

                    try:
                        cmd = [
                            'soffice',
                            '--headless',
                            '--norestore',
                            '--nologo',
                            '--nofirststartwizard',
                            '--convert-to', 'pdf',
                            '--outdir', str(pdf_output_dir),
                            str(docx_output_path)
                        ]

                        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)

                        # LibreOffice creates file with same stem
                        expected_output = pdf_output_dir / f"{docx_output_path.stem}.pdf"

                        if expected_output.exists():
                            if expected_output != pdf_output_path:
                                shutil.move(str(expected_output), str(pdf_output_path))

                            size_mb = pdf_output_path.stat().st_size / (1024 * 1024)
                            st.success(f"✓ PDF created ({size_mb:.1f} MB)")

                            results.append({
                                'filename': original_name,
                                'replacements': replacements,
                                'images': images,
                                'pdf_status': 'Success',
                                'pdf_size_mb': size_mb
                            })
                        else:
                            st.warning("⚠ PDF conversion failed - LibreOffice error")
                            results.append({
                                'filename': original_name,
                                'replacements': replacements,
                                'images': images,
                                'pdf_status': 'Failed',
                                'pdf_size_mb': 0
                            })

                    except subprocess.TimeoutExpired:
                        st.warning("⚠ PDF conversion timeout (5 min exceeded)")
                        results.append({
                            'filename': original_name,
                            'replacements': replacements,
                            'images': images,
                            'pdf_status': 'Timeout',
                            'pdf_size_mb': 0
                        })
                    except Exception as e:
                        st.warning(f"⚠ PDF conversion error: {e}")
                        results.append({
                            'filename': original_name,
                            'replacements': replacements,
                            'images': images,
                            'pdf_status': 'Error',
                            'pdf_size_mb': 0
                        })

                except Exception as e:
                    st.error(f"❌ Error processing DOCX: {e}")
                    results.append({
                        'filename': original_name,
                        'replacements': 0,
                        'images': 0,
                        'pdf_status': 'DOCX Error',
                        'pdf_size_mb': 0
                    })

            progress_bar.progress((i + 1) / len(files_to_process))

        status_text.text("✓ All files processed!")

        # Save results to session state
        st.session_state.results = results
        st.session_state.total_files = len(files_to_process)
        st.session_state.total_replacements = total_replacements
        st.session_state.total_images = total_images
        st.session_state.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # Create ZIP files and store in session state
        timestamp = st.session_state.timestamp

        # ZIP anonymized DOCX files
        docx_zip_path = temp_path / f"anonymized_docx_{timestamp}.zip"
        with zipfile.ZipFile(docx_zip_path, 'w') as zipf:
            for docx_file in docx_output_dir.glob('*.docx'):
                zipf.write(docx_file, docx_file.name)

        with open(docx_zip_path, 'rb') as f:
            st.session_state.docx_zip_data = f.read()

        # ZIP PDFs
        pdf_zip_path = temp_path / f"anonymized_pdf_{timestamp}.zip"
        with zipfile.ZipFile(pdf_zip_path, 'w') as zipf:
            for pdf_file in pdf_output_dir.glob('*.pdf'):
                zipf.write(pdf_file, pdf_file.name)

        with open(pdf_zip_path, 'rb') as f:
            st.session_state.pdf_zip_data = f.read()

        st.session_state.processing_complete = True

# Display results from session state (persists after download)
if st.session_state.processing_complete:
    st.divider()
    st.subheader("📊 Summary")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Files", st.session_state.total_files)
    with col2:
        st.metric("Total Replacements", st.session_state.total_replacements)
    with col3:
        st.metric("Total Images Removed", st.session_state.total_images)

    # Show results table
    st.dataframe(st.session_state.results, use_container_width=True)

    st.divider()
    st.subheader("📥 Download Results")
    st.info("Files are ready to download. Results remain visible after downloading.")

    col1, col2, col3 = st.columns([1, 1, 1])

    with col1:
        if st.session_state.docx_zip_data:
            st.download_button(
                label="📦 Download DOCX (ZIP)",
                data=st.session_state.docx_zip_data,
                file_name=f"anonymized_docx_{st.session_state.timestamp}.zip",
                mime="application/zip"
            )

    with col2:
        if st.session_state.pdf_zip_data:
            st.download_button(
                label="📦 Download PDFs (ZIP)",
                data=st.session_state.pdf_zip_data,
                file_name=f"anonymized_pdf_{st.session_state.timestamp}.zip",
                mime="application/zip"
            )

    with col3:
        if st.button("🔄 Start New Batch"):
            st.session_state.processing_complete = False
            st.session_state.results = []
            st.session_state.docx_zip_data = None
            st.session_state.pdf_zip_data = None
            st.rerun()

# Sidebar info
with st.sidebar:
    st.header("Stats")
    if docx_files:
        st.metric("DOCX Files", len(docx_files))
    if excel_file:
        st.success("✓ Requirements loaded")

    st.divider()
    st.caption("Excel format: Column 1 = Before, Column 2 = After")
    st.caption("Results persist after download")
