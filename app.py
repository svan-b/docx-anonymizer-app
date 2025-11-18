#!/usr/bin/env python3
"""
DOCX Anonymizer + PDF Converter
Professional document anonymization tool for financial data rooms
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

# Import anonymization functions
from process_adobe_word_files import (
    load_aliases_from_excel,
    categorize_and_sort_aliases,
    process_single_docx,
    precompile_patterns
)
from process_powerpoint import process_single_pptx
from process_excel import process_single_xlsx
import logging

# Page configuration
st.set_page_config(
    page_title="DOCX Anonymizer - xAI",
    page_icon="xai_logo.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS - xAI Soft Aesthetic
st.markdown("""
<style>
    /* Main container */
    .block-container {
        padding-top: 1rem;
        padding-bottom: 3rem;
        max-width: 1400px;
    }

    /* Headers */
    h1 {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        font-weight: 300;
        letter-spacing: -0.03em;
        border-bottom: 1px solid rgba(255, 255, 255, 0.1);
        padding-bottom: 1rem;
        margin-bottom: 2rem;
        color: #FFFFFF;
    }

    h2, h3 {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        font-weight: 400;
        letter-spacing: -0.02em;
        color: rgba(255, 255, 255, 0.95);
    }

    h3 {
        font-size: 1.1rem;
        color: rgba(255, 255, 255, 0.8);
        font-weight: 500;
    }

    /* Metrics */
    [data-testid="stMetricValue"] {
        font-size: 2rem;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        font-weight: 300;
        color: #FFFFFF;
    }

    [data-testid="stMetricLabel"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 0.1em;
        color: rgba(255, 255, 255, 0.5);
        font-weight: 500;
    }

    /* Upload boxes */
    [data-testid="stFileUploader"] {
        border: 1px solid rgba(255, 255, 255, 0.15);
        border-radius: 16px;
        padding: 2rem;
        background: rgba(255, 255, 255, 0.02);
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
        backdrop-filter: blur(10px);
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
    }

    [data-testid="stFileUploader"]:hover {
        border-color: rgba(255, 255, 255, 0.25);
        box-shadow: 0 12px 48px rgba(0, 0, 0, 0.4);
        background: rgba(255, 255, 255, 0.04);
        transform: translateY(-2px);
    }

    /* Upload button styling */
    [data-testid="stFileUploader"] section button {
        background-color: rgba(255, 255, 255, 0.08) !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
        color: rgba(255, 255, 255, 0.9) !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important;
        font-weight: 400 !important;
        border-radius: 8px !important;
        padding: 0.5rem 1rem !important;
        transition: all 0.3s ease !important;
    }

    [data-testid="stFileUploader"] section button:hover {
        background-color: rgba(255, 255, 255, 0.15) !important;
        border-color: rgba(255, 255, 255, 0.3) !important;
        transform: translateY(-1px) !important;
    }

    /* Buttons */
    .stButton>button {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        font-weight: 400;
        letter-spacing: 0.01em;
        border-radius: 12px;
        border: 2px solid rgba(255, 255, 255, 0.4) !important;
        padding: 0.75rem 2rem;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        background: rgba(255, 255, 255, 0.12) !important;
        color: #FFFFFF !important;
    }

    .stButton>button:hover {
        border-color: rgba(255, 255, 255, 0.7) !important;
        box-shadow: 0 8px 24px rgba(255, 255, 255, 0.15);
        background: rgba(255, 255, 255, 0.2) !important;
        transform: translateY(-2px);
    }

    /* Data tables */
    [data-testid="stDataFrame"] {
        border: 1px solid rgba(255, 255, 255, 0.1);
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        border-radius: 12px;
        overflow: hidden;
    }

    /* Progress bar */
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, rgba(255, 255, 255, 0.8) 0%, rgba(255, 255, 255, 0.95) 100%);
        border-radius: 10px;
    }

    /* Expanders */
    [data-testid="stExpander"] {
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 12px;
        background-color: rgba(255, 255, 255, 0.02);
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.2);
    }

    /* Dividers */
    hr {
        border-color: rgba(255, 255, 255, 0.1);
        margin: 3rem 0;
        opacity: 0.5;
    }

    /* Info boxes */
    .stAlert {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        border-radius: 12px;
        border: 1px solid rgba(255, 255, 255, 0.15);
        backdrop-filter: blur(10px);
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: rgba(0, 0, 0, 0.6);
        border-right: 1px solid rgba(255, 255, 255, 0.1);
        backdrop-filter: blur(20px);
    }

    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
        color: rgba(255, 255, 255, 0.9);
    }

    /* Section containers */
    .section-container {
        background: rgba(255, 255, 255, 0.02);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 20px;
        padding: 2.5rem;
        margin: 2rem 0;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
        backdrop-filter: blur(10px);
        transition: all 0.3s ease;
    }

    .section-container:hover {
        box-shadow: 0 12px 48px rgba(0, 0, 0, 0.4);
        border-color: rgba(255, 255, 255, 0.15);
    }

    /* Status indicator */
    .status-box {
        padding: 1rem 1.5rem;
        border-left: 2px solid rgba(255, 255, 255, 0.3);
        background-color: rgba(255, 255, 255, 0.03);
        border-radius: 8px;
        margin: 0.75rem 0;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        font-weight: 300;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.2);
    }

    /* xAI Logo styling */
    .xai-logo-header {
        position: relative;
        padding: 1rem 0 2rem 0;
        margin-bottom: 1rem;
        display: inline-block;
    }

    .xai-logo-img {
        height: 60px;
        width: auto;
        background-color: #FFFFFF;
        padding: 12px 20px;
        border-radius: 12px;
        box-shadow: 0 4px 16px rgba(255, 255, 255, 0.1);
    }

    /* Checkboxes */
    [data-testid="stCheckbox"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        font-weight: 400;
    }

    /* Success/Info messages */
    .stSuccess, .stInfo {
        border-radius: 12px;
        backdrop-filter: blur(10px);
    }

    /* Sticky results container */
    .sticky-results {
        position: sticky;
        top: 0;
        z-index: 1000;
        background: rgba(0, 0, 0, 0.98);
        backdrop-filter: blur(20px);
        box-shadow: 0 4px 24px rgba(0, 0, 0, 0.4);
        border-bottom: 2px solid rgba(255, 255, 255, 0.2);
        margin-bottom: 2rem;
        padding: 1.5rem;
        border-radius: 16px;
    }

    /* Compact table styling */
    .compact-table {
        max-height: 400px;
        overflow-y: auto;
    }

    /* Download button emphasis */
    [data-testid="stDownloadButton"] button {
        font-size: 1.1rem !important;
        font-weight: 600 !important;
        padding: 1rem 1.5rem !important;
        background: linear-gradient(135deg, rgba(255, 255, 255, 0.15) 0%, rgba(255, 255, 255, 0.25) 100%) !important;
        border: 2px solid rgba(255, 255, 255, 0.5) !important;
    }

    [data-testid="stDownloadButton"] button:hover {
        background: linear-gradient(135deg, rgba(255, 255, 255, 0.25) 0%, rgba(255, 255, 255, 0.35) 100%) !important;
        border-color: rgba(255, 255, 255, 0.8) !important;
        transform: translateY(-3px) !important;
        box-shadow: 0 12px 32px rgba(255, 255, 255, 0.2) !important;
    }
</style>
""", unsafe_allow_html=True)

# Session state initialization
for key, default in [
    ('processing_complete', False),
    ('results', []),
    ('total_files', 0),
    ('total_replacements', 0),
    ('total_images', 0),
    ('originals_zip_data', None),
    ('pdf_zip_data', None),
    ('timestamp', None),
    ('processing_logs', [])  # Store detailed logs
]:
    if key not in st.session_state:
        st.session_state[key] = default

# xAI Logo and Header
import base64
from pathlib import Path

# Use relative path to work on both local and Streamlit Cloud
logo_path = Path(__file__).parent / "xai_logo.png"
with open(logo_path, "rb") as f:
    logo_data = base64.b64encode(f.read()).decode()
    st.markdown(f"""
    <div class="xai-logo-header">
        <img src="data:image/png;base64,{logo_data}" alt="xAI" class="xai-logo-img">
    </div>
    """, unsafe_allow_html=True)

# Header with version indicator
col1, col2 = st.columns([3, 1])
with col1:
    st.title("DOCX ANONYMIZER")
    st.caption("PROFESSIONAL DOCUMENT ANONYMIZATION SYSTEM")
with col2:
    st.markdown("""
    <div style='text-align: right; padding-top: 1rem;'>
        <p style='font-size: 0.7rem; color: rgba(255, 255, 255, 0.4); margin: 0;'>
            v1.7 - Detailed Tracking<br>
            <span style='font-size: 0.65rem;'>Updated: Nov 18, 2025</span>
        </p>
    </div>
    """, unsafe_allow_html=True)

# Sidebar configuration
with st.sidebar:
    st.header("SYSTEM STATUS")

    if docx_files := st.session_state.get('docx_files_uploaded'):
        st.metric("FILES QUEUED", len(docx_files))
    else:
        st.metric("FILES QUEUED", 0)

    if st.session_state.get('excel_loaded'):
        st.success("✓ MAPPINGS LOADED")
    else:
        st.warning("○ AWAITING MAPPINGS")

    st.divider()
    st.markdown("### OPERATION GUIDE")
    st.caption("""
    **STEP 1** → Upload source documents
    **STEP 2** → Upload Excel mappings
    **STEP 3** → Configure options
    **STEP 4** → Execute anonymization
    **STEP 5** → Download results
    """)

    st.divider()
    st.markdown("### TECHNICAL SPECS")
    st.caption("""
    **Format Support:** Word, Excel, PowerPoint
    **Output:** Original Format + PDF
    **Max File Size:** 200MB
    **Batch Processing:** Mixed formats supported
    **PDF Engine:** LibreOffice
    """)

# Main interface
st.markdown('<div class="section-container">', unsafe_allow_html=True)
st.markdown("### INPUT CONFIGURATION")

col1, col2 = st.columns([3, 2])

with col1:
    st.markdown("#### SOURCE DOCUMENTS")
    docx_files = st.file_uploader(
        "Upload Documents (Word, Excel, PowerPoint)",
        type=['docx', 'doc', 'pptx', 'ppt', 'xlsx', 'xls'],
        accept_multiple_files=True,
        key="docx_upload",
        help="Supports batch processing: Word, PowerPoint, Excel"
    )
    if docx_files:
        st.session_state.docx_files_uploaded = docx_files
        st.success(f"✓ {len(docx_files)} file(s) loaded")

with col2:
    st.markdown("#### ANONYMIZATION MAPPINGS")
    excel_file = st.file_uploader(
        "Upload Excel requirements",
        type=['xlsx'],
        key="excel_upload",
        help="Column 1: Before | Column 2: After"
    )
    if excel_file:
        st.session_state.excel_loaded = True
        st.success("✓ Mappings ready")
st.markdown('</div>', unsafe_allow_html=True)

st.divider()

# Processing options
st.markdown('<div class="section-container">', unsafe_allow_html=True)
st.markdown("### PROCESSING OPTIONS")
col1, col2 = st.columns(2)

with col1:
    remove_images = st.checkbox(
        "REMOVE ALL IMAGES",
        value=True,
        key="remove_images",
        help="Strips all embedded images from documents"
    )

with col2:
    clear_headers_footers = st.checkbox(
        "CLEAR HEADERS/FOOTERS",
        value=False,
        key="clear_headers_footers",
        help="Removes logo and text from headers/footers"
    )
st.markdown('</div>', unsafe_allow_html=True)

st.divider()

# Execute button
st.markdown('<div class="section-container">', unsafe_allow_html=True)
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    execute_btn = st.button(
        "EXECUTE ANONYMIZATION",
        type="primary",
        disabled=(not docx_files or not excel_file),
        width='stretch'
    )
st.markdown('</div>', unsafe_allow_html=True)

if execute_btn:
    # Reset state
    st.session_state.processing_complete = False
    st.session_state.results = []
    st.session_state.originals_zip_data = None
    st.session_state.pdf_zip_data = None

    # Validate LibreOffice
    with st.spinner("Validating PDF conversion engine..."):
        try:
            result = subprocess.run(['soffice', '--version'], capture_output=True, timeout=5)
            if result.returncode != 0:
                st.error("❌ LibreOffice not found")
                st.info("Install: `sudo apt-get install libreoffice`")
                st.stop()
        except (FileNotFoundError, Exception) as e:
            st.error(f"❌ PDF engine error: {e}")
            st.stop()

    # Validate files
    if not docx_files or not excel_file:
        st.error("❌ Missing required files")
        st.stop()

    # Processing pipeline
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        input_dir = temp_path / "input"
        originals_output_dir = temp_path / "originals_output"  # Preserves format
        pdf_output_dir = temp_path / "pdf_output"

        input_dir.mkdir()
        originals_output_dir.mkdir()
        pdf_output_dir.mkdir()

        # Save Excel requirements
        excel_path = temp_path / "requirements.xlsx"
        with open(excel_path, 'wb') as f:
            f.write(excel_file.getbuffer())

        # Save input files and determine type
        files_to_process = []
        for uploaded_file in docx_files:
            # SECURITY: Sanitize filename to prevent path traversal
            from pathlib import Path
            safe_filename = Path(uploaded_file.name).name  # Strips any directory components
            file_path = input_dir / safe_filename
            with open(file_path, 'wb') as f:
                f.write(uploaded_file.getbuffer())

            # Detect file type by extension
            file_ext = file_path.suffix.lower()

            # Convert legacy formats to modern ones
            if file_ext == '.doc':
                # Convert .doc to .docx
                with st.spinner(f"Converting {safe_filename} to DOCX..."):
                    try:
                        cmd = [
                            'soffice', '--headless', '--norestore', '--nologo',
                            '--nofirststartwizard', '--convert-to', 'docx',
                            '--outdir', str(input_dir), str(file_path)
                        ]
                        subprocess.run(cmd, capture_output=True, text=True, timeout=120)
                        converted_path = file_path.with_suffix('.docx')
                        if converted_path.exists():
                            files_to_process.append((safe_filename, converted_path, 'word', '.docx'))
                        else:
                            st.error(f"Conversion failed: {safe_filename}")
                    except Exception as e:
                        st.error(f"Conversion error: {e}")

            elif file_ext == '.ppt':
                # Convert .ppt to .pptx
                with st.spinner(f"Converting {safe_filename} to PPTX..."):
                    try:
                        cmd = [
                            'soffice', '--headless', '--norestore', '--nologo',
                            '--nofirststartwizard', '--convert-to', 'pptx',
                            '--outdir', str(input_dir), str(file_path)
                        ]
                        subprocess.run(cmd, capture_output=True, text=True, timeout=120)
                        converted_path = file_path.with_suffix('.pptx')
                        if converted_path.exists():
                            files_to_process.append((safe_filename, converted_path, 'powerpoint', '.pptx'))
                        else:
                            st.error(f"Conversion failed: {safe_filename}")
                    except Exception as e:
                        st.error(f"Conversion error: {e}")

            elif file_ext == '.xls':
                # Convert .xls to .xlsx (LibreOffice can do this)
                with st.spinner(f"Converting {safe_filename} to XLSX..."):
                    try:
                        cmd = [
                            'soffice', '--headless', '--norestore', '--nologo',
                            '--nofirststartwizard', '--convert-to', 'xlsx',
                            '--outdir', str(input_dir), str(file_path)
                        ]
                        subprocess.run(cmd, capture_output=True, text=True, timeout=120)
                        converted_path = file_path.with_suffix('.xlsx')
                        if converted_path.exists():
                            files_to_process.append((safe_filename, converted_path, 'excel', '.xlsx'))
                        else:
                            st.error(f"Conversion failed: {safe_filename}")
                    except Exception as e:
                        st.error(f"Conversion error: {e}")

            elif file_ext == '.docx':
                files_to_process.append((safe_filename, file_path, 'word', '.docx'))
            elif file_ext == '.pptx':
                files_to_process.append((safe_filename, file_path, 'powerpoint', '.pptx'))
            elif file_ext == '.xlsx':
                files_to_process.append((safe_filename, file_path, 'excel', '.xlsx'))
            elif file_ext == '.xlsm':
                files_to_process.append((safe_filename, file_path, 'excel', '.xlsm'))
            else:
                st.warning(f"Unsupported file type: {safe_filename}")

        # Check if any files were successfully prepared
        if not files_to_process:
            st.error("❌ No files could be processed. Check file formats and conversion errors above.")
            st.stop()

        # Load mappings and precompile patterns
        with st.spinner("Loading anonymization mappings..."):
            try:
                alias_map = load_aliases_from_excel(excel_path)
                sorted_keys = categorize_and_sort_aliases(alias_map)
                compiled_patterns = precompile_patterns(alias_map)  # Precompile for performance
                st.success(f"✓ {len(alias_map)} mappings loaded")
            except Exception as e:
                st.error(f"Mapping error: {e}")
                st.stop()

        st.divider()

        # Create a container for the entire processing section that we can clear later
        processing_container = st.container()

        with processing_container:
            st.markdown('<div class="section-container">', unsafe_allow_html=True)
            st.markdown('<h2 style="margin: 0 0 1rem 0; font-size: 1.4rem; font-weight: 500; letter-spacing: 0.05em;">PROCESSING FILES...</h2>', unsafe_allow_html=True)

            # Simple progress indicators using empty containers for in-place updates
            progress_bar = st.progress(0)
            status_text = st.empty()

            # Create empty containers for metrics
            metrics_cols = st.columns(4)
            metric_containers = []
            for col in metrics_cols:
                with col:
                    metric_containers.append(st.empty())

            # Initialize display
            metric_containers[0].metric("FILES", f"0/{len(files_to_process)}")
            metric_containers[1].metric("REPLACEMENTS", "0")
            metric_containers[2].metric("IMAGES REMOVED", "0")
            metric_containers[3].metric("PDF STATUS", "⏳")

            st.markdown('</div>', unsafe_allow_html=True)

        # Initialize counters and logs
        total_replacements = 0
        total_images = 0
        results = []
        replacement_details = []  # NEW: Track what was actually replaced
        st.session_state.processing_logs = []

        logger = logging.getLogger(__name__)

        # Process files with compact updates
        for i, (original_name, input_path, file_type, output_ext) in enumerate(files_to_process):
            # Update status line
            status_text.text(f"Processing: {original_name}")

            # Output path preserves original format
            output_filename = Path(original_name).stem + output_ext
            original_output_path = originals_output_dir / output_filename

            log_entry = {
                'filename': original_name,
                'file_type': file_type,
                'status': 'processing',
                'details': []
            }

            try:
                # Route to appropriate processor based on file type (with detailed tracking)
                if file_type == 'word':
                    replacements, images, details = process_single_docx(
                        input_path, original_output_path, alias_map, sorted_keys, logger,
                        remove_images=remove_images,
                        clear_headers_footers_flag=clear_headers_footers,
                        track_details=True  # NEW: Enable detailed replacement tracking
                    )
                    log_entry['details'].append(f"Word: {replacements} replacements, {images} images removed")

                    # Store replacement details for this file
                    if details:
                        for original, count in details.items():
                            replacement_details.append({
                                'File': original_name,
                                'Original': original,
                                'Replacement': alias_map.get(original, alias_map.get(original.lower(), '?')),
                                'Count': count
                            })

                elif file_type == 'powerpoint':
                    replacements, images, details = process_single_pptx(
                        input_path, original_output_path, alias_map, sorted_keys,
                        compiled_patterns, logger, remove_images=remove_images,
                        track_details=True  # NEW: Enable detailed replacement tracking
                    )
                    log_entry['details'].append(f"PowerPoint: {replacements} replacements, {images} images removed")

                    # Store replacement details for this file
                    if details:
                        for original, count in details.items():
                            replacement_details.append({
                                'File': original_name,
                                'Original': original,
                                'Replacement': alias_map.get(original, alias_map.get(original.lower(), '?')),
                                'Count': count
                            })

                elif file_type == 'excel':
                    replacements, images, details = process_single_xlsx(
                        input_path, original_output_path, alias_map, sorted_keys,
                        compiled_patterns, logger, remove_images=False,
                        track_details=True  # NEW: Enable detailed replacement tracking
                    )
                    log_entry['details'].append(f"Excel: {replacements} replacements")

                    # Store replacement details for this file
                    if details:
                        for original, count in details.items():
                            replacement_details.append({
                                'File': original_name,
                                'Original': original,
                                'Replacement': alias_map.get(original, alias_map.get(original.lower(), '?')),
                                'Count': count
                            })

                else:
                    raise ValueError(f"Unsupported file type: {file_type}")

                total_replacements += replacements
                total_images += images

                # Convert to PDF (works for all file types via LibreOffice)
                pdf_output_path = pdf_output_dir / Path(original_name).with_suffix('.pdf').name

                try:
                    cmd = [
                        'soffice', '--headless', '--norestore', '--nologo',
                        '--nofirststartwizard', '--convert-to', 'pdf',
                        '--outdir', str(pdf_output_dir), str(original_output_path)
                    ]
                    subprocess.run(cmd, capture_output=True, text=True, timeout=300)

                    expected_output = pdf_output_dir / f"{original_output_path.stem}.pdf"

                    if expected_output.exists():
                        if expected_output != pdf_output_path:
                            shutil.move(str(expected_output), str(pdf_output_path))

                        size_kb = pdf_output_path.stat().st_size / 1024
                        log_entry['details'].append(f"PDF: Success ({size_kb:.0f} KB)")
                        log_entry['status'] = 'success'

                        results.append({
                            'filename': original_name,
                            'file_type': file_type.capitalize(),
                            'replacements': replacements,
                            'images': images,
                            'pdf_status': '✓ Success',
                            'pdf_size_kb': round(size_kb)
                        })
                    else:
                        log_entry['details'].append("PDF: Conversion failed")
                        log_entry['status'] = 'warning'
                        results.append({
                            'filename': original_name,
                            'file_type': file_type.capitalize(),
                            'replacements': replacements,
                            'images': images,
                            'pdf_status': '✗ Failed',
                            'pdf_size_kb': 0
                        })

                except subprocess.TimeoutExpired:
                    log_entry['details'].append("PDF: Timeout (5min exceeded)")
                    log_entry['status'] = 'warning'
                    results.append({
                        'filename': original_name,
                        'file_type': file_type.capitalize(),
                        'replacements': replacements,
                        'images': images,
                        'pdf_status': '⚠ Timeout',
                        'pdf_size_kb': 0
                    })
                except Exception as e:
                    log_entry['details'].append(f"PDF: Error - {str(e)[:100]}")
                    log_entry['status'] = 'warning'
                    results.append({
                        'filename': original_name,
                        'file_type': file_type.capitalize(),
                        'replacements': replacements,
                        'images': images,
                        'pdf_status': '✗ Error',
                        'pdf_size_kb': 0
                    })

            except Exception as e:
                log_entry['details'].append(f"Processing Error: {str(e)[:100]}")
                log_entry['status'] = 'error'
                results.append({
                    'filename': original_name,
                    'file_type': file_type.capitalize(),
                    'replacements': 0,
                    'images': 0,
                    'pdf_status': f'✗ {file_type.capitalize()} Error',
                    'pdf_size_kb': 0
                })

            # Store log entry
            st.session_state.processing_logs.append(log_entry)

            # Update progress and metrics IN PLACE
            progress_bar.progress((i + 1) / len(files_to_process))

            # Update metric containers instead of creating new metrics
            metric_containers[0].metric("FILES", f"{i+1}/{len(files_to_process)}")
            metric_containers[1].metric("REPLACEMENTS", f"{total_replacements:,}")
            metric_containers[2].metric("IMAGES REMOVED", f"{total_images:,}")
            pdf_success = sum(1 for r in results if '✓' in r.get('pdf_status', ''))
            metric_containers[3].metric("PDF SUCCESS", f"{pdf_success}/{i+1}")

        # Clear status and show completion
        status_text.success("✓ Processing Complete!")

        # Save results to session state
        st.session_state.results = results
        st.session_state.replacement_details = replacement_details  # NEW: Store detailed replacements
        st.session_state.total_files = len(files_to_process)
        st.session_state.total_replacements = total_replacements
        st.session_state.total_images = total_images
        st.session_state.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # Create ZIP archives
        timestamp = st.session_state.timestamp

        # ZIP 1: Original formats (preserves .docx, .pptx, .xlsx)
        originals_zip_path = temp_path / f"anonymized_originals_{timestamp}.zip"
        with zipfile.ZipFile(originals_zip_path, 'w') as zipf:
            # Add all files from originals_output_dir (mixed formats)
            for file in originals_output_dir.glob('*'):
                if file.is_file():
                    zipf.write(file, file.name)

        with open(originals_zip_path, 'rb') as f:
            st.session_state.originals_zip_data = f.read()

        # ZIP 2: PDFs (all files converted to PDF)
        pdf_zip_path = temp_path / f"anonymized_pdf_{timestamp}.zip"
        with zipfile.ZipFile(pdf_zip_path, 'w') as zipf:
            for pdf_file in pdf_output_dir.glob('*.pdf'):
                zipf.write(pdf_file, pdf_file.name)

        with open(pdf_zip_path, 'rb') as f:
            st.session_state.pdf_zip_data = f.read()

        st.session_state.processing_complete = True
        st.rerun()  # Reload page to show only results, hiding processing section

# Results display
if st.session_state.processing_complete:
    st.divider()

    # Sticky results container with prominent downloads
    st.markdown('<div class="sticky-results">', unsafe_allow_html=True)

    # Success header
    st.markdown('''
        <div style="text-align: center; margin-bottom: 1.5rem;">
            <h2 style="margin: 0 0 0.5rem 0; font-size: 2rem; color: #4ade80;">✓ PROCESSING COMPLETE</h2>
            <p style="color: rgba(255,255,255,0.7); font-size: 1.1rem; margin: 0;">
                Your files are ready for download
            </p>
        </div>
    ''', unsafe_allow_html=True)

    # Download buttons row - prominent and centered
    col1, col2, col3 = st.columns([1, 3, 1])

    with col2:
        download_cols = st.columns(2)

        with download_cols[0]:
            if st.session_state.originals_zip_data:
                st.download_button(
                    label="📄 DOWNLOAD ORIGINAL FORMATS",
                    data=st.session_state.originals_zip_data,
                    file_name=f"anonymized_originals_{st.session_state.timestamp}.zip",
                    mime="application/zip",
                    width='stretch',
                    type="primary",
                    help="Anonymized files in original formats (.docx, .xlsx, .pptx)"
                )

        with download_cols[1]:
            if st.session_state.pdf_zip_data:
                st.download_button(
                    label="📑 DOWNLOAD AS PDF",
                    data=st.session_state.pdf_zip_data,
                    file_name=f"anonymized_pdf_{st.session_state.timestamp}.zip",
                    mime="application/zip",
                    width='stretch',
                    type="primary",
                    help="Anonymized files converted to PDF"
                )

    # Summary stats in a compact row
    st.markdown('<div style="margin-top: 2rem; padding-top: 1.5rem; border-top: 1px solid rgba(255,255,255,0.1);">', unsafe_allow_html=True)

    stats_cols = st.columns(5)

    with stats_cols[0]:
        st.metric("FILES", st.session_state.total_files, delta=None)

    with stats_cols[1]:
        st.metric("REPLACEMENTS", f"{st.session_state.total_replacements:,}", delta=None)

    with stats_cols[2]:
        st.metric("IMAGES", st.session_state.total_images, delta="Removed" if st.session_state.total_images > 0 else None)

    with stats_cols[3]:
        pdf_success = sum(1 for r in st.session_state.results if '✓' in r.get('pdf_status', ''))
        st.metric("PDF SUCCESS", f"{pdf_success}/{st.session_state.total_files}", delta=None)

    with stats_cols[4]:
        if st.button("🔄 NEW BATCH", width='stretch'):
            st.session_state.processing_complete = False
            st.session_state.results = []
            st.session_state.originals_zip_data = None
            st.session_state.pdf_zip_data = None
            st.session_state.processing_logs = []
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Detailed results section below (optional viewing)
    st.markdown('<div class="section-container" style="margin-top: 2rem;">', unsafe_allow_html=True)

    # Tabs for different views
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Results Table", "🔍 Replacement Details", "📝 Processing Logs", "ℹ️ File Details"])

    with tab1:
        if st.session_state.results:
            st.dataframe(
                st.session_state.results,
                width='stretch',
                hide_index=True,
                height=400
            )

    with tab2:
        # NEW: Detailed replacement tracking
        if st.session_state.get('replacement_details'):
            st.markdown("### What Was Replaced")
            st.caption(f"Showing {len(st.session_state.replacement_details)} unique replacements across all files")

            # Group by file for cleaner display
            import pandas as pd
            df_replacements = pd.DataFrame(st.session_state.replacement_details)

            # Sort by File, then Count (descending)
            df_replacements = df_replacements.sort_values(['File', 'Count'], ascending=[True, False])

            # Display with nice formatting
            st.dataframe(
                df_replacements,
                width='stretch',
                hide_index=True,
                height=500,
                column_config={
                    'File': st.column_config.TextColumn('File', width='medium'),
                    'Original': st.column_config.TextColumn('Original Text', width='medium'),
                    'Replacement': st.column_config.TextColumn('Anonymized To', width='medium'),
                    'Count': st.column_config.NumberColumn('Times Found', format='%d')
                }
            )

            # Summary stats by file
            st.markdown("---")
            st.markdown("### Summary by File")
            summary = df_replacements.groupby('File').agg({
                'Count': 'sum',
                'Original': 'count'
            }).rename(columns={'Count': 'Total Replacements', 'Original': 'Unique Terms'})
            st.dataframe(summary, width='stretch')

        else:
            st.info("No replacements were made in this batch.")

    with tab3:
        if st.session_state.get('processing_logs'):
            for log in st.session_state.processing_logs:
                status_icon = "✓" if log['status'] == 'success' else "⚠" if log['status'] == 'warning' else "❌"
                with st.expander(f"{status_icon} {log['filename']}", expanded=False):
                    for detail in log['details']:
                        st.text(detail)

    with tab4:
        # Show individual file sizes and details
        for result in st.session_state.results:
            col1, col2 = st.columns([3, 1])
            with col1:
                st.text(f"📄 {result['filename']}")
            with col2:
                if result.get('pdf_size_kb', 0) > 0:
                    st.text(f"{result['pdf_size_kb']} KB")

    st.markdown('</div>', unsafe_allow_html=True)
