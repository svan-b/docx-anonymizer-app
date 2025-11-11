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
    process_single_docx
)
import logging

# Page configuration
st.set_page_config(
    page_title="DOCX Anonymizer",
    page_icon="🔒",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS - xAI Black & White Aesthetic
st.markdown("""
<style>
    /* Main container */
    .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
        max-width: 1400px;
    }

    /* Headers */
    h1 {
        font-family: sans-serif;
        font-weight: 700;
        letter-spacing: -0.02em;
        border-bottom: 2px solid #FFFFFF;
        padding-bottom: 0.5rem;
        margin-bottom: 1.5rem;
        color: #FFFFFF;
    }

    h2, h3 {
        font-family: sans-serif;
        font-weight: 600;
        letter-spacing: -0.01em;
        color: #FFFFFF;
    }

    /* Metrics */
    [data-testid="stMetricValue"] {
        font-size: 2rem;
        font-family: sans-serif;
        font-weight: 700;
        color: #FFFFFF;
    }

    [data-testid="stMetricLabel"] {
        font-family: sans-serif;
        font-size: 0.9rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #CCCCCC;
    }

    /* Upload boxes */
    [data-testid="stFileUploader"] {
        border: 2px solid #FFFFFF;
        border-radius: 6px;
        padding: 1.25rem;
        background: linear-gradient(135deg, rgba(255, 255, 255, 0.05) 0%, rgba(26, 26, 26, 0.95) 100%);
        box-shadow: 0 0 15px rgba(255, 255, 255, 0.1);
        transition: all 0.3s ease;
    }

    [data-testid="stFileUploader"]:hover {
        border-color: #FFFFFF;
        box-shadow: 0 0 20px rgba(255, 255, 255, 0.2);
        background: linear-gradient(135deg, rgba(255, 255, 255, 0.08) 0%, rgba(26, 26, 26, 0.95) 100%);
    }

    /* Upload button styling */
    [data-testid="stFileUploader"] section button {
        background-color: rgba(255, 255, 255, 0.1) !important;
        border: 1px solid #FFFFFF !important;
        color: #FFFFFF !important;
        font-family: sans-serif !important;
        font-weight: 600 !important;
    }

    [data-testid="stFileUploader"] section button:hover {
        background-color: rgba(255, 255, 255, 0.2) !important;
        border-color: #FFFFFF !important;
    }

    /* Buttons */
    .stButton>button {
        font-family: sans-serif;
        font-weight: 600;
        letter-spacing: 0.02em;
        border-radius: 4px;
        border: 1px solid #FFFFFF;
        transition: all 0.2s;
    }

    .stButton>button:hover {
        border-color: #FFFFFF;
        box-shadow: 0 0 10px rgba(255, 255, 255, 0.3);
    }

    /* Data tables */
    [data-testid="stDataFrame"] {
        border: 1px solid #333333;
        font-family: sans-serif;
    }

    /* Progress bar */
    .stProgress > div > div > div > div {
        background-color: #FFFFFF;
    }

    /* Expanders */
    [data-testid="stExpander"] {
        border: 1px solid #333333;
        border-radius: 4px;
        background-color: #1A1A1A;
    }

    /* Dividers */
    hr {
        border-color: #333333;
        margin: 2rem 0;
    }

    /* Info boxes */
    .stAlert {
        font-family: sans-serif;
        border-radius: 4px;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, rgba(0, 0, 0, 0.95) 0%, rgba(255, 255, 255, 0.05) 100%);
        border-right: 2px solid rgba(255, 255, 255, 0.2);
        backdrop-filter: blur(10px);
    }

    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
        color: #FFFFFF;
    }

    /* Section containers */
    .section-container {
        background: linear-gradient(135deg, rgba(26, 26, 26, 0.6) 0%, rgba(0, 0, 0, 0.8) 100%);
        border: 1px solid #333333;
        border-radius: 8px;
        padding: 1.5rem;
        margin: 1rem 0;
        box-shadow: 0 2px 8px rgba(255, 255, 255, 0.1);
    }

    /* Status indicator */
    .status-box {
        padding: 0.75rem 1rem;
        border-left: 3px solid #FFFFFF;
        background-color: #1A1A1A;
        border-radius: 4px;
        margin: 0.5rem 0;
        font-family: sans-serif;
    }

    /* xAI Logo styling */
    .xai-logo {
        text-align: center;
        padding: 1rem 0 0.5rem 0;
        margin-bottom: 1rem;
    }

    .xai-logo img {
        height: 50px;
        filter: brightness(0) invert(1);
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
    ('docx_zip_data', None),
    ('pdf_zip_data', None),
    ('timestamp', None)
]:
    if key not in st.session_state:
        st.session_state[key] = default

# xAI Logo and Header
st.markdown("""
<div class="xai-logo">
    <img src="https://x.ai/static/images/logo.svg" alt="xAI" onerror="this.onerror=null; this.src='https://www.x.ai/favicon.ico'; this.style.height='40px';">
</div>
""", unsafe_allow_html=True)

st.title("DOCX ANONYMIZER")
st.caption("PROFESSIONAL DOCUMENT ANONYMIZATION SYSTEM")

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
    **Format Support:** DOCX, DOC
    **Output:** DOCX + PDF
    **Max File Size:** 200MB
    **Batch Processing:** Enabled
    **PDF Engine:** LibreOffice
    """)

# Main interface
st.markdown('<div class="section-container">', unsafe_allow_html=True)
st.markdown("### INPUT CONFIGURATION")

col1, col2 = st.columns([3, 2])

with col1:
    st.markdown("#### SOURCE DOCUMENTS")
    docx_files = st.file_uploader(
        "Upload DOCX or DOC files",
        type=['docx', 'doc'],
        accept_multiple_files=True,
        key="docx_upload",
        help="Supports batch processing of multiple files"
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
        "⚡ EXECUTE ANONYMIZATION",
        type="primary",
        disabled=(not docx_files or not excel_file),
        use_container_width=True
    )
st.markdown('</div>', unsafe_allow_html=True)

if execute_btn:
    # Reset state
    st.session_state.processing_complete = False
    st.session_state.results = []
    st.session_state.docx_zip_data = None
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
        docx_output_dir = temp_path / "docx_output"
        pdf_output_dir = temp_path / "pdf_output"

        input_dir.mkdir()
        docx_output_dir.mkdir()
        pdf_output_dir.mkdir()

        # Save Excel requirements
        excel_path = temp_path / "requirements.xlsx"
        with open(excel_path, 'wb') as f:
            f.write(excel_file.getbuffer())

        # Save and convert input files
        files_to_process = []
        for docx_file in docx_files:
            file_path = input_dir / docx_file.name
            with open(file_path, 'wb') as f:
                f.write(docx_file.getbuffer())

            # Convert .doc to .docx if needed
            if docx_file.name.lower().endswith('.doc') and not docx_file.name.lower().endswith('.docx'):
                with st.spinner(f"Converting {docx_file.name} to DOCX..."):
                    try:
                        cmd = [
                            'soffice', '--headless', '--norestore', '--nologo',
                            '--nofirststartwizard', '--convert-to', 'docx',
                            '--outdir', str(input_dir), str(file_path)
                        ]
                        subprocess.run(cmd, capture_output=True, text=True, timeout=120)
                        converted_path = file_path.with_suffix('.docx')
                        if converted_path.exists():
                            files_to_process.append((docx_file.name, converted_path))
                        else:
                            st.error(f"Conversion failed: {docx_file.name}")
                    except Exception as e:
                        st.error(f"Conversion error: {e}")
            else:
                files_to_process.append((docx_file.name, file_path))

        # Load mappings
        with st.spinner("Loading anonymization mappings..."):
            try:
                alias_map = load_aliases_from_excel(excel_path)
                sorted_keys = categorize_and_sort_aliases(alias_map)
                st.success(f"✓ {len(alias_map)} mappings loaded")
            except Exception as e:
                st.error(f"Mapping error: {e}")
                st.stop()

        st.divider()
        st.markdown('<div class="section-container">', unsafe_allow_html=True)
        st.markdown("### PROCESSING PIPELINE")

        total_replacements = 0
        total_images = 0
        results = []

        progress_bar = st.progress(0)
        status_container = st.empty()

        logger = logging.getLogger(__name__)

        for i, (original_name, input_path) in enumerate(files_to_process):
            status_container.markdown(
                f'<div class="status-box">PROCESSING [{i+1}/{len(files_to_process)}]: {original_name}</div>',
                unsafe_allow_html=True
            )

            docx_output_path = docx_output_dir / Path(original_name).with_suffix('.docx').name

            with st.expander(f"📄 {original_name}", expanded=(i == 0)):
                try:
                    # Anonymize DOCX
                    replacements, images = process_single_docx(
                        input_path, docx_output_path, alias_map, sorted_keys, logger,
                        remove_images=remove_images,
                        clear_headers_footers_flag=clear_headers_footers
                    )

                    total_replacements += replacements
                    total_images += images

                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("REPLACEMENTS", replacements)
                    with col2:
                        st.metric("IMAGES REMOVED", images)
                    with col3:
                        st.metric("STATUS", "DOCX ✓")

                    # Convert to PDF
                    pdf_output_path = pdf_output_dir / Path(original_name).with_suffix('.pdf').name

                    try:
                        cmd = [
                            'soffice', '--headless', '--norestore', '--nologo',
                            '--nofirststartwizard', '--convert-to', 'pdf',
                            '--outdir', str(pdf_output_dir), str(docx_output_path)
                        ]
                        subprocess.run(cmd, capture_output=True, text=True, timeout=300)

                        expected_output = pdf_output_dir / f"{docx_output_path.stem}.pdf"

                        if expected_output.exists():
                            if expected_output != pdf_output_path:
                                shutil.move(str(expected_output), str(pdf_output_path))

                            size_mb = pdf_output_path.stat().st_size / (1024 * 1024)
                            st.success(f"✓ PDF created ({size_mb:.1f}MB)")

                            results.append({
                                'filename': original_name,
                                'replacements': replacements,
                                'images': images,
                                'pdf_status': '✓ Success',
                                'pdf_size_mb': round(size_mb, 1)
                            })
                        else:
                            st.warning("⚠ PDF conversion failed")
                            results.append({
                                'filename': original_name,
                                'replacements': replacements,
                                'images': images,
                                'pdf_status': '✗ Failed',
                                'pdf_size_mb': 0
                            })

                    except subprocess.TimeoutExpired:
                        st.warning("⚠ PDF timeout (5min exceeded)")
                        results.append({
                            'filename': original_name,
                            'replacements': replacements,
                            'images': images,
                            'pdf_status': '⚠ Timeout',
                            'pdf_size_mb': 0
                        })
                    except Exception as e:
                        st.warning(f"⚠ PDF error: {e}")
                        results.append({
                            'filename': original_name,
                            'replacements': replacements,
                            'images': images,
                            'pdf_status': '✗ Error',
                            'pdf_size_mb': 0
                        })

                except Exception as e:
                    st.error(f"❌ DOCX error: {e}")
                    results.append({
                        'filename': original_name,
                        'replacements': 0,
                        'images': 0,
                        'pdf_status': '✗ DOCX Error',
                        'pdf_size_mb': 0
                    })

            progress_bar.progress((i + 1) / len(files_to_process))

        status_container.markdown(
            '<div class="status-box">✓ PIPELINE COMPLETE</div>',
            unsafe_allow_html=True
        )
        st.markdown('</div>', unsafe_allow_html=True)

        # Save results to session state
        st.session_state.results = results
        st.session_state.total_files = len(files_to_process)
        st.session_state.total_replacements = total_replacements
        st.session_state.total_images = total_images
        st.session_state.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # Create ZIP archives
        timestamp = st.session_state.timestamp

        docx_zip_path = temp_path / f"anonymized_docx_{timestamp}.zip"
        with zipfile.ZipFile(docx_zip_path, 'w') as zipf:
            for docx_file in docx_output_dir.glob('*.docx'):
                zipf.write(docx_file, docx_file.name)

        with open(docx_zip_path, 'rb') as f:
            st.session_state.docx_zip_data = f.read()

        pdf_zip_path = temp_path / f"anonymized_pdf_{timestamp}.zip"
        with zipfile.ZipFile(pdf_zip_path, 'w') as zipf:
            for pdf_file in pdf_output_dir.glob('*.pdf'):
                zipf.write(pdf_file, pdf_file.name)

        with open(pdf_zip_path, 'rb') as f:
            st.session_state.pdf_zip_data = f.read()

        st.session_state.processing_complete = True

# Results display
if st.session_state.processing_complete:
    st.divider()
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown("### PROCESSING SUMMARY")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("FILES PROCESSED", st.session_state.total_files)
    with col2:
        st.metric("TOTAL REPLACEMENTS", st.session_state.total_replacements)
    with col3:
        st.metric("IMAGES REMOVED", st.session_state.total_images)
    with col4:
        st.metric("BATCH ID", st.session_state.timestamp[-6:])

    # Results table
    st.markdown("#### DETAILED RESULTS")
    st.dataframe(
        st.session_state.results,
        use_container_width=True,
        hide_index=True
    )
    st.markdown('</div>', unsafe_allow_html=True)

    st.divider()
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown("### DOWNLOAD ARCHIVES")

    col1, col2, col3 = st.columns([2, 2, 1])

    with col1:
        if st.session_state.docx_zip_data:
            st.download_button(
                label="📦 DOWNLOAD DOCX ARCHIVE",
                data=st.session_state.docx_zip_data,
                file_name=f"anonymized_docx_{st.session_state.timestamp}.zip",
                mime="application/zip",
                use_container_width=True
            )

    with col2:
        if st.session_state.pdf_zip_data:
            st.download_button(
                label="📦 DOWNLOAD PDF ARCHIVE",
                data=st.session_state.pdf_zip_data,
                file_name=f"anonymized_pdf_{st.session_state.timestamp}.zip",
                mime="application/zip",
                use_container_width=True
            )

    with col3:
        if st.button("🔄 NEW BATCH", use_container_width=True):
            for key in ['processing_complete', 'results', 'docx_zip_data', 'pdf_zip_data',
                       'docx_files_uploaded', 'excel_loaded']:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
