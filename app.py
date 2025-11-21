#!/usr/bin/env python3
"""
Streamlit Cloud Entrypoint (app.py)
This file serves as the entrypoint for Streamlit Cloud deployment.
The actual application code is in src/streamlit_app.py
"""

# Add src to path for imports
import sys
from pathlib import Path

# Add both root and src to path
root_dir = Path(__file__).parent
sys.path.insert(0, str(root_dir))
sys.path.insert(0, str(root_dir / "src"))

# Import and run the actual app using importlib to avoid caching issues
import importlib.util
spec = importlib.util.spec_from_file_location("streamlit_app", root_dir / "src" / "streamlit_app.py")
streamlit_app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(streamlit_app)
