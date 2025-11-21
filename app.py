#!/usr/bin/env python3
"""
Streamlit Cloud Entrypoint (app.py)
This file serves as the entrypoint for Streamlit Cloud deployment.
The actual application code is in src/streamlit_app.py
"""

# Add src to path for imports
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

# Import and run the actual app
import src.streamlit_app
