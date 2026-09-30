"""
app.py - Main entry point for the REPOTIC Automation Web Application.
Compatible with local execution and serverless deployments (e.g. Vercel, AWS Lambda).
"""

import sys
from pathlib import Path

# Ensure project root is in sys.path
_ROOT = Path(__file__).resolve().parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

# Marketplace prefix configuration (editable dict)
from src.config import PREFIX
from src.web.routes import app

if __name__ == "__main__":
    app.run(debug=True, port=5000)
