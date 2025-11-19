import sys
from pathlib import Path

# Ensure parent directory (backend/) is on sys.path so we can import app.py
sys.path.append(str(Path(__file__).resolve().parents[1]))

from app import app  # FastAPI instance

# Vercel serverless function handler
handler = app
