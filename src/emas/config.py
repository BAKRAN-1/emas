"""Paths and credentials supplied locally through environment variables."""
import os
from pathlib import Path
PROJECT_ROOT = Path(__file__).resolve().parents[2]

def database_config_path():
    return Path(os.environ.get('EMAS_DB_CONFIG', str(PROJECT_ROOT / 'config' / 'db.json')))

def claims_settings():
    base = os.environ.get('EMAS_API_BASE_URL', '').rstrip('/')
    key = os.environ.get('EMAS_API_KEY', '')
    if not base or not key:
        raise RuntimeError('Set EMAS_API_BASE_URL and EMAS_API_KEY before extracting claim data.')
    return (base, key)
