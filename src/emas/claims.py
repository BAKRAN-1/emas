"""Claim lookup with an externally configured endpoint and API key."""
import pandas as pd
import requests
from .config import claims_settings

def fetch_claim_data(claim_number):
    if pd.isna(claim_number):
        return None
    base, key = claims_settings()
    response = requests.get(base + '/Billrecord', headers={'X-Api-Key': key}, params={'select': 'claimnumber', 'maxSize': 5, 'offset': 0, 'where[0][type]': 'equals', 'where[0][attribute]': 'claimnumber', 'where[0][value]': str(claim_number)}, timeout=30)
    response.raise_for_status()
    records = response.json().get('list', [])
    return records[0]['id'] if records else None
