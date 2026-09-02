#!/usr/bin/env python3
import re
import json
import os
from datetime import datetime
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import base64

class HarmoniumLeadsSync:
    def __init__(self):
        self.gmail_service = self._authenticate_gmail()
        self.drive_service = self._authenticate_drive()
        self.processed_emails = set()
        self.load_processed_emails()
    
    def _load_credentials_from_json(self):
        token_file = Path('token.json')
        if token_file.exists():
            with open(token_file, 'r') as f:
                token_data = json.load(f)
                creds = Credentials(
                    token=token_data['token'],
                    refresh_token=token_data.get('refresh_token'),
                    token_uri=token_data.get('token_uri'),
                    client_id=token_data.get('client_id'),
                    client_secret=token_data.get('client_secret'),
                    scopes=token_data.get('scopes')
                )
                return creds
        return None
    
    def _authenticate_gmail(self):
        creds = self._load_credentials_from_json()
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        if not creds:
            raise ValueError("No credentials available")
        return build('gmail', 'v1', credentials=creds)
    
    def _authenticate_drive(self):
        creds = self._load_credentials_from_json()
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        if not creds:
            raise ValueError("No credentials available")
        return build('drive', 'v3',
