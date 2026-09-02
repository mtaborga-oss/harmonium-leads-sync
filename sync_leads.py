#!/usr/bin/env python3
"""
Harmonium Leads Auto-Sync para GitHub Actions
"""

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
    GMAIL_SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 
                    'https://www.googleapis.com/auth/gmail.modify']
    DRIVE_SCOPES = ['https://www.googleapis.com/auth/drive']
    
    def __init__(self):
        self.gmail_service = self._authenticate_gmail()
        self.drive_service = self._authenticate_drive()
        self.processed_emails = set()
        self.load_processed_emails()
    
    def _load_credentials_from_json(self):
        """Carga credenciales desde token.json"""
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
        """Autentica con Gmail API"""
        creds = self._load_credentials_from_json()
        
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        
        if not creds:
            raise ValueError("No credentials available. Run setup locally first.")
        
        return build('gmail', 'v1', credentials=creds)
    
    def _authenticate_drive(self):
        """Autentica con Google Drive API"""
        creds = self._load_credentials_from_json()
        
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        
        if not creds:
            raise ValueError("No credentials available")
        
        return build('drive', 'v3', credentials=creds)
    
    def load_processed_emails(self):
        """Carga los IDs de emails ya procesados"""
        log_file = Path('.processed_emails.json')
        if log_file.exists():
            with open(log_file, 'r') as f:
                data = json.load(f)
                self.processed_emails = set(data.get('ids', []))
    
    def save_processed_emails(self):
        """Guarda los IDs de emails procesados"""
        with open('.processed_emails.json', 'w') as f:
            json.dump({'ids': list(self.processed_emails)}, f)
    
    def fetch_new_emails(self):
        """Obtiene emails sin leer"""
        try:
            results = self.gmail_service.users().messages().list(
                userId='me',
                q='from:no-reply@harmonium.design
