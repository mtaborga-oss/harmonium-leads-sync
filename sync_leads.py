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
            raise ValueError("No credentials")
        return build('gmail', 'v1', credentials=creds)
    
    def _authenticate_drive(self):
        creds = self._load_credentials_from_json()
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        if not creds:
            raise ValueError("No credentials")
        return build('drive', 'v3', credentials=creds)
    
    def load_processed_emails(self):
        log_file = Path('.processed_emails.json')
        if log_file.exists():
            with open(log_file, 'r') as f:
                data = json.load(f)
                self.processed_emails = set(data.get('ids', []))
    
    def save_processed_emails(self):
        with open('.processed_emails.json', 'w') as f:
            json.dump({'ids': list(self.processed_emails)}, f)
    
    def fetch_new_emails(self):
        try:
            results = self.gmail_service.users().messages().list(
                userId='me',
                q='from:no-reply@harmonium.design is:unread',
                maxResults=50
            ).execute()
            return results.get('messages', [])
        except Exception as e:
            print(f"Error: {e}")
            return []
    
    def extract_lead_data(self, message_id):
        try:
            msg = self.gmail_service.users().messages().get(userId='me', id=message_id, format='full').execute()
            headers = msg['payload']['headers']
            date_str = next(h['value'] for h in headers if h['name'] == 'Date')
            body_data = msg['payload'].get('parts', [{}])[0].get('body', {})
            body_text = body_data.get('data', '')
            if not body_text and 'parts' in msg['payload']:
                for part in msg['payload']['parts']:
                    if 'body' in part:
                        body_text = part['body'].get('data', '')
                        break
            if body_text:
                body_text = base64.urlsafe_b64decode(body_text).decode('utf-8')
            lead = self._parse_lead_text(body_text)
            lead['fecha'] = self._parse_date(date_str)
            lead['email_id'] = message_id
            return lead
        except Exception as e:
            print(f"Error: {e}")
            return None
    
    def _parse_lead_text(self, text):
        lead = {'nombre': '', 'email': '', 'telefono': '', 'tipo_cliente': '', 'ubicacion': '', 'mensaje': ''}
        patterns = {
            'nombre': r'Nombre:\s*(.+?)(?:\n|$)',
            'email': r'Email:\s*(.+?)(?:\n|$)',
            'telefono': r'Teléfono:\s*(.+?)(?:\n|$)',
            'tipo_cliente': r'Tipo de cliente:\s*(.+?)(?:\n|$)',
            'ubicacion': r'Ubicación:\s*(.+?)(?:\n|$)',
            'mensaje': r'Mensaje:\s*(.+?)(?:\n\n|$)'
        }
        for key, pattern in patterns.items():
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                lead[key] = match.group(1).strip()
        return lead
    
    def _parse_date(self, date_str):
        try:
            from email.utils import parsedate_to_datetime
            dt = parsedate_to_datetime(date_str)
            return dt.strftime('%Y-%m-%d')
        except:
            return datetime.now().strftime('%Y-%m-%d')
    
    def find_file_in_drive(self, filename):
        try:
            results = self.drive_service.files().list(q=f"name='{filename}' and trashed=false", spaces='drive', fields='files(id, name)', pageSize=1).execute()
            files = results.get('files', [])
            return files[0]['id'] if files else None
        except Exception as e:
            print(f"Error: {e}")
            return None
    
    def download_excel_from_drive(self, file_id, local_path):
        try:
            request = self.drive_service.files().get_media(fileId=file_id)
            with open(local_path, 'wb') as f:
                f.write(request.execute())
            print("✓ Excel descargado")
        except Exception as e:
            print(f"Error: {e}")
    
    def upload_excel_to_drive(self, file_id, local_path):
        try:
            from googleapiclient.http import MediaFileUpload
            media = MediaFileUpload(local_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.drive_service.files().update(fileId=file_id, media_body=media).execute()
            print("✓ Excel actualizado")
        except Exception as e:
            print(f"Error: {e}")
    
    def update_excel(self, leads, excel_path):
        if not leads:
            return 0
        if not Path(excel_path).exists():
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Leads"
            headers = ['Fecha lead', 'Nombre', 'Email', 'Teléfono', 'Tipo de cliente', 'Ubicación', 'Interés / Producto', 'Mensaje', 'URL página', 'Estado', 'Fecha registro', 'ID Gmail', 'GCLID', 'Estado comercial', 'Fecha conversión', 'Valor €']
            ws.append(headers)
            for cell in ws[1]:
                cell.fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
                cell.font = Font(color='FFFFFF', bold=True)
            wb.save(excel_path)
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        count = 0
        for lead in leads:
            if lead['email_id'] in self.processed_emails:
                continue
            row = [lead.get('fecha'), lead.get('nombre'), lead.get('email'), lead.get('telefono'), lead.get('tipo_cliente'), lead.get('ubicacion'), lead.get('mensaje'), '', '', '', '', lead.get('email_id'), '', 'Nuevo', '', '']
            ws.append(row)
            self.processed_emails.add(lead['email_id'])
            count += 1
        if count > 0:
            wb.save(excel_path)
            print(f"✓ {count} leads agregados")
        return count
    
    def mark_emails_read(self, message_ids):
        try:
            for msg_id in message_ids:
                self.gmail_service.users().messages().modify(userId='me', id=msg_id, body={'removeLabelIds': ['UNREAD']}).execute()
        except Exception as e:
            print(f"Error: {e}")
    
    def sync(self):
        print(f"\n🔄 Sincronización: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        excel_filename = 'Leads_Harmonium_marcado_amarillo.xlsx'
        file_id = self.find_file_in_drive(excel_filename)
        if file_id:
            self.download_excel_from_drive(file_id, excel_filename)
        messages = self.fetch_new_emails()
        if not messages:
            print("✓ Sin emails nuevos")
            return 0
        print(f"📧 {len(messages)} emails encontrados")
        leads = []
        msg_ids = []
        for msg in messages:
            lead = self.extract_lead_data(msg['id'])
            if lead:
                leads.append(lead)
                msg_ids.append(msg['id'])
        count = self.update_excel(leads, excel_filename)
        if count > 0 and file_id:
            self.upload_excel_to_drive(file_id, excel_filename)
            self.mark_emails_read(msg_ids)
            self.save_processed_emails()
        return count

if __name__ == "__main__":
    sync = HarmoniumLeadsSync()
    sync.sync()
