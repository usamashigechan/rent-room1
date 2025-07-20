import os
import pickle
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

class GmailSender:
    def __init__(self):
        self.SCOPES = ['https://www.googleapis.com/auth/gmail.send']
        self.service = self.authenticate()
    
    def authenticate(self):
        creds = None
        
        # 既存トークン読み込み
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        
        # 認証が必要な場合
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                # 初回認証
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', self.SCOPES)
                creds = flow.run_local_server(port=0)
            
            # トークン保存
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
        
        return build('gmail', 'v1', credentials=creds)
    
    def send_email(self, to_email, subject, body):
        """メール送信"""
        try:
            message = MIMEMultipart()
            message['to'] = to_email
            message['subject'] = subject
            message.attach(MIMEText(body, 'plain', 'utf-8'))
            
            # base64エンコード
            raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
            
            # 送信
            result = self.service.users().messages().send(
                userId='me', body={'raw': raw_message}).execute()
            
            print(f"メール送信成功: {to_email}")
            print(f"メッセージID: {result['id']}")
            return True
            
        except Exception as e:
            print(f"エラー: {e}")
            return False