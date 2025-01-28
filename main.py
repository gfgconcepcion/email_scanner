import os
import pickle
import base64
from typing import List, Dict
from datetime import datetime
from pathlib import Path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Gmail API scopes
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def get_gmail_service():
    """Sets up and returns the Gmail service."""
    creds = None
    
    # Check if token.pickle exists
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    # If credentials are invalid or don't exist, get new ones
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save credentials for future use
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    return build('gmail', 'v1', credentials=creds)

def get_emails_from_sender(service, sender_email: str) -> List[Dict]:
    """Retrieves all emails from a specific sender."""
    results = service.users().messages().list(
        userId='me',
        q=f'from:{sender_email}'
    ).execute()

    messages = results.get('messages', [])
    emails = []

    for message in messages:
        email = service.users().messages().get(
            userId='me',
            id=message['id'],
            format='full'
        ).execute()
        emails.append(email)
    
    return emails

def filter_emails_by_subject(emails: List[Dict], keyword: str) -> List[Dict]:
    """Filters emails based on a keyword in the subject."""
    filtered_emails = []
    
    for email in emails:
        headers = email['payload']['headers']
        subject = next(
            (header['value'] for header in headers if header['name'].lower() == 'subject'),
            ''
        )
        
        if keyword.lower() in subject.lower():
            filtered_emails.append(email)
    
    return filtered_emails

def decode_email_content(content: str) -> str:
    """Decodes base64 encoded email content."""
    try:
        # Add padding if needed
        padding = 4 - (len(content) % 4) if len(content) % 4 else 0
        content += '=' * padding
        
        # Replace URL-safe characters
        content = content.replace('-', '+').replace('_', '/')
        
        # Decode content
        decoded = base64.b64decode(content).decode('utf-8', errors='ignore')
        return decoded
    except Exception as e:
        return f"Error decoding content: {str(e)}"

def get_email_content(email: Dict) -> str:
    """Extracts and processes email content."""
    content = []
    payload = email['payload']

    def extract_content(part):
        """Helper function to extract content from a message part."""
        if 'body' in part:
            if 'data' in part['body']:
                decoded = decode_email_content(part['body']['data'])
                if decoded:
                    content.append(decoded)
            elif 'attachmentId' in part['body']:
                content.append('[Attachment not downloaded]')

    # Handle all parts recursively
    def process_parts(part):
        """Recursively process message parts."""
        if part.get('mimeType', '').startswith('text/'):
            extract_content(part)
        
        # Handle multipart messages
        if 'parts' in part:
            for p in part['parts']:
                process_parts(p)

    # Process the email content
    process_parts(payload)
    
    # If no content was found in parts, try the main payload
    if not content and 'body' in payload:
        extract_content(payload)

    return '\n'.join(content) if content else 'No readable content available'

def save_emails(emails: List[Dict]):
    """Saves filtered emails to a directory."""
    output_dir = Path('filtered_emails')
    output_dir.mkdir(exist_ok=True)
    
    for email in emails:
        headers = email['payload']['headers']
        subject = next(
            (header['value'] for header in headers if header['name'].lower() == 'subject'),
            'No Subject'
        )
        date = next(
            (header['value'] for header in headers if header['name'].lower() == 'date'),
            'No Date'
        )
        
        # Create a safe filename from the subject
        safe_subject = "".join(x for x in subject if x.isalnum() or x in (' ', '-', '_'))
        filename = f"{safe_subject[:50]}_{email['id']}.txt"
        
        # Get and decode email content
        content = get_email_content(email)
        
        with open(output_dir / filename, 'w', encoding='utf-8') as f:
            f.write(f"Subject: {subject}\n")
            f.write(f"Date: {date}\n")
            f.write("\nMessage:\n")
            f.write(content)

def main():
    # Get configuration from environment variables
    sender_email = os.getenv('SENDER_EMAIL')
    subject_keyword = os.getenv('SUBJECT_KEYWORD')
    
    if not sender_email or not subject_keyword:
        print("Please set SENDER_EMAIL and SUBJECT_KEYWORD in your .env file")
        return
    
    try:
        # Initialize Gmail API service
        service = get_gmail_service()
        
        # Get emails from sender
        print(f"Fetching emails from {sender_email}...")
        emails = get_emails_from_sender(service, sender_email)
        
        # Filter emails by subject keyword
        print(f"Filtering emails with keyword '{subject_keyword}' in subject...")
        filtered_emails = filter_emails_by_subject(emails, subject_keyword)
        
        # Save filtered emails
        print("Saving filtered emails...")
        save_emails(filtered_emails)
        
        print(f"Successfully processed {len(filtered_emails)} emails")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == '__main__':
    main()