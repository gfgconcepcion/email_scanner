import os
import pickle
import base64
import csv
import re
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

def get_emails_from_senders(service, sender_emails: List[str]) -> List[Dict]:
    """Retrieves all emails from specified senders."""
    all_emails = []
    
    for sender_email in sender_emails:
        print(f"Fetching emails from {sender_email}...")
        results = service.users().messages().list(
            userId='me',
            q=f'from:{sender_email}'
        ).execute()

        messages = results.get('messages', [])
        
        for message in messages:
            email = service.users().messages().get(
                userId='me',
                id=message['id'],
                format='full'
            ).execute()
            all_emails.append(email)
    
    return all_emails

def filter_emails_by_subjects(emails: List[Dict], keywords: List[str]) -> List[Dict]:
    """Filters emails based on keywords in the subject."""
    filtered_emails = []
    
    for email in emails:
        headers = email['payload']['headers']
        subject = next(
            (header['value'] for header in headers if header['name'].lower() == 'subject'),
            ''
        ).lower()
        
        # Check if any of the keywords match
        if any(keyword.lower() in subject for keyword in keywords):
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

def create_word_frequency_table(content: str) -> Dict[str, int]:
    """Creates a frequency table of unique words in the content."""
    # Remove HTML tags if any exist
    content = re.sub(r'<[^>]+>', '', content)
    
    # Remove URLs
    content = re.sub(r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', '', content)
    
    # Split content into lines and keep only content before "Subscription Information"
    lines = content.split('\n')
    content_lines = []
    for line in lines:
        if line.strip() == "Subscription Information":
            break
        content_lines.append(line)
    
    # Rejoin the filtered content
    filtered_content = '\n'.join(content_lines)
    
    # Convert to lowercase and split into words
    # Only keep alphanumeric words (removes punctuation, special characters)
    words = re.findall(r'\b\w+\b', filtered_content.lower())
    
    # Create frequency table
    word_freq = {}
    for word in words:
        if word.isalnum():  # Additional check to ensure word is alphanumeric
            word_freq[word] = word_freq.get(word, 0) + 1
    
    return word_freq

def save_word_frequency_csv(word_freq: Dict[str, int], filename: str, output_dir: Path):
    """Saves word frequency table as CSV file."""
    csv_path = output_dir / f"{filename}.csv"
    
    with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Word', 'Frequency'])  # Header
        
        # Sort by frequency in descending order
        sorted_words = sorted(word_freq.items(), key=lambda x: x[1], reverse=True)
        writer.writerows(sorted_words)

def save_emails(emails: List[Dict]):
    """Saves filtered emails to a directory and creates word frequency tables."""
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
        filename = f"{safe_subject[:50]}_{email['id']}"
        
        # Get and decode email content
        content = get_email_content(email)
        
        # Save email content as text file
        with open(output_dir / f"{filename}.txt", 'w', encoding='utf-8') as f:
            f.write(f"Subject: {subject}\n")
            f.write(f"Date: {date}\n")
            f.write("\nMessage:\n")
            f.write(content)
        
        # Create and save word frequency table
        word_freq = create_word_frequency_table(content)
        save_word_frequency_csv(word_freq, filename, output_dir)

def main():
    # Get configuration from environment variables
    sender_emails_str = os.getenv('SENDER_EMAILS')
    subject_keywords_str = os.getenv('SUBJECT_KEYWORDS')
    
    if not sender_emails_str or not subject_keywords_str:
        print("Please set SENDER_EMAILS and SUBJECT_KEYWORDS in your .env file")
        print("Format: comma-separated values (e.g., SENDER_EMAILS=email1@domain.com,email2@domain.com)")
        return
    
    # Split the comma-separated strings into lists
    sender_emails = [email.strip() for email in sender_emails_str.split(',')]
    subject_keywords = [keyword.strip() for keyword in subject_keywords_str.split(',')]
    
    try:
        # Initialize Gmail API service
        service = get_gmail_service()
        
        # Get emails from all specified senders
        emails = get_emails_from_senders(service, sender_emails)
        print(f"Found {len(emails)} total emails from specified senders")
        
        # Filter emails by subject keywords
        print(f"Filtering emails with keywords: {', '.join(subject_keywords)}")
        filtered_emails = filter_emails_by_subjects(emails, subject_keywords)
        
        # Save filtered emails
        print("Saving filtered emails...")
        save_emails(filtered_emails)
        
        print(f"Successfully processed {len(filtered_emails)} emails")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == '__main__':
    main()