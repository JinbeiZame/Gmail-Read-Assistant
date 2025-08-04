import os
import time
import string
from datetime import datetime
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import platform

SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

# For sound notification
def play_sound():
    try:
        if platform.system() == "Windows":
            import winsound
            winsound.Beep(1000, 400)  # frequency, duration(ms)
        else:
            # Try playsound if available, otherwise fallback to system beep
            try:
                from playsound import playsound
                playsound('/System/Library/Sounds/Glass.aiff')  # Mac default sound
            except Exception:
                print('\a', end='')  # ASCII Bell
    except Exception as e:
        print("Couldn't play sound:", e)

def hard_sanitize(text):
    allowed = string.ascii_letters + string.digits + " .,!?"
    return ''.join([c for c in text if c in allowed])

def gmail_authenticate():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def get_today_unread_messages(service):
    today = datetime.now().strftime('%Y/%m/%d')
    query = f'is:unread after:{today}'
    print(f"Gmail search query: {query}")
    result = service.users().messages().list(userId='me', q=query).execute()
    messages = result.get('messages', [])
    print(f"Found {len(messages)} unread emails from today.")
    return messages

def get_message_subject(service, msg_id):
    msg = service.users().messages().get(
        userId='me',
        id=msg_id,
        format='metadata',
        metadataHeaders=['Subject']
    ).execute()
    headers = msg.get('payload', {}).get('headers', [])
    subject = next((h['value'] for h in headers if h['name'] == 'Subject'), '(No Subject)')
    return subject

def mark_as_read(service, msg_id):
    service.users().messages().modify(
        userId='me',
        id=msg_id,
        body={'removeLabelIds': ['UNREAD']}
    ).execute()

def main():
    creds = gmail_authenticate()
    service = build('gmail', 'v1', credentials=creds)
    print("Listening for new emails (today only)...")
    notified_ids = set()
    while True:
        try:
            messages = get_today_unread_messages(service)
            new_subjects = []
            new_ids = []
            for msg in messages:
                msg_id = msg['id']
                if msg_id not in notified_ids:
                    subject = get_message_subject(service, msg_id)
                    sanitized_subject = hard_sanitize(subject)
                    new_subjects.append(sanitized_subject)
                    new_ids.append(msg_id)
                    notified_ids.add(msg_id)
            if new_subjects:  # Only output if at least one new email
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                count = len(new_subjects)
                print(f"\n📬 [{timestamp}] You have {count} unread email(s) today!")
                for i, subj in enumerate(new_subjects, 1):
                    print(f"   {i}. 📧 {subj}")
                print("-" * 40)
                play_sound()  # Play sound after showing output
                for msg_id in new_ids:
                    mark_as_read(service, msg_id)
            time.sleep(15)
        except Exception as e:
            print("Error:", e)
            time.sleep(30)

if __name__ == '__main__':
    main()
