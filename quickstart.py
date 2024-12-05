import pandas as pd
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import io

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.send',
          'https://www.googleapis.com/auth/drive.readonly']


# Step 1: Authenticate and build services
def authenticate_google():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    gmail_service = build("gmail", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)

    return gmail_service, drive_service


# Step 2: Read the CSV
def read_csv_file(csv_file):
    df = pd.read_csv(csv_file)
    return df


# Step 3: Download the file from Google Drive
def download_file_from_drive(drive_service, file_id, destination):
    request = drive_service.files().get_media(fileId=file_id)
    fh = io.FileIO(destination, 'wb')
    downloader = MediaIoBaseDownload(fh, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()
        print(f"Download {int(status.progress() * 100)}%.")

    return destination


# Step 4: Send the email
def send_email(gmail_service, recipient_email, subject, body, attachment_path, name):
    message = MIMEMultipart()
    message["to"] = recipient_email
    message["subject"] = subject

    # Attach the body with the msg instance
    message.attach(MIMEText(body, "plain"))

    # Attach the file
    with open(attachment_path, "rb") as f:
        part = MIMEText(f.read(), "base64", "utf-8")
        part.add_header("Content-Disposition", f'attachment; filename="{name}_Certificate.pdf"')
        message.attach(part)

    # Encode the message
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

    # Send the message
    send_message = {
        'raw': raw_message
    }
    gmail_service.users().messages().send(userId="me", body=send_message).execute()


# Step 5: Process each row and send emails
def process_and_send_emails(csv_file):
    gmail_service, drive_service = authenticate_google()
    df = read_csv_file(csv_file)

    for index, row in df.iterrows():
        name = row["name"]
        recipient_email = row["email"]
        team_name = row["team name"]
        doc_id = row["doc_id"]
        subject = "Navdhara 2024 Certificates Distribution"
        body = (f"Dear {team_name},\n\nWe hope this email finds you well.\n\nWe would like to extend our heartfelt gratitude for your dedication and hard work during the Navdhara event. The success of the event was a collective achievement, and each of you played a vital role in making it happen.\n\nAttached to this email, you will find the official certificates, which formally acknowledge your participation and efforts."
                f"\n\nPlease feel free to reach out if you have any questions or require further assistance.\n\nThank you once again for your commitment and support. We look forward to working with you on future endeavors.\n\nBest regards,\nTeam Navdhara")

        try:
            # Download the file from Google Drive
            attachment_path = f"{name}_Certificate.pdf"
            download_file_from_drive(drive_service, doc_id, attachment_path)

            # Send the email with the attachment
            send_email(gmail_service, recipient_email, subject, body, attachment_path, name)

            # Update status to "Sent"
            df.at[index, "status"] = "Sent"

            # Clean up: remove the downloaded file
            os.remove(attachment_path)

        except Exception as e:
            df.at[index, "status"] = f"Failed: {str(e)}"

    df.to_csv(csv_file, index=False)
    print("Emails processed and CSV updated.")


# Provide the CSV file path
csv_file = "311-410.csv"
process_and_send_emails(csv_file)
