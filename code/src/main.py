import imaplib
import smtplib
import email
import os
import pytesseract
import google.generativeai as genai
from PIL import Image
from email import policy
from email.parser import BytesParser
from PyPDF2 import PdfReader
from io import BytesIO
from docx import Document
import uuid
import re
from datetime import datetime
import gspread
import google.oauth2.service_account
from google.oauth2.service_account import Credentials
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Email credentials (use environment variables for security)
EMAIL_USER = "technologyhackathon5@gmail.com"
EMAIL_PASS = "jxjhckdkklfbrfkk"
IMAP_SERVER = "imap.gmail.com"
IMAP_FOLDER = "INBOX"
GEMINI_API_KEY = "***" # Update Gemini API Key 
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

genai.configure(api_key=GEMINI_API_KEY)


# Connect to email server
def connect_email():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select(IMAP_FOLDER)
    return mail

# Fetch latest emails
def fetch_emails(mail, num=10):
    result, data = mail.search(None, "UNSEEN")
    email_ids = data[0].split()[-num:]
    emails = []
    for e_id in email_ids:
        result, msg_data = mail.fetch(e_id, "(RFC822)")
        raw_email = msg_data[0][1]
        msg = BytesParser(policy=policy.default).parsebytes(raw_email)
        emails.append(msg)
    return emails

# Extract text from email
def extract_email_content(msg):
    subject = msg["subject"]
    sender = msg["from"]
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                body = part.get_payload(decode=True).decode(errors="ignore")
    else:
        body = msg.get_payload(decode=True).decode(errors="ignore")

    return {"subject": subject, "sender": sender, "body": body}

# Extract attachments and process text
def extract_attachments(msg, save_path="attachments"):
    os.makedirs(save_path, exist_ok=True)
    attachments = []
    extracted_text = ""
    eml_messages = []

    for part in msg.walk():
        if part.get_content_disposition() == "attachment":
            filename = part.get_filename()
            payload = part.get_payload(decode=True)

            # Handle cases where the payload is None
            if not payload:
                print(f"Skipping empty attachment: {filename}")
                continue
            if filename:
                filepath = os.path.join(save_path, filename)
                with open(filepath, "wb") as f:
                    f.write(part.get_payload(decode=True))
                attachments.append(filepath)

                # Extract text if it's a PDF
                if filepath.endswith(".pdf"):
                    extracted_text += extract_text_from_pdf(filepath) + "\n"

                # Extract text if it's an image
                elif filepath.lower().endswith((".jpg", ".jpeg", ".png")):
                    extracted_text += extract_text_from_image(filepath) + "\n"

                elif filepath.endswith(".docx"):
                    with open(filepath, "rb") as docx_file:
                        extracted_text += extract_text_from_docx(docx_file.read()) + "\n"

                elif filepath.endswith(".eml"):
                    with open(filepath, "rb") as eml_file:
                        eml_msg = BytesParser(policy=policy.default).parse(eml_file)
                        eml_messages.append(eml_msg)  # Store parsed .eml message

    return attachments, extracted_text, eml_messages

# Extract text from PDF using PyPDF2
def extract_text_from_pdf(pdf_path):
    text = ""
    with open(pdf_path, "rb") as f:
        pdf_reader = PdfReader(f)
        for page in pdf_reader.pages:
            text += page.extract_text() + "\n"
    return text

# Extract text from images using OCR
def extract_text_from_image(image_path):
    try:
        image = Image.open(image_path)
        text = pytesseract.image_to_string(image)
        return text
    except Exception as e:
        return f"Error extracting text from image: {e}"

# Extract text from .docx files
def extract_text_from_docx(docx_bytes):
    """Extracts text from a .docx file"""
    docx_file = BytesIO(docx_bytes)
    doc = Document(docx_file)
    return "\n".join([para.text for para in doc.paragraphs])

# Classify email using Gemini AI
def classify_email_with_gemini(email_text, email_data):
    prompt = f"""
    You are an AI email assistant. Classify the following email and extract key details.

    Email:
    ---
    {email_data}
    {email_text}
    ---

    Extract the following with a precise words:
    1. Request Type (Complaint, Support, Adjustment, AU Transfer, Closing Notice, Commitment Change, Fee Payment, Money Movement - Inbound, Money Movement - Outbound)
    2. Sub Request Type 
    3. Sender
    4. Subject

    Do NOT include extra text. If Request Type doesn't matches the option mentioned above please return it as None
    Example:
    1. Request Type: Request
    2. Sub Request Type: Sub Request
    3. Sender: bobbichetty tribhuvankumar <bobbichettytribhuvankumar@gmail.com>
    4. Subject: Fraudulent transaction
    """

    model = genai.GenerativeModel("gemini-2.0-flash")
    response = model.generate_content(prompt)
    return response.text

def extract_request_details(ai_response):
    # Define regex patterns to match request type and sub request type
    request_type_pattern = r"Request Type:\s*(.+)"
    sub_request_type_pattern = r"Sub Request Type:\s*(.+)"
    email_pattern = r"<([^>]+)>"  # Extracts email inside < >
    subject_pattern = r"Subject:\s*(.+)"  # Extracts subject text

    # Search for request type
    request_type_match = re.search(request_type_pattern, ai_response)
    sub_request_type_match = re.search(sub_request_type_pattern, ai_response)
    email_match = re.search(email_pattern, ai_response)
    subject_match = re.search(subject_pattern, ai_response)

    # Extract values or set default if not found
    request_type = request_type_match.group(1) if request_type_match else "Unknown"
    sub_request_type = sub_request_type_match.group(1) if sub_request_type_match else "Unknown"
    email = email_match.group(1).strip() if email_match else "Unknown"
    subject = subject_match.group(1).strip() if subject_match else "Unknown"

    return request_type, sub_request_type, email, subject

# Google Sheets Credentials
SERVICE_ACCOUNT_FILE = "C:/Users/Admin/Downloads/high-office.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Authenticate
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(creds)

# Open the Google Sheet
SPREADSHEET_ID = "1yLsOlV0Z3_JXjsRcAlSisIycs1MFJ-xzI8fUa2Xd0eI"
SHEET_NAME = "MainSheet"  # Change to your sheet name
sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)

# Function to update Google Sheet
def update_google_sheet(request_type, sub_request_type, sender_email, subject, unique_id):
    data = [datetime.now().isoformat(), unique_id, sender_email, request_type, sub_request_type, subject]
    sheet.append_row(data)  # Appends data as a new row
    print("✅ Data added to Google Sheets")

def send_acknowledgment_email(request_type, sub_request_type, sender_email, subject, unique_id):
    try:
        msg = MIMEMultipart()
        msg["From"] = EMAIL_USER
        msg["To"] = sender_email
        msg["Subject"] = f"Acknowledgment: {subject} - {unique_id}"

        body = f"""
Dear customer,

Thank you for reaching out to us. We have received your email regarding "{subject}".

Please find your ticket details -
Ticket No: {unique_id}
Request Type: {request_type}
Sub Request Type: {sub_request_type}

Will process your request soon. Thank you for banking with us.

Best Regards,
Technology Hackathon Team
        """
        msg.attach(MIMEText(body, "plain"))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_USER, EMAIL_PASS)
            server.sendmail(EMAIL_USER, sender_email, msg.as_string())

        print(f"Acknowledgment email sent to {sender_email}")

    except Exception as e:
        print(f"Error sending acknowledgment email: {e}")

# Main execution
if __name__ == "__main__":
    mail = connect_email()
    emails = fetch_emails(mail, num=5)

    for msg in emails:
        email_data = extract_email_content(msg)
        attachments, extracted_text, eml_messages = extract_attachments(msg)
        combined_text = email_data["body"] + "\n" + extracted_text
        print("Email:", email_data)
        classification = classify_email_with_gemini(combined_text, email_data)
        print(classification)
        Request_type, Sub_request_type, sender_email, subject = extract_request_details(classification)
        print("Request Type: ", Request_type)
        print("Sub Request Type: ", Sub_request_type)
        print("Email: ", sender_email)
        print("Subject: ", subject)
        unique_id = str(uuid.uuid4())[:16]
        if Request_type != "None":
            update_google_sheet(Request_type.replace("*", ""), Sub_request_type.replace("*", ""), sender_email,
                                subject.replace("*", ""),unique_id)
            send_acknowledgment_email(Request_type.replace("*", ""), Sub_request_type.replace("*", ""), sender_email,
                                subject.replace("*", ""), unique_id)
        for eml_msg in eml_messages:
            eml_data = extract_email_content(eml_msg)
            classification = classify_email_with_gemini(eml_data["body"], eml_data)
            Request_type, Sub_request_type, sender_email, subject = extract_request_details(classification)
            if Request_type != "None":
                update_google_sheet(Request_type, Sub_request_type, sender_email, subject, unique_id)
        for attachment in attachments:
            print(f"Processed attachment: {attachment}")
