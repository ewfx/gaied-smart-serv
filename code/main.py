import imaplib
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
from datetime import datetime
# Email credentials (use environment variables for security)
EMAIL_USER = "technologyhackathon5@gmail.com"
EMAIL_PASS = "jxjhckdkklfbrfkk"
IMAP_SERVER = "imap.gmail.com"
IMAP_FOLDER = "INBOX"
GEMINI_API_KEY = "*****" # Add Api Key here

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

    for part in msg.walk():
        if part.get_content_disposition() == "attachment":
            filename = part.get_filename()
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

    return attachments, extracted_text


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
def classify_email_with_gemini(email_text):
    prompt = f"""
        Analyze the following text, which is related to loan servicing at a bank.

        Based on the content, determine the following:

        1.  *Request Type:* Categorize the overall request. Choose from the following list:
        2.  *Sub Request Type(s):* If applicable, identify specific sub-requests within the main request. Choose from the sub-request types provided in the above list. If no sub-requests are identifiable, state "None".
        3.  *Prioritization:* Assign a prioritization level to the request (High, Medium, Low) and explain the reasoning behind your choice. Consider factors like urgency, potential financial impact, and customer sensitivity.
        4.  *Confidence Score:* Provide a confidence score (0-100) indicating how certain you are about the categorization and prioritization. Explain your reasoning for the assigned score.
        5.  *Unique Case Keys:* Identify the unique key or combination of keys that would help in creating a service case and grouping similar texts. Examples include: Account Number, Loan ID, Customer Name + Loan ID, etc. Explain why this combination would be unique.
        6.  *Key Entities:* Extract and list any important entities mentioned in the text, including:
            * Names (e.g., John Doe)
            * Amounts (e.g., $10,000)
            * Dates (if any)
            * Loan IDs
            * Company Names
            * Any other relevant information.
            For each entity, explain why it is considered important.

        Text:

        {email_text}

        Output (JSON format):
        """

    model = genai.GenerativeModel("gemini-2.0-pro-exp")
    response = model.generate_content(prompt)
    print(response.text)
    return response.text

import re

def extract_request_details(ai_response):
    # Define regex patterns to match request type and sub request type
    request_type_pattern = r"\*\*Request Type:\*\*\s*(.+)"
    sub_request_type_pattern = r"\*\*Sub Request Type:\*\*\s*(.+)"

    # Search for request type
    request_type_match = re.search(request_type_pattern, ai_response)
    sub_request_type_match = re.search(sub_request_type_pattern, ai_response)

    # Extract values or set default if not found
    request_type = request_type_match.group(1) if request_type_match else "Unknown"
    sub_request_type = sub_request_type_match.group(1) if sub_request_type_match else "Unknown"

    return request_type, sub_request_type

import gspread
import google.oauth2.service_account
from google.oauth2.service_account import Credentials

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

def generate_unique_id():
    return str(uuid.uuid4())[:16]

# Function to update Google Sheet
def update_google_sheet(request_type, sub_request_type):
    unique_id = generate_unique_id()
    data = [datetime.now().isoformat(), unique_id, request_type, sub_request_type,]
    sheet.append_row(data)  # Appends data as a new row
    print("✅ Data added to Google Sheets")


# Main execution
if __name__ == "__main__":
    mail = connect_email()
    emails = fetch_emails(mail, num=5)

    for msg in emails:
        email_data = extract_email_content(msg)
        attachments, extracted_text = extract_attachments(msg)
        combined_text = email_data["body"] + "\n" + extracted_text

        print("Email:", email_data)
        classification = classify_email_with_gemini(combined_text)
        print(classification)
        Request_type, Sub_request_type = extract_request_details(classification)
        print("Request Type: ", Request_type,"Sub Request Type: ",Sub_request_type)
        update_google_sheet(Request_type,Sub_request_type)
        for attachment in attachments:
            print(f"Processed attachment: {attachment}")

