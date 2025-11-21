import imaplib
import email
from email.header import decode_header
import os
import smtplib
from email.mime.text import MIMEText
import pandas as pd
import requests
import time
from imapclient import IMAPClient


os.username()

IMAP_SERVER = "imap.gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

EMAIL_USER = "email_id"
EMAIL_PASS = "pass_key" #!important- two step verification must be enabled and use. Gmail App Password
username = os.getlogin()
SAVE_ATTACHMENTS = rf"C:\\Users\\{username}\\OneDrive\\Desktop\\Intelligent_email_Automation"
OUTPUT_EXCEL = rf"C:\\Users\\{username}\\OneDrive\\Desktop\\Intelligent_email_Automation\\email_log.xlsx"

# Perplexity AI
PERPLEXITY_API_KEY = "your_perplexity_api_key"
PERPLEXITY_URL = "https://api.perplexity.ai/chat/completions"

# Folders to monitor
FOLDERS_TO_MONITOR = ["SPAM", "InvoiceFolder"]

HOST = "imap.gmail.com"
SIGNATURE = "\n\nBest regards,\nYour Name\nYour Position\nYour Company"

print("Email Automation Started...")



# HELPER FUNCTIONS

def clean_header(value):
    if value is None:
        return ""
    decoded, charset = decode_header(value)[0]
    print("Decoded:", decoded, "Charset:", charset)
    if isinstance(decoded, bytes):
        print("Decoding bytes...",charset or "utf-8", errors="ignore")
        return decoded.decode(charset or "utf-8", errors="ignore")
    return decoded


def save_attachment(part, subject):
    filename = part.get_filename()
    if filename:
        filename = clean_header(filename)
        filepath = os.path.join(SAVE_ATTACHMENTS, filename)
        with open(filepath, "wb") as f:
            f.write(part.get_payload(decode=True))
        print(f"Attachment saved: {filepath}")
        return filepath
    return None



# FETCH EMAILS

def fetch_unread_emails(server, folder):
    emails = []
    try:
        server.select_folder(folder)
        messages = server.search(["UNSEEN"])
        print(f"Found {len(messages)} unread emails in {folder}.")

        for msg_id in messages:
            msg_data = server.fetch(msg_id, ["RFC822"])[msg_id] #["RFC822"]: tells the server to return the full raw email content (headers + body + attachments) in RFC822 format.
            msg = email.message_from_bytes(msg_data[b"RFC822"])

            subject = clean_header(msg["Subject"])
            sender = clean_header(msg["From"])
            date = msg["Date"]

            body = ""
            attachments = []

            if msg.is_multipart():
                for part in msg.walk():
                    ctype = part.get_content_type()

                    # Attachments
                    if "attachment" in str(part.get("Content-Disposition")).lower():
                        attachment_path = save_attachment(part, subject)
                        if attachment_path:
                            attachments.append(attachment_path)

                    # Body
                    if ctype == "text/plain":
                        body += part.get_payload(decode=True).decode("utf-8", errors="ignore")
            else:
                body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")

            emails.append({
                "Date": date,
                "Sender": sender,
                "Subject": subject,
                "Body": body,
                "Attachments": ", ".join(attachments),
                "Folder": folder
            })
    except Exception as e:
        print(f"⚠ Error fetching emails from folder {folder}: {e}")

    return emails



# CLASSIFICATION

def classify_email(subject, body):
    s = subject.lower()
    b = body.lower()

    spam_words = ["lottery", "win money", "prize", "urgent click", "bitcoin", "crypto", "xxx"]
    for w in spam_words:
        if w in s or w in b:
            return "Spam"

    if "invoice" in s or "invoice" in b:
        return "Invoice"
    if "leave" in s or "vacation" in b:
        return "Leave Request"
    if "request" in s:
        return "General Request"

    return "Uncategorized"



# OPTIONAL: PERPLEXITY AI

def perplexity_chat(prompt):
    if not PERPLEXITY_API_KEY:
        return ""
    headers = {"Content-Type": "application/json",
               "Authorization": f"Bearer {PERPLEXITY_API_KEY}"}
    payload = {"model": "sonar-pro", "messages": [{"role": "user", "content": prompt}]}

    response = requests.post(PERPLEXITY_URL, json=payload, headers=headers)
    if response.status_code != 200:
        print("Perplexity Error:", response.text)
        return ""
    return response.json()["choices"][0]["message"]["content"]


def generate_summary(body):
    return perplexity_chat(f"Summarize this email:\n\n{body}")


def generate_auto_reply(body):
    return perplexity_chat(f"Write a polite lke human email reply to this email:\n\n{body}")

def generate_auto_category(category):
    return perplexity_chat(f"Categorize this email into one of the following categories: Invoice, Leave Request, Complaint, General Request, Uncategorized:\n\n{category}")



# SEND EMAIL AUTO-REPLY

def send_email(to, subject, body):

    if not body or not isinstance(body, str):
        body = "Thank you for your email. I will get back to you soon."
    #body += SIGNATURE

    msg = MIMEText(body)
    msg["From"] = EMAIL_USER
    msg["To"] = to
    msg["Subject"] = subject

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASS)
        server.sendmail(EMAIL_USER, to, msg.as_string())
    print(f"Auto-reply sent to {to}")



# SAVE LOG

def save_to_excel(data):
    df = pd.DataFrame(data)
    if os.path.exists(OUTPUT_EXCEL):
        old_df = pd.read_excel(OUTPUT_EXCEL)
        df = pd.concat([old_df, df], ignore_index=True)
    df.to_excel(OUTPUT_EXCEL, index=False)
    print(f"Log updated: {OUTPUT_EXCEL}")



# PROCESS SINGLE EMAIL

def process_email(email_data):
    subject = email_data["Subject"]
    body = email_data["Body"]
    sender = email_data["Sender"]

    if "<" in sender:
        sender = sender.split("<")[1].split(">")[0]

    category = classify_email(subject, body)

    summary = generate_summary(body)
    
    
    reply = generate_auto_reply(body) if PERPLEXITY_API_KEY else "Thank you for your email. I will get back to you soon."
    reply="[Your Name]" in reply and reply.replace("[Your Name]", "XYZ") or "[Your Position]" in reply and reply.replace("[Your Position]", "") or "Here’s a polite and professional reply you could send to*" in reply and reply.replace("Here’s a polite and professional reply you could send to*", "")
    #reply = generate_auto_reply(body) + signature if PERPLEXITY_API_KEY else "Thank you for your email. I will get back to you soon." + signature

    # Send auto reply
    print(generate_auto_category(category))
    send_email(sender, "Re: " + subject, reply)

    # Save details
    email_data.update({
        "Category": category,
        "Summary": summary,
        "AI_Reply": reply
    })

    save_to_excel([email_data])



# REAL-TIME MULTI-FOLDER POLLING

def real_time_monitor():
    print(" Starting REAL-TIME multi-folder polling listener...")
    with IMAPClient(HOST, use_uid=True, ssl=True) as server:
        server.login(EMAIL_USER, EMAIL_PASS)

        while True:
            print("Checking folders for new emails...")
            for folder in FOLDERS_TO_MONITOR:
                emails = fetch_unread_emails(server, folder)
                for email_data in emails:
                    print(f" New email detected in {folder}: {email_data['Subject']}")
                    process_email(email_data)
            time.sleep(10)  # poll every 10 seconds



# START SCRIPT

if __name__ == "__main__":
    real_time_monitor()
