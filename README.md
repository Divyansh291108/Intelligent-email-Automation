# Intelligent-email-Automation

Problem:-

Managing emails manually is time-consuming and error-prone, especially when handling repetitive tasks like processing invoices, leave requests, or support tickets. This project automates email reading, classification, and response generation to save time and reduce errors.

Tech Stack:-

Python 3.10+ — Core scripting language
IMAPClient / IMAP/SMTP — Real-time email monitoring and sending
Pandas — Data extraction, parsing, and storage
OpenAI API  — Summarization or AI-generated email replies

Workflow:-

Monitor incoming emails in real-time using Python IMAPClient.
Extract email subject, sender, body, and attachments.
Classify emails based on predefined rules (e.g., invoice, leave request, support).
Store extracted data in Excel or database for tracking.
Use AI(OpenAI, prpelexity) to generate summaries or smart replies.
