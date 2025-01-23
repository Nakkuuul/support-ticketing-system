import os
import time
from dotenv import load_dotenv
from datetime import datetime
from pymongo import MongoClient
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import imaplib
import smtplib
import email


class SupportTicketingSystem:
    def __init__(self):
        # Load environment variables from the .env file
        load_dotenv()

        # Storing EMAIL ID and APP PASSWORD into two different variables
        self.mail = os.getenv('EMAIL_ID', False)
        self.password = os.getenv('APP_PASS', False)

        # Database connection
        self.client = MongoClient('mongodb://localhost:27017')
        self.db = self.client['support_ticketing_system']
        self.collection = self.db['threads']

        # Email service setup
        self.imap_url = 'imap.gmail.com'
        self.smtp_host = 'smtp.gmail.com'
        self.smtp_port = 587

    def connect_to_mail_service(self):
        """ Connect to the Gmail IMAP server and log in """
        try:
            self.mailservice = imaplib.IMAP4_SSL(self.imap_url)
            self.mailservice.login(self.mail, self.password)
        except Exception as e:
            print("Error Occurred during login:", e)
            return False
        return True

    def fetch_unseen_emails(self):
        """ Fetch unseen emails from Gmail """
        status, total_messages = self.mailservice.select('Inbox')  # Response: Status, Total No. of Mails
        if status != "OK":
            print("Failed to connect to Inbox")
            return []
        
        status, unseen_emails = self.mailservice.search(None, 'UNSEEN')
        email_ids = unseen_emails[0].split()

        if not email_ids:
            print("No new threads found.")
            return []
        
        return email_ids

    def parse_email(self, email_id):
        """ Parse an individual email and extract details """
        status, data = self.mailservice.fetch(email_id, '(RFC822)')
        if status != "OK":
            print(f"Failed to fetch email {email_id}.")
            return None, None, None

        # Parse the raw email
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        # Get the sender and subject information
        sender = msg.get("From")
        subject = msg.get("Subject")
        if subject:
            ticket_subject, encoding = decode_header(subject)[0]
            if isinstance(ticket_subject, bytes):
                ticket_subject = ticket_subject.decode(encoding if encoding else "utf-8")
        else:
            ticket_subject = "No Subject"

        # Get the email message
        ticket_message = None
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                if content_type == "text/plain" and "attachment" not in content_disposition:
                    ticket_message = part.get_payload(decode=True).decode()
                    break
        else:
            ticket_message = msg.get_payload(decode=True).decode()

        # Return parsed data
        return sender, ticket_subject, ticket_message

    def generate_ticket_id(self):
        """ Generate a new ticket ID """
        empty_check = self.collection.find_one()
        if empty_check is None:
            # Initialize year, month, date, and counter
            year = datetime.now().year
            month = datetime.now().month
            date = datetime.now().strftime("%d")
            counter = 100000

            initial_ticket_id = int(f"{year}{str(month).zfill(2)}{date}{counter}")
            self.collection.insert_one({
                "_id": 10000000000000,
                "ticket_id": initial_ticket_id,
                "sender": "",
                "subject": "",
                "message": "",
                "status": "",
                "created_at": "",
                "updated_at": ""
                })
            return initial_ticket_id
        else:
            latest_ticket_id = self.collection.find_one(sort=[('_id', -1)])['_id']
            return latest_ticket_id + 1

    def create_ticket(self, sender, ticket_subject, ticket_message):
        """ Create a new ticket and insert it into the database """
        new_ticket_id = self.generate_ticket_id()

        schema = {
            "_id": new_ticket_id,
            "ticket_id": new_ticket_id,
            "sender": sender or "Unknown Sender",
            "subject": ticket_subject or "No Subject",
            "message": ticket_message or "No Content",
            "status": "open",
            "created_at": datetime.now(),
            "updated_at": datetime.now()
        }

        self.collection.insert_one(schema)
        print(f"Ticket created with ID: {new_ticket_id}")
        return new_ticket_id

    def send_ticket_email(self, sender, new_ticket_id):
        """ Send an email to the sender with the new ticket ID """
        if not sender:
            print("Sender not defined. Email not sent.")
            return

        msg = MIMEMultipart()
        msg['From'] = os.getenv("RECIPIENT_NAME", False)
        msg['To'] = sender
        msg['Subject'] = "Support Ticket Update"

        # Load and update the HTML file
        try:
            with open("index.html", "r", encoding="utf-8") as file:
                html_content = file.read()

            # Replace placeholders with actual values using string replacement
            print(sender)
            html_content = html_content.replace("[CLIENT_NAME]", str(sender))
            html_content = html_content.replace("[TICKET_NUMBER]", str(new_ticket_id))

            # Add the modified HTML to the email
            msg.attach(MIMEText(html_content, 'html'))

            # Connect to SMTP server and send the email
            with smtplib.SMTP(self.smtp_host, self.smtp_port) as smtp_server:
                smtp_server.starttls()  # Secure the connection
                smtp_server.login(self.mail, self.password)
                smtp_server.send_message(msg)
        except Exception as e:
            print("Failed to send email:", e)

    def process_emails(self):
        """ Main loop for processing emails and creating tickets """
        if not self.connect_to_mail_service():
            return

        email_ids = self.fetch_unseen_emails()
        if not email_ids:
            return

        for email_id in email_ids:
            sender, ticket_subject, ticket_message = self.parse_email(email_id)

            # Create a new ticket
            new_ticket_id = self.create_ticket(sender, ticket_subject, ticket_message)

            # Send email to the client
            self.send_ticket_email(sender, new_ticket_id)

def run_system():
    """ Run the support ticketing system """
    system = SupportTicketingSystem()

    while True:
        system.process_emails()
        time.sleep(0.5)

if __name__ == "__main__":
    run_system()
