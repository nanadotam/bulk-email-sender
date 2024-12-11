import pandas as pd
import smtplib
from email.message import EmailMessage
import os
from dotenv import load_dotenv
import logging
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# Configure logging
log_filename = f'email_sender_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_filename),
        logging.StreamHandler()  # This will also print to console
    ]
)

# Load environment variables
load_dotenv()
logging.info("Environment variables loaded")

try:
    # Load the CSV file
    csv_path = 'test_list_nana.csv'
    logging.info(f"Attempting to read CSV file from: {csv_path}")
    data = pd.read_csv(csv_path)
    logging.info(f"Successfully loaded CSV with {len(data)} entries")

    # Email server configuration
    smtp_server = 'smtp.office365.com'
    smtp_port = 587
    sender_email = os.getenv('EMAIL_USER')
    sender_password = os.getenv('EMAIL_PASSWORD')

    if not sender_email or not sender_password:
        logging.error("Email credentials not found in environment variables")
        raise ValueError("Missing email credentials")

    # Add these email template definitions
    subject = "Your Good Deed Recognition"
    body_template = """
    Dear {name},

    This is a test email. I'll be in your office shortly.

    Best regards,
    Your Team
    """

    # Initialize the SMTP server
    logging.info(f"Attempting to connect to SMTP server: {smtp_server}:{smtp_port}")
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        logging.info("TLS connection established")
        
        server.login(sender_email, sender_password)
        logging.info("Successfully logged into SMTP server")

        # Loop through each row in the CSV file
        for index, row in data.iterrows():
            try:
                recipient = row['Email']
                name = row['Name']
                pdf_path = row['PDFPath']
                
                logging.info(f"Processing email for: {recipient}")

                # Create the email
                msg = EmailMessage()
                msg['From'] = sender_email
                msg['To'] = recipient

                # Email subject and template
                msg['Subject'] = subject
                msg.set_content(body_template.format(name=name))

                # Attach the PDF if it exists
                if os.path.exists(pdf_path):
                    with open(pdf_path, 'rb') as pdf:
                        msg.add_attachment(
                            pdf.read(), 
                            maintype='application', 
                            subtype='pdf', 
                            filename=os.path.basename(pdf_path)
                        )
                    logging.info(f"PDF attached successfully: {pdf_path}")
                else:
                    logging.warning(f"Attachment not found: {pdf_path}")

                # Create HTML content with embedded images
                html = f"""
                <html>
                    <body>
                        <!-- Method 1: Using CID (embedded images) -->
                        <img src="cid:logo_image" width="200" height="100">
                        
                        <!-- Method 2: Using external URL -->
                        <img src="https://your-hosted-image-url.com/image.jpg" width="200" height="100">
                        
                        {body_template.format(name=name)}
                    </body>
                </html>
                """
                
                msg_html = MIMEText(html, 'html')
                msg.attach(msg_html)

                # Attach local image file
                with open('path/to/your/logo.png', 'rb') as f:
                    img = MIMEImage(f.read())
                    img.add_header('Content-ID', '<logo_image>')  # Reference this in HTML with cid:logo_image
                    msg.attach(img)

                # Send the email
                server.send_message(msg)
                logging.info(f"Email successfully sent to: {recipient}")

            except Exception as e:
                logging.error(f"Error processing email for {recipient}: {str(e)}")
                continue

    logging.info("Email sending process completed successfully")

except Exception as e:
    logging.error(f"An error occurred: {str(e)}")
    raise
