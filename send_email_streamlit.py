import streamlit as st
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
from email.mime.application import MIMEApplication
from jinja2 import Template

# Create logs directory if it doesn't exist
if not os.path.exists('logs'):
    os.makedirs('logs')

# Configure logging
log_filename = f'logs/email_sender_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_filename),
        logging.StreamHandler()
    ]
)

def initialize_app():
    st.set_page_config(page_title="Email Sender App", layout="wide")
    st.title("📧 Bulk Email Sender")
    
    if 'preview_data' not in st.session_state:
        st.session_state.preview_data = None

def upload_csv():
    uploaded_file = st.file_uploader("Upload your CSV file", type=['csv'])
    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file)
            required_columns = ['Nominator Email(CC)', 'Nominator Name', 'Nominee Email', 'Name of Nominee', 'PDFPath']
            if not all(col in df.columns for col in required_columns):
                st.error(f"CSV must contain columns: {', '.join(required_columns)}")
                return None
            
            # Store data in session state but don't preview here
            st.session_state.preview_data = df
            return df
        except Exception as e:
            st.error(f"Error reading CSV: {str(e)}")
            logging.error(f"CSV upload error: {str(e)}")
    return None

def email_settings():
    with st.expander("Email Settings", expanded=True):
        st.subheader("Email Header Image")
        uploaded_header = st.file_uploader("Upload header image", type=['png', 'jpg', 'jpeg'])
        if uploaded_header:
            st.session_state.header_image_data = uploaded_header.getvalue()
            st.image(uploaded_header, caption="Preview of header image", width=600)
        else:
            st.warning("Please upload a header image")
            return None
        
        col1, col2 = st.columns(2)
        with col1:
            email = st.text_input("Your Email Address", key="email")
        with col2:
            password = st.text_input("Email Password", type="password", key="password")
        
        subject = st.text_input("Email Subject", value="Caught Being Good Nomination!")
        
        # Create a sample preview using the template
        sample_name = "Sample Recipient"
        sample_nominator = "Sample Nominator"
        
        # Create HTML content with proper escaping of curly braces for CSS
        template_content = """<!DOCTYPE html>
<html>
    <head>
        <style>
            .email-container {{
                max-width: 600px;
                margin: 0 auto;
                font-family: Arial, sans-serif;
                padding: 20px;
                background-color: #ffffff;
            }}
            .hero-image {{
                width: 100%;
                max-width: 600px;
                height: auto;
                aspect-ratio: 3/1;
                object-fit: cover;
                margin-bottom: 30px;
            }}
            .content {{
                padding: 25px;
                line-height: 1.8;
                color: #333;
                background-color: #ffffff;
            }}
            .greeting {{
                font-size: 20px;
                margin-bottom: 25px;
                font-weight: 500;
            }}
            .message {{
                margin-bottom: 25px;
                font-size: 16px;
            }}
            .signature {{
                margin-top: 35px;
                padding-top: 20px;
                border-top: 1px solid #eee;
                font-style: italic;
                color: #555;
            }}
        </style>
    </head>
    <body style="background-color: #f5f5f5; margin: 0; padding: 20px;">
        <div class="email-container">
            <img src="cid:header_image" 
                 alt="Header Image" 
                 class="hero-image"
                 width="600"
                 height="200">
            <div class="content">
                <div class="greeting">Dear {name} 🤩,</div>
                <div class="message">
                    Congratulations on being recognized as a Caught Being Good nominee! 🎉 Your actions and dedication to making Ashesi a better community have not gone unnoticed thanks to {nominator_name}, and we are so proud of you. 🎉🤩
                </div>
                <div class="message">
                    This attached certificate is a small token of appreciation for the kindness, leadership, and positive impact you continue to share with those around you. You inspire us all to do better and be better. 🤗❤️
                </div>
                <div class="message">
                    Keep shining your light and making a difference; you are truly appreciated!
                </div>
                <div class="signature">
                    Yours Sincerely,<br>
                    <strong>The SLE Team</strong>
                </div>
            </div>
        </div>
    </body>
</html>"""

        # Add preview section
        st.subheader("Email Preview")
        if st.session_state.get('header_image_data'):
            st.image(st.session_state.header_image_data, caption="Header Image Preview", width=600)
        
        # Format the template with both required parameters
        preview_html = template_content.format(
            name=sample_name,
            nominator_name=sample_nominator
        )
        
        # Display the formatted content in a container with custom CSS
        st.markdown("""
        <style>
        .email-preview {
            border: 1px solid #ddd;
            border-radius: 5px;
            padding: 20px;
            margin: 10px 0;
            background-color: white;
        }
        </style>
        """, unsafe_allow_html=True)
        
        st.markdown("<div class='email-preview'>", unsafe_allow_html=True)
        st.markdown(preview_html, unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Add a note about the preview
        st.info("👆 This is how your email will appear to recipients. The actual email will include the uploaded header image and personalized name.")

        return {
            'email': email,
            'password': password,
            'subject': subject,
            'template': template_content,
            'header_image_data': st.session_state.get('header_image_data') if 'header_image_data' in st.session_state else None
        }

def send_emails(settings, data):
    try:
        smtp_server = 'smtp.office365.com'
        smtp_port = 587
        SLE_EMAIL = 'sle@ashesi.edu.gh'  # Add constant
        logging.info(f"Attempting to connect to SMTP server: {smtp_server}:{smtp_port}")
        
        if not settings.get('header_image_data'):
            raise ValueError("Header image is required")
        
        with st.spinner('Sending emails...'):
            progress_bar = st.progress(0)
            total_emails = len(data)
            
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                logging.info("TLS connection established")
                
                server.login(settings['email'], settings['password'])
                logging.info("Successfully logged into SMTP server")
                
                for index, row in data.iterrows():
                    try:
                        msg = MIMEMultipart('related')
                        msg['From'] = settings['email']
                        msg['To'] = row['Nominee Email']
                        
                        # Handle CC recipients - now including SLE email
                        cc_recipients = [SLE_EMAIL]  # Start with SLE email
                        nominator_cc = row['Nominator Email(CC)']
                        if pd.notna(nominator_cc) and isinstance(nominator_cc, str) and '@' in nominator_cc:
                            cc_recipients.append(nominator_cc.strip())
                        
                        msg['Cc'] = ', '.join(cc_recipients)  # Join all CC recipients
                        all_recipients = [row['Nominee Email']] + cc_recipients
                        logging.info(f"CC recipients: {cc_recipients}")
                        
                        msg['Subject'] = settings['subject']
                        
                        logging.info(f"Processing email for: {row['Nominee Email']}")
                        
                        # Create alternative part for HTML
                        alt_part = MIMEMultipart('alternative')
                        msg.attach(alt_part)
                        
                        # HTML Content
                        html_content = settings['template'].format(
                            image_url="cid:header_image",
                            name=row['Name of Nominee'],
                            nominator_name=row['Nominator Name']
                        )
                        html_part = MIMEText(html_content, 'html')
                        alt_part.attach(html_part)
                        
                        # Attach header image
                        image = MIMEImage(settings['header_image_data'])
                        image.add_header('Content-ID', '<header_image>')
                        image.add_header('Content-Disposition', 'inline')
                        msg.attach(image)
                        logging.info("Header image attached successfully")
                        
                        # Attach PDF
                        pdf_path_with_extension = row['PDFPath'] + ".pdf"
                        if os.path.exists(pdf_path_with_extension):
                            with open(pdf_path_with_extension, 'rb') as pdf:
                                pdf_attachment = MIMEApplication(pdf.read(), _subtype='pdf')
                                pdf_attachment.add_header(
                                    'Content-Disposition', 
                                    'attachment', 
                                    filename=os.path.basename(pdf_path_with_extension)
                                )
                                msg.attach(pdf_attachment)
                                logging.info(f"PDF attached successfully: {pdf_path_with_extension}")
                        else:
                            logging.warning(f"PDF path not found for {row['Nominee Email']}: {pdf_path_with_extension}")

                        
                        # Send email to all recipients
                        server.send_message(msg, to_addrs=all_recipients)
                        progress_bar.progress((index + 1) / total_emails)
                        
                        # Log success with recipient details - using cc_recipients instead of cc_email
                        cc_info = f" with CC: {', '.join(cc_recipients)}" if cc_recipients else ""
                        logging.info(f"Email successfully sent to: {row['Nominee Email']}{cc_info}")
                        
                        # Show success in Streamlit
                        st.write(f"✅ Sent to {row['Name of Nominee']} ({row['Nominator Name']} - {row['Nominee Email']}){cc_info}")
                        
                    except Exception as e:
                        error_msg = f"Error sending to {row['Nominee Email']}: {str(e)}"
                        st.error(error_msg)
                        logging.error(error_msg)
                        continue
                        
        st.success("All emails sent successfully!")
        logging.info("Email sending process completed successfully")
        
    except Exception as e:
        error_msg = f"Error: {str(e)}"
        st.error(error_msg)
        logging.error(error_msg)

def main():
    initialize_app()
    
    # Sidebar for help and information
    with st.sidebar:
        st.header("Instructions")
        st.markdown("""
        1. Upload your CSV file (must have 'Nominator Name', 'Nominator Email(CC)', 'Nominee Email', 'Name of Nominee', and 'PDFPath' columns)
        2. Enter your email credentials
        3. Choose email type and compose your message
        4. Preview the CSV data
        5. Send emails
        """)
        
        st.header("CSV Format")
        st.markdown("""
        Required columns:
        - Nominator Name
        - Nominator Email(CC)
        - Nominee Email
        - Name of Nominee
        - PDFPath
        """)
    
    # Main app layout
    data = upload_csv()
    if data is not None:
        # Show CSV preview only once
        st.subheader("CSV Preview")
        st.dataframe(data.head())
        
        settings = email_settings()
        
        if st.button("Send Emails", type="primary"):
            if not all([settings['email'], settings['password'], settings['subject']]):
                st.error("Please fill in all required fields")
            else:
                send_emails(settings, data)

if __name__ == "__main__":
    main()
