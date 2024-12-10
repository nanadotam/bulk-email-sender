import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage
import os
from dotenv import load_dotenv
import logging
from datetime import datetime
import requests
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def initialize_app():
    st.set_page_config(page_title="Email Sender App", layout="wide")
    st.title("📧 Bulk Email Sender")
    
    if 'email_type' not in st.session_state:
        st.session_state.email_type = 'regular'
    if 'preview_data' not in st.session_state:
        st.session_state.preview_data = None

def upload_csv():
    uploaded_file = st.file_uploader("Upload your CSV file", type=['csv'])
    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file)
            st.session_state.preview_data = df
            return df
        except Exception as e:
            st.error(f"Error reading CSV: {str(e)}")
    return None

def preview_data():
    if st.session_state.preview_data is not None:
        st.subheader("CSV Preview")
        st.dataframe(st.session_state.preview_data.head())

def email_settings():
    with st.expander("Email Settings", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            email = st.text_input("Your Email Address", key="email")
        with col2:
            password = st.text_input("Email Password", type="password", key="password")
        
        email_type = st.radio(
            "Select Email Type",
            ["Regular Text", "HTML with Images"],
            key="email_type"
        )
        
        subject = st.text_input("Email Subject")
        if email_type == "Regular Text":
            body = st.text_area("Email Body (Use {name} for recipient's name)")
        else:
            body = st.text_area("HTML Body (Use {name} for recipient's name)")
            image_url = st.text_input("Image URL (ImgBB or other hosted image)")
        
        return {
            'email': email,
            'password': password,
            'subject': subject,
            'body': body,
            'type': email_type,
            'image_url': image_url if email_type == "HTML with Images" else None
        }

def send_emails(settings, data):
    try:
        smtp_server = 'smtp.office365.com'
        smtp_port = 587
        
        with st.spinner('Sending emails...'):
            progress_bar = st.progress(0)
            total_emails = len(data)
            
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(settings['email'], settings['password'])
                
                for index, row in data.iterrows():
                    try:
                        msg = MIMEMultipart('alternative') if settings['type'] == "HTML with Images" else EmailMessage()
                        msg['From'] = settings['email']
                        msg['To'] = row['Email']
                        msg['Subject'] = settings['subject']
                        
                        if settings['type'] == "HTML with Images":
                            html_content = f"""
                            <html>
                                <body>
                                    <img src="{settings['image_url']}" alt="Header Image" width="600" height="auto">
                                    {settings['body'].format(name=row['Name'])}
                                </body>
                            </html>
                            """
                            msg.attach(MIMEText(html_content, 'html'))
                        else:
                            msg.set_content(settings['body'].format(name=row['Name']))
                        
                        server.send_message(msg)
                        progress_bar.progress((index + 1) / total_emails)
                        
                    except Exception as e:
                        st.error(f"Error sending to {row['Email']}: {str(e)}")
                        continue
                        
        st.success("Emails sent successfully!")
        
    except Exception as e:
        st.error(f"Error: {str(e)}")

def main():
    initialize_app()
    
    # Sidebar for help and information
    with st.sidebar:
        st.header("Instructions")
        st.markdown("""
        1. Upload your CSV file (must have 'Name' and 'Email' columns)
        2. Enter your email credentials
        3. Choose email type and compose your message
        4. Preview the CSV data
        5. Send emails
        """)
        
        st.header("CSV Format")
        st.markdown("""
        Required columns:
        - Name
        - Email
        """)
    
    # Main app layout
    data = upload_csv()
    if data is not None:
        preview_data()
        settings = email_settings()
        
        if st.button("Send Emails", type="primary"):
            if not all([settings['email'], settings['password'], settings['subject'], settings['body']]):
                st.error("Please fill in all required fields")
            else:
                send_emails(settings, data)

if __name__ == "__main__":
    main()
