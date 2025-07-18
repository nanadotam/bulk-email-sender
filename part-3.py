from tqdm.notebook import tqdm
import time
import smtplib
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from ipywidgets import Button, Output, HTML, VBox, HBox
from IPython.display import display

# --- Send Button + Output Widgets ---
send_all_button = Button(description="🚀 Send All Emails", button_style='danger')
stop_button = Button(description="⏹️ Stop Sending", button_style='warning', disabled=True)
confirm_button = Button(description="⚠️ Confirm Send to ALL Recipients", button_style='danger', disabled=True)
cancel_button = Button(description="❌ Cancel", button_style='')
send_output = Output()
confirmation_output = Output()

# --- Global flags ---
stop_sending = False
confirmed = False

# --- Confirmation Functions ---
def show_confirmation(b):
    global confirmed
    confirmed = False
    confirmation_output.clear_output()
    
    if df.empty:
        with confirmation_output:
            print("❌ No CSV data loaded. Please complete Step 1 first.")
        return
    
    recipient_count = len(df)
    
    with confirmation_output:
        print(f"⚠️ CONFIRMATION REQUIRED")
        print(f"You're about to send {recipient_count} emails!")
        print(f"📧 From: {sender_email_input.value}")
        print(f"📝 Subject: {subject_input.value}")
        print(f"👥 Recipients: {recipient_count}")
        
        if recipient_count > 50:
            print("\n🚨 WARNING: Large recipient list detected!")
            print("   Consider breaking this into smaller batches")
            print("   Sending too many emails at once may trigger spam filters")
        
        print("\n💡 This action cannot be undone!")
        print("   Make sure you've tested your email first")
        print("   Double-check your recipient list")
        
    # Enable/disable buttons
    send_all_button.disabled = True
    confirm_button.disabled = False
    cancel_button.disabled = False

def confirm_send(b):
    global confirmed
    confirmed = True
    confirmation_output.clear_output()
    
    with confirmation_output:
        print("✅ Confirmed! Starting bulk email send...")
    
    # Disable confirmation buttons, proceed with sending
    confirm_button.disabled = True
    cancel_button.disabled = True
    send_bulk_emails(None)

def cancel_send(b):
    global confirmed
    confirmed = False
    confirmation_output.clear_output()
    
    with confirmation_output:
        print("❌ Send cancelled. You can modify your email and try again.")
    
    # Re-enable send button, disable confirmation buttons
    send_all_button.disabled = False
    confirm_button.disabled = True
    cancel_button.disabled = True

# --- Stop Function ---
def stop_sending_emails(b):
    global stop_sending
    stop_sending = True
    with send_output:
        print("\n🛑 Stop requested... finishing current email and stopping.")

# --- Bulk Email Function ---
def send_bulk_emails(b):
    global stop_sending
    stop_sending = False
    
    send_output.clear_output()
    if df.empty:
        with send_output:
            print("❌ No CSV data loaded. Please complete Step 1 first.")
        return

    # Pre-flight checks
    if not sender_email_input.value or not password_input.value:
        with send_output:
            print("❌ Please enter your email credentials in Step 2.")
        return

    # Enable stop button, disable other buttons
    send_all_button.disabled = True
    confirm_button.disabled = True
    cancel_button.disabled = True
    stop_button.disabled = False

    name_col = column_selector_name.value
    email_col = column_selector_email.value
    total_recipients = len(df)

    failed_list = []
    sent_count = 0
    start_time = time.time()

    try:
        with send_output:
            print("🔗 Connecting to email server...")
        
        server = smtplib.SMTP("smtp.office365.com", 587)
        server.starttls()
        server.login(sender_email_input.value, password_input.value)

        with send_output:
            print("✅ Successfully connected to email server!")
            print("📊 Email Campaign Details:")
            print(f"   📧 From: {sender_email_input.value}")
            print(f"   📝 Subject: {subject_input.value}")
            print(f"   👥 Total Recipients: {total_recipients}")
            print("📨 Starting email send... (Click 'Stop Sending' to cancel)\n")

        for index, row in tqdm(df.iterrows(), total=len(df), desc="Sending emails"):
            # Check if stop was requested
            if stop_sending:
                with send_output:
                    print(f"\n🛑 Sending stopped by user at {sent_count}/{total_recipients} emails sent.")
                break
                
            name = str(row[name_col]).strip()
            recipient = str(row[email_col]).strip()
            
            # Validate email format
            if not re.match(r"[^@]+@[^@]+\.[^@]+", recipient):
                failed_list.append((name, recipient, "Invalid email format"))
                continue

            try:
                # Create email message
                msg = MIMEMultipart("alternative")
                msg["Subject"] = subject_input.value
                msg["From"] = sender_email_input.value
                msg["To"] = recipient
                msg.attach(MIMEText(generate_email_html(row), "html"))

                # Send email
                server.sendmail(sender_email_input.value, recipient, msg.as_string())
                sent_count += 1
                
                # Progress update every 10 emails
                if sent_count % 10 == 0:
                    elapsed_time = time.time() - start_time
                    avg_time_per_email = elapsed_time / sent_count
                    remaining_emails = total_recipients - sent_count
                    estimated_time_remaining = remaining_emails * avg_time_per_email
                    
                    with send_output:
                        print(f"📈 Progress: {sent_count}/{total_recipients} sent ({(sent_count/total_recipients)*100:.1f}%)")
                        print(f"⏱️ Estimated time remaining: {estimated_time_remaining/60:.1f} minutes")
                
                time.sleep(0.5)  # Delay to avoid spam filters

            except smtplib.SMTPRecipientsRefused:
                failed_list.append((name, recipient, "Recipient email refused by server"))
                continue
            except smtplib.SMTPDataError as e:
                failed_list.append((name, recipient, f"SMTP Data Error: {str(e)}"))
                continue
            except Exception as e:
                failed_list.append((name, recipient, f"Unexpected error: {str(e)}"))
                continue

        server.quit()

        # Final results
        elapsed_time = time.time() - start_time
        success_rate = (sent_count / total_recipients) * 100 if total_recipients > 0 else 0

        with send_output:
            print("\n" + "="*50)
            print("📊 CAMPAIGN SUMMARY")
            print("="*50)
            print(f"✅ Successfully sent: {sent_count}/{total_recipients} emails ({success_rate:.1f}%)")
            print(f"⏱️ Total time: {elapsed_time/60:.1f} minutes")
            print(f"📧 Average: {elapsed_time/sent_count:.1f} seconds per email" if sent_count > 0 else "")
            
            if failed_list:
                print(f"\n❌ Failed to send: {len(failed_list)} emails")
                print("Failed recipients:")
                for entry in failed_list:
                    print(f"   • {entry[0]} ({entry[1]}): {entry[2]}")
                    
                # Suggest retry for certain errors
                retry_candidates = [entry for entry in failed_list if "timeout" in entry[2].lower() or "connection" in entry[2].lower()]
                if retry_candidates:
                    print(f"\n💡 {len(retry_candidates)} failures may be due to temporary network issues. Consider retrying those recipients.")
            
            if not stop_sending and sent_count == total_recipients:
                print("\n🎉 Campaign completed successfully!")

    except smtplib.SMTPAuthenticationError:
        with send_output:
            print("❌ Authentication failed!")
            print("💡 Please check your email and password.")
            print("   Make sure you're using an Office 365 account.")
            print("   You may need to enable 'Less secure app access' or use an App Password.")
    except smtplib.SMTPConnectError:
        with send_output:
            print("❌ Cannot connect to email server!")
            print("💡 Please check your internet connection.")
    except Exception as e:
        with send_output:
            print(f"🚨 Unexpected error: {str(e)}")
            print("💡 Please try again or contact support if the issue persists.")
    
    finally:
        # Re-enable buttons
        send_all_button.disabled = False
        stop_button.disabled = True
        confirm_button.disabled = True
        cancel_button.disabled = True

# --- Hook Up Buttons ---
send_all_button.on_click(show_confirmation)
stop_button.on_click(stop_sending_emails)
confirm_button.on_click(confirm_send)
cancel_button.on_click(cancel_send)

# --- Display Enhanced UI ---
display(HTML(value="""
<div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 10px; margin: 10px 0;">
<h2 style="margin: 0; text-align: center;">🚀 Personal SMTP Email Service - Step 3</h2>
<p style="text-align: center; margin: 5px 0;">Send your emails to all recipients</p>
</div>
"""))

display(HTML(value="""
<div style="background: #fff3cd; padding: 15px; border-radius: 8px; margin: 10px 0; border-left: 4px solid #ffc107;">
<h3 style="color: #856404; margin-top: 0;">⚠️ Safety Guidelines</h3>
<ul style="margin: 10px 0; color: #856404;">
<li><strong>Test First:</strong> Always send a test email from Step 2 before bulk sending</li>
<li><strong>Check Recipients:</strong> Verify your CSV data is correct</li>
<li><strong>Small Batches:</strong> For large lists (50+), consider breaking into smaller batches</li>
<li><strong>Monitor Progress:</strong> Watch for errors and use the stop button if needed</li>
<li><strong>Backup Plan:</strong> Keep a record of failed sends for follow-up</li>
</ul>
</div>
"""))

display(VBox([
    HTML(value="<h3 style='color: #dc3545;'>🚀 Bulk Email Sending</h3>"),
    HTML(value="<p style='color: #666; margin: 10px 0;'>Click the button below to start the confirmation process:</p>"),
    HBox([send_all_button, stop_button]),
    confirmation_output,
    HTML(value="<h4 style='color: #dc3545;'>Confirmation Required</h4>"),
    HTML(value="<p style='color: #666; margin: 5px 0;'>After clicking 'Send All Emails', you'll need to confirm:</p>"),
    HBox([confirm_button, cancel_button]),
    HTML(value="<hr style='border: 1px solid #dee2e6; margin: 20px 0;'>"),
    HTML(value="<h4 style='color: #007bff;'>📊 Sending Progress & Results</h4>"),
    send_output
]))
