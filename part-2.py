import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
import ipywidgets as widgets
from IPython.display import display, clear_output
import re
import io

# -- Globals from Step 1 --
# Assume df is already loaded, and st_placeholders contains dynamic columns
# E.g., st_placeholders = ['event', 'location', 'date']

# --- Widgets ---
sender_email_input = widgets.Text(
    placeholder='yourname@example.com', description='Email:', layout=widgets.Layout(width='100%')
)
password_input = widgets.Password(description='Password:', layout=widgets.Layout(width='100%'))

subject_input = widgets.Text(
    value="🎉 Special Invite Just for You", description='Subject:', layout=widgets.Layout(width='100%')
)

salutation_input = widgets.Textarea(
    value="Dear {{name}},", description="Salutation:", layout=widgets.Layout(width='100%', height='60px')
)

body_input = widgets.Textarea(
    value="This is a sample email body. Edit here...",
    description="Body:", layout=widgets.Layout(width='100%', height='200px')
)

signature_input = widgets.Textarea(
    value="— Your ASC Family", description="Signature:", layout=widgets.Layout(width='100%', height='50px')
)

logo_uploader = widgets.FileUpload(accept='image/*', multiple=False, description="🏢 Upload Custom Logo (Optional)")
logo_preview_output = widgets.Output()
embedded_images_uploader = widgets.FileUpload(accept='image/*', multiple=True, description="📷 Upload Embedded Images")
attachments_uploader = widgets.FileUpload(accept='.pdf,image/*', multiple=True, description="📎 Upload Attachments (≤5MB)")
preview_button = widgets.Button(description="👀 Preview Email", button_style='info')
test_button = widgets.Button(description="📨 Send Test Email", button_style='warning')
test_email_input = widgets.Text(placeholder='Enter your email', description='Test To:')
preview_output = widgets.Output()
test_output = widgets.Output()

# --- Logo Preview Handler ---
def update_logo_preview(change):
    logo_preview_output.clear_output()
    with logo_preview_output:
        if logo_uploader.value:
            import base64
            logo_file = list(logo_uploader.value.values())[0]
            logo_content = logo_file['content']
            logo_base64 = base64.b64encode(logo_content).decode('utf-8')
            logo_name = logo_file['metadata']['name']
            
            if logo_name.lower().endswith('.png'):
                data_url = f"data:image/png;base64,{logo_base64}"
            elif logo_name.lower().endswith(('.jpg', '.jpeg')):
                data_url = f"data:image/jpeg;base64,{logo_base64}"
            else:
                data_url = f"data:image/png;base64,{logo_base64}"
            
            print(f"📄 Custom logo: {logo_name}")
            display(widgets.HTML(value=f'<img src="{data_url}" style="max-width: 120px; height: auto; border: 1px solid #ddd; padding: 5px;" />'))
        else:
            print("📄 Default ASC logo will be used")
            display(widgets.HTML(value='<img src="https://raw.githubusercontent.com/nanadotam/1nri-photo/main/temp/ASC%20LOGO%20PNG.png" style="max-width: 120px; height: auto; border: 1px solid #ddd; padding: 5px;" />'))

logo_uploader.observe(update_logo_preview, names='value')

# Initialize logo preview with default
update_logo_preview(None)

# --- Markdown Formatter ---
def markdown_to_html(text):
    # Bold + Italic (***text***)
    text = re.sub(r'\*\*\*(.*?)\*\*\*', r'<strong><em>\1</em></strong>', text)
    # Bold (**text**)
    text = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', text)
    # Italic (_text_)
    text = re.sub(r'_(.*?)_', r'<em>\1</em>', text)
    # Strikethrough (~~text~~)
    text = re.sub(r'~~(.*?)~~', r'<del>\1</del>', text)
    # Monospace (`text`)
    text = re.sub(r'`(.*?)`', r'<code>\1</code>', text)
    # Horizontal Rule (---)
    text = re.sub(r'^---$', r'<hr>', text, flags=re.MULTILINE)
    # Blockquote (> text)
    text = re.sub(r'^> (.*)$', r'<blockquote>\1</blockquote>', text, flags=re.MULTILINE)
    # Links ([text](url))
    text = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'<a href="\2">\1</a>', text)
    # Images (![alt](url))
    text = re.sub(r'!\[([^\]]*)\]\(([^)]+)\)', r'<img src="\2" alt="\1" style="max-width: 100%;">', text)
    # Unordered list (- item)
    text = re.sub(r'^- (.*)$', r'<ul><li>\1</li></ul>', text, flags=re.MULTILINE)
    # Ordered list (1. item)
    text = re.sub(r'^\d+\. (.*)$', r'<ol><li>\1</li></ol>', text, flags=re.MULTILINE)
    # Line breaks
    text = re.sub(r'\n', r'<br>', text)
    return text

# --- Dynamic HTML Email Generator ---
def generate_email_html(row, embedded_images=None):
    import base64
    
    # Check if custom logo is uploaded, otherwise use default
    if logo_uploader.value:
        # Use uploaded custom logo
        logo_file = list(logo_uploader.value.values())[0]
        logo_content = logo_file['content']
        logo_base64 = base64.b64encode(logo_content).decode('utf-8')
        # Detect image type
        if logo_file['metadata']['name'].lower().endswith('.png'):
            logo_url = f"data:image/png;base64,{logo_base64}"
        elif logo_file['metadata']['name'].lower().endswith(('.jpg', '.jpeg')):
            logo_url = f"data:image/jpeg;base64,{logo_base64}"
        else:
            logo_url = f"data:image/png;base64,{logo_base64}"  # Default to PNG
        logo_alt = "Custom Logo"
    else:
        # Use default ASC logo
        logo_url = "https://raw.githubusercontent.com/nanadotam/1nri-photo/main/temp/ASC%20LOGO%20PNG.png"
        logo_alt = "ASC Logo"

    def apply_placeholders(text):
        for col in df.columns:
            placeholder = f"{{{{{col}}}}}"
            value = str(row.get(col, ''))
            text = text.replace(placeholder, value)
        return markdown_to_html(text)

    salutation = apply_placeholders(salutation_input.value)
    body = apply_placeholders(body_input.value)

    embedded_html = ""
    if embedded_images:
        for idx in range(len(embedded_images)):
            embedded_html += f'<img src="cid:image{idx}" style="max-width: 100%; margin-top: 20px;" /><br/>'

    return f"""
    <html>
      <head>
        <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@1,400&display=swap" rel="stylesheet">
      </head>
      <body style="margin: 0; padding: 0; background-color: #f8f8f8; font-family: 'Times New Roman', serif;">
        <div style="max-width: 700px; margin: 40px auto; background-color: white; padding: 14px;">
          <div style="border: 2px solid #83142A; padding: 30px;">

            <div style="text-align: center; margin-bottom: 30px;">
              <img src="{logo_url}" alt="{logo_alt}" style="max-width: 120px; height: auto;" />
            </div>

            <div style="font-family: 'Playfair Display', 'Times New Roman', serif; font-style: italic; font-size: 20px; margin-bottom: 20px; color: #262626;">
              {salutation}
            </div>

            <div style="font-size: 16px; line-height: 1.7; color: #222;">
              {body}
            </div>

            {embedded_html}

            <div style="margin-top: 30px; font-family: 'Playfair Display', 'Times New Roman', serif; font-style: italic; font-size: 16px; color: #262626;">
              {apply_placeholders(signature_input.value)}
            </div>

          </div>
        </div>
      </body>
    </html>
    """

# --- Format Button Logic ---
def insert_format(tag):
    original = body_input.value
    insert = ""

    if tag == 'bold':
        insert = "**bold text**"
    elif tag == 'italic':
        insert = "_italic text_"
    elif tag == 'bolditalic':
        insert = "***bold italic***"
    elif tag == 'strike':
        insert = "~~strikethrough~~"
    elif tag == 'mono':
        insert = "`monospace`"
    elif tag == 'hrule':
        insert = "\n---\n"
    elif tag == 'blockquote':
        insert = "> quoted text"
    elif tag == 'ulist':
        insert = "- item"
    elif tag == 'olist':
        insert = "1. item"
    elif tag == 'link':
        insert = "[link text](https://example.com)"
    elif tag == 'image':
        insert = "![alt text](https://example.com/image.png)"
    elif tag == 'linebreak':
        insert = "\n"

    # Append the formatting to the end of current content
    body_input.value = original + insert

# --- Buttons ---
bold_btn = widgets.Button(description="B", button_style='primary')
italic_btn = widgets.Button(description="I", button_style='info')
bolditalic_btn = widgets.Button(description="B+I", button_style='primary')
strike_btn = widgets.Button(description="S", button_style='warning')
mono_btn = widgets.Button(description="Mono", button_style='')
hrule_btn = widgets.Button(description="HR", button_style='')
blockquote_btn = widgets.Button(description="❝", button_style='')
ulist_btn = widgets.Button(description="•", button_style='')
olist_btn = widgets.Button(description="1.", button_style='')
link_btn = widgets.Button(description="🔗", button_style='')
image_btn = widgets.Button(description="🖼️", button_style='')
line_btn = widgets.Button(description="↵", button_style='')

# --- Bind Clicks ---
bold_btn.on_click(lambda b: insert_format('bold'))
italic_btn.on_click(lambda b: insert_format('italic'))
bolditalic_btn.on_click(lambda b: insert_format('bolditalic'))
strike_btn.on_click(lambda b: insert_format('strike'))
mono_btn.on_click(lambda b: insert_format('mono'))
hrule_btn.on_click(lambda b: insert_format('hrule'))
blockquote_btn.on_click(lambda b: insert_format('blockquote'))
ulist_btn.on_click(lambda b: insert_format('ulist'))
olist_btn.on_click(lambda b: insert_format('olist'))
link_btn.on_click(lambda b: insert_format('link'))
image_btn.on_click(lambda b: insert_format('image'))
line_btn.on_click(lambda b: insert_format('linebreak'))

# --- Preview Handler ---
def preview_email(b):
    preview_output.clear_output()
    if df.empty:
        with preview_output:
            print("⚠️ Upload a CSV first.")
        return
    email_html = generate_email_html(df.iloc[0], embedded_images_uploader.value)
    with preview_output:
        # Show logo status
        if logo_uploader.value:
            logo_name = list(logo_uploader.value.keys())[0]
            print(f"🏢 Using custom logo: {logo_name}")
        else:
            print("🏢 Using default ASC logo")
        
        display(widgets.HTML(value=email_html))
        
        if embedded_images_uploader.value:
            print(f"✅ Embedded {len(embedded_images_uploader.value)} image(s).")
        if attachments_uploader.value:
            print(f"📎 {len(attachments_uploader.value)} file attachment(s) ready.")

# --- Test Email Sender ---
def send_test_email(b):
    test_output.clear_output()
    recipient = test_email_input.value.strip()
    if not re.match(r"[^@]+@[^@]+\.[^@]+", recipient):
        with test_output:
            print("❌ Invalid test email")
        return

    msg = MIMEMultipart("related")
    msg["Subject"] = subject_input.value
    msg["From"] = sender_email_input.value
    msg["To"] = recipient

    html_content = generate_email_html(df.iloc[0], embedded_images_uploader.value)
    alt_part = MIMEMultipart("alternative")
    alt_part.attach(MIMEText(html_content, "html"))
    msg.attach(alt_part)

    for idx, (fname, filedata) in enumerate(embedded_images_uploader.value.items()):
        img = MIMEImage(filedata['content'])
        img.add_header('Content-ID', f'<image{idx}>')
        img.add_header('Content-Disposition', 'inline', filename=fname)
        msg.attach(img)

    for fname, filedata in attachments_uploader.value.items():
        if len(filedata['content']) <= 5 * 1024 * 1024:
            part = MIMEApplication(filedata['content'])
            part.add_header('Content-Disposition', 'attachment', filename=fname)
            msg.attach(part)
        else:
            with test_output:
                print(f"⚠️ Skipped large file: {fname}")

    try:
        server = smtplib.SMTP("smtp.office365.com", 587)
        server.starttls()
        server.login(sender_email_input.value, password_input.value)
        server.sendmail(sender_email_input.value, recipient, msg.as_string())
        server.quit()
        with test_output:
            print("✅ Test email sent successfully!")
    except Exception as e:
        with test_output:
            print(f"❌ Failed to send test email: {e}")

# --- Hooks ---
preview_button.on_click(preview_email)
test_button.on_click(send_test_email)

# --- Display UI with enhanced organization ---
display(widgets.HTML(value="""
<div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 10px; margin: 10px 0;">
<h2 style="margin: 0; text-align: center;">✍️ Personal - ASC SMTP Email Service - Step 2</h2>
<p style="text-align: center; margin: 5px 0;">Compose and customize your email template</p>
</div>
"""))

# Quick Start Guide
display(widgets.HTML(value="""
<div style="background: #e8f5e8; padding: 15px; border-radius: 8px; margin: 10px 0; border-left: 4px solid #28a745;">
<h3 style="color: #28a745; margin-top: 0;">🎯 Quick Start Guide</h3>
<ol style="margin: 10px 0;">
<li><strong>Email Settings:</strong> Enter your Office 365 email and password</li>
<li><strong>Subject Line:</strong> Write your email subject</li>
<li><strong>Logo (Optional):</strong> Upload your company logo or leave empty for ASC logo</li>
<li><strong>Email Content:</strong> Write your message using placeholders from Step 1</li>
<li><strong>Test First:</strong> Always send a test email to yourself before bulk sending</li>
</ol>
</div>
"""))

display(widgets.VBox([
    # Email Settings Section
    widgets.HTML(value="<hr style='border: 2px solid #dc3545; margin: 20px 0;'>"),
    widgets.HTML(value="<h3 style='color: #dc3545;'>🔐 Email Settings</h3>"),
    widgets.HTML(value="<p style='color: #666; margin: 10px 0;'>Enter your Office 365 email credentials (your password is secure and not stored):</p>"),
    sender_email_input,
    password_input,
    
    # Subject Section
    widgets.HTML(value="<hr style='border: 2px solid #6f42c1; margin: 20px 0;'>"),
    widgets.HTML(value="<h3 style='color: #6f42c1;'>📝 Email Subject</h3>"),
    widgets.HTML(value="<p style='color: #666; margin: 10px 0;'>Write a compelling subject line for your email:</p>"),
    subject_input,
    
    # Branding Section
    widgets.HTML(value="<hr style='border: 2px solid #fd7e14; margin: 20px 0;'>"),
    widgets.HTML(value="<h3 style='color: #fd7e14;'>🎨 Branding & Content</h3>"),
    widgets.HTML(value="<p style='color: #666; margin: 10px 0;'>Customize your email's appearance and content:</p>"),
    logo_uploader,
    logo_preview_output,
    widgets.HTML(value="<small style='color: #666;'>💡 Leave empty to use default ASC logo. Recommended size: 120px width</small>"),
    
    # Salutation Section
    widgets.HTML(value="<br><h4 style='color: #fd7e14;'>👋 Salutation</h4>"),
    widgets.HTML(value="<p style='color: #666; margin: 5px 0;'>How to greet your recipients (use placeholders from Step 1):</p>"),
    salutation_input,
    
    # Formatting Toolbar
    widgets.HTML(value="<br><h4 style='color: #fd7e14;'>🔧 Formatting Toolbar</h4>"),
    widgets.HTML(value="<p style='color: #666; margin: 5px 0;'>Click buttons to add formatting to your email body:</p>"),
    widgets.HBox([
        bold_btn, italic_btn, bolditalic_btn, strike_btn, mono_btn,
        hrule_btn, blockquote_btn, ulist_btn, olist_btn, link_btn, image_btn, line_btn
    ]),
    
    # Email Body
    widgets.HTML(value="<br><h4 style='color: #fd7e14;'>📄 Email Body</h4>"),
    widgets.HTML(value="""
    <div style="background: #fff3cd; padding: 10px; border-radius: 5px; margin: 5px 0;">
    <strong>💡 Formatting Tips:</strong>
    <ul style="margin: 5px 0;">
    <li>Use placeholders like {{Name}}, {{Email}}, {{Event}} from your CSV</li>
    <li>Use **text** for bold, _text_ for italic</li>
    <li>Use [link text](https://example.com) for links</li>
    <li>Click formatting buttons above to add styling</li>
    </ul>
    </div>
    """),
    body_input,
    
    # Signature Section
    widgets.HTML(value="<br><h4 style='color: #fd7e14;'>✍️ Email Signature</h4>"),
    widgets.HTML(value="<p style='color: #666; margin: 5px 0;'>Add your closing signature:</p>"),
    signature_input,
    
    # Attachments Section
    widgets.HTML(value="<hr style='border: 2px solid #20c997; margin: 20px 0;'>"),
    widgets.HTML(value="<h3 style='color: #20c997;'>📎 Attachments & Media</h3>"),
    widgets.HTML(value="<p style='color: #666; margin: 10px 0;'>Add images and files to your email:</p>"),
    embedded_images_uploader,
    widgets.HTML(value="<small style='color: #666;'>Images will appear in the email body</small>"),
    attachments_uploader,
    widgets.HTML(value="<small style='color: #666;'>Files will be attached to the email (max 5MB each)</small>"),
    
    # Testing Section
    widgets.HTML(value="<hr style='border: 2px solid #007bff; margin: 20px 0;'>"),
    widgets.HTML(value="<h3 style='color: #007bff;'>🧪 Testing & Preview</h3>"),
    widgets.HTML(value="""
    <div style="background: #d1ecf1; padding: 10px; border-radius: 5px; margin: 10px 0;">
    <strong>⚠️ Important:</strong> Always preview and test your email before sending to all recipients!
    </div>
    """),
    widgets.HBox([preview_button, test_button]),
    widgets.HTML(value="<p style='color: #666; margin: 5px 0;'>Enter your email address for testing:</p>"),
    test_email_input,
    preview_output,
    test_output
]))
