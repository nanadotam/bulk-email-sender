# ImgBB Integration Guide

## Overview
This guide explains how to use ImgBB as an image hosting service for email attachments. ImgBB provides a simple API to upload images and get permanent URLs for email embedding.

## Prerequisites
- Python 3.x
- `requests` library (`pip install requests`)
- ImgBB API key (get it from [https://api.imgbb.com/](https://api.imgbb.com/))

## Setup
1. Get your API key:
   - Register at [ImgBB](https://imgbb.com/)
   - Go to [https://api.imgbb.com/](https://api.imgbb.com/)
   - Click "Get API key"

2. Install required package:
```bash
pip install requests
```

## Usage Example

```python
import requests

def upload_to_imgbb(image_path, api_key):
    """
    Upload an image to ImgBB and return the URL.
    
    Args:
        image_path (str): Path to the image file
        api_key (str): Your ImgBB API key
    
    Returns:
        str: Direct URL to the uploaded image
    """
    url = "https://api.imgbb.com/1/upload"
    
    with open(image_path, "rb") as file:
        payload = {
            "key": api_key,
            "image": file
        }
        response = requests.post(url, files=payload)
        
        if response.status_code == 200:
            return response.json()['data']['url']
        else:
            raise Exception("Upload failed")

# Example usage
API_KEY = "YOUR_API_KEY"
image_url = upload_to_imgbb("path/to/your/image.jpg", API_KEY)
```

## Using in Email HTML

```python
def create_email_content(image_url):
    html_content = f"""
    <html>
        <body>
            <img src="{image_url}" alt="Embedded Image" width="200" height="auto">
            <p>Your email content here...</p>
        </body>
    </html>
    """
    return html_content
```

## Important Notes

### Rate Limits
- Free tier allows up to 100 uploads/hour
- Maximum file size: 32 MB
- Supported formats: JPG, PNG, BMP, GIF

### Best Practices
1. Store image URLs after upload
2. Implement error handling
3. Compress images before upload
4. Use appropriate image dimensions for emails

### Security Considerations
- Never expose your API key in client-side code
- Store API key in environment variables:
```python
import os
API_KEY = os.getenv('IMGBB_API_KEY')
```

## Error Handling Example

```python
def safe_upload_to_imgbb(image_path, api_key):
    try:
        url = upload_to_imgbb(image_path, api_key)
        return url
    except Exception as e:
        print(f"Upload failed: {str(e)}")
        return None
```

## Complete Implementation Example

```python
import os
import requests
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def send_email_with_imgbb_image(to_email, subject, image_path):
    # Upload image
    api_key = os.getenv('IMGBB_API_KEY')
    image_url = safe_upload_to_imgbb(image_path, api_key)
    
    if not image_url:
        raise Exception("Failed to upload image")
    
    # Create email
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = "your@email.com"
    msg['To'] = to_email
    
    html_content = create_email_content(image_url)
    msg.attach(MIMEText(html_content, 'html'))
    
    # Send email using your existing send method
    # ... your email sending code here ...
```

## Troubleshooting

Common issues and solutions:
1. **Upload fails**: Check API key and file size
2. **Image not showing in email**: Verify URL is direct image link
3. **Rate limit exceeded**: Implement delay between uploads
