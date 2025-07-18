import pandas as pd
import re
from ipywidgets import FileUpload, Dropdown, VBox, Output, Label, Button, HTML, HBox
from IPython.display import display, clear_output

# Widgets
upload_widget = FileUpload(accept='.csv', multiple=False)
output_preview = Output()
column_selector_name = Dropdown(description="Name column")
column_selector_email = Dropdown(description="Email column")
submit_columns_btn = Button(description="✅ Confirm Columns", button_style='success')
validation_output = Output()

# Global DataFrame
df = pd.DataFrame()

def handle_upload(change):
  clear_output(wait=True)
  output_preview.clear_output()
  global df
  if upload_widget.value:
      try:
          file_info = list(upload_widget.value.values())[0]
          content = file_info['content']
          df = pd.read_csv(pd.io.common.BytesIO(content))
          
          with output_preview:
              print("✅ CSV File Uploaded Successfully!")
              print(f"📊 Found {len(df)} rows and {len(df.columns)} columns")
              
              # Auto-detect likely name and email columns
              name_candidates = [col for col in df.columns if any(word in col.lower() for word in ['name', 'full', 'first', 'last'])]
              email_candidates = [col for col in df.columns if 'email' in col.lower() or 'mail' in col.lower()]
              
              if name_candidates:
                  print(f"💡 Auto-detected Name column: '{name_candidates[0]}'")
              if email_candidates:
                  print(f"💡 Auto-detected Email column: '{email_candidates[0]}'")
              if not name_candidates or not email_candidates:
                  print("⚠️ Couldn't auto-detect name/email columns. Please select manually below.")
              
              print("\n👀 Preview of your data:")
              display(df.head())
              
              print("\n📝 Available placeholders for your email template:")
              for col in df.columns:
                  print(f"   {{{{ {col} }}}}")
              print("\n💡 Copy these exact placeholder names to use in Step 2!")
              
          # Update dropdown options
          columns = df.columns.tolist()
          column_selector_name.options = columns
          column_selector_email.options = columns
          
          # Auto-select if we found candidates
          if name_candidates:
              column_selector_name.value = name_candidates[0]
          if email_candidates:
              column_selector_email.value = email_candidates[0]
              
          display(VBox([
              output_preview,
              HTML(value="<hr style='border: 2px solid #007acc; margin: 20px 0;'>"),
              HTML(value="<h3 style='color: #007acc;'>📋 Step 1.2: Select Your Columns</h3>"),
              HTML(value="<p style='color: #666; margin: 10px 0;'>Choose which columns contain the name and email addresses for your recipients:</p>"),
              HBox([column_selector_name, column_selector_email]),
              submit_columns_btn,
              validation_output
          ]))
          
      except Exception as e:
          with output_preview:
              print("❌ Error reading CSV file!")
              print(f"Error details: {str(e)}")
              print("\n💡 Troubleshooting steps:")
              print("   ✓ Make sure your file is a valid CSV format")
              print("   ✓ Ensure the file has column headers in the first row")
              print("   ✓ Check that the file isn't corrupted or empty")
              print("   ✓ Try saving your Excel file as CSV if needed")
              
  # Detect available placeholders from columns
  if not df.empty:
      placeholder_candidates = [col for col in df.columns if col not in [column_selector_name.value, column_selector_email.value]]
      st_placeholders = placeholder_candidates  # Save for Step 2 usage


def validate_email_list(name_col, email_col):
  validation_output.clear_output()
  if df.empty:
      return
  errors = []
  for idx, row in df.iterrows():
      email = str(row[email_col]).strip()
      name = str(row[name_col]).strip()
      if not re.match(r"[^@]+@[^@]+\.[^@]+", email):
          errors.append((idx+1, name, email))
  with validation_output:
      if errors:
          print("❌ Invalid email addresses found:")
          for row_num, name, email in errors:
              print(f"Row {row_num}: {name} - {email}")
      else:
          print("✅ All emails are valid!")

submit_columns_btn.on_click(lambda b: validate_email_list(
    column_selector_name.value, column_selector_email.value
))
upload_widget.observe(handle_upload, names='value')

# Display upload UI with enhanced guidance
display(HTML(value="""
<div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 10px; margin: 10px 0;">
<h2 style="margin: 0; text-align: center;">📧 Personal - ASC SMTP Email Service - Step 1</h2>
<p style="text-align: center; margin: 5px 0;">Upload your contact list to get started</p>
</div>
"""))

display(HTML(value="""
<div style="background: #f8f9fa; padding: 15px; border-radius: 8px; margin: 10px 0; border-left: 4px solid #007acc;">
<h3 style="color: #007acc; margin-top: 0;">📋 CSV File Requirements</h3>
<ul style="margin: 10px 0;">
<li><strong>Format:</strong> Must be a .csv file (not .xlsx or .xls)</li>
<li><strong>Headers:</strong> First row should contain column names</li>
<li><strong>Required columns:</strong> At least one for names and one for email addresses</li>
<li><strong>Example format:</strong></li>
</ul>
<pre style="background: #e9ecef; padding: 10px; border-radius: 4px; font-family: monospace;">
Name,Email,Event,Location
John Doe,john@example.com,Conference,New York
Jane Smith,jane@example.com,Conference,New York
</pre>
</div>
"""))

display(HTML(value="<h3 style='color: #007acc;'>📤 Step 1.1: Upload Your CSV File</h3>"))
display(upload_widget)
