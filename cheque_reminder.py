import os
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
import requests
import io
import logging
import re
from urllib.parse import unquote

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class SharePointSharedLinkReminder:
    def __init__(self, sharepoint_shared_url, smtp_server, smtp_port, email_username, email_password, recipient_email):
        """
        Initialize the SharePoint Shared Link Reminder system
        
        Args:
            sharepoint_shared_url: SharePoint shared link URL
            smtp_server: Email server (e.g., 'smtp.gmail.com' for Gmail)
            smtp_port: Email port (587 for most)
            email_username: Email account username
            email_password: Email account password (or app password)
            recipient_email: Email address to send reminders to
        """
        self.sharepoint_shared_url = sharepoint_shared_url
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.email_username = email_username
        self.email_password = email_password
        self.recipient_email = recipient_email
        
    def convert_sharepoint_url_to_direct_download(self, shared_url):
        """
        Convert SharePoint shared URL to direct download URL
        """
        try:
            # Extract the file ID from the SharePoint URL
            # Pattern for OneDrive/SharePoint shared links
            if "sharepoint.com" in shared_url and ":x:" in shared_url:
                # Extract the direct download URL
                # Replace the view URL with download URL
                if "?e=" in shared_url:
                    base_url = shared_url.split("?e=")[0]
                else:
                    base_url = shared_url
                
                # Convert to direct download URL
                download_url = base_url.replace(":x:/g/personal/", ":x:/g/personal/").replace(":x:", ":b:")
                
                # Alternative method: try to get direct download link
                if "/_layouts/15/Doc.aspx?sourcedoc=" in shared_url:
                    # Already in the right format
                    download_url = shared_url
                else:
                    # Convert the sharing URL to download URL
                    # This might need adjustment based on the exact SharePoint setup
                    download_url = shared_url.replace("/personal/", "/personal/").replace(":x:", ":b:")
                    
                logging.info(f"Converted URL to: {download_url}")
                return download_url
            else:
                # If it's already a direct download URL or different format
                return shared_url
                
        except Exception as e:
            logging.error(f"Error converting SharePoint URL: {e}")
            return shared_url
    
    def download_excel_file(self):
        """Download Excel file from SharePoint shared link"""
        try:
            # Convert to download URL
            download_url = self.convert_sharepoint_url_to_direct_download(self.sharepoint_shared_url)
            
            # Set headers to mimic a browser request
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,*/*',
                'Accept-Language': 'en-US,en;q=0.9',
                'Accept-Encoding': 'gzip, deflate, br',
                'Connection': 'keep-alive',
                'Upgrade-Insecure-Requests': '1',
            }
            
            # Try multiple URL formats
            urls_to_try = [
                download_url,
                self.sharepoint_shared_url,
                # Add &download=1 parameter
                f"{download_url}&download=1" if "?" in download_url else f"{download_url}?download=1",
            ]
            
            for url in urls_to_try:
                try:
                    logging.info(f"Trying to download from: {url}")
                    response = requests.get(url, headers=headers, timeout=30, allow_redirects=True)
                    
                    if response.status_code == 200:
                        # Check if the response contains Excel data
                        content_type = response.headers.get('content-type', '').lower()
                        if ('excel' in content_type or 
                            'spreadsheet' in content_type or 
                            'vnd.openxmlformats' in content_type or
                            len(response.content) > 1000):  # Assume it's Excel if content is substantial
                            
                            bytes_file_obj = io.BytesIO(response.content)
                            logging.info("Successfully downloaded Excel file from SharePoint")
                            return bytes_file_obj
                        else:
                            logging.warning(f"Response doesn't seem to be Excel file. Content-Type: {content_type}")
                    else:
                        logging.warning(f"HTTP {response.status_code} for URL: {url}")
                        
                except requests.exceptions.RequestException as e:
                    logging.warning(f"Request failed for {url}: {e}")
                    continue
            
            # If all direct methods fail, try to extract the actual file URL from the page
            return self.extract_file_from_sharepoint_page()
            
        except Exception as e:
            logging.error(f"Failed to download Excel file: {e}")
            return None
    
    def extract_file_from_sharepoint_page(self):
        """Try to extract the actual file URL from SharePoint page"""
        try:
            logging.info("Attempting to extract file URL from SharePoint page...")
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
            
            response = requests.get(self.sharepoint_shared_url, headers=headers, timeout=30)
            
            if response.status_code == 200:
                # Look for download links in the page content
                content = response.text
                
                # Try to find direct download URLs in the page
                download_patterns = [
                    r'"downloadUrl":"([^"]+)"',
                    r'"@microsoft.graph.downloadUrl":"([^"]+)"',
                    r'data-downloadurl="([^"]+)"',
                ]
                
                for pattern in download_patterns:
                    matches = re.findall(pattern, content)
                    if matches:
                        download_url = matches[0].replace('\\u0026', '&')
                        logging.info(f"Found download URL: {download_url}")
                        
                        # Try to download from this URL
                        file_response = requests.get(download_url, headers=headers, timeout=30)
                        if file_response.status_code == 200:
                            return io.BytesIO(file_response.content)
            
            logging.error("Could not extract download URL from SharePoint page")
            return None
            
        except Exception as e:
            logging.error(f"Failed to extract file from SharePoint page: {e}")
            return None
    
    def parse_excel_data(self, excel_file):
        """Parse Excel file and return relevant data"""
        try:
            # Try reading the Excel file
            try:
                df = pd.read_excel(excel_file, engine='openpyxl')
            except Exception as e:
                logging.warning(f"Failed with openpyxl, trying xlrd: {e}")
                excel_file.seek(0)  # Reset file pointer
                df = pd.read_excel(excel_file, engine='xlrd')
            
            logging.info(f"Excel file loaded successfully. Shape: {df.shape}")
            logging.info(f"Columns found: {list(df.columns)}")
            
            # Check if required columns exist (case-insensitive search)
            required_columns = ['Mode of Payment', 'Date of Transfer']
            column_mapping = {}
            
            for req_col in required_columns:
                found_col = None
                for df_col in df.columns:
                    if str(df_col).lower().strip() == req_col.lower().strip():
                        found_col = df_col
                        break
                
                if found_col is None:
                    # Try partial matching
                    for df_col in df.columns:
                        if req_col.lower().replace(' ', '').replace('_', '') in str(df_col).lower().replace(' ', '').replace('_', ''):
                            found_col = df_col
                            break
                
                if found_col:
                    column_mapping[req_col] = found_col
                    logging.info(f"Mapped '{req_col}' to column '{found_col}'")
                else:
                    logging.error(f"Required column '{req_col}' not found in Excel file")
                    logging.error(f"Available columns: {list(df.columns)}")
                    return None
            
            # Rename columns to standardized names
            df = df.rename(columns=column_mapping)
            
            # Filter for cheque payments only (case-insensitive)
            df['Mode of Payment'] = df['Mode of Payment'].astype(str)
            cheque_df = df[df['Mode of Payment'].str.lower().str.contains('cheque|check', na=False)].copy()
            
            if cheque_df.empty:
                logging.info("No cheque payments found")
                return pd.DataFrame()
            
            # Convert Date of Transfer to datetime with multiple format attempts
            date_formats = ['%d-%b-%y', '%d-%b-%Y', '%d/%m/%y', '%d/%m/%Y', '%Y-%m-%d', '%m/%d/%Y']
            
            for date_format in date_formats:
                try:
                    cheque_df['Date of Transfer'] = pd.to_datetime(cheque_df['Date of Transfer'], 
                                                                 format=date_format, errors='coerce')
                    valid_dates = cheque_df['Date of Transfer'].notna().sum()
                    if valid_dates > 0:
                        logging.info(f"Successfully parsed {valid_dates} dates using format {date_format}")
                        break
                except:
                    continue
            else:
                # If no format worked, try pandas' automatic parsing
                try:
                    cheque_df['Date of Transfer'] = pd.to_datetime(cheque_df['Date of Transfer'], errors='coerce')
                    logging.info("Used pandas automatic date parsing")
                except:
                    logging.error("Failed to parse any dates")
            
            # Remove rows with invalid dates
            initial_count = len(cheque_df)
            cheque_df = cheque_df.dropna(subset=['Date of Transfer'])
            final_count = len(cheque_df)
            
            logging.info(f"Found {final_count} cheque payments with valid dates (filtered from {initial_count})")
            return cheque_df
            
        except Exception as e:
            logging.error(f"Failed to parse Excel file: {e}")
            return None
    
    def find_reminders_needed(self, df):
        """Find cheques that need reminders (3 days before transfer date)"""
        if df is None or df.empty:
            return pd.DataFrame()
        
        today = datetime.now().date()
        target_date = today + timedelta(days=3)
        
        # Find records where transfer date is exactly 3 days from today
        reminders_needed = df[df['Date of Transfer'].dt.date == target_date].copy()
        
        logging.info(f"Found {len(reminders_needed)} reminders needed for {target_date}")
        return reminders_needed
    
    def send_email_reminder(self, reminder_data):
        """Send email reminder"""
        if reminder_data.empty:
            logging.info("No reminders to send")
            return True
        
        try:
            # Create email message
            msg = MIMEMultipart()
            msg['From'] = self.email_username
            msg['To'] = self.recipient_email
            msg['Subject'] = f"Cheque Transfer Reminder - {len(reminder_data)} payment(s) due in 3 days"
            
            # Create email body
            body = self.create_email_body(reminder_data)
            msg.attach(MIMEText(body, 'html'))
            
            # Send email
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.email_username, self.email_password)
            server.send_message(msg)
            server.quit()
            
            logging.info(f"Email reminder sent successfully to {self.recipient_email}")
            return True
            
        except Exception as e:
            logging.error(f"Failed to send email: {e}")
            return False
    
    def create_email_body(self, reminder_data):
        """Create HTML email body"""
        html_body = """
        <html>
        <body>
            <h2>🏦 Cheque Transfer Reminder</h2>
            <p>The following cheque payments are due for transfer in <strong>3 days</strong>:</p>
            <table border="1" style="border-collapse: collapse; width: 100%; margin: 20px 0;">
                <thead>
                    <tr style="background-color: #f2f2f2;">
        """
        
        # Add headers
        for col in reminder_data.columns:
            html_body += f"<th style='padding: 12px; text-align: left; font-weight: bold;'>{col}</th>"
        
        html_body += """
                    </tr>
                </thead>
                <tbody>
        """
        
        # Add data rows
        for i, (_, row) in enumerate(reminder_data.iterrows()):
            row_style = "background-color: #f9f9f9;" if i % 2 == 0 else ""
            html_body += f"<tr style='{row_style}'>"
            for col in reminder_data.columns:
                value = row[col]
                if pd.isna(value):
                    value = ""
                elif col == 'Date of Transfer':
                    value = value.strftime('%d-%b-%Y') if hasattr(value, 'strftime') else str(value)
                html_body += f"<td style='padding: 10px; border-bottom: 1px solid #ddd;'>{value}</td>"
            html_body += "</tr>"
        
        html_body += f"""
                </tbody>
            </table>
            <div style="margin-top: 20px; padding: 15px; background-color: #e8f4fd; border-left: 4px solid #2196F3;">
                <p><strong>📅 Reminder Date:</strong> {datetime.now().strftime('%d-%b-%Y')}</p>
                <p><strong>🎯 Target Transfer Date:</strong> {(datetime.now() + timedelta(days=3)).strftime('%d-%b-%Y')}</p>
            </div>
            <br>
            <p style="color: #666; font-size: 14px;">
                <em>⚡ This is an automated reminder generated from your SharePoint file.</em><br>
                <em>Please ensure these cheques are processed on time to avoid any delays.</em>
            </p>
        </body>
        </html>
        """
        
        return html_body
    
    def run_reminder_check(self):
        """Main method to run the reminder check"""
        logging.info("🚀 Starting SharePoint cheque reminder check...")
        
        # Download Excel file
        excel_file = self.download_excel_file()
        if not excel_file:
            logging.error("❌ Failed to download Excel file")
            return False
        
        # Parse Excel data
        df = self.parse_excel_data(excel_file)
        if df is None:
            logging.error("❌ Failed to parse Excel data")
            return False
        
        # Find reminders needed
        reminders = self.find_reminders_needed(df)
        
        # Send email if reminders needed
        if not reminders.empty:
            success = self.send_email_reminder(reminders)
            if success:
                logging.info("✅ Reminder email sent successfully")
            else:
                logging.error("❌ Failed to send reminder email")
            return success
        else:
            logging.info("ℹ️  No cheque transfer reminders needed today")
            return True

def main():
    """Main function - reads configuration from environment variables"""
    logging.info("📋 SharePoint Cheque Reminder Starting...")
    
    # Read configuration from environment variables
    config = {
        'sharepoint_shared_url': os.getenv('SHAREPOINT_SHARED_URL'),
        'smtp_server': os.getenv('SMTP_SERVER'),
        'smtp_port': int(os.getenv('SMTP_PORT', 587)),
        'email_username': os.getenv('EMAIL_USERNAME'),
        'email_password': os.getenv('EMAIL_PASSWORD'),
        'recipient_email': os.getenv('RECIPIENT_EMAIL')
    }
    
    # Validate that all required config is present
    missing_config = [key for key, value in config.items() if value is None or value == '']
    if missing_config:
        logging.error(f"❌ Missing required environment variables: {missing_config}")
        return False
    
    logging.info("✅ Configuration loaded successfully")
    
    # Create and run reminder system
    reminder_system = SharePointSharedLinkReminder(**config)
    success = reminder_system.run_reminder_check()
    
    if success:
        logging.info("🎉 Reminder check completed successfully")
    else:
        logging.error("💥 Reminder check failed")
        
    return success

if __name__ == "__main__":
    success = main()
    # Exit with error code if failed (helps GitHub Actions detect failures)
    exit(0 if success else 1)