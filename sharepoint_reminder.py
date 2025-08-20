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
from urllib.parse import unquote, quote
import time

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
        Multiple methods to handle different SharePoint URL formats
        """
        try:
            logging.info(f"Original URL: {shared_url}")
            
            # Method 1: Try to extract direct download URL by replacing view with download
            if "sharepoint.com" in shared_url and ("/:x:/" in shared_url or "/:b:/" in shared_url):
                # Replace :x: with :b: and add download=1
                if "/:x:/" in shared_url:
                    download_url = shared_url.replace("/:x:/", "/:b:/")
                else:
                    download_url = shared_url
                    
                # Add download parameter if not present
                if "download=1" not in download_url:
                    separator = "&" if "?" in download_url else "?"
                    download_url = f"{download_url}{separator}download=1"
                    
                logging.info(f"Method 1 - Download URL: {download_url}")
                return download_url
                
            # Method 2: Handle different SharePoint URL formats
            elif "_layouts/15/Doc.aspx" in shared_url:
                # Already in a usable format, just add download parameter
                download_url = shared_url
                if "download=1" not in download_url:
                    separator = "&" if "?" in download_url else "?"
                    download_url = f"{download_url}{separator}download=1"
                logging.info(f"Method 2 - Download URL: {download_url}")
                return download_url
                
            # Method 3: Try to construct download URL from sharing URL
            elif "sharepoint.com" in shared_url:
                # Extract the file ID and construct download URL
                parts = shared_url.split('/')
                for i, part in enumerate(parts):
                    if part == "personal":
                        try:
                            # Reconstruct with download format
                            base_parts = parts[:i+2]  # Include domain and personal/username
                            base_url = "/".join(base_parts)
                            download_url = f"{base_url}/_layouts/15/download.aspx?SourceUrl={quote(shared_url)}"
                            logging.info(f"Method 3 - Download URL: {download_url}")
                            return download_url
                        except:
                            pass
            
            # If no conversion worked, return original URL
            logging.info("No URL conversion applied, using original URL")
            return shared_url
                
        except Exception as e:
            logging.error(f"Error converting SharePoint URL: {e}")
            return shared_url
    
    def download_excel_file(self):
        """Download Excel file from SharePoint shared link with multiple strategies"""
        try:
            # Strategy 1: Try different URL formats
            urls_to_try = [
                self.sharepoint_shared_url,
                self.convert_sharepoint_url_to_direct_download(self.sharepoint_shared_url),
            ]
            
            # Add more URL variations
            base_url = self.sharepoint_shared_url.split('?')[0] if '?' in self.sharepoint_shared_url else self.sharepoint_shared_url
            urls_to_try.extend([
                f"{base_url}?download=1",
                f"{self.sharepoint_shared_url}&download=1" if "?" in self.sharepoint_shared_url else f"{self.sharepoint_shared_url}?download=1",
            ])
            
            # Remove duplicates while preserving order
            urls_to_try = list(dict.fromkeys(urls_to_try))
            
            # Set headers to mimic a browser request
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,application/octet-stream,*/*',
                'Accept-Language': 'en-US,en;q=0.9',
                'Accept-Encoding': 'gzip, deflate, br',
                'Connection': 'keep-alive',
                'Upgrade-Insecure-Requests': '1',
                'Sec-Fetch-Dest': 'document',
                'Sec-Fetch-Mode': 'navigate',
                'Sec-Fetch-Site': 'none',
            }
            
            session = requests.Session()
            
            for i, url in enumerate(urls_to_try):
                try:
                    logging.info(f"Attempt {i+1}: Trying to download from: {url}")
                    
                    response = session.get(url, headers=headers, timeout=30, allow_redirects=True)
                    logging.info(f"Response status: {response.status_code}")
                    logging.info(f"Content-Type: {response.headers.get('content-type', 'N/A')}")
                    logging.info(f"Content-Length: {len(response.content)} bytes")
                    
                    if response.status_code == 200:
                        content_type = response.headers.get('content-type', '').lower()
                        
                        # Check if response is HTML (SharePoint login/error page)
                        if 'text/html' in content_type or response.content.startswith(b'<!DOCTYPE') or response.content.startswith(b'<html'):
                            logging.warning(f"Received HTML instead of Excel file. First 200 chars: {response.content[:200]}")
                            continue
                        
                        # Check if the response looks like an Excel file
                        if (len(response.content) > 1000 and  # Must have substantial content
                            ('excel' in content_type or 
                             'spreadsheet' in content_type or 
                             'vnd.openxmlformats' in content_type or
                             'application/octet-stream' in content_type or
                             response.content.startswith(b'PK') or  # ZIP/Excel signature
                             response.content.startswith(b'\xd0\xcf\x11\xe0'))):  # Old Excel signature
                            
                            logging.info(f"✅ Successfully downloaded Excel file ({len(response.content)} bytes)")
                            return io.BytesIO(response.content)
                        else:
                            logging.warning(f"Content doesn't appear to be Excel file. Content starts with: {response.content[:50]}")
                    else:
                        logging.warning(f"HTTP {response.status_code} for URL: {url}")
                        
                except requests.exceptions.RequestException as e:
                    logging.warning(f"Request failed for {url}: {e}")
                    continue
                
                # Add small delay between attempts
                time.sleep(1)
            
            # Strategy 2: Try to extract direct download link from SharePoint page
            logging.info("Attempting to extract download link from SharePoint page...")
            return self.extract_file_from_sharepoint_page(session, headers)
            
        except Exception as e:
            logging.error(f"Failed to download Excel file: {e}")
            return None
    
    def extract_file_from_sharepoint_page(self, session, headers):
        """Try to extract the actual file download URL from SharePoint page"""
        try:
            response = session.get(self.sharepoint_shared_url, headers=headers, timeout=30)
            
            if response.status_code == 200:
                content = response.text
                logging.info("Searching for download URLs in page content...")
                
                # Multiple patterns to search for download URLs
                download_patterns = [
                    r'"downloadUrl":"([^"]+)"',
                    r'"@microsoft\.graph\.downloadUrl":"([^"]+)"',
                    r'data-downloadurl="([^"]+)"',
                    r'"downloadUrl"\s*:\s*"([^"]+)"',
                    r'downloadUrl["\']?\s*:\s*["\']([^"\']+)["\']',
                    r'href="([^"]*download[^"]*)"',
                ]
                
                for pattern in download_patterns:
                    matches = re.findall(pattern, content, re.IGNORECASE)
                    if matches:
                        for match in matches:
                            download_url = match.replace('\\u0026', '&').replace('\\/', '/')
                            if 'sharepoint.com' in download_url:
                                logging.info(f"Found potential download URL: {download_url}")
                                
                                try:
                                    file_response = session.get(download_url, headers=headers, timeout=30)
                                    if file_response.status_code == 200 and len(file_response.content) > 1000:
                                        content_type = file_response.headers.get('content-type', '').lower()
                                        if not ('text/html' in content_type or file_response.content.startswith(b'<!DOCTYPE')):
                                            logging.info("✅ Successfully downloaded Excel file using extracted URL")
                                            return io.BytesIO(file_response.content)
                                except:
                                    continue
            
            logging.error("Could not extract valid download URL from SharePoint page")
            return None
            
        except Exception as e:
            logging.error(f"Failed to extract file from SharePoint page: {e}")
            return None
    
    def parse_excel_data(self, excel_file):
        """Parse Excel file and return relevant data"""
        try:
            # Reset file pointer to beginning
            excel_file.seek(0)
            
            # Try different engines and methods
            engines_to_try = ['openpyxl', 'xlrd']
            df = None
            
            for engine in engines_to_try:
                try:
                    excel_file.seek(0)  # Reset file pointer
                    df = pd.read_excel(excel_file, engine=engine)
                    logging.info(f"Successfully loaded Excel file using {engine} engine")
                    break
                except Exception as e:
                    logging.warning(f"Failed with {engine} engine: {e}")
                    continue
            
            if df is None:
                logging.error("Failed to read Excel file with any engine")
                return None
            
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
                        col_clean = str(df_col).lower().replace(' ', '').replace('_', '').replace('-', '')
                        req_clean = req_col.lower().replace(' ', '').replace('_', '').replace('-', '')
                        if req_clean in col_clean or col_clean in req_clean:
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
            date_formats = ['%d-%b-%y', '%d-%b-%Y', '%d/%m/%y', '%d/%m/%Y', '%Y-%m-%d', '%m/%d/%Y', '%d.%m.%Y', '%Y/%m/%d']
            
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