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
        """
        self.sharepoint_shared_url = sharepoint_shared_url
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.email_username = email_username
        self.email_password = email_password
        self.recipient_email = recipient_email
        
    def convert_sharepoint_url_to_direct_download(self, shared_url):
        """Convert SharePoint shared URL to direct download URL"""
        try:
            logging.info(f"Original URL: {shared_url}")
            
            if "sharepoint.com" in shared_url and ("/:x:/" in shared_url or "/:b:/" in shared_url):
                if "/:x:/" in shared_url:
                    download_url = shared_url.replace("/:x:/", "/:b:/")
                else:
                    download_url = shared_url
                    
                if "download=1" not in download_url:
                    separator = "&" if "?" in download_url else "?"
                    download_url = f"{download_url}{separator}download=1"
                    
                logging.info(f"Download URL: {download_url}")
                return download_url
                
            return shared_url
                
        except Exception as e:
            logging.error(f"Error converting SharePoint URL: {e}")
            return shared_url
    
    def download_excel_file(self):
        """Download Excel file from SharePoint shared link"""
        try:
            urls_to_try = [
                self.sharepoint_shared_url,
                self.convert_sharepoint_url_to_direct_download(self.sharepoint_shared_url),
            ]
            
            base_url = self.sharepoint_shared_url.split('?')[0] if '?' in self.sharepoint_shared_url else self.sharepoint_shared_url
            urls_to_try.extend([
                f"{base_url}?download=1",
                f"{self.sharepoint_shared_url}&download=1" if "?" in self.sharepoint_shared_url else f"{self.sharepoint_shared_url}?download=1",
            ])
            
            urls_to_try = list(dict.fromkeys(urls_to_try))
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,application/octet-stream,*/*',
                'Accept-Language': 'en-US,en;q=0.9',
            }
            
            session = requests.Session()
            
            for i, url in enumerate(urls_to_try):
                try:
                    logging.info(f"Attempt {i+1}: Trying to download from: {url}")
                    
                    response = session.get(url, headers=headers, timeout=30, allow_redirects=True)
                    logging.info(f"Response status: {response.status_code}")
                    logging.info(f"Content-Type: {response.headers.get('content-type', 'N/A')}")
                    
                    if response.status_code == 200:
                        content_type = response.headers.get('content-type', '').lower()
                        
                        if 'text/html' in content_type or response.content.startswith(b'<!DOCTYPE'):
                            logging.warning("Received HTML instead of Excel file")
                            continue
                        
                        if (len(response.content) > 1000 and  
                            ('excel' in content_type or 
                             'spreadsheet' in content_type or 
                             'vnd.openxmlformats' in content_type or
                             'application/octet-stream' in content_type or
                             response.content.startswith(b'PK'))):
                            
                            logging.info(f"✅ Successfully downloaded Excel file ({len(response.content)} bytes)")
                            return io.BytesIO(response.content)
                            
                except requests.exceptions.RequestException as e:
                    logging.warning(f"Request failed for {url}: {e}")
                    continue
                
                time.sleep(1)
            
            logging.error("Failed to download Excel file")
            return None
            
        except Exception as e:
            logging.error(f"Failed to download Excel file: {e}")
            return None
    
    def find_header_row_and_columns(self, excel_file):
        """
        Intelligently find the header row and identify required columns
        """
        try:
            excel_file.seek(0)
            
            # Read first few rows to analyze structure
            sample_df = pd.read_excel(excel_file, header=None, nrows=10)
            logging.info(f"Sample data shape: {sample_df.shape}")
            
            # Look for rows that contain our target column names
            target_columns = ['mode of payment', 'date of transfer']
            header_candidates = []
            
            for row_idx in range(min(5, len(sample_df))):
                row_data = sample_df.iloc[row_idx].astype(str).str.lower()
                matches = 0
                
                for target in target_columns:
                    target_words = target.split()
                    for cell_value in row_data:
                        if pd.notna(cell_value) and cell_value != 'nan':
                            # Check if all words of target are in the cell
                            if all(word in cell_value for word in target_words):
                                matches += 1
                                break
                            # Or check if any target word is in the cell (partial match)
                            elif any(word in cell_value for word in target_words):
                                matches += 0.5
                
                if matches > 0:
                    header_candidates.append((row_idx, matches))
                    logging.info(f"Row {row_idx} has {matches} column matches: {row_data.tolist()}")
            
            # Sort by best match
            header_candidates.sort(key=lambda x: x[1], reverse=True)
            
            if header_candidates:
                best_header_row = header_candidates[0][0]
                logging.info(f"Selected header row: {best_header_row}")
                
                # Now read the file properly with the identified header row
                excel_file.seek(0)
                df = pd.read_excel(excel_file, header=best_header_row)
                
                # Clean column names
                df.columns = [str(col).strip().replace('\n', ' ').replace('\r', ' ') for col in df.columns]
                df.columns = [' '.join(col.split()) for col in df.columns]  # Remove extra spaces
                
                return df, best_header_row
            else:
                # Fallback: try each row as header until we find data
                for header_row in range(min(4, len(sample_df))):
                    try:
                        excel_file.seek(0)
                        df = pd.read_excel(excel_file, header=header_row)
                        if len(df) > 0 and not df.empty:
                            logging.info(f"Using header row {header_row} as fallback")
                            return df, header_row
                    except:
                        continue
                
                return None, None
                
        except Exception as e:
            logging.error(f"Error finding header row: {e}")
            return None, None
    
    def parse_excel_data(self, excel_file):
        """Parse Excel file with intelligent header detection"""
        try:
            # Find the correct header row and read data
            df, header_row = self.find_header_row_and_columns(excel_file)
            
            if df is None:
                logging.error("Could not identify proper header row")
                return None
            
            logging.info(f"Excel file loaded successfully. Shape: {df.shape}")
            logging.info(f"Columns found: {list(df.columns)}")
            
            # Find required columns with flexible matching
            required_columns = ['Mode of Payment', 'Date of Transfer']
            column_mapping = {}
            
            for req_col in required_columns:
                found_col = None
                req_words = req_col.lower().split()
                
                # Try different matching strategies
                for df_col in df.columns:
                    df_col_str = str(df_col).lower()
                    
                    # Exact match
                    if df_col_str == req_col.lower():
                        found_col = df_col
                        break
                    
                    # All words present
                    if all(word in df_col_str for word in req_words):
                        found_col = df_col
                        break
                    
                    # Key word matching
                    if req_col.lower() == 'mode of payment':
                        if any(keyword in df_col_str for keyword in ['mode', 'payment', 'pay', 'method']):
                            found_col = df_col
                            break
                    elif req_col.lower() == 'date of transfer':
                        if any(keyword in df_col_str for keyword in ['date', 'transfer', 'due']):
                            found_col = df_col
                            break
                
                if found_col:
                    column_mapping[req_col] = found_col
                    logging.info(f"✅ Mapped '{req_col}' to column '{found_col}'")
                else:
                    logging.error(f"❌ Required column '{req_col}' not found")
                    logging.error(f"Available columns: {list(df.columns)}")
                    
                    # Show sample of first few rows to help debug
                    logging.info("First few rows of data:")
                    logging.info(df.head().to_string())
                    return None
            
            # Rename columns
            df = df.rename(columns=column_mapping)
            
            # Remove rows where both key columns are empty/null
            df = df.dropna(subset=['Mode of Payment', 'Date of Transfer'], how='all')
            
            # Filter for cheque payments
            df['Mode of Payment'] = df['Mode of Payment'].astype(str)
            cheque_mask = df['Mode of Payment'].str.lower().str.contains('cheque|check', na=False)
            cheque_df = df[cheque_mask].copy()
            
            if cheque_df.empty:
                logging.info("No cheque payments found")
                return pd.DataFrame()
            
            logging.info(f"Found {len(cheque_df)} cheque payments")
            
            # Parse dates
            date_formats = [
                '%d-%b-%y', '%d-%b-%Y', '%d/%m/%y', '%d/%m/%Y', 
                '%Y-%m-%d', '%m/%d/%Y', '%d.%m.%Y', '%Y/%m/%d',
                '%d-%m-%y', '%d-%m-%Y'
            ]
            
            for date_format in date_formats:
                try:
                    cheque_df['Date of Transfer'] = pd.to_datetime(
                        cheque_df['Date of Transfer'], format=date_format, errors='coerce'
                    )
                    valid_dates = cheque_df['Date of Transfer'].notna().sum()
                    if valid_dates > 0:
                        logging.info(f"Parsed {valid_dates} dates using format {date_format}")
                        break
                except:
                    continue
            else:
                # Automatic parsing
                try:
                    cheque_df['Date of Transfer'] = pd.to_datetime(
                        cheque_df['Date of Transfer'], errors='coerce'
                    )
                    logging.info("Used automatic date parsing")
                except:
                    logging.error("Failed to parse dates")
            
            # Remove invalid dates
            initial_count = len(cheque_df)
            cheque_df = cheque_df.dropna(subset=['Date of Transfer'])
            final_count = len(cheque_df)
            
            logging.info(f"Final: {final_count} cheque payments with valid dates (from {initial_count})")
            
            if final_count > 0:
                logging.info("Sample processed data:")
                logging.info(cheque_df[['Mode of Payment', 'Date of Transfer']].head().to_string())
            
            return cheque_df
            
        except Exception as e:
            logging.error(f"Failed to parse Excel file: {e}")
            import traceback
            logging.error(traceback.format_exc())
            return None
    
    def find_reminders_needed(self, df):
        """Find cheques that need reminders (3 days before transfer date)"""
        if df is None or df.empty:
            return pd.DataFrame()
        
        today = datetime.now().date()
        target_date = today + timedelta(days=3)
        
        reminders_needed = df[df['Date of Transfer'].dt.date == target_date].copy()
        
        logging.info(f"Checking for reminders needed for {target_date}")
        logging.info(f"Found {len(reminders_needed)} reminders needed")
        
        if len(reminders_needed) > 0:
            logging.info("Reminders for:")
            for _, row in reminders_needed.iterrows():
                logging.info(f"  - {row.get('Mode of Payment', 'N/A')} on {row['Date of Transfer'].strftime('%Y-%m-%d')}")
        
        return reminders_needed
    
    def send_email_reminder(self, reminder_data):
        """Send email reminder"""
        if reminder_data.empty:
            logging.info("No reminders to send")
            return True
        
        try:
            msg = MIMEMultipart()
            msg['From'] = self.email_username
            msg['To'] = self.recipient_email
            msg['Subject'] = f"Cheque Transfer Reminder - {len(reminder_data)} payment(s) due in 3 days"
            
            body = self.create_email_body(reminder_data)
            msg.attach(MIMEText(body, 'html'))
            
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.email_username, self.email_password)
            server.send_message(msg)
            server.quit()
            
            logging.info(f"✅ Email reminder sent successfully to {self.recipient_email}")
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
        
        for col in reminder_data.columns:
            html_body += f"<th style='padding: 12px; text-align: left; font-weight: bold;'>{col}</th>"
        
        html_body += """
                    </tr>
                </thead>
                <tbody>
        """
        
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
        
        excel_file = self.download_excel_file()
        if not excel_file:
            logging.error("❌ Failed to download Excel file")
            return False
        
        df = self.parse_excel_data(excel_file)
        if df is None:
            logging.error("❌ Failed to parse Excel data")
            return False
        
        reminders = self.find_reminders_needed(df)
        
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
    """Main function"""
    logging.info("📋 SharePoint Cheque Reminder Starting...")
    
    config = {
        'sharepoint_shared_url': os.getenv('SHAREPOINT_SHARED_URL'),
        'smtp_server': os.getenv('SMTP_SERVER'),
        'smtp_port': int(os.getenv('SMTP_PORT', 587)),
        'email_username': os.getenv('EMAIL_USERNAME'),
        'email_password': os.getenv('EMAIL_PASSWORD'),
        'recipient_email': os.getenv('RECIPIENT_EMAIL')
    }
    
    missing_config = [key for key, value in config.items() if value is None or value == '']
    if missing_config:
        logging.error(f"❌ Missing required environment variables: {missing_config}")
        return False
    
    logging.info("✅ Configuration loaded successfully")
    
    reminder_system = SharePointSharedLinkReminder(**config)
    success = reminder_system.run_reminder_check()
    
    if success:
        logging.info("🎉 Reminder check completed successfully")
    else:
        logging.error("💥 Reminder check failed")
        
    return success

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)