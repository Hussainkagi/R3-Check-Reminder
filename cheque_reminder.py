import os
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import io
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class SharePointChequeReminder:
    def __init__(self, sharepoint_url, sharepoint_file_path, username, password, 
                 smtp_server, smtp_port, email_username, email_password, recipient_email):
        """
        Initialize the SharePoint Cheque Reminder system
        
        Args:
            sharepoint_url: Your SharePoint site URL (e.g., 'https://yourcompany.sharepoint.com/sites/yoursite')
            sharepoint_file_path: Path to Excel file in SharePoint (e.g., '/Shared Documents/filename.xlsx')
            username: SharePoint username
            password: SharePoint password
            smtp_server: Email server (e.g., 'smtp.gmail.com' for Gmail, 'smtp-mail.outlook.com' for Outlook)
            smtp_port: Email port (587 for most)
            email_username: Email account username
            email_password: Email account password (or app password)
            recipient_email: Email address to send reminders to
        """
        self.sharepoint_url = sharepoint_url
        self.sharepoint_file_path = sharepoint_file_path
        self.username = username
        self.password = password
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.email_username = email_username
        self.email_password = email_password
        self.recipient_email = recipient_email
        
    def connect_to_sharepoint(self):
        """Establish connection to SharePoint"""
        try:
            auth_context = AuthenticationContext(self.sharepoint_url)
            auth_context.acquire_token_for_user(self.username, self.password)
            ctx = ClientContext(self.sharepoint_url, auth_context)
            logging.info("Successfully connected to SharePoint")
            return ctx
        except Exception as e:
            logging.error(f"Failed to connect to SharePoint: {e}")
            return None
    
    def download_excel_file(self, ctx):
        """Download Excel file from SharePoint"""
        try:
            response = File.open_binary(ctx, self.sharepoint_file_path)
            bytes_file_obj = io.BytesIO()
            bytes_file_obj.write(response.content)
            bytes_file_obj.seek(0)
            logging.info("Successfully downloaded Excel file from SharePoint")
            return bytes_file_obj
        except Exception as e:
            logging.error(f"Failed to download Excel file: {e}")
            return None
    
    def parse_excel_data(self, excel_file):
        """Parse Excel file and return relevant data"""
        try:
            df = pd.read_excel(excel_file)
            
            # Check if required columns exist
            required_columns = ['Mode of Payment', 'Date of Transfer']
            if not all(col in df.columns for col in required_columns):
                logging.error(f"Required columns not found. Available columns: {list(df.columns)}")
                return None
            
            # Filter for cheque payments only
            cheque_df = df[df['Mode of Payment'].str.lower() == 'cheque'].copy()
            
            if cheque_df.empty:
                logging.info("No cheque payments found")
                return pd.DataFrame()
            
            # Convert Date of Transfer to datetime
            cheque_df['Date of Transfer'] = pd.to_datetime(cheque_df['Date of Transfer'], 
                                                         format='%d-%b-%y', errors='coerce')
            
            # Remove rows with invalid dates
            cheque_df = cheque_df.dropna(subset=['Date of Transfer'])
            
            logging.info(f"Found {len(cheque_df)} cheque payments with valid dates")
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
            <h2>Cheque Transfer Reminder</h2>
            <p>The following cheque payments are due for transfer in 3 days:</p>
            <table border="1" style="border-collapse: collapse; width: 100%;">
                <thead>
                    <tr style="background-color: #f2f2f2;">
        """
        
        # Add headers
        for col in reminder_data.columns:
            html_body += f"<th style='padding: 8px; text-align: left;'>{col}</th>"
        
        html_body += """
                    </tr>
                </thead>
                <tbody>
        """
        
        # Add data rows
        for _, row in reminder_data.iterrows():
            html_body += "<tr>"
            for col in reminder_data.columns:
                value = row[col]
                if pd.isna(value):
                    value = ""
                elif col == 'Date of Transfer':
                    value = value.strftime('%d-%b-%y') if hasattr(value, 'strftime') else str(value)
                html_body += f"<td style='padding: 8px;'>{value}</td>"
            html_body += "</tr>"
        
        html_body += """
                </tbody>
            </table>
            <br>
            <p>Please ensure these cheques are processed on time.</p>
            <p><em>This is an automated reminder.</em></p>
        </body>
        </html>
        """
        
        return html_body
    
    def run_reminder_check(self):
        """Main method to run the reminder check"""
        logging.info("Starting cheque reminder check...")
        
        # Connect to SharePoint
        ctx = self.connect_to_sharepoint()
        if not ctx:
            return False
        
        # Download Excel file
        excel_file = self.download_excel_file(ctx)
        if not excel_file:
            return False
        
        # Parse Excel data
        df = self.parse_excel_data(excel_file)
        if df is None:
            return False
        
        # Find reminders needed
        reminders = self.find_reminders_needed(df)
        
        # Send email if reminders needed
        if not reminders.empty:
            return self.send_email_reminder(reminders)
        else:
            logging.info("No cheque transfer reminders needed today")
            return True

# Configuration and usage
def main():
    # Configuration - UPDATE THESE VALUES
    config = {
        'sharepoint_url': 'https://yourcompany.sharepoint.com/sites/yoursite',  # Your SharePoint site URL
        'sharepoint_file_path': '/Shared Documents/your_file.xlsx',  # Path to your Excel file
        'username': 'your-email@company.com',  # SharePoint username
        'password': 'your-sharepoint-password',  # SharePoint password
        
        # Email configuration
        'smtp_server': 'smtp.gmail.com',  # Gmail: smtp.gmail.com, Outlook: smtp-mail.outlook.com
        'smtp_port': 587,
        'email_username': 'your-email@gmail.com',  # Email to send from
        'email_password': 'your-email-password',  # Use app password for Gmail/Outlook
        'recipient_email': 'recipient@company.com'  # Email to send reminders to
    }
    
    # Create and run reminder system
    reminder_system = SharePointChequeReminder(**config)
    success = reminder_system.run_reminder_check()
    
    if success:
        logging.info("Reminder check completed successfully")
    else:
        logging.error("Reminder check failed")

if __name__ == "__main__":
    main()