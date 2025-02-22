import win32com.client
import pandas as pd
from pathlib import Path
import sys

# Define the base path for all attachments
BASE_PATH = Path('C:/Users/zacharywe/Documents/T&C')

# Map buyer names to their corresponding attachment files
BUYER_ATTACHMENTS = {
    'AH4R': 'AH4R - Terms and Conditions Addendum to Real Estate Purchase Agreement 08.09.2021.docx',
    'Atlas': 'Atlas - Terms and Conditions Addendum to Real Estate Purchase Agreement 6.3.21.docx',
    'Divvy': 'Divvy - Terms and Conditions Addendum to Real Estate Purchase Agreement 5.25.21.docx',
    'Tricon': 'Tricon - Terms and Conditions Addendum to Real Estate Purchase Agreement 10.30.2019.docx',
    'Hudson Homes': 'Hudson Homes - Terms and Conditions Addendum to Real Estate Purchase Agreement 2.16.21.docx',
    'Invitation': 'Invitation Homes - Terms and Conditions Addendum to Real Estate Purchase Agreement [12.8.2020].docx',
    'MCH': 'MCH - Terms and Conditions Addendum to Real Estate Purchase Agreement 06.16.2021.docx',
    'Open House': 'Open House - Terms and Conditions Addendum to Real Estate Purchase Agreement [05.25.21].docx',
    'Progress': 'Progress - Terms and Conditions Addendum to Real Estate Purchase Agreement [12.11.20].docx',
    'Cerberus': 'Progress - Terms and Conditions Addendum to Real Estate Purchase Agreement [12.11.20].docx',
    'Second Avenue': 'Second Avenue - Terms and Conditions Addendum to Real Estate Purchase Agreement 4.12.21.docx',
    'Sparrow': 'Sparrow - Terms and Conditions Addendum to Real Estate Purchase Agreement 5.25.21.docx',
    'Sylvan Road': 'Sylvan Road - Terms and Conditions Addendum to Real Estate Purchase Agreement 03.27.20 .docx',
    'Starwood': 'Starwood (Tiber) - Terms and Conditions Addendum to Real Estate Purchase Agreement 12.14.21 [Georgia, North Carolina, Florida, Tennessee].docx',
    'Starwood (Roofstock)': 'Starwood (Roofstock) - Terms and Conditions Addendum to Resale Agreement 12.14.21 [Texas, Nevada, Colorado].docx',
    'Westport Capital': 'Westport Capital - Terms and Conditions to Real Estate Purchase Agreement 12.12.21.docx'
}

def create_and_send_email(recipient, cc, subject, body, attachment_path):
    """Create and send an email with the specified details and attachment."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        email = outlook.CreateItem(0x0)  # 0x0 is the code for mail item
        
        email.Subject = subject
        email.HTMLBody = body  # Changed to HTMLBody since the template contains HTML
        email.To = recipient
        email.CC = cc
        
        if attachment_path.exists():
            email.Attachments.Add(str(attachment_path))
        else:
            print(f"Warning: Attachment not found - {attachment_path}")
        
        email.Display()  # Show the email before sending
        email.Send()
        return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return False

def main():
    try:
        # Check if excel file is provided as command line argument
        excel_file = 'Bulk_Template_PSAs.xlsx' if len(sys.argv) < 2 else sys.argv[1]
        
        # Read the Excel file - specifically the 'Basic' sheet which contains the email data
        df = pd.read_excel(excel_file, sheet_name='Basic')
        
        # Process each row in the dataframe
        for _, row in df.iterrows():
            # Skip if required fields are missing
            if pd.isnull(row['Email Addresses']) or pd.isnull(row['Subject']) or pd.isnull(row['Email HTML Body']):
                print(f"Skipping row for {row.get('Short Address', 'Unknown Address')} due to missing required information")
                continue
            
            buyer = str(row['Buyer'])
            if buyer in BUYER_ATTACHMENTS:
                attachment_path = BASE_PATH / BUYER_ATTACHMENTS[buyer]
                
                # Clean up email addresses (remove any extra spaces)
                recipient_emails = row['Email Addresses'].strip()
                cc_emails = str(row['CC']).strip() if not pd.isnull(row['CC']) else ''
                
                # Format subject if it doesn't already contain the address
                subject = row['Subject']
                if 'Draft Contract' not in subject and row.get('Short Address'):
                    subject = f"Draft Contract - {row['Short Address']}"
                
                # Send the email
                success = create_and_send_email(
                    recipient=recipient_emails,
                    cc=cc_emails,
                    subject=subject,
                    body=row['Email HTML Body'],
                    attachment_path=attachment_path
                )
                
                if success:
                    print(f"Email sent successfully for {row['Short Address']}")
            else:
                print(f"Warning: No attachment found for buyer '{buyer}' - {row.get('Short Address', 'Unknown Address')}")

    except Exception as e:
        print(f"Error processing Excel file: {e}")

if __name__ == "__main__":
    main()