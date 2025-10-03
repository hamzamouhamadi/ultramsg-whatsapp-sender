import http.client
import ssl
import pandas as pd
from urllib.parse import urlencode
from dotenv import load_dotenv
import os

# Load environment variables from .env file
load_dotenv()

# UltraMsg API credentials
TOKEN = os.getenv("ULTRAMSG_TOKEN")
INSTANCE_ID = os.getenv("ULTRAMSG_INSTANCE_ID")
API_URL = "api.ultramsg.com"

# Message to send
MESSAGE_BODY = "Test message from Python script."

# Read Excel file
def read_phone_numbers(file_path, column_name="PhoneNumber"):
    try:
        df = pd.read_excel(file_path)
        if column_name not in df.columns:
            raise ValueError(f"Column '{column_name}' not found in Excel file.")
        return df[column_name].dropna().tolist()
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

# Send WhatsApp message to a single phone number
def send_whatsapp_message(phone_number):
    try:
        conn = http.client.HTTPSConnection(API_URL, context=ssl._create_unverified_context())
        
        # Prepare payload
        payload = {
            "token": TOKEN,
            "to": phone_number,
            "body": MESSAGE_BODY
        }
        payload_str = urlencode(payload)
        
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        
        # Send request
        endpoint = f"/{INSTANCE_ID}/messages/chat"
        conn.request("POST", endpoint, payload_str, headers)
        
        # Get response
        res = conn.getresponse()
        data = res.read().decode("utf-8")
        
        if res.status == 200:
            print(f"Message sent to {phone_number}: {data}")
            return True
        else:
            print(f"Failed to send message to {phone_number}: {data}")
            return False
            
    except Exception as e:
        print(f"Error sending message to {phone_number}: {e}")
        return False
    finally:
        conn.close()

# Main function
def main():
    # Path to Excel file in the same directory as the script
    excel_file = os.path.join(os.path.dirname(__file__), "contacts.xlsx")
    
    # Read phone numbers
    phone_numbers = read_phone_numbers(excel_file)
    
    if not phone_numbers:
        print("No phone numbers found or error reading the file.")
        return
    
    # Send messages to all phone numbers
    success_count = 0
    for number in phone_numbers:
        if send_whatsapp_message(number):
            success_count += 1
    
    print(f"\nSummary: {success_count}/{len(phone_numbers)} messages sent successfully.")

if __name__ == "__main__":
    main()