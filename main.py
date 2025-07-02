import email
import imaplib
import os
from email.message import EmailMessage
import magic
from dotenv import load_dotenv
import smtplib
import pandas as pd


#Load the environment variables
load_dotenv('.env')
my_email = os.getenv('EMAIL')            #Load the email it will read mails from
my_password = os.getenv('APP_PASSWORD')  #Load the app password
email_server = "imap.gmail.com"          #the server you want to connect to

download_folder = r"your download folder path"

#Define a function to connect to the server
def connect():
    try:
        imap = imaplib.IMAP4_SSL(email_server)
        imap.login(my_email,my_password)
        imap.select("Inbox")
        return imap
    except Exception as e:
        print("Error connecting to the server.",e)
        return None

#Function to search the required email ids
def search_emails(imap):
    _,mail_ids = imap.search(None,'FROM "email id you want to read from" SINCE "01-Jul-2025"')
    return [mail_id.decode() for mail_id in mail_ids[0].split()]

#Function to save email attachments in their original format
def save_attachments(part, file_name):
    try:
        original_file_path = os.path.join(download_folder, file_name)
        if not os.path.exists(original_file_path):
            with open(original_file_path, 'wb') as f:
                f.write(part.get_payload(decode=True))
            return original_file_path
        else:
            print(f"File already exists: {file_name}")
            return original_file_path  # or return None if you want to skip
    except Exception as e:
        print(f"Error saving attachment: {file_name}", e)
        return None


#Function to convert these attachments
def conversion(file_path):
    file_name = os.path.basename(file_path)
    mime_type = magic.from_file(file_path,mime=True)
    new_path = os.path.splitext(file_path)[0] + '.xlsx'

    try:
        if mime_type == "text/csv":
            df = pd.read_csv(file_path)
            df.to_excel(new_path,index=False)
            os.remove(file_path)
            return new_path
        elif mime_type == "text/html":
            tables = pd.read_html(file_path)
            if tables:
                tables[0].to_excel(new_path, index=False)
                os.remove(file_path)
                return new_path
        elif mime_type == "application/vnd.ms-excel":
            df = pd.read_excel(file_path, engine='xlrd')  # Explicit engine for .xls
            df.to_excel(new_path, index=False)
            os.remove(file_path)
            return new_path
    except Exception as e:
        print(f"Conversion Failed : {file_name}",e)

#Function which will send emails
def send_email(to_address,subject,body,attachment):
    try:
        msg = EmailMessage()
        msg['Subject'] = ''.join(subject.splitlines())
        msg['From'] = my_email
        msg['To'] = to_address
        msg.set_content(body)
        maintype , subtype = magic.from_file(attachment,mime=True).split('/')
        filename = os.path.basename(attachment)

        with open(attachment,'rb') as f:
            file_data = f.read()

        msg.add_attachment(file_data, maintype= maintype, subtype=subtype, filename = filename)

        with smtplib.SMTP(email_server,587) as connection:
            connection.starttls()
            connection.login(my_email,my_password)
            connection.send_message(msg)

    except Exception as e:
        print("Email could not be sent.",e)

def process_email(message):         #Function which reads the mails, save attachments and converts them accordingly
    try:
        print("-----------------------------------")
        print(f"From : {message.get('From')}")
        print(f"To : {message.get('To')}")
        print(f"Bcc : {message.get('Bcc')}")
        print(f"Date : {message.get('Date')}")
        print(f"Subject: {message.get('Subject')}")

        body_data =  None
        for part in message.walk():
            #Read the body of the email
            if body_data is None and (part.get_content_type() in ['text/plain','text/html']) and part.get_content_disposition() is None:
                body_data = part.get_payload(decode=True).decode(errors='ignore')
                print("Body: \n", body_data)

            if part.get_filename():
                original_file_path = save_attachments(part, part.get_filename())
                if original_file_path:
                    converted_file_path = conversion(file_path=original_file_path)
                    if converted_file_path:
                        send_email(
                            to_address=message.get("From"),
                            subject=message['Subject'],
                            body= body_data or " ",
                            attachment = converted_file_path
                        )
    except Exception as e:
        print("Error processing message:", e)

def main():
    imap = connect()
    if not imap:
        print("Could not connect to server.")
        return

    try:
        email_ids = search_emails(imap)
        print(f"Found {len(email_ids)} emails.")
        for mail_id in email_ids:
            try:
                status, msg_data = imap.fetch(mail_id, '(RFC822)')
                if status == 'OK':
                    raw_email = msg_data[0][1]
                    message = email.message_from_bytes(raw_email)
                    process_email(message)
            except Exception as inner_e:
                print(f"Failed to process email ID {mail_id}: {inner_e}")
    finally:
        try:
            imap.logout()
        except Exception as logout_err:
            print("Error during logout:", logout_err)

if __name__ == "__main__":
    main()
