import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import time
import random
import ssl
from tkinter import filedialog
from tkinter import Tk
import mimetypes
import re
# from email_validator import validate_email, EmailNotValidError

class EmailManager:
    def __init__(self, credentials):
        self.credentials = credentials
        self.log_file = filedialog.askopenfilename(title="Select Sent Log File", filetypes=[("CSV files", "*.csv")])
        self.error_file = filedialog.askopenfilename(title="Select Error Log File", filetypes=[("CSV files", "*.csv")])
        self.invalid_file = filedialog.askopenfilename(title="Select Invalid Email Log File", filetypes=[("CSV files", "*.csv")])
        self.nocert_file = filedialog.askopenfilename(title="Select No Certificate Log File", filetypes=[("CSV files", "*.csv")])

    def get_current_credential(self):
        return self.credentials

    ## if using attachment directory, uncomment this line
    # def send_email(self, receiver_email, receiver_id, subject, message, attachment_dir=None):
    def send_email(self, receiver_email, subject, message, attachment_paths=None):
        if self.read_log(receiver_email, self.log_file):
            print(f"Email already sent to: {receiver_email}")
            return True
        elif self.read_log(receiver_email, self.error_file):
            print(f"Email previously faced an issue: {receiver_email}")
            self.remove_from_error_log(receiver_email)
        elif self.read_log(receiver_email, self.invalid_file):
            print(f"Invalid email: {receiver_email}")
            return False  # Skip sending for invalid email addresses
        
        if not self.is_valid_email(receiver_email):
            print(f"Invalid email address: {receiver_email}")
            self.update_log(self.invalid_file, receiver_email)
            return False
        cc_names = ['ashad001sp@gmail.com', 'ahadaziz4@gmail.com']
        cc_names = ",".join(cc_names)
        try:
            current_credential = self.get_current_credential()
            smtp_server = "smtp.gmail.com"
            smtp_port = 465

            msg = MIMEMultipart()
            msg["Subject"] = subject
            msg["From"] = current_credential["email"]
            msg["To"] = receiver_email
            # msg['CC'] = cc_names
            msg.attach(MIMEText(message, "html"))

            if attachment_paths:
                for attachment_path in attachment_paths:
                    with open(attachment_path, "rb") as attachment:
                        mime_type, _ = mimetypes.guess_type(attachment_path)
                        main_type, sub_type = mime_type.split('/', 1)
                        file_content = MIMEBase(main_type, sub_type)
                        file_content.set_payload(attachment.read())
                        encoders.encode_base64(file_content)
                        file_content.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
                        msg.attach(file_content)

            # # if using attachment directory, uncomment the following lines
            # if attachment_dir:
            #     # Generate attachment path based on receiver_email or any other attribute from CSV
            #     attachment_path = os.path.join(attachment_dir, f"{receiver_id}_certificate.pdf")       # change the name of the file and extension as you like 
            #     if os.path.exists(attachment_path):
            #         with open(attachment_path, "rb") as attachment:
            #             mime_type, _ = mimetypes.guess_type(attachment_path)
            #             main_type, sub_type = mime_type.split('/', 1)
            #             file_content = MIMEBase(main_type, sub_type)
            #             file_content.set_payload(attachment.read())
            #             encoders.encode_base64(file_content)
            #             file_content.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
            #             msg.attach(file_content)
            #     else:
            #         print("Attachment file does not exist.")
            #         self.update_log(self.nocert_file,receiver_id)
            #         return False

            context = ssl.create_default_context()
            server = smtplib.SMTP_SSL(smtp_server, smtp_port, context=context)
            server.login(current_credential["email"], current_credential["password"])
            server.sendmail(current_credential["email"], receiver_email, msg.as_string())
            server.quit()

            self.update_log(self.log_file, receiver_email)
            print(f"Email sent successfully to: {receiver_email}")
            return True    
        except Exception as e:
            import traceback
            print(f"Error when sending mail to: {receiver_email}")
            print(f"Exception details: {traceback.format_exc()}")
            self.update_log(self.error_file, receiver_email)
            return False

    def update_log(self, log_file, data):
        with open(log_file, "a") as log:
            log.write(data + "\n")

    def is_valid_email(self, email):
        # Perform basic email format validation
        return re.match(r"[^@]+@[^@]+\.[^@]+", email)

    # def is_valid_email(self, email):
    #     try:
    #         validate_email(email)
    #         return True
    #     except EmailNotValidError as e:
    #         return False

    def read_log(self, email, log_file):
        try:
            with open(log_file, "r") as log:
                data = set(log.read().splitlines())
            return email in data
        except FileNotFoundError:
            return False

    def remove_from_error_log(self, email):
        try:
            with open(self.error_file, "r") as f:
                lines = f.readlines()
            with open(self.error_file, "w") as f:
                for line in lines:
                    if line.strip("\n") != email:
                        f.write(line)
        except FileNotFoundError:
            pass

if __name__ == "__main__":
    root = Tk()
    root.withdraw()

    excel_file_path = filedialog.askopenfilename(title="Select Main Excel File", filetypes=[("CSV files", "*.csv")])

    ## If you want to send the same attachment to every address in the file, uncomment the following line
    attachment_path = filedialog.askopenfilenames(title="Select Attachment Files", filetypes=[("All files", "*.*")])

    ## If you want to open a folder/directory for attachment files
    # attachment_dir = filedialog.askdirectory(title="Select Attachment Folder")

    email_manager = EmailManager(credentials={"email": "", "password": ""})

    df = pd.read_csv(excel_file_path)
    successful_emails = []

    for _, row in df.iterrows():
        member_name = row["Name"]
        receiver_email = row["Email"]
        # receiver_id = row["ID"]
        subject = "PROCOM: Zero Hour Invitation"

        # Check if email is empty
        if pd.isna(receiver_email) or not receiver_email.strip():
            print("No email detected in this row. Skipping...")
            continue

        html_message = f"""
            <html>
                <body>

                    # Your HTML code goes here
                    
                </body>
            </html>
        """
        ## If you want to send a separate attachment to every address in file, uncomment the following line
        # attachment_path = filedialog.askopenfilenames(title="Select Attachment File", filetypes=[("All files", "*.*")])  

        ## for using attachment directory instead attachement file, uncomment next line 
        # if email_manager.send_email(receiver_email, receiver_id, subject, html_message, attachment_dir):
        if email_manager.send_email(receiver_email, subject, html_message, attachment_path):
            successful_emails.append(receiver_email)
            time.sleep(random.randint(1, 3))
