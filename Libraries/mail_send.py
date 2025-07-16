import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta
import os
from Variables import variables
import paramiko
import pandas as pd
from pathlib import Path
from robot.api  import logger

def prepare_email_body_output(start_date, report_name, end_date, text_message):
    try:
        # Convert to datetime if needed
        report_name = "Job Card Creation Process Completed"
        if isinstance(start_date, str):
            start_date = datetime.strptime(start_date, "%Y-%m-%d")  
        if isinstance(end_date, str):
            end_date = datetime.strptime(end_date, "%Y-%m-%d")  

        formatted_start_date = start_date.strftime("%d-%m-%Y")
        formatted_end_date = end_date.strftime("%d-%m-%Y")

        if not text_message or text_message is None:

            body = f"""
            Hi,<br><br>
            {report_name} for {formatted_start_date}.<br><br>
            Thanks,<br>
            Bot
            """
            subject = f"{report_name}"
            return subject, body, "Success"
        else:
            body = f"""
            Hi,<br><br>
            {report_name} for {formatted_start_date}. {text_message} <br><br>
            Thanks,<br>
            Bot
            """
            subject = f"{report_name}"
            return subject, body, "Success"

    except Exception as e:
        print(f"Error preparing email body: {e}")
        return None, None, "Failure"


def read_email_credentials():

    # config_file = Path(__file__).resolve().parent.parent / "Config" / "Popular_Credentials.xlsx"
    config_file = Path(r"C:\JobcardOpeningIntegrated") / "Config" / "Popular_Credentials.xlsx"
    # config_file = Path(r"C:\Users\popular\Desktop\JobcardOpeningIntegrated\Jobcard-Opening") / "Config" / "Popular_Credentials.xlsx"
    
    df = pd.read_excel(config_file)

    # Initialize lists
    RECIPIENT_EMAIL = []
    # CC_EMAIL = []
    for column in df.columns:
        if column.strip().lower().startswith('recipient email'):
            for cell in df[column].dropna():
                # Split each cell by semicolon, strip spaces, and add to list
                emails = [email.strip() for email in str(cell).split(';') if email.strip()]
                RECIPIENT_EMAIL.extend(emails)

    return RECIPIENT_EMAIL


    # Iterate through columns
    # for column in df.columns:
    #     if column.strip().lower().startswith('recipient email'):
    #         RECIPIENT_EMAIL.extend(df[column].dropna().tolist())
    #     elif column.strip().lower().startswith('cc email'):
    #         CC_EMAIL.extend(df[column].dropna().tolist())

    # return RECIPIENT_EMAIL, CC_EMAIL

# recipient_emails = read_email_credentials()
# print("Recipient Emails:", recipient_emails)
# print("CC Emails:", cc_emails)

def send_email_output(start_date, report_name, text_message, attachment_path=None):
    try:
        
        # Prepare email content
        email_body_data = prepare_email_body_output(start_date, report_name, start_date, text_message)
        if not email_body_data or email_body_data[0] is None:
            print("Failed to prepare email body. Email not sent.")
            return False

        subject, body, stat = email_body_data

        smtp_server = variables.smtp_server
        smtp_port = variables.smtp_port
        sender_email = variables.SENDER_EMAIL
        password = variables.SENDER_PASSWORD
        recipient_emails = read_email_credentials()
        logger.info(recipient_emails)
        # logger.info(cc_emails)

        # recipient_emails = variables.RECIPIENT_EMAIL  # <-- Expecting a list now
        # cc_emails = variables.CC_EMAIL  # <-- Also expecting a list

        # Combine To and CC for final send list
        to_addrs = recipient_emails

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = ', '.join(recipient_emails)  # Multiple recipients
        # msg['Cc'] = ', '.join(cc_emails)
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html'))

        if attachment_path and os.path.exists(attachment_path) and os.path.isfile(attachment_path):
            with open(attachment_path, "rb") as attachment_file:
                attach_part = MIMEApplication(attachment_file.read(), _subtype="xlsx")
                attach_part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
                msg.attach(attach_part)
        else:
            if attachment_path:
                print(f"Attachment not found: {attachment_path}")

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(sender_email, to_addrs, msg.as_string())
            print("Final Email sent successfully!")
            return True

    except smtplib.SMTPException as e:
        print(f"SMTP error occurred: {e}")
        return False
    except Exception as e:
        print(f"An error occurred while sending the email: {e}")
        return False


def read_sftp_credentials(file_path):
    try:
        
        df = pd.read_excel(file_path, engine="openpyxl")

        # Extract values from the first (and only) row
        remote_host = df.loc[0, 'SFTP Remote Host']
        remote_port = df.loc[0, 'SFTP Remote Port']
        username = df.loc[0, 'SFTP Username']  
        password = df.loc[0, 'SFTP Password']

        return remote_host, remote_port, username, password
    except Exception as e:
        print(f"An error occurred: {e}")
        return None, None
    
def consolidated_report_copy_to_central_machine(curr_date, consolidated_report_path, config_file, job_card_no_to_check):
    # Load the source Excel file
    report_df = pd.read_excel(consolidated_report_path)

    # Filter: Job Card No, Recall Status = Yes, DMS Execution Status = Success
    filtered_df = report_df[
        (report_df['Job Card No'].astype(str) == str(job_card_no_to_check)) &
        (report_df['Recall Status'].astype(str).str.lower() == 'yes') &
        (report_df['DMS Execution Status'].astype(str).str.lower() == 'success')
    ]

    if filtered_df.empty:
        print(f"No matching row found for Job Card No: {job_card_no_to_check} with Recall Status='Yes' and DMS Execution Status='Success'")
        return

    # Setup for SFTP
    remote_host, remote_port, username, password = read_sftp_credentials(config_file)

    remote_base_path = "/test"
    remote_folder_path = f"{remote_base_path}/{curr_date}"
    remote_downloads_path = f"{remote_base_path}/{curr_date}/Downloads"

    # Generate the timestamped file name
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    output_file_name = f"consolidated_jobcard_report{timestamp}.xlsx"
    local_file_path = os.path.join(os.getcwd(), output_file_name)  # Save in current directory
    remote_file_path = f"{remote_folder_path}/{output_file_name}"  # Remote path for uploading

    try:
        # Save the filtered row directly to the new Excel file
        filtered_df.to_excel(local_file_path, index=False)
        print(f"Created local file: {local_file_path}")

        # Connect to SFTP
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(remote_host, remote_port, username, password)
        sftp = ssh.open_sftp()

        # Ensure remote folders exist
        try:
            sftp.stat(remote_folder_path)
            print(f"Remote folder '{remote_folder_path}' already exists.")
        except FileNotFoundError:
            sftp.mkdir(remote_folder_path)
            print(f"Created remote folder: {remote_folder_path}")

        try:
            sftp.stat(remote_downloads_path)
            print(f"Remote folder '{remote_downloads_path}' already exists.")
        except FileNotFoundError:
            sftp.mkdir(remote_downloads_path)
            print(f"Created remote folder: {remote_downloads_path}")

        # Upload the new Excel file
        sftp.put(local_file_path, remote_file_path)
        print(f"File successfully uploaded to: {remote_file_path}")

        # Clean up
        sftp.close()
        ssh.close()

        # Remove the local file after uploading
        os.remove(local_file_path)
        print(f"Deleted local file: {local_file_path}")

    except Exception as e:
        print(f"Error during file transfer: {e}")
    
# def consolidated_report_copy_to_central_machine(curr_date, consolidated_report_path, config_file, job_card_no_to_check):
#     # Load the source Excel file
#     report_df = pd.read_excel(consolidated_report_path)

#     # Filter: Job Card No, Recall Status = Yes, DMS Execution Status = Success
#     filtered_df = report_df[
#         (report_df['Job Card No'].astype(str) == str(job_card_no_to_check)) &
#         (report_df['Recall Status'].astype(str).str.lower() == 'yes') &
#         (report_df['DMS Execution Status'].astype(str).str.lower() == 'success')
#     ]

#     if filtered_df.empty:
#         print(f"No matching row found for Job Card No: {job_card_no_to_check} with Recall Status='Yes' and DMS Execution Status='Success'")
#         return

#     # Setup for SFTP
#     remote_host, remote_port, username, password = read_sftp_credentials(config_file)

#     remote_base_path = "/test"
#     remote_folder_path = f"{remote_base_path}/{curr_date}"
#     output_file_name = "consolidated_jobcard_report.xlsx"
#     remote_file_path = f"{remote_folder_path}/{output_file_name}"
#     remote_downloads_path = f"{remote_base_path}/{curr_date}/Downloads"

#     try:
#         ssh = paramiko.SSHClient()
#         ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
#         ssh.connect(remote_host, remote_port, username, password)
#         sftp = ssh.open_sftp()

#         # Ensure the folder exists
#         try:
#             sftp.stat(remote_folder_path)
#         except FileNotFoundError:
#             sftp.mkdir(remote_folder_path)
#             print(f"Created remote folder: {remote_folder_path}")

#         try:
#             sftp.stat(remote_downloads_path)  # Check if folder exists
#             print(f"Folder '{remote_downloads_path}' already exists.")
#         except FileNotFoundError:
#             sftp.mkdir(remote_downloads_path)  # Create folder if missing
#             print(f"Folder '{remote_downloads_path}' created successfully.")

#         # Use a local temp file to prepare the Excel content
#         local_temp_file = "temp_consolidated_jobcard_report.xlsx"

#         try:
#             # Check if the destination file exists
#             sftp.stat(remote_file_path)
#             sftp.get(remote_file_path, local_temp_file)
#             existing_df = pd.read_excel(local_temp_file)
#             updated_df = pd.concat([existing_df, filtered_df], ignore_index=True)
#             print("Appending to existing consolidated report.")
#         except FileNotFoundError:
#             # File doesn't exist, so create it fresh
#             updated_df = filtered_df
#             print("Creating new consolidated report.")

#         # Save and upload the updated file
#         updated_df.to_excel(local_temp_file, index=False)
#         sftp.put(local_temp_file, remote_file_path)
#         print(f"File successfully updated at: {remote_file_path}")

#         # Clean up
#         os.remove(local_temp_file)
#         sftp.close()
#         ssh.close()

#     except Exception as e:
#         print(f"Error during file transfer: {e}")


# def consolidated_report_copy_to_central_machine(curr_date, consolidated_report, config_file):
    
#     consolidated_report_df = pd.read_excel(consolidated_report)

#     # Check if any row has 'yes' in the 'recall status' column
#     if (consolidated_report_df['Recall Status'].astype(str).str.lower() == 'yes').any():

#         remote_host, remote_port, username, password = read_sftp_credentials(config_file)

#         # Local file path
#         # consolidated_report = "E:\\JobcardOpeningIntegrated\\ProcessRelatedFolders\\2025-03-13\\ProcessedDMS\\consolidated_jobcard_report20250307221941.xlsx"
#         file_name = os.path.basename(consolidated_report)  # Extract the file name

#         # Correct remote paths (use forward slashes `/`)
#         remote_base_path = "/test"
#         remote_folder_path = f"{remote_base_path}/{curr_date}"  
#         remote_file_path = f"{remote_folder_path}/{file_name}"  # Correct path
#         remote_downloads_path = f"{remote_base_path}/{curr_date}/Downloads"

#         try:
#             # Establish SSH and SFTP connection
#             ssh = paramiko.SSHClient()
#             ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
#             ssh.connect(remote_host, remote_port, username, password)
#             sftp = ssh.open_sftp()

#             # Ensure the folder exists
#             try:
#                 sftp.stat(remote_folder_path)  # Check if folder exists
#                 print(f"Folder '{remote_folder_path}' already exists.")
#             except FileNotFoundError:
#                 sftp.mkdir(remote_folder_path)  # Create folder if missing
#                 print(f"Folder '{remote_folder_path}' created successfully.")

#             try:
#                 sftp.stat(remote_downloads_path)  # Check if folder exists
#                 print(f"Folder '{remote_downloads_path}' already exists.")
#             except FileNotFoundError:
#                 sftp.mkdir(remote_downloads_path)  # Create folder if missing
#                 print(f"Folder '{remote_downloads_path}' created successfully.")

#             # Upload the file
#             sftp.put(consolidated_report, remote_file_path)
#             print(f"Upload successful! File saved at: {remote_file_path}")

#             # Close SFTP & SSH connection
#             sftp.close()
#             ssh.close()
#         except Exception as e:
#             print(f"Error: {e}")
#     else:
#         print("No rows with 'yes' in 'Recall Status'.")

    

# curr_date = "2025-03-14"
# consolidated_report_copy_to_central_machine(curr_date)

