import win32com.client
import os
import pandas as pd
import shutil

# Function to load the email log and attachments log
def load_logs(email_log_filepath):
    if os.path.exists(email_log_filepath):
        email_log = pd.read_excel(email_log_filepath, sheet_name='EmailLog')
        attachments_log = pd.read_excel(email_log_filepath, sheet_name='AttachmentsLog')
        last_email_serial = email_log['EmailSerial'].max()
        last_attachment_serial = attachments_log['AttachmentSerial'].max()
        email_serial_counter = last_email_serial + 1
        attachment_serial_counter = last_attachment_serial + 1
    else:
        email_log = pd.DataFrame(columns=['EmailSerial', 'Sender', 'Receiver', 'Time', 'Subject', 'MessageID'])
        attachments_log = pd.DataFrame(columns=['EmailSerial', 'AttachmentSerial', 'FileType', 'RenamedAttachment'])
        email_serial_counter = 1
        attachment_serial_counter = 1
    return email_log, attachments_log, email_serial_counter, attachment_serial_counter

# Function to access Outlook and retrieve inbox messages
def retrieve_inbox_messages():
    outlook = win32com.client.Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")
    inbox = mapi.GetDefaultFolder(6)  # Inbox folder
    messages = inbox.Items
    return messages

# Function to download attachments and update the log
def download_attachments(messages, email_log, attachments_log, email_serial_counter, attachment_serial_counter, attachments_folder):
    for message in messages:
        if message.Unread and message.Attachments.Count > 0:
            sender = message.Sender
            receiver = message.ReceivedByName
            time = message.ReceivedTime
            subject = message.Subject
            if (email_log['MessageID'] == message.EntryID).any():
                continue
            email_log = email_log.append({
                'EmailSerial': email_serial_counter,
                'Sender': sender,
                'Receiver': receiver,
                'Time': str(time),
                'Subject': subject,
                'MessageID': message.EntryID
            }, ignore_index=True)
            attachments_info = []
            for attachment in message.Attachments:
                filename = attachment.FileName
                file_extension = os.path.splitext(filename)[1]
                new_filename = f"{email_serial_counter}_{attachment_serial_counter}_{filename}"
                filepath = os.path.abspath(os.path.join(attachments_folder, new_filename))
                if (attachments_log['RenamedAttachment'] == new_filename).any():
                    continue
                attachment.SaveAsFile(filepath)
                attachments_info.append({
                    'EmailSerial': email_serial_counter,
                    'AttachmentSerial': attachment_serial_counter,
                    'FileType': file_extension,
                    'RenamedAttachment': new_filename
                })
                attachment_serial_counter += 1
            for attachment_info in attachments_info:
                attachments_log = attachments_log.append(attachment_info, ignore_index=True)
            email_serial_counter += 1
    return email_log, attachments_log, email_serial_counter, attachment_serial_counter

# Function to sort attachments by extension
def sort_attachments_by_extension(attachments_folder):
    files = os.listdir(attachments_folder)
    extension_folders = {}
    for file in files:
        file_path = os.path.join(attachments_folder, file)
        if not os.path.isfile(file_path) or os.path.dirname(file_path) != attachments_folder:
            continue
        file_extension = os.path.splitext(file)[1]
        if file_extension not in extension_folders:
            extension_folder = os.path.join(attachments_folder, f"{file_extension}_folder")
            os.makedirs(extension_folder, exist_ok=True)
            extension_folders[file_extension] = extension_folder
        destination_folder = extension_folders[file_extension]
        destination_path = os.path.join(destination_folder, file)
        shutil.move(file_path, destination_path)

# Main function to execute the email attachment download and sorting process
def execute_email_attachment_download_sorting(email_log_filepath, attachments_folder):
    email_log, attachments_log, email_serial_counter, attachment_serial_counter = load_logs(email_log_filepath)
    messages = retrieve_inbox_messages()
    email_log, attachments_log, email_serial_counter, attachment_serial_counter = download_attachments(
        messages, email_log, attachments_log, email_serial_counter, attachment_serial_counter, attachments_folder
    )
    with pd.ExcelWriter(email_log_filepath) as writer:
        email_log.to_excel(writer, sheet_name='EmailLog', index=False)
        attachments_log.to_excel(writer, sheet_name='AttachmentsLog', index=False)
    sort_attachments_by_extension(attachments_folder)

# Execute the main function
execute_email_attachment_download_sorting(
    email_log_filepath=r"#path to log file",
    attachments_folder=r"#path to attachments folder"
)
