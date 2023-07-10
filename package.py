import win32com.client
import os
import pandas as pd
import shutil

def load_email_log(filepath):
    # Load the existing email log from the file if it exists, otherwise create empty logs
    if os.path.exists(filepath):
        email_log = pd.read_excel(filepath, sheet_name='EmailLog')
        attachments_log = pd.read_excel(filepath, sheet_name='AttachmentsLog')
    else:
        email_log = pd.DataFrame(columns=['EmailSerial', 'Sender', 'Receiver', 'Time', 'Subject', 'MessageID'])
        attachments_log = pd.DataFrame(columns=['EmailSerial', 'AttachmentSerial', 'FileType', 'RenamedAttachment'])
    
    return email_log, attachments_log



def save_email_log(email_log, attachments_log, filepath):
    # Save the email log and attachments log to the same Excel file
    with pd.ExcelWriter(filepath) as writer:
        email_log.to_excel(writer, sheet_name='EmailLog', index=False)
        attachments_log.to_excel(writer, sheet_name='AttachmentsLog', index=False)



def access_outlook():
    # Access Outlook and retrieve the inbox folder
    return win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(6)



def get_attachments_folder():
    # Create a folder to save the attachments
    attachments_folder = r"D:\GIKI\CAA_intern\attachments"
    os.makedirs(attachments_folder, exist_ok=True)
    return attachments_folder



def download_attachments(messages, email_log, attachments_log, attachments_folder):
    # Get the last serial numbers from the existing logs
    email_serial_counter = email_log['EmailSerial'].max() + 1 if not email_log.empty else 1
    attachment_serial_counter = attachments_log['AttachmentSerial'].max() + 1 if not attachments_log.empty else 1

    for message in messages:
        if message.Unread and message.Attachments.Count > 0:
            # Get email information
            sender = message.Sender
            receiver = message.ReceivedByName
            time = message.ReceivedTime
            subject = message.Subject
            message_id = message.EntryID

            # Skip the email if it is already in the log
            if (email_log['MessageID'] == message_id).any():
                continue

            # Add email information to the email log
            email_log = email_log.append({
                'EmailSerial': email_serial_counter,
                'Sender': sender,
                'Receiver': receiver,
                'Time': str(time),
                'Subject': subject,
                'MessageID': message_id
            }, ignore_index=True)

            attachments_info = []

            for attachment in message.Attachments:
                # Get attachment information
                filename = attachment.FileName
                file_extension = os.path.splitext(filename)[1]
                new_filename = f"{email_serial_counter}_{attachment_serial_counter}_{filename}"
                filepath = os.path.abspath(os.path.join(attachments_folder, new_filename))

                # Skip the attachment if it is already in the log
                if (attachments_log['RenamedAttachment'] == new_filename).any():
                    continue

                # Save the attachment to the attachments folder
                attachment.SaveAsFile(filepath)

                # Add attachment information to the attachments log
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

    return email_log, attachments_log



def sort_attachments_by_extension(attachments_folder):
    # Get all files in the attachments folder
    files = os.listdir(attachments_folder)

    # Create a dictionary to store folders for each extension
    extension_folders = {}

    for file in files:
        file_path = os.path.join(attachments_folder, file)
        
        # Skip sub-folders and files within extension folders from sorting
        if not os.path.isfile(file_path) or os.path.dirname(file_path) != attachments_folder:
            continue

        file_extension = os.path.splitext(file)[1]

        # Create a folder for the extension if it doesn't exist
        if file_extension not in extension_folders:
            extension_folder = os.path.join(attachments_folder, f"{file_extension}_folder")
            os.makedirs(extension_folder, exist_ok=True)
            extension_folders[file_extension] = extension_folder

        # Move the file to the respective extension folder
        destination_folder = extension_folders[file_extension]
        destination_path = os.path.join(destination_folder, file)
        shutil.move(file_path, destination_path)


# MAIN FUNCTION :


if __name__ == "__main__":
    email_log_filepath = os.path.abspath(r"D:\GIKI\CAA_intern\email_log_with_attachments.xlsx")
    email_log, attachments_log = load_email_log(email_log_filepath)
    print("email log acessed")

    outlook = access_outlook()
    print("outlook access check")
    messages = outlook.Items
    attachments_folder = get_attachments_folder()

    email_log, attachments_log = download_attachments(messages, email_log, attachments_log, attachments_folder)
    save_email_log(email_log, attachments_log, email_log_filepath)

    sort_attachments_by_extension(attachments_folder)
