from win32com.client import Dispatch
import os
import re

def open_pst(pst_file_path):
    outlook = Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    namespace.Addstore(pst_file_path)
    print(f"PST File {pst_file_path} opened successfully.")
    root_folder = namespace.Folders[1]
    file_name = os.path.splitext(os.path.basename(pst_file_path))[0]
    root_dest_path = os.path.join(os.getcwd(), file_name)
    if not os.path.exists(root_dest_path):
        os.mkdir(root_dest_path)
    total_attachment = process_folders(root_folder, root_dest_path)
    return total_attachment
    

def process_folders(folder, root_dest_path):
    total_attachment = 0
    pattern = rf'\buk\b|\busa\b'
    print(f"\nFolder: {folder.Name}")
    for item in folder.Items:
        if item.Class == 43:
            subject = item.Subject.lower()
            if re.search(pattern, subject):
                print(f"Subject: {subject}")
                sub_dir = re.sub(r'[\\\/\:\*<>|?]',' ',subject)
                dest_path = os.path.join(root_dest_path, re.sub(r'\s+',' ',sub_dir))
                if not os.path.exists(dest_path):
                    os.mkdir(dest_path)
                attachment_count = process_email(item, dest_path)
                if attachment_count==0:
                    os.removedirs(dest_path)
                total_attachment += attachment_count

    for subfolder in folder.Folders:
        total_attachment += process_folders(subfolder, root_dest_path)
    return total_attachment

def process_email(email,dest_path):
    attachment_count = 0
    for attachment in email.Attachments:
        attachment_name = os.path.join(dest_path.strip(), re.sub(r'\s+','',attachment.FileName))
        attachment.SaveAsFile(attachment_name)
        attachment_count += 1
    print(f"\tAttachment Downloaded : {attachment_count}")
    return attachment_count

if __name__ == "__main__":
    attachment_count = 0
    pst_folder = input("Enter Folder Path: ").strip()
    for root, dirs, files in os.walk(pst_folder):
        for file in files:
            pst_file_path = os.path.join(root,file)
            attachment_count = open_pst(pst_file_path)
            attachment_count += attachment_count
    print(f"Attachments: {attachment_count}")