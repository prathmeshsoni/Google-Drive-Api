import os
import pandas as pd
from googleapiclient.http import MediaFileUpload
from Google import Create_Service
import docx
import random
import string


def upload(file_name, folder_name):
    # Login Process Start
    CLIENT_SECRET_FILE = 'client-secret.json' # Download From console.cloud.google.com
    API_NAME = 'drive'
    API_VERSION = 'v3'
    SCOPES = ['https://www.googleapis.com/auth/drive']
    service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
    # Login Process End

    # Create a Folder
    folder_type = 'application/vnd.google-apps.folder'
    folder_metadata = {
        'name': folder_name,
        'mimeType': folder_type,
    }
    folder = service.files().create(
        body=folder_metadata,
        fields='id'
    ).execute()
    folder_id = folder.get('id')

    # Folder Permission
    request_body = {
        'role': 'reader',
        'type': 'anyone',
    }
    permission_folder = service.permissions().create(
        fileId=folder_id,
        body=request_body
    ).execute()

    # Print Sharing URL FOLDER
    response_share_link_folder = service.files().get(
        fileId=folder_id,
        fields='webViewLink'
    ).execute()
    folder_url = response_share_link_folder['webViewLink']
    print('Google Folder Url :: ', folder_url, ', Folder Name :: ', folder_name)

    # Upload a File
    file_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    file_metadata = {
        'name': f'{file_name}',
        'parents': [folder_id],
    }
    media_content = MediaFileUpload(file_name, mimetype=file_type)
    file = service.files().create(
        body=file_metadata,
        media_body=media_content,
        fields='id'
    ).execute()
    file_id = file["id"]

    # File Permission
    permission_file = service.permissions().create(
        fileId=file_id,
        body=request_body
    ).execute()

    # Print Sharing URL File
    response_share_link_file = service.files().get(
        fileId=file_id,
        fields='webViewLink'
    ).execute()
    file_url = response_share_link_file['webViewLink']
    print('Google Document Url :: ', file_url, ', File Name :: ', file_name)

    # Create CSV
    items = {
        'Folder Name': folder_name,
        'Folder Url': folder_url,
        'File Name': file_name,
        'File Url': file_url,
    }
    alldata = [items]
    df = pd.DataFrame(alldata)
    csv_name = 'File_info.csv'
    if os.path.exists(csv_name):
        df.to_csv(csv_name, mode='a', index=False, header=False)
    else:
        df.to_csv(csv_name, mode='a', index=False, header=True)


def create_docs(text, filee_name, folder_name):
    # Create a Document File
    doc = docx.Document()
    doc.add_paragraph(f"{text}")
    doc.save(filee_name)

    upload(filee_name, folder_name)


if __name__ == '__main__':
    texts = 'Hello Fam mm!!'
    filename = ''.join(random.choice(string.ascii_lowercase) for i in range(9))
    folders_name = 'Google testing Document'
    create_docs(texts, filename, folders_name)

# Create a Document File
# Login Process Start
# Create a Folder
# Folder Permission
# Print Sharing URL Folder
# Upload a File
# File Permission
# Print Sharing URL File
# Create CSV
