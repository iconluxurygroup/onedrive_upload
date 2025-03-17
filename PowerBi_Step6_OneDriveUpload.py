
import argparse
import mimetypes
import os.path
import requests

from class_installation import Installation

# parser = argparse.ArgumentParser()
# parser.add_argument('file', metavar='FILE')
# parser.add_argument('--path', metavar='PATH', default='')
# args = parser.parse_args()

# file_path = args.file
# drive_path = args.path

FileFolder = '/source'

installation = Installation()
#access_token = installation.get_access_token()
# In PowerBi_Step6_OneDriveUpload.py
access_token = installation.get_access_token(v=True)

# List all files in the folder
files = os.listdir(folder_path)

# Filter out directories and only list files
file_list = [file for file in files if os.path.isfile(os.path.join(FileFolder, file))]


# Loop through the files and rename each one
for file in file_list:
    current_file_path = os.path.join(FileFolder, file)

    file_name = os.path.basename(current_file_path)
    file_size = os.path.getsize(current_file_path)
    file_mime_type = mimetypes.guess_type(current_file_path)[0]

    file_obj = open(current_file_path, 'rb')
    file_name = 'PowerBI/' + str(file_name)
    url = 'https://graph.microsoft.com/v1.0/me/drive/root:/{:s}:/content'.format(file_name)
    # url = 'https://graph.microsoft.com/v1.0/me/drive/root:/{:s}:/content'.format(drive_path + file_name)
    print(url)
    r = requests.put(url, headers={
        'Authorization': 'Bearer ' + access_token,
        'Content-Type': file_mime_type,
    }, data=file_obj.read())
    r = r.json()
    assert 'error' not in r, r['error']['message']
    file_obj.close()

    new_file_name = file.replace(".xlsx", "_fid_" + r['id'] + ".xlsx")  # Example: appending "_renamed" before file extension
    new_file_path = os.path.join(FileFolder, new_file_name)
    
    # Rename the file
    os.rename(current_file_path, new_file_path)

    print(r['id'])
