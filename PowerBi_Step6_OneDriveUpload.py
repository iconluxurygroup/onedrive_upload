import _settings as cfg

import mimetypes
import os.path
import requests
from typing import List
import re

from class_installation import Installation

def get_site_id(access_token: str, hostname: str, site_path: str) -> str:
    """Get the SharePoint site ID."""
    try:
        url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}"
        headers = {'Authorization': f'Bearer {access_token}'}
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        site_data = response.json()
        print(f"Site ID retrieved: {site_data['id']}")
        return site_data['id']
    except requests.exceptions.HTTPError as e:
        error_msg = e.response.text if e.response else str(e)
        print(f"Failed to get site ID: {error_msg}")
        return None

def ensure_folder_exists(access_token: str, site_id: str, folder_path: str, is_sharepoint: bool = True) -> bool:
    """Ensure the specified folder exists in the SharePoint site or OneDrive."""
    try:
        folder_path = folder_path.strip('/')
        base_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drive' if is_sharepoint else 'https://graph.microsoft.com/v1.0/me/drive'
        check_url = f'{base_url}/root:/{folder_path}'
        headers = {'Authorization': f'Bearer {access_token}'}
        
        response = requests.get(check_url, headers=headers)
        
        if response.status_code == 200:
            return True
        elif response.status_code == 404:
            create_url = f'{base_url}/root/children'
            folder_data = {
                "name": folder_path.split('/')[-1],
                "folder": {},
                "@microsoft.graph.conflictBehavior": "rename"
            }
            if '/' in folder_path:
                parent_path = '/'.join(folder_path.split('/')[:-1])
                parent_response = requests.get(f'{base_url}/root:/{parent_path}', headers=headers)
                if parent_response.status_code == 200:
                    folder_data["parentReference"] = {"path": f"/drive/root:/{parent_path}"}
            
            create_response = requests.post(
                create_url,
                headers={**headers, 'Content-Type': 'application/json'},
                json=folder_data
            )
            create_response.raise_for_status()
            print(f"Created folder: {folder_path}")
            return True
        else:
            response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Failed to ensure folder exists: {str(e)}")
        return False


def upload_and_rename_files(source_folder: str, hostname: str = 'iconluxurygroup.sharepoint.com', 
                          site_path: str = '/sites/icon-powerbi', 
                          drive_path: str = 'powerbi/') -> None:  # Changed default drive_path to 'powerbi/'
    """Upload files to SharePoint document library (powerbi folder) or fall back to OneDrive."""
    try:
        installation = Installation()
        access_token = installation.get_access_token(v=True)
        if not access_token:
            raise ValueError("Failed to obtain access token")

        # Try SharePoint first
        site_id = get_site_id(access_token, hostname, site_path)
        is_sharepoint = bool(site_id)
        
        if not is_sharepoint:
            print("Falling back to OneDrive due to SharePoint access failure")
            site_id = None  # For OneDrive, we don't need a site ID
        else:
            if drive_path and not ensure_folder_exists(access_token, site_id, drive_path, is_sharepoint=True):
                raise ValueError(f"Could not create or access folder in SharePoint: {drive_path}")

        files: List[str] = [
            f for f in os.listdir(source_folder) 
            if os.path.isfile(os.path.join(source_folder, f))
        ]

        if not files:
            print("No files found in the source folder")
            return

        for file in files:
            current_file_path = os.path.join(source_folder, file)
            
            try:
                file_name = os.path.basename(current_file_path)
                if not file_name:
                    raise ValueError(f"Invalid filename for path: {current_file_path}")
                
                file_mime_type = mimetypes.guess_type(current_file_path)[0] or 'application/octet-stream'
                file_size = os.path.getsize(current_file_path)

                base_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drive' if is_sharepoint else 'https://graph.microsoft.com/v1.0/me/drive'
                url = f'{base_url}/root:/{drive_path}{file_name}:/content'
                print(f"Uploading {file_name} ({file_size} bytes) to: {url}")

                with open(current_file_path, 'rb') as file_obj:
                    response = requests.put(
                        url,
                        headers={
                            'Authorization': f'Bearer {access_token}',
                            'Content-Type': file_mime_type,
                        },
                        data=file_obj
                    )
                    response.raise_for_status()
                    result = response.json()

                if 'id' not in result:
                    raise KeyError("No file ID returned")
                
                new_file_name = file.replace(".xlsx", "_fid_" + result['id'] + ".xlsx")  # Example: appending "_renamed" before file extension
                new_file_path = os.path.join(FileFolder, new_file_name)
                
                # Rename the file
                os.rename(current_file_path, new_file_path)
                
                print(f"Successfully uploaded and renamed: {new_file_name} (ID: {result['id']})")

            except requests.exceptions.HTTPError as e:
                error_msg = e.response.text if e.response else str(e)
                print(f"Upload failed for {file}: {error_msg}")
            except KeyError as e:
                print(f"Invalid response format for {file}: {str(e)}")
            except OSError as e:
                print(f"File operation failed for {file}: {str(e)}")
            except Exception as e:
                print(f"Unexpected error processing {file}: {str(e)}")

    except Exception as e:
        print(f"Initialization failed: {str(e)}")


if __name__ == "__main__":

    FileFolder = cfg.FileFolder + "powerbi/ready/"

    upload_and_rename_files(FileFolder)