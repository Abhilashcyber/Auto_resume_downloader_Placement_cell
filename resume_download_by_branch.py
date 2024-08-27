import os
import requests
import pandas as pd
from urllib.parse import urlparse, parse_qs

# Function to get the file ID from the Google Drive link
def get_drive_file_id(drive_url):
    parsed_url = urlparse(drive_url)
    path_parts = parsed_url.path.split('/')
    
    # Check for known URL patterns to extract the file ID
    if 'drive.google.com' in parsed_url.netloc:
        if 'd' in path_parts:  # Standard drive link format
            return path_parts[path_parts.index('d') + 1]
        elif 'view' in path_parts:  # Another possible format
            return path_parts[path_parts.index('view') - 1]
        elif 'open' in parsed_url.path:  # Older open links
            return parse_qs(parsed_url.query).get('id', [None])[0]
    return None

# Function to check if a link is a folder
def is_folder_link(drive_url):
    parsed_url = urlparse(drive_url)
    return 'folders' in parsed_url.path

# Function to download a file from Google Drive using file ID
def download_file_from_google_drive(file_id, dest_path):
    url = f"https://drive.google.com/uc?id={file_id}&export=download"
    session = requests.Session()
    response = session.get(url, stream=True)

    # Check for specific Google Drive responses related to access and missing files
    if response.status_code == 401:
        raise Exception("Access Denied: You do not have permission to download this file.")
    elif response.status_code == 404:
        raise Exception("File Not Found: The requested file does not exist on Google Drive.")
    elif "Request access" in response.text or "You need permission" in response.text:
        raise Exception("Request Access: You do not have permission and need to request access to this file.")

    # Handle the case where Google Drive shows a warning page for large files
    for key, value in response.cookies.items():
        if key.startswith('download_warning'):
            params = {'id': file_id, 'confirm': value}
            response = session.get(url, params=params, stream=True)
            break

    if response.status_code == 200:
        with open(dest_path, "wb") as file:
            for chunk in response.iter_content(32768):
                file.write(chunk)
    else:
        raise Exception(f"Failed to download file with status code {response.status_code}")

# Load the Excel file
file_path = 'D:\python\GE VERNOVA-short (1).xlsx'
excel_data = pd.read_excel(file_path)

# Define the local directory where files will be stored
base_dir = 'D:\GE Vernova'

# Create the base directory if it doesn't exist
if not os.path.exists(base_dir):
    os.makedirs(base_dir)

# Prepare lists to store failed downloads and folder links
failed_downloads = []
folder_links = []

# Process each row in the Excel file
for index, row in excel_data.iterrows():
    branch = row['Branch']
    resume_link = row['Resume Link']
    full_name = row['Full Name']
    usn = row['University Roll Number']
    email = row['Email']
    contact_number = row['Contact Number']

    # Skip rows where the resume link is missing or invalid
    if not isinstance(resume_link, str):
        print(f"Invalid or missing URL for {full_name}. Logging for manual review.")
        failed_downloads.append({
            'University Roll Number': usn,
            'Full Name': full_name,
            'Email': email,
            'Contact Number': contact_number,
            'Branch': branch,
            'Resume Link': resume_link,
            'Error': 'Invalid or missing URL'
        })
        continue

    # Check if the link is a folder
    if is_folder_link(resume_link):
        print(f"Found a folder link for {full_name}. Logging for manual download.")
        folder_links.append({
            'University Roll Number': usn,
            'Full Name': full_name,
            'Email': email,
            'Contact Number': contact_number,
            'Branch': branch,
            'Folder Link': resume_link
        })
        failed_downloads.append({
            'University Roll Number': usn,
            'Full Name': full_name,
            'Email': email,
            'Contact Number': contact_number,
            'Branch': branch,
            'Folder Link': resume_link
        })
        continue  # Skip further processing for folder links

    # Create a folder for the branch if it doesn't exist
    branch_folder = os.path.join(base_dir, branch)
    if not os.path.exists(branch_folder):
        os.makedirs(branch_folder)

    # Extract the file ID from the Google Drive link
    file_id = get_drive_file_id(resume_link)
    
    if file_id:
        try:
            # Save the resume with a meaningful name
            resume_filename = os.path.join(branch_folder, f"{full_name}_Resume.pdf")
            download_file_from_google_drive(file_id, resume_filename)
            print(f"Downloaded resume for {full_name} and saved in {branch_folder}")
        
        except Exception as e:
            print(f"Failed to download resume for {full_name}: {e}")
            # Add the failed download information to the list
            failed_downloads.append({
                'University Roll Number': usn,
                'Full Name': full_name,
                'Email': email,
                'Contact Number': contact_number,
                'Branch': branch,
                'Resume Link': resume_link,
                'Error': str(e)
            })
    else:
        print(f"Invalid Google Drive link for {full_name}: {resume_link}")
        failed_downloads.append({
            'University Roll Number': usn,
            'Full Name': full_name,
            'Email': email,
            'Contact Number': contact_number,
            'Branch': branch,
            'Resume Link': resume_link,
            'Error': 'Invalid Google Drive link'
        })

# If there were any failed downloads, save them to a new Excel file
if failed_downloads:
    failed_df = pd.DataFrame(failed_downloads)
    failed_df.to_excel(os.path.join(base_dir, 'Failed_Resume_Downloads.xlsx'), index=False)
    print("Some resumes failed to download. Details have been saved to 'Failed_Resume_Downloads.xlsx'.")

# If there were any folder links, save them to a new Excel file
if folder_links:
    folder_df = pd.DataFrame(folder_links)
    folder_df.to_excel(os.path.join(base_dir, 'Folder_Links.xlsx'), index=False)
    print("Some links point to folders. These have been saved to 'Folder_Links.xlsx'.")
