import os
import sys
from office365.sharepoint.client import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

# --- Get environment variables ---
# SharePoint Site URL (e.g., https://yourtenant.sharepoint.com/sites/Jenkins-testing)
sharepoint_site_url = os.environ.get("SHAREPOINT_SITE_URL")
# Client ID of your Azure AD App Registration
client_id = os.environ.get("SHAREPOINT_CLIENT_ID")
# Client Secret of your Azure AD App Registration
client_secret = os.environ.get("SHAREPOINT_CLIENT_SECRET")
# Relative path to the file on SharePoint (e.g., /sites/Jenkins-testing/Shared Documents/Tarfiles/one.zip)
file_server_relative_url = os.environ.get("FILE_SERVER_RELATIVE_URL")
# Local path to download the file to (e.g., .)
download_path = os.environ.get("DOWNLOAD_PATH")

if not all([sharepoint_site_url, client_id, client_secret, file_server_relative_url, download_path]):
    print("Error: One or more required environment variables are missing.")
    print(f"SHAREPOINT_SITE_URL: {sharepoint_site_url}")
    print(f"SHAREPOINT_CLIENT_ID: {client_id}")
    print(f"SHAREPOINT_CLIENT_SECRET: {'*' * len(client_secret) if client_secret else 'None'}") # Mask secret in output
    print(f"FILE_SERVER_RELATIVE_URL: {file_server_relative_url}")
    print(f"DOWNLOAD_PATH: {download_path}")
    sys.exit(1)

# --- Construct local file path ---
local_file_name = os.path.basename(file_server_relative_url)
local_download_full_path = os.path.join(download_path, local_file_name)

print(f"Attempting to download file from: {file_server_relative_url}")
print(f"To local path: {local_download_full_path}")

try:
    # 1. Create ClientCredential object
    app_credential = ClientCredential(client_id, client_secret)

    # 2. Create ClientContext with SharePoint site URL and ClientCredential
    ctx = ClientContext(sharepoint_site_url).with_credentials(app_credential)

    # 3. Get the file object by its server relative URL
    # Ensure the file exists before trying to download
    file_obj = ctx.web.get_file_by_server_relative_url(file_server_relative_url)
    
    # 4. Download the file content
    file_obj.download_content(local_download_full_path).execute_query()

    print(f"Successfully downloaded file: {local_file_name} to {local_download_full_path}")

except Exception as e:
    print(f"Error downloading file from SharePoint: {e}", file=sys.stderr)
    print("Please ensure:", file=sys.stderr)
    print("1. The SharePoint file path and name are correct.", file=sys.stderr)
    print("2. The SharePoint App Registration has 'Sites.Read.All' or 'Sites.FullControl.All' permissions (and admin consent).", file=sys.stderr)
    print("3. The Client ID, Client Secret are correct and active.", file=sys.stderr)
    sys.exit(1)
