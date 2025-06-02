# download_sharepoint_file.py
import os
import sys
from office366.runtime.auth.client_credential import ClientCredential
from office366.sharepoint.client_context import ClientContext

# Retrieve all necessary variables from environment
sharepoint_site_url = os.environ.get("SHAREPOINT_SITE_URL")
tenant_id = os.environ.get("SHAREPOINT_TENANT_ID")
client_id = os.environ.get("SHAREPOINT_CLIENT_ID")
client_secret = os.environ.get("SHAREPOINT_CLIENT_SECRET")
file_server_relative_url = os.environ.get("FILE_SERVER_RELATIVE_URL") # This now includes the dynamic file name
download_path = os.environ.get("DOWNLOAD_PATH") # This will now be '.'

# --- Input Validation (Crucial for robustness) ---
if not sharepoint_site_url:
    print("Error: SHAREPOINT_SITE_URL environment variable is not set.")
    sys.exit(1)
if not tenant_id:
    print("Error: SHAREPOINT_TENANT_ID environment variable is not set.")
    sys.exit(1)
if not client_id:
    print("Error: SHAREPOINT_CLIENT_ID environment variable is not set.")
    sys.exit(1)
if not client_secret:
    print("Error: SHAREPOINT_CLIENT_SECRET environment variable is not set.")
    sys.exit(1)
if not file_server_relative_url:
    print("Error: FILE_SERVER_RELATIVE_URL environment variable is not set.")
    sys.exit(1)
# No need to check if download_path is empty if it's always '.' or a valid path

print(f"Attempting to download file from: {file_server_relative_url}")
print(f"To local path: {download_path}")

try:
    # Authenticate
    ctx_auth = ClientCredential(tenant_id, client_id, client_secret)
    ctx = ClientContext(sharepoint_site_url, ctx_auth)

    # Get file object
    file_obj = ctx.web.get_file_by_server_relative_url(file_server_relative_url)

    # Resolve the file to ensure it exists and get its name (and handle potential errors)
    ctx.load(file_obj)
    ctx.execute_query() # This will raise an exception if the file doesn't exist or access is denied

    # Extract file name from the URL
    file_name = os.path.basename(file_server_relative_url)
    local_file_path = os.path.join(download_path, file_name)

    # Ensure the download directory exists locally (if download_path is not '.' this would still be useful)
    # Since download_path is '.', os.makedirs('.') is harmless and just verifies the current directory exists.
    os.makedirs(download_path, exist_ok=True)

    print(f"Downloading '{file_name}' to '{local_file_path}'...")
    # Download the file
    with open(local_file_path, "wb") as local_file:
        file_obj.download(local_file).execute_query()

    print(f"Successfully downloaded '{file_name}' to '{local_file_path}'")

except Exception as e:
    print(f"Error downloading file from SharePoint: {e}", file=sys.stderr)
    print("Please ensure the file name is correct and the SharePoint App Registration has permissions to the site/folder.", file=sys.stderr)
    sys.exit(1) # Exit with an error code to make Jenkins job fail
