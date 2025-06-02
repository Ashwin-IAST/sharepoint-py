import os
import sys
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext

# Retrieve all necessary variables from environment
sharepoint_site_url = os.environ.get("SHAREPOINT_SITE_URL")
tenant_id = os.environ.get("SHAREPOINT_TENANT_ID") # Keep this, it might be implicitly used or good for debugging
client_id = os.environ.get("SHAREPOINT_CLIENT_ID")
client_secret = os.environ.get("SHAREPOINT_CLIENT_SECRET")
file_server_relative_url = os.environ.get("FILE_SERVER_RELATIVE_URL")
download_path = os.environ.get("DOWNLOAD_PATH")

# --- Input Validation ---
if not sharepoint_site_url:
    print("Error: SHAREPOINT_SITE_URL environment variable is not set.", file=sys.stderr)
    sys.exit(1)
if not tenant_id:
    print("Error: SHAREPOINT_TENANT_ID environment variable is not set.", file=sys.stderr)
    sys.exit(1)
if not client_id:
    print("Error: SHAREPOINT_CLIENT_ID environment variable is not set.", file=sys.stderr)
    sys.exit(1)
if not client_secret:
    print("Error: SHAREPOINT_CLIENT_SECRET environment variable is not set.", file=sys.stderr)
    sys.exit(1)
if not file_server_relative_url:
    print("Error: FILE_SERVER_RELATIVE_URL environment variable is not set.", file=sys.stderr)
    sys.exit(1)
if not download_path:
    print("Error: DOWNLOAD_PATH environment variable is not set.", file=sys.stderr)
    sys.exit(1)

print(f"Attempting to download file from: {file_server_relative_url}")
print(f"To local path: {download_path}")

try:
    # Authenticate with SharePoint using Client Credentials
    # FIX: Only pass client_id and client_secret to ClientCredential
    ctx_auth = ClientCredential(client_id, client_secret)
    # The ClientContext will use the full SharePoint site URL to determine the tenant
    ctx = ClientContext(sharepoint_site_url, ctx_auth)

    # Get the file object by its server-relative URL
    file_obj = ctx.web.get_file_by_server_relative_url(file_server_relative_url)

    # Load file properties and execute the query to ensure the file exists and is accessible
    ctx.load(file_obj)
    ctx.execute_query()

    # Extract file name from the URL
    file_name = os.path.basename(file_server_relative_url)
    local_file_path = os.path.join(download_path, file_name)

    # Ensure the download directory exists. 'exist_ok=True' prevents an error if it already does.
    os.makedirs(download_path, exist_ok=True)

    print(f"Downloading '{file_name}' to '{local_file_path}'...")
    # Download the file content
    with open(local_file_path, "wb") as local_file:
        file_obj.download(local_file).execute_query()

    print(f"Successfully downloaded '{file_name}' to '{local_file_path}'")

except Exception as e:
    print(f"Error downloading file from SharePoint: {e}", file=sys.stderr)
    print("Please ensure:", file=sys.stderr)
    print("1. The SharePoint file path and name are correct.", file=sys.stderr)
    print("2. The SharePoint App Registration has 'Sites.Read.All' or 'Sites.FullControl.All' permissions.", file=sys.stderr)
    print("3. The Client ID, Client Secret, and Tenant ID are correct and active (though Tenant ID might not be directly used in this ClientCredential call).", file=sys.stderr)
    sys.exit(1) # Exit with an error code to signal failure to Jenkins
