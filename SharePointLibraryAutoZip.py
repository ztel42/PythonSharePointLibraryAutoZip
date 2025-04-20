import os
import shutil
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext

# Configuration
client_id = "your_client_id"
client_secret = "your_client_secret"
site_url = "https://yourorganization.sharepoint.com/sites/yoursite"
library_title = "YourLibraryTitle"
temp_dir = "temp_download"
zip_file = "output.zip"

# Authenticate
ctx = ClientContext(site_url).with_credentials(ClientCredential(client_id, client_secret))

# Get the library and root folder
library = ctx.web.lists.get_by_title(library_title)
root_folder = library.root_folder
ctx.load(root_folder)
ctx.execute_query()

def download_folder(ctx, folder, local_dir):
    # Create local directory
    os.makedirs(local_dir, exist_ok=True)
    # Get files in the folder
    files = folder.files.get().execute_query()
    for file in files:
        file_name = file.name
        local_file_path = os.path.join(local_dir, file_name)
        with open(local_file_path, 'wb') as local_file:
            file.download(local_file).execute_query()
            print(f"Downloaded {file_name} to {local_file_path}")
    # Get subfolders
    subfolders = folder.folders.get().execute_query()
    for subfolder in subfolders:
        subfolder_name = subfolder.name
        sub_local_dir = os.path.join(local_dir, subfolder_name)
        download_folder(ctx, subfolder, sub_local_dir)

# Remove existing temp_dir if exists
shutil.rmtree(temp_dir, ignore_errors=True)
# Download files
download_folder(ctx, root_folder, temp_dir)
# Create zip file
base_name = os.path.splitext(zip_file)[0]
shutil.make_archive(base_name, 'zip', temp_dir)
print(f"Zip file created at {zip_file}")
