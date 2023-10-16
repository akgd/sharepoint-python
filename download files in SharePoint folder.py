from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

# SharePoint site URL
site_url = "https://YOURTENANT.sharepoint.com/sites/YOURSITE"
folder_server_relative_url = "/sites/YOURSITE/Some Library/Some Sub Folder/Another Folder"

# Output directory where files will be saved
output_dir = "downloaded_files"

# Create the output directory if it doesn't exist
import os
os.makedirs(output_dir, exist_ok=True)

# App-only authentication
app_principal = {
    'client_id': 'YOUR CLIENT ID',
    'client_secret': 'YOUR CLIENT SECRET',
}

# Initialize the SharePoint client context with app-only authentication
credentials = ClientCredential(app_principal['client_id'], app_principal['client_secret'])
ctx = ClientContext(site_url).with_credentials(credentials)

# Get the folder by server-relative URL
folder = ctx.web.get_folder_by_server_relative_url(folder_server_relative_url)
ctx.load(folder)
files = folder.files
ctx.load(files)
ctx.execute_query()

# Download files from the folder
print(len(files))
for file in files:
    file_name = os.path.join(output_dir, file.properties["Name"])
    with open(file_name, 'wb') as local_file:
        ctx.load(file)
        file.download(local_file)
        ctx.execute_query()

print("Download completed.")
