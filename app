import streamlit as st
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext

# Function to authenticate to SharePoint
def authenticate_to_sharepoint(site_url, username, password):
    context_auth = AuthenticationContext(site_url)
    if context_auth.acquire_token_for_user(username, password):
        return ClientContext(site_url, context_auth)
    else:
        st.error("Authentication failed.")
        return None

# Function to list files in a folder
def list_files_in_folder(ctx, folder_url):
    folder = ctx.web.get_folder_by_server_relative_url(folder_url)
    ctx.load(folder.files)
    ctx.execute_query()
    return [file.properties["Name"] for file in folder.files]

# Function to delete version history for selected files
def delete_file_versions(ctx, folder_url, filenames):
    folder = ctx.web.get_folder_by_server_relative_url(folder_url)
    for filename in filenames:
        file = folder.get_file_by_server_relative_url(f"{folder_url}/{filename}")
        file_versions = file.versions
        ctx.load(file_versions)
        ctx.execute_query()
        
        for version in file_versions:
            version.delete_object()
        ctx.execute_query()
        st.success(f"Deleted all versions for file: {filename}")

# Streamlit UI
st.title("SharePoint File Version History Deletion Tool")

# Input fields
site_url = st.text_input("Enter SharePoint site URL")
folder_url = st.text_input("Enter the folder URL relative to the site (e.g., /Shared Documents)")
username = st.text_input("Enter your SharePoint username")
password = st.text_input("Enter your SharePoint password", type="password")

# Authenticate and list files
if st.button("Authenticate and List Files"):
    if site_url and folder_url and username and password:
        ctx = authenticate_to_sharepoint(site_url, username, password)
        if ctx:
            files = list_files_in_folder(ctx, folder_url)
            if files:
                st.session_state["files"] = files
                st.success("Files retrieved successfully. Select the files below.")
            else:
                st.warning("No files found in the specified folder.")
    else:
        st.error("Please fill in all the fields.")

# Display files and allow user to select
if "files" in st.session_state:
    selected_files = st.multiselect("Select files to delete version history", st.session_state["files"])

    # Delete version history for selected files
    if st.button("Delete Version History"):
        if selected_files:
            delete_file_versions(ctx, folder_url, selected_files)
        else:
            st.error("Please select at least one file.")
