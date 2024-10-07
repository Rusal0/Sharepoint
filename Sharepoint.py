import streamlit as st
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

# Streamlit UI for entering SharePoint details
st.title("SharePoint Version History Deletion Tool")

sharepoint_url = st.text_input("Enter the SharePoint URL:")
username = st.text_input("Enter your SharePoint Username:")
password = st.text_input("Enter your SharePoint Password:", type="password")

if st.button("Connect"):
    if sharepoint_url and username and password:
        try:
            # Connect to SharePoint
            ctx_auth = AuthenticationContext(sharepoint_url)
            if ctx_auth.acquire_token_for_user(username, password):
                ctx = ClientContext(sharepoint_url, ctx_auth)
                web = ctx.web
                ctx.load(web)
                ctx.execute_query()
                st.success(f"Connected to {web.properties['Title']}")
                
                # Load all files eligible for version history deletion
                folder = ctx.web.get_folder_by_server_relative_url("/")
                files = folder.files
                ctx.load(files)
                ctx.execute_query()

                file_options = []
                for file in files:
                    file_options.append(file.properties["Name"])

                if file_options:
                    selected_files = st.multiselect(
                        "Select files to delete version history:",
                        file_options
                    )

                    if st.button("Delete Version History"):
                        if selected_files:
                            for file_name in selected_files:
                                file = folder.files.get_by_url(file_name)
                                ctx.load(file)
                                ctx.execute_query()
                                
                                # Deleting the version history
                                versions = file.versions
                                ctx.load(versions)
                                ctx.execute_query()
                                
                                for version in versions:
                                    version.delete_object()
                                    ctx.execute_query()
                                
                            st.success("Version history for selected files has been deleted successfully.")
                        else:
                            st.warning("Please select at least one file to delete the version history.")
                else:
                    st.warning("No files found in the specified SharePoint folder.")
            else:
                st.error("Authentication failed. Please check your credentials.")
        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.warning("Please enter all required details.")
