import streamlit as st
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext

def delete_version_history(site_url, client_id, client_secret, library_url):
    # Authenticate
    ctx_auth = AuthenticationContext(site_url)
    if not ctx_auth.acquire_token_for_client(client_id, client_secret):
        return "Failed to authenticate"

    ctx = ClientContext(site_url, ctx_auth)

    # Get the list
    document_library = ctx.web.lists.get_by_title(library_url)
    ctx.load(document_library)
    ctx.execute_query()

    # Get all items in the library
    items = document_library.items
    ctx.load(items)
    ctx.execute_query()

    deletion_messages = []

    for item in items:
        # Load the item's version history
        versions = item.versions
        ctx.load(versions)
        ctx.execute_query()

        # Delete each version
        for version in versions:
            version.delete_object()
            deletion_messages.append(f"Deleted version: {version.version_label} for item: {item.properties['FileLeafRef']}")

        # Execute the deletion of versions
        ctx.execute_query()

    return "\n".join(deletion_messages) if deletion_messages else "No versions to delete."

# Streamlit UI
def main():
    st.title("SharePoint Version History Deletion Tool")

    # Input fields for user credentials and library details
    site_url = st.text_input("SharePoint Site URL")
    client_id = st.text_input("Client ID")
    client_secret = st.text_input("Client Secret", type="password")
    library_url = st.text_input("Document Library Name")

    if st.button("Delete Version History"):
        if site_url and client_id and client_secret and library_url:
            result = delete_version_history(site_url, client_id, client_secret, library_url)
            st.success(result)
        else:
            st.error("Please fill all the fields.")

if __name__ == "__main__":
    main()
