import pandas as pd
import requests
from io import BytesIO
import json

def fetch_sharepoint_excel(
    tenant_id: str,
    client_id: str,
    client_secret: str,
    sharepoint_domain: str,
    site_name: str,
    file_path: str,
    sheet_name: str,
    debug: bool = False
) -> pd.DataFrame:
    """
    Downloads an Excel file from SharePoint and returns a specified sheet as a Pandas DataFrame.

    Parameters:
        tenant_id (str): Azure tenant ID
        client_id (str): Azure app client ID
        client_secret (str): Azure app client secret
        sharepoint_domain (str): Your SharePoint domain (e.g., 'yourcompany.sharepoint.com')
        site_name (str): SharePoint site name
        file_path (str): Path to the Excel file in SharePoint document library (e.g., '/Shared Documents/folder/file.xlsx')
        sheet_name (str): Sheet name to load from Excel file
        debug (bool): Whether to print step-by-step debugging output

    Returns:
        pd.DataFrame: Data from the specified Excel sheet
    """
    # Step 1: Authenticate with Microsoft Graph
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    token_data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default"
    }
    response = requests.post(token_url, data=token_data)
    if debug:
        print("🔐 Token response:", response.status_code, response.text)
    response.raise_for_status()
    access_token = response.json().get("access_token")
    if debug:
        print("✅ Access token received:", access_token[:80], "...")

    headers = {"Authorization": f"Bearer {access_token}"}

    # Step 2: Get SharePoint Site ID
    site_url = f"https://graph.microsoft.com/v1.0/sites/{sharepoint_domain}:/sites/{site_name}"
    response = requests.get(site_url, headers=headers)
    response.raise_for_status()
    site_id = response.json()["id"]
    if debug:
        print(f"📍 Site ID for '{site_name}': {site_id}")

    # Step 3: Explore directory contents (optional)
    if debug:
        children_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
        children_response = requests.get(children_url, headers=headers)
        print("📁 Root folder contents:", json.dumps(children_response.json(), indent=2))

        # Optionally, inspect folder-level contents where the file is supposed to be
        if "/" in file_path:
            folder_path = "/".join(file_path.split("/")[:-1])
            folder_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{folder_path}:/children"
            folder_response = requests.get(folder_url, headers=headers)
            print(f"📂 Contents of folder '{folder_path}':", json.dumps(folder_response.json(), indent=2))

    # Step 4: Download Excel file content
    file_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}:/content"
    if debug:
        print("⬇️  File download URL:", file_url)
    response = requests.get(file_url, headers=headers)
    response.raise_for_status()

    # Step 5: Load Excel content into DataFrame
    xls = BytesIO(response.content)
    try:
        df = pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl")
        print(f"✅ Excel loaded successfully. Rows: {df.shape[0]}")
        return df
    except Exception as e:
        raise ValueError(f"❌ Failed to load Excel file: {e}")
