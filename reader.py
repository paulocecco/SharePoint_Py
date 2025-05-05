import pandas as pd
import requests
from io import BytesIO

def fetch_sharepoint_excel(
    tenant_id: str,
    client_id: str,
    client_secret: str,
    sharepoint_domain: str,
    site_name: str,
    file_path: str,
    sheet_name: str
) -> pd.DataFrame:
    """
    Downloads an Excel file from SharePoint and returns a specified sheet as a Pandas DataFrame.

    Parameters:
        tenant_id (str): Azure tenant ID
        client_id (str): Azure app client ID
        client_secret (str): Azure app client secret
        sharepoint_domain (str): Your SharePoint domain (e.g., 'yourcompany.sharepoint.com')
        site_name (str): SharePoint site name
        file_path (str): Path to the Excel file in SharePoint document library
        sheet_name (str): Sheet name to load from Excel file

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
    response.raise_for_status()
    access_token = response.json().get("access_token")

    # Step 2: Get SharePoint Site ID
    headers = {"Authorization": f"Bearer {access_token}"}
    site_url = f"https://graph.microsoft.com/v1.0/sites/{sharepoint_domain}:/sites/{site_name}"
    response = requests.get(site_url, headers=headers)
    response.raise_for_status()
    site_id = response.json()["id"]

    # Step 3: Download Excel file
    file_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}:/content"
    response = requests.get(file_url, headers=headers)
    response.raise_for_status()

    # Step 4: Load Excel content into DataFrame
    xls = BytesIO(response.content)
    df = pd.read_excel(xls, sheet_name=sheet_name)
    return df
