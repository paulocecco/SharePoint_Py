import pandas as pd
import msal
import requests
import io

def upload_dataframes_to_sharepoint_excel(
    tenant_id: str,
    client_id: str,
    client_secret: str,
    sharepoint_domain: str,
    site_name: str,
    file_path: str,
    sheet_data: dict
):
    """
    Uploads one or more pandas DataFrames to specific sheets in an existing SharePoint Excel file.

    Parameters:
        tenant_id (str): Azure tenant ID
        client_id (str): App registration client ID
        client_secret (str): App registration secret
        sharepoint_domain (str): Your SharePoint domain, e.g. 'yourcompany.sharepoint.com'
        site_name (str): SharePoint site name (not full URL)
        file_path (str): Path to the Excel file within the document library
        sheet_data (dict): A dictionary of {sheet_name: DataFrame} pairs to upload
    """

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scope = ["https://graph.microsoft.com/.default"]

    # Authenticate using MSAL
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret
    )
    token_result = app.acquire_token_for_client(scopes=scope)
    if 'access_token' not in token_result:
        raise Exception(f"Failed to get access token: {token_result.get('error_description')}")

    access_token = token_result['access_token']
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    graph_url = "https://graph.microsoft.com/v1.0"

    # Get site ID
    site_resp = requests.get(
        f"{graph_url}/sites/{sharepoint_domain}:/sites/{site_name}",
        headers=headers
    )
    site_resp.raise_for_status()
    site_id = site_resp.json()['id']

    # Get file ID
    file_resp = requests.get(
        f"{graph_url}/sites/{site_id}/drive/root:/{file_path}",
        headers=headers
    )
    file_resp.raise_for_status()
    file_id = file_resp.json()['id']

    # Download existing Excel file
    download_url = f"{graph_url}/sites/{site_id}/drive/items/{file_id}/content"
    file_content = requests.get(download_url, headers=headers).content

    with pd.ExcelFile(io.BytesIO(file_content)) as xls:
        df_existing = pd.read_excel(xls, sheet_name=None)

    # Update or add new sheets
    for sheet_name, df in sheet_data.items():
        df_existing[sheet_name] = df

    # Save all sheets to memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, df in df_existing.items():
            df.to_excel(writer, sheet_name=name, index=False)
    output.seek(0)

    # Upload updated Excel file back to SharePoint
    upload_url = f"{graph_url}/sites/{site_id}/drive/items/{file_id}/content"
    upload_resp = requests.put(upload_url, headers={
        'Authorization': f'Bearer {access_token}'
    }, data=output)

    if upload_resp.status_code == 200:
        print("✅ File uploaded successfully to SharePoint.")
    else:
        raise Exception(f"❌ Upload failed: {upload_resp.status_code} - {upload_resp.text}")
