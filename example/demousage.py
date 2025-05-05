from reader import fetch_sharepoint_excel
from writer import upload_dataframes_to_sharepoint_excel

# Authentication and SharePoint setup
tenant_id = "your-tenant-id"
client_id = "your-client-id"
client_secret = "your-client-secret"
domain = "yourcompany.sharepoint.com"
site = "YourSiteName"
path = "Shared Documents/Reports/your_excel.xlsx"

# Read sheet
df_sales = fetch_sharepoint_excel(tenant_id, client_id, client_secret, domain, site, path, sheet_name="Sales")

# Modify or create new DataFrame
df_forecast = df_sales.copy()

# Upload
upload_dataframes_to_sharepoint_excel(
    tenant_id, client_id, client_secret,
    domain, site, path,
    sheet_data={"Sales": df_sales, "Forecast": df_forecast}
)
