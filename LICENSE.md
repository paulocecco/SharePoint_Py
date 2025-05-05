
---

### 📗 `README_Read_SharePoint_excel2pandas.md`

```markdown
# 📥 Read SharePoint Excel to pandas

This script allows you to read a specific sheet from an Excel file stored in SharePoint directly into a pandas DataFrame using Microsoft Graph API.

## 📌 Function

```python
fetch_sharepoint_excel(
    tenant_id,
    client_id,
    client_secret,
    sharepoint_domain,
    site_name,
    file_path,
    sheet_name
)
