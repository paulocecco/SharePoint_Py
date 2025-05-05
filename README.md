# 📁 SharePoint Excel Integration with Python

This repository provides two Python utilities to help you integrate Microsoft SharePoint Excel files with your data workflows using pandas and Microsoft Graph API.

---

## 🔧 Modules

### 1. `Read_SharePoint_excel2pandas.py`

Fetch a sheet from an Excel file stored in SharePoint and load it into a pandas DataFrame.
📝 You'll be needing the following permissions:
1) Microsoft Graph - User.Read
2) SharePoint - Sites.Read.All

### 2. `Save_SharePoint_pandas2excel.py`

Update or add one or more sheets to an Excel file on SharePoint from pandas DataFrames.
📝 You'll be needing the following permissions:
1) Microsoft Graph - User.ReadWrite
2) SharePoint - Sites.ReadWrite.All

---

## 🚀 Installation

```bash
pip install pandas openpyxl requests msal
