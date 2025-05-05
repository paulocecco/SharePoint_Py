from setuptools import setup, find_packages

setup(
    name="sharepoint_excel_tools",
    version="0.1.0",
    description="Read and write SharePoint Excel files using Graph API and pandas",
    author="Paulo Alvarez Cecco",
    author_email="paualvcec@gmail.com",
    url="https://github.com/paulocecco/sharepoint-excel-tools",
    packages=find_packages(),
    install_requires=[
        "pandas",
        "requests",
        "openpyxl",
        "msal"
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.7',
)
