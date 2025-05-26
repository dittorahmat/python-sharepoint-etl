# Python SharePoint ETL

This project provides a Python script to extract and consolidate lab test results from Excel files stored in SharePoint Online. It automates the process of downloading files, reading data, and combining it into a single dataset.

## Setup and Usage

1.  **Install Dependencies:**
    Install the required Python packages using pip:
    ```bash
    pip install -r requirements.txt
    ```

2.  **Configure SharePoint Authentication:**
    Set the following environment variables for SharePoint authentication:
    ```bash
    SHAREPOINT_SITE_URL=YOUR_SHAREPOINT_SITE_URL
    SHAREPOINT_CLIENT_ID=YOUR_CLIENT_ID
    SHAREPOINT_CLIENT_SECRET=YOUR_CLIENT_SECRET
    ```
    You can use a `.env` file and a library like `python-dotenv` to manage these variables, or set them directly in your environment.

3.  **Configure Script Variables:**
    Update the following variables within the [`sharepoint_etl.py`](sharepoint_etl.py) script to match your specific SharePoint document library, folder path, and data format:
    *   `SHAREPOINT_DOC_LIBRARY`
    *   `SHAREPOINT_FOLDER_PATH`
    *   `key_columns`
    *   `qc_patterns`

4.  **Run the Script:**
    Execute the script using the following command:
    ```bash
    python sharepoint_etl.py
    ```
