# sharepoint_etl.py

import os
import pandas as pd
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.sharepoint.listitems.listitem import ListItem
from office365.sharepoint.lists.list import List
from office365.sharepoint.folders.folder import Folder
import io
import json
import datetime

# --- Configuration ---
SHAREPOINT_SITE_URL = os.environ.get("SHAREPOINT_SITE_URL")
SHAREPOINT_CLIENT_ID = os.environ.get("SHAREPOINT_CLIENT_ID")
SHAREPOINT_CLIENT_SECRET = os.environ.get("SHAREPOINT_CLIENT_SECRET")
SHAREPOINT_DOC_LIBRARY = os.environ.get("SHAREPOINT_DOC_LIBRARY", "Documents") # Default to "Documents" if not set
SHAREPOINT_FOLDER_PATH = os.environ.get("SHAREPOINT_FOLDER_PATH", "Shared Documents/Your/Target/Folder") # Default path
MASTER_OUTPUT_FILE = os.environ.get("MASTER_OUTPUT_FILE", "master_lab_results.xlsx") # Default output file name
PROCESSED_FILES_LOG = os.environ.get("PROCESSED_FILES_LOG", "processed_files.json") # Default log file name

# --- SharePoint Connection ---
def connect_to_sharepoint():
    """Connects to SharePoint Online using client credentials."""
    if not all([SHAREPOINT_SITE_URL, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET]):
        print("Error: SharePoint credentials not found in environment variables.")
        print("Please set SHAREPOINT_SITE_URL, SHAREPOINT_CLIENT_ID, and SHAREPOINT_CLIENT_SECRET.")
        return None

    try:
        ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(
            ClientCredential(SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET)
        )
        # Verify connection by getting the web title
        ctx.web.get().execute_query()
        print(f"Successfully connected to SharePoint site: {ctx.web.title}")
        return ctx
    except Exception as e:
        print(f"Error connecting to SharePoint: {e}")
        return None

# --- File Discovery ---
def list_excel_files_recursive(ctx, folder_path):
    """Recursively lists all Excel files in a SharePoint folder."""
    excel_files = []
    try:
        root_folder = ctx.web.get_folder_by_server_relative_url(folder_path)
        ctx.load(root_folder)
        ctx.load(root_folder.files)
        ctx.load(root_folder.folders)
        ctx.execute_query()

        # Add Excel files in the current folder
        for file in root_folder.files:
            if file.name.lower().endswith(('.xlsx', '.xls')):
                excel_files.append(file.server_relative_url)

        # Recursively process subfolders
        for folder in root_folder.folders:
            # Skip system folders like "Forms"
            if not folder.name.startswith('_') and folder.name != 'Forms':
                 excel_files.extend(list_excel_files_recursive(ctx, folder.server_relative_url))

    except Exception as e:
        print(f"Error listing files in {folder_path}: {e}")

    return excel_files

# --- Data Extraction and Processing ---
def read_excel_from_sharepoint(ctx, file_url):
    """Reads an Excel file from SharePoint into a pandas DataFrame."""
    try:
        file = File.open_binary(ctx, file_url)
        ctx.execute_query()
        file_content = io.BytesIO(file.content)
        return file_content
    except Exception as e:
        print(f"Error reading file {file_url}: {e}")
        return None

def process_excel_file(file_content):
    """Processes a single Excel file, extracting data from specified sheets."""
    all_sheet_data = pd.DataFrame()
    sheets_to_process = ["Batch Sheet", "Product Info"] # Sheets to look for

    try:
        xls = pd.ExcelFile(file_content)
        available_sheets = xls.sheet_names

        for sheet_name in sheets_to_process:
            # Simple check for sheet name variations (can be expanded)
            actual_sheet_name = None
            for available_sheet in available_sheets:
                if sheet_name.lower() in available_sheet.lower():
                    actual_sheet_name = available_sheet
                    break

            if actual_sheet_name:
                print(f"  Processing sheet: {actual_sheet_name}")
                try:
                    # Read the sheet, assuming the first row is headers
                    df = xls.parse(actual_sheet_name)
                    # Add a column to indicate the source sheet
                    df['_SourceSheet'] = actual_sheet_name
                    all_sheet_data = pd.concat([all_sheet_data, df], ignore_index=True)
                except Exception as e:
                    print(f"    Error reading sheet {actual_sheet_name}: {e}")
            else:
                print(f"  Warning: Sheet '{sheet_name}' not found in the file.")

    except Exception as e:
        print(f"Error processing Excel content: {e}")

    return all_sheet_data

# --- Data Cleaning and Filtering ---
def clean_and_filter_data(df):
    """Cleans and filters the consolidated data."""
    initial_rows = len(df)
    print(f"Initial rows before cleaning: {initial_rows}")

    # 1. Skip blank rows (check if all values in a row are NaN or None)
    df.dropna(how='all', inplace=True)
    print(f"Rows after dropping blank rows: {len(df)}")

    # 2. Skip rows identified as QC rows (example patterns)
    # You might need to expand these patterns based on your data
    qc_patterns = ["CCV", "MB", "Blank", "Check"]
    # Create a regex pattern that matches any of the QC patterns in any column
    qc_regex = '|'.join(qc_patterns)
    # Check if any column in a row contains a QC pattern (case-insensitive)
    df = df[~df.astype(str).apply(lambda row: row.str.contains(qc_regex, case=False, na=False)).any(axis=1)]
    print(f"Rows after dropping QC rows: {len(df)}")

    # 3. Skip partial entries (define key columns that must be present)
    # Replace 'KeyColumn1', 'KeyColumn2' with actual column names that must be non-null
    key_columns = ['Sample ID', 'Result'] # Example key columns - **UPDATE THIS**
    # Ensure key columns exist before checking
    key_columns_exist = [col for col in key_columns if col in df.columns]
    if key_columns_exist:
        df.dropna(subset=key_columns_exist, inplace=True)
        print(f"Rows after dropping partial entries (based on {key_columns_exist}): {len(df)}")
    else:
        print(f"Warning: Key columns for partial entry check not found: {key_columns}")


    # Add more cleaning steps as needed (e.g., data type conversion, removing extra whitespace)
    # Example: Strip whitespace from string columns
    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].str.strip()

    print(f"Final rows after cleaning: {len(df)}")
    return df

# --- Daily Update Logic ---
def load_processed_files_log(log_file):
    """Loads the log of processed files and their timestamps."""
    if os.path.exists(log_file):
        with open(log_file, 'r') as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                print(f"Warning: Could not decode JSON from {log_file}. Starting with empty log.")
                return {}
    return {}

def save_processed_files_log(log_file, processed_files):
    """Saves the log of processed files and their timestamps."""
    with open(log_file, 'w') as f:
        json.dump(processed_files, f, indent=4)

def get_sharepoint_file_modified_time(ctx, file_url):
    """Gets the last modified time of a SharePoint file."""
    try:
        file = ctx.web.get_file_by_server_relative_url(file_url)
        ctx.load(file, ["TimeLastModified"])
        ctx.execute_query()
        # TimeLastModified is returned as a string like '2023-10-27T10:00:00Z'
        # Convert to datetime object
        return datetime.datetime.fromisoformat(file.time_last_modified.replace('Z', '+00:00'))
    except Exception as e:
        print(f"Error getting modified time for {file_url}: {e}")
        return None

# --- Main ETL Process ---
def run_etl():
    """Runs the main ETL process."""
    ctx = connect_to_sharepoint()
    if not ctx:
        return

    print(f"Searching for Excel files in: {SHAREPOINT_DOC_LIBRARY}/{SHAREPOINT_FOLDER_PATH}")
    # Construct the full server relative path
    full_sharepoint_path = f"/{SHAREPOINT_DOC_LIBRARY}/{SHAREPOINT_FOLDER_PATH}"
    all_excel_files = list_excel_files_recursive(ctx, full_sharepoint_path)
    print(f"Found {len(all_excel_files)} potential Excel files.")

    processed_files_log = load_processed_files_log(PROCESSED_FILES_LOG)
    data_to_append = pd.DataFrame()
    updated_processed_files_log = processed_files_log.copy()

    for file_url in all_excel_files:
        last_modified_time = get_sharepoint_file_modified_time(ctx, file_url)

        if last_modified_time:
            # Convert datetime to string for comparison and storage
            last_modified_str = last_modified_time.isoformat()

            if file_url in processed_files_log:
                # File was processed before, check if modified
                if processed_files_log[file_url] == last_modified_str:
                    print(f"Skipping {file_url}: Not modified since last run.")
                    continue
                else:
                    print(f"Processing {file_url}: Modified since last run.")
            else:
                print(f"Processing {file_url}: New file.")

            # Process the file
            file_content = read_excel_from_sharepoint(ctx, file_url)
            if file_content:
                file_data = process_excel_file(file_content)
                if not file_data.empty:
                    cleaned_data = clean_and_filter_data(file_data)
                    data_to_append = pd.concat([data_to_append, cleaned_data], ignore_index=True)
                    # Update the log with the new modified time
                    updated_processed_files_log[file_url] = last_modified_str
                else:
                    print(f"  No data extracted from {file_url}.")
            else:
                 print(f"  Could not read content for {file_url}. Skipping processing.")
        else:
            print(f"Could not get modified time for {file_url}. Skipping for now.")


    if not data_to_append.empty:
        print(f"\nAppending {len(data_to_append)} new/modified rows to {MASTER_OUTPUT_FILE}")
        try:
            # Check if master file exists to decide between writing and appending
            if os.path.exists(MASTER_OUTPUT_FILE):
                # Read existing data and append
                existing_data = pd.read_excel(MASTER_OUTPUT_FILE)
                consolidated_data = pd.concat([existing_data, data_to_append], ignore_index=True)
            else:
                # No existing file, this is the first run or file was deleted
                consolidated_data = data_to_append

            # Write the consolidated data to the master file
            consolidated_data.to_excel(MASTER_OUTPUT_FILE, index=False)
            print(f"Successfully updated {MASTER_OUTPUT_FILE} with {len(data_to_append)} new/modified rows.")

            # Save the updated processed files log
            save_processed_files_log(PROCESSED_FILES_LOG, updated_processed_files_log)
            print(f"Updated processed files log: {PROCESSED_FILES_LOG}")

        except Exception as e:
            print(f"Error writing to master Excel file {MASTER_OUTPUT_FILE}: {e}")
    else:
        print("\nNo new or modified data to append.")
        # If no new data, still save the log to record files that were checked
        save_processed_files_log(PROCESSED_FILES_LOG, updated_processed_files_log)
        print(f"Updated processed files log: {PROCESSED_FILES_LOG}")


if __name__ == "__main__":
    run_etl()