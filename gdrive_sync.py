#!/usr/bin/env python3
"""
Google Drive and Sheets synchronization module for the ops-sim project.
This module provides functionality for syncing Excel files with Google Drive and Sheets.
"""

# Standard library imports
import json
import os
import pickle
from pathlib import Path
from typing import Optional, Dict, Any, Tuple, List, Union

# Third-party imports
import gspread
import numpy as np
import pandas as pd
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Local imports
import config


# Define the scopes
SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
]


def save_sheet_id(sheet_id: str) -> None:
    """Save the Google Sheet ID to the config module and to a persistent file.
    
    Args:
        sheet_id: The ID of the Google Sheet
        
    Raises:
        IOError: If there's an error saving the Sheet ID
    """
    try:
        # Update the config module in memory
        config.SHEET_ID = sheet_id
        
        # Save to the persistent store
        config_data = {}
        if os.path.exists(config.CONFIG_STORE_FILE):
            with open(config.CONFIG_STORE_FILE, 'r') as f:
                config_data = json.load(f)
        
        config_data['SHEET_ID'] = sheet_id
        
        with open(config.CONFIG_STORE_FILE, 'w') as f:
            json.dump(config_data, f, indent=2)
        
        print(f"Saved Google Sheet ID: {sheet_id} to config")
    except Exception as e:
        raise IOError(f"Could not save Sheet ID to persistent store: {str(e)}")


def load_saved_config() -> None:
    """Load saved configuration from the persistent store.
    
    Raises:
        IOError: If there's an error loading the configuration
    """
    try:
        if os.path.exists(config.CONFIG_STORE_FILE):
            with open(config.CONFIG_STORE_FILE, 'r') as f:
                config_data = json.load(f)
            
            if 'SHEET_ID' in config_data and config_data['SHEET_ID'] and not config.SHEET_ID:
                config.SHEET_ID = config_data['SHEET_ID']
                print(f"Loaded saved Sheet ID: {config.SHEET_ID}")
    except Exception as e:
        raise IOError(f"Could not load from persistent store: {str(e)}")


def get_credentials() -> Optional[Any]:
    """Get valid user credentials for Google API access.
    
    First tries to load credentials from token.pickle file.
    If that doesn't work, it will either refresh the token or
    initiate the OAuth2 flow.
    
    Returns:
        Credentials object, or None if authentication fails
        
    Raises:
        IOError: If there's an error with authentication
    """
    creds = None
    
    # Check if token.pickle exists
    if os.path.exists(config.TOKEN_PICKLE_FILE):
        with open(config.TOKEN_PICKLE_FILE, 'rb') as token:
            creds = pickle.load(token)
    
    # If there are no (valid) credentials, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            try:
                # Try to use service account if credentials.json exists
                if os.path.exists(config.CREDENTIALS_FILE):
                    creds = service_account.Credentials.from_service_account_file(
                        config.CREDENTIALS_FILE, scopes=SCOPES)
                elif os.path.exists(config.CLIENT_SECRET_FILE):
                    # Use OAuth flow if client_secret.json exists
                    flow = InstalledAppFlow.from_client_secrets_file(
                        config.CLIENT_SECRET_FILE, SCOPES)
                    creds = flow.run_local_server(port=0)
                else:
                    raise IOError(
                        "No authentication files found. Need either "
                        f"{config.CREDENTIALS_FILE} (service account) or "
                        f"{config.CLIENT_SECRET_FILE} (OAuth)"
                    )
            except Exception as e:
                raise IOError(
                    f"Error with authentication: {str(e)}\n"
                    f"Please make sure you have either a {config.CREDENTIALS_FILE} "
                    f"(service account) or {config.CLIENT_SECRET_FILE} file."
                )
        
        # Save the credentials for the next run
        with open(config.TOKEN_PICKLE_FILE, 'wb') as token:
            pickle.dump(creds, token)
    
    return creds


def _clean_dataframe(df: pd.DataFrame) -> List[List[Any]]:
    """Clean a DataFrame for Google Sheets upload.
    
    Args:
        df: DataFrame to clean
        
    Returns:
        List of lists representing the cleaned data
    """
    # Replace NaN values with empty strings
    df = df.fillna('')
    
    # Convert DataFrame to list of lists
    values = [df.columns.tolist()]  # First row is headers
    
    # Convert all values to strings if they're not already
    rows_data = []
    for _, row in df.iterrows():
        row_data = []
        for val in row:
            if pd.isna(val):  # Double-check for NaN/None values
                row_data.append('')
            elif isinstance(val, (int, float)):
                # For numbers, make sure they're within JSON range
                if np.isnan(val) or np.isinf(val):
                    row_data.append('')
                else:
                    row_data.append(val)
            else:
                # Convert to string for everything else
                row_data.append(str(val))
        rows_data.append(row_data)
    
    values.extend(rows_data)
    return values


def update_google_sheet(file_path: Union[str, Path], sheet_id: Optional[str] = None,
                       sheet_name: Optional[str] = None) -> Optional[str]:
    """Updates or creates a Google Sheet with data from an Excel file.
    
    Args:
        file_path: Path to the Excel file
        sheet_id: ID of the Google Sheet to update (if None, creates a new sheet)
        sheet_name: Name of the sheet file in Google Drive
        
    Returns:
        The ID of the created or updated Google Sheet, or None if failed
        
    Raises:
        FileNotFoundError: If the Excel file doesn't exist
        IOError: If there's an error updating the sheet
    """
    creds = get_credentials()
    if not creds:
        return None
    
    # Use config values if not provided
    if not sheet_name:
        sheet_name = config.SHEET_NAME
    
    # Connect to Google Sheets API
    gc = gspread.authorize(creds)
    
    try:
        # Load all sheets from the Excel file
        excel_data = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        
        if sheet_id:
            # Open existing spreadsheet
            try:
                spreadsheet = gc.open_by_key(sheet_id)
                print(f"Updating existing Google Sheet: {spreadsheet.title}")
            except Exception as e:
                print(f"Error opening Google Sheet: {str(e)}")
                print("Creating a new spreadsheet instead.")
                spreadsheet = gc.create(sheet_name)
                sheet_id = spreadsheet.id
        else:
            # Create new spreadsheet
            spreadsheet = gc.create(sheet_name)
            sheet_id = spreadsheet.id
            print(f"Created new Google Sheet with ID: {sheet_id}")
            
            # Move the new spreadsheet to the configured folder
            if config.GDRIVE_FOLDER_ID:
                drive_service = build('drive', 'v3', credentials=creds)
                try:
                    # First verify folder exists
                    folder = drive_service.files().get(
                        fileId=config.GDRIVE_FOLDER_ID,
                        fields='id, name, mimeType'
                    ).execute()
                    
                    if folder.get('mimeType') == 'application/vnd.google-apps.folder':
                        # Move the file to the folder
                        drive_service.files().update(
                            fileId=sheet_id,
                            addParents=config.GDRIVE_FOLDER_ID,
                            removeParents='root',
                            fields='id, parents'
                        ).execute()
                        print(f"Moved spreadsheet to folder: {folder.get('name')}")
                    else:
                        print(f"Warning: The provided folder ID is not a folder. Spreadsheet created in root.")
                except Exception as e:
                    print(f"Warning: Could not move spreadsheet to folder. Error: {str(e)}")
        
        # Get existing worksheets
        existing_worksheets = {worksheet.title: worksheet for worksheet in spreadsheet.worksheets()}
        excel_sheet_names = list(excel_data.keys())
        
        # First, make sure we have at least one worksheet
        # If spreadsheet has no worksheets, add a temporary one
        if not existing_worksheets:
            temp_worksheet = spreadsheet.add_worksheet(title="Temp", rows=1, cols=1)
            existing_worksheets["Temp"] = temp_worksheet
        
        # Process each worksheet from the Excel file
        processed_sheets = []
        
        for sheet_name, df in excel_data.items():
            # Skip certain sheets like graph sheets
            if sheet_name.endswith('-Graphs'):
                print(f"Skipping graph sheet '{sheet_name}'")
                continue
            
            processed_sheets.append(sheet_name)
            
            # Clean the dataframe
            values = _clean_dataframe(df)
            
            if sheet_name in existing_worksheets:
                # Update existing worksheet
                print(f"Updating existing worksheet: {sheet_name}")
                worksheet = existing_worksheets[sheet_name]
                try:
                    worksheet.clear()
                    worksheet.update(values)
                except Exception as e:
                    print(f"Error updating existing worksheet: {str(e)}")
                    # Try deleting and recreating the worksheet
                    try:
                        spreadsheet.del_worksheet(worksheet)
                        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=len(values), cols=len(values[0]))
                        worksheet.update(values)
                        print(f"Recreated worksheet: {sheet_name}")
                    except Exception as inner_e:
                        print(f"Failed to recreate worksheet: {str(inner_e)}")
            else:
                # Create new worksheet
                print(f"Adding new worksheet: {sheet_name}")
                worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=len(values), cols=len(values[0]))
                worksheet.update(values)
        
        # Remove unused worksheets
        for ws_title, worksheet in existing_worksheets.items():
            if ws_title not in processed_sheets:
                print(f"Removing unused worksheet: {ws_title}")
                spreadsheet.del_worksheet(worksheet)
        
        # Share with user if email is provided
        if config.USER_EMAIL:
            drive_service = build('drive', 'v3', credentials=creds)
            try:
                drive_service.permissions().create(
                    fileId=sheet_id,
                    body={'type': 'user', 'role': 'writer', 'emailAddress': config.USER_EMAIL},
                    fields='id'
                ).execute()
                print(f"Shared spreadsheet with {config.USER_EMAIL}")
            except Exception as e:
                print(f"Warning: Could not share with user: {str(e)}")
        
        return sheet_id
        
    except Exception as e:
        raise IOError(f"Error updating Google Sheet: {str(e)}")


def upload_to_gdrive(file_path: Union[str, Path], file_name: Optional[str] = None) -> Optional[str]:
    """Upload a file to Google Drive.
    
    Args:
        file_path: Path to the file to upload
        file_name: Name to give the file in Google Drive (if None, uses the original filename)
        
    Returns:
        The ID of the uploaded file, or None if failed
        
    Raises:
        FileNotFoundError: If the file doesn't exist
        IOError: If there's an error uploading the file
    """
    creds = get_credentials()
    if not creds:
        return None
    
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    
    if not file_name:
        file_name = os.path.basename(file_path)
    
    try:
        # Create Drive API service
        drive_service = build('drive', 'v3', credentials=creds)
        
        # Create file metadata
        file_metadata = {
            'name': file_name,
            'parents': [config.GDRIVE_FOLDER_ID] if config.GDRIVE_FOLDER_ID else None
        }
        
        # Create media
        media = MediaFileUpload(
            file_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            resumable=True
        )
        
        # Upload file
        file = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()
        
        # Share the file with the user
        if config.USER_EMAIL:
            drive_service.permissions().create(
                fileId=file.get('id'),
                body={'type': 'user', 'role': 'writer', 'emailAddress': config.USER_EMAIL},
                fields='id'
            ).execute()
        
        print(f"Uploaded file to Google Drive with ID: {file.get('id')}")
        return file.get('id')
        
    except Exception as e:
        raise IOError(f"Error uploading to Google Drive: {str(e)}")


def sync_to_google(sheet_id: Optional[str] = None, upload_to_drive: bool = False) -> Tuple[Optional[str], Optional[str]]:
    """Sync the master Excel file to Google Drive and Sheets.
    
    Args:
        sheet_id: ID of the Google Sheet to update (if None, uses the one from config)
        upload_to_drive: Whether to upload the Excel file to Google Drive (default: False)
        
    Returns:
        Tuple containing:
        - sheet_id: ID of the updated/created Google Sheet
        - file_id: ID of the uploaded file in Google Drive (if upload_to_drive is True)
        
    Raises:
        FileNotFoundError: If the master file doesn't exist
        IOError: If there's an error during sync
    """
    if not os.path.exists(config.MASTER_FILE):
        raise FileNotFoundError(f"Master file not found: {config.MASTER_FILE}")
    
    # Use the sheet_id from config if none is provided and config has one
    if not sheet_id and config.SHEET_ID:
        sheet_id = config.SHEET_ID
        print(f"Using existing Google Sheet ID from config: {sheet_id}")
    
    try:
        # Update Google Sheet
        sheet_id = update_google_sheet(config.MASTER_FILE, sheet_id)
        if not sheet_id:
            raise IOError("Failed to update Google Sheet")
        
        # Save the sheet ID for future use
        save_sheet_id(sheet_id)
        
        # Upload to Google Drive if requested
        file_id = None
        if upload_to_drive:
            file_id = upload_to_gdrive(config.MASTER_FILE)
            if not file_id:
                raise IOError("Failed to upload to Google Drive")
        
        return sheet_id, file_id
        
    except Exception as e:
        raise IOError(f"Error during sync: {str(e)}")


# Load saved configuration at import time
load_saved_config()

if __name__ == "__main__":
    # Only update Google Sheet by default, without uploading to Drive
    import argparse
    
    parser = argparse.ArgumentParser(description="Sync master Excel file to Google Sheets/Drive")
    parser.add_argument("--upload-to-drive", action="store_true", 
                       help="Also upload Excel file to Google Drive")
    args = parser.parse_args()
    
    sheet_id, file_id = sync_to_google(upload_to_drive=args.upload_to_drive)
    if sheet_id:
        print(f"Synced to Google Sheet: {sheet_id}")
        if file_id:
            print(f"Uploaded to Google Drive: {file_id}") 