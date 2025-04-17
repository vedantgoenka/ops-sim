#!/usr/bin/env python3
"""
Test script for Google Drive and Sheets integration.
This will help diagnose authentication issues and test the sync functionality.
"""

import os
import sys
from pathlib import Path
from gdrive_sync import get_credentials, sync_to_google, update_google_sheet, upload_to_gdrive
import pandas as pd
import tempfile
import argparse
import config  # Import configuration settings

def create_test_excel():
    """Create a simple test Excel file with some data."""
    # Create a temporary file
    temp_dir = tempfile.gettempdir()
    test_file = Path(temp_dir) / "test_gdrive_sync.xlsx"
    
    # Create a simple DataFrame
    df1 = pd.DataFrame({
        'Name': ['Test 1', 'Test 2', 'Test 3'],
        'Value': [100, 200, 300]
    })
    
    df2 = pd.DataFrame({
        'ID': [1, 2, 3],
        'Description': ['Description 1', 'Description 2', 'Description 3']
    })
    
    # Write to Excel with multiple sheets
    with pd.ExcelWriter(test_file, engine='openpyxl') as writer:
        df1.to_excel(writer, sheet_name='Sheet1', index=False)
        df2.to_excel(writer, sheet_name='Sheet2', index=False)
    
    print(f"Created test Excel file at: {test_file}")
    return test_file

def test_auth():
    """Test Google API authentication."""
    print("\n=== Testing Google API Authentication ===")
    
    creds = get_credentials()
    if creds:
        print("✅ Authentication successful!")
        if hasattr(creds, 'service_account_email'):
            print(f"Using service account: {creds.service_account_email}")
        else:
            print("Using OAuth credentials")
        return True
    else:
        print("❌ Authentication failed!")
        print("Please check that you have either:")
        print("  - A credentials.json file (for service accounts)")
        print("  - A client_secret.json file (for OAuth)")
        return False

def test_sheets():
    """Test Google Sheets integration."""
    print("\n=== Testing Google Sheets Integration ===")
    
    test_file = create_test_excel()
    
    # Try to create a new sheet
    sheet_id = update_google_sheet(test_file, sheet_name="Test Sheet")
    
    if sheet_id:
        print(f"✅ Successfully created Google Sheet with ID: {sheet_id}")
        print(f"View at: https://docs.google.com/spreadsheets/d/{sheet_id}")
        return True
    else:
        print("❌ Failed to create Google Sheet")
        return False

def test_drive():
    """Test Google Drive integration."""
    print("\n=== Testing Google Drive Integration ===")
    
    test_file = create_test_excel()
    
    # Try to upload to Google Drive
    file_id = upload_to_gdrive(test_file)
    
    if file_id:
        print(f"✅ Successfully uploaded file to Google Drive with ID: {file_id}")
        print(f"View at: https://drive.google.com/file/d/{file_id}/view")
        return True
    else:
        print("❌ Failed to upload to Google Drive")
        return False

def test_sync(upload_to_drive=True):
    """Test the sync_to_google function.
    
    Args:
        upload_to_drive: Whether to upload to Google Drive (default: True)
    """
    print("\n=== Testing Full Google Sync ===")
    if not upload_to_drive:
        print("Note: Drive upload is disabled for this test")
    
    test_file = create_test_excel()
    
    # Temporarily override the FOLDER_PATH used by sync_to_google
    import gdrive_sync
    original_path = gdrive_sync.FOLDER_PATH
    temp_dir = Path(tempfile.gettempdir())
    
    # Copy the test file to a file named "Master.xlsx" in the temp directory
    master_file = temp_dir / "Master.xlsx"
    import shutil
    shutil.copy(test_file, master_file)
    
    try:
        # Override the FOLDER_PATH
        gdrive_sync.FOLDER_PATH = temp_dir
        
        # Run the sync function
        sheet_id, file_id = sync_to_google(upload_to_drive=upload_to_drive)
        
        if sheet_id:
            print("✅ Successfully updated Google Sheet")
            print(f"Sheet ID: {sheet_id}")
            print(f"Sheet URL: https://docs.google.com/spreadsheets/d/{sheet_id}")
            
            if upload_to_drive:
                if file_id:
                    print("✅ Successfully uploaded to Google Drive")
                    print(f"File ID: {file_id}")
                    print(f"File URL: https://drive.google.com/file/d/{file_id}/view")
                    return True
                else:
                    print("⚠️ Google Sheet update succeeded but Drive upload failed")
                    return False
            else:
                # If we're not uploading to Drive, file_id should be None
                if file_id is None:
                    return True
                else:
                    print("⚠️ Unexpected file_id when upload_to_drive=False")
                    return False
        else:
            print("❌ Failed to sync to Google")
            return False
    finally:
        # Restore the original FOLDER_PATH
        gdrive_sync.FOLDER_PATH = original_path
        
        # Clean up the temporary master file
        if master_file.exists():
            master_file.unlink()

def main():
    """Run all tests."""
    # Parse command line arguments
    parser = argparse.ArgumentParser(description="Test Google Drive and Sheets integration")
    parser.add_argument("--folder-id", type=str, help="Google Drive folder ID to use for testing (overrides config)")
    parser.add_argument("--user-email", type=str, help="Email address to share access with (overrides config)")
    parser.add_argument("--test-sheets-only", action="store_true", help="Only test Google Sheets integration")
    parser.add_argument("--test-drive-only", action="store_true", help="Only test Google Drive integration")
    parser.add_argument("--test-sync-only", action="store_true", help="Only test full sync functionality")
    parser.add_argument("--no-drive-upload", action="store_true", help="Skip uploading to Drive during sync test")
    
    args = parser.parse_args()
    
    # Store original config values to restore later
    original_folder_id = config.GDRIVE_FOLDER_ID
    original_user_email = config.USER_EMAIL
    
    # Update config values if provided in arguments
    if args.folder_id:
        config.GDRIVE_FOLDER_ID = args.folder_id
    if args.user_email:
        config.USER_EMAIL = args.user_email
    
    print("Google Drive and Sheets Integration Test")
    print("=======================================")
    
    print(f"Using folder ID: {config.GDRIVE_FOLDER_ID or 'Not set (will use root)'}")
    print(f"Files will be shared with: {config.USER_EMAIL}")
    if args.no_drive_upload:
        print("Drive upload is disabled for sync test")
    
    try:
        # First test authentication
        if not test_auth():
            print("\n❌ Authentication failed, cannot proceed with other tests.")
            sys.exit(1)
        
        # Determine which tests to run
        run_all = not (args.test_sheets_only or args.test_drive_only or args.test_sync_only)
        
        # Run the selected tests
        sheets_result = True
        drive_result = True
        sync_result = True
        
        if run_all or args.test_sheets_only:
            sheets_result = test_sheets()
            
        if run_all or args.test_drive_only:
            drive_result = test_drive()
            
        if run_all or args.test_sync_only:
            # Pass the upload_to_drive parameter
            sync_result = test_sync(upload_to_drive=not args.no_drive_upload)
        
        # Print summary
        print("\n=== Test Summary ===")
        print(f"Authentication: ✅ Passed")
        
        if run_all or args.test_sheets_only:
            print(f"Google Sheets:  {'✅ Passed' if sheets_result else '❌ Failed'}")
            
        if run_all or args.test_drive_only:
            print(f"Google Drive:   {'✅ Passed' if drive_result else '❌ Failed'}")
            
        if run_all or args.test_sync_only:
            print(f"Full Sync:      {'✅ Passed' if sync_result else '❌ Failed'}")
        
        # Return exit code based on results
        tests_run = []
        if run_all or args.test_sheets_only:
            tests_run.append(sheets_result)
        if run_all or args.test_drive_only:
            tests_run.append(drive_result)
        if run_all or args.test_sync_only:
            tests_run.append(sync_result)
        
        if all(tests_run):
            print("\n✅ All tests passed! Your Google Drive and Sheets integration is working.")
            return 0
        else:
            print("\n❌ Some tests failed. Check the error messages above.")
            return 1
    finally:
        # Restore original config values
        config.GDRIVE_FOLDER_ID = original_folder_id
        config.USER_EMAIL = original_user_email

if __name__ == "__main__":
    sys.exit(main()) 