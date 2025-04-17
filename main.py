#!/usr/bin/env python3
"""
Main entry point for the ops-sim data processing and sync operations.
This script orchestrates the data processing pipeline and Google Drive/Sheets sync.
"""

# Standard library imports
import argparse
from typing import Optional, Tuple

# Local imports
import config
from append import append_to_master
from analysis import DataAnalyzer
from gdrive_sync import sync_to_google


def run_all(sync_to_cloud: bool = True, sheet_id: Optional[str] = None, 
            upload_to_drive: bool = False) -> Tuple[Optional[str], Optional[str]]:
    """Run all data processing and sync operations.
    
    This function orchestrates the entire data processing pipeline:
    1. Appends new data to the master file
    2. Analyzes the data (adds current price and capacity allocation)
    3. Optionally syncs to Google Drive/Sheets
    
    Args:
        sync_to_cloud: Whether to sync to Google Drive/Sheets
        sheet_id: ID of Google Sheet to update (None to use the one in config.py)
        upload_to_drive: Whether to upload the Excel file to Google Drive (default: False)
        
    Returns:
        Tuple containing:
        - sheet_id: ID of the updated/created Google Sheet
        - file_id: ID of the uploaded file in Google Drive (if upload_to_drive is True)
    """
    try:
        print(f"Using data folder: {config.DATA_FOLDER_PATH}")
        
        # Run the append_to_master function
        append_to_master()
        
        # Create an instance of DataAnalyzer and run analysis
        analyzer = DataAnalyzer()
        analyzer.add_current_price()
        analyzer.add_capacity_allocation()
        analyzer.add_batch_sizes()
        
        # Sync to Google Drive/Sheets if requested
        if sync_to_cloud:
            # Use the sheet_id from config if none is provided
            if not sheet_id and config.SHEET_ID:
                sheet_id = config.SHEET_ID
                
            print(f"\nSyncing data to Google Sheets{' and Drive' if upload_to_drive else ''}...")
            print(f"Using Google Sheet ID: {sheet_id or 'New sheet will be created'}")
            
            if upload_to_drive:
                print(f"Will upload to Drive folder: {config.GDRIVE_FOLDER_ID or 'Root folder'}")
            
            sheet_id, file_id = sync_to_google(sheet_id, upload_to_drive=upload_to_drive)
            
            if upload_to_drive:
                if sheet_id and file_id:
                    print("Data sync completed successfully!")
                    return sheet_id, file_id
                else:
                    print("Data sync failed. Check error messages above.")
            else:
                if sheet_id:
                    print("Google Sheet update completed successfully!")
                    return sheet_id, None
                else:
                    print("Google Sheet update failed. Check error messages above.")
        
        return None, None
        
    except Exception as e:
        print(f"Error in run_all: {str(e)}")
        return None, None


def parse_arguments() -> argparse.Namespace:
    """Parse command line arguments.
    
    Returns:
        argparse.Namespace: Parsed command line arguments
    """
    parser = argparse.ArgumentParser(description="Process and sync operational simulation data")
    parser.add_argument("--sync-to-cloud", action="store_true", 
                       help="Enable syncing to Google Drive/Sheets")
    parser.add_argument("--sheet-id", type=str, 
                       help="Google Sheet ID to update (uses config value if not provided)")
    parser.add_argument("--folder-id", type=str, 
                       help="Google Drive folder ID to upload to (uses config value if not provided)")
    parser.add_argument("--user-email", type=str, 
                       help="Email address to share access with (uses config value if not provided)")
    parser.add_argument("--upload-to-drive", action="store_true", 
                       help="Upload Excel file to Google Drive (default: only update Sheets)")
    return parser.parse_args()


def main() -> None:
    """Main entry point for the script."""
    args = parse_arguments()
    
    # Update config values if provided as arguments
    if args.folder_id:
        config.GDRIVE_FOLDER_ID = args.folder_id
    if args.user_email:
        config.USER_EMAIL = args.user_email
    
    # Run all operations with arguments
    run_all(
        sync_to_cloud=args.sync_to_cloud,
        sheet_id=args.sheet_id,
        upload_to_drive=args.upload_to_drive
    )


if __name__ == "__main__":
    main()
