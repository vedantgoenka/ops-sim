# Ops Simulation Data Sync

This project provides tools for managing and syncing operational simulation data, including the ability to sync Excel files to Google Drive and Google Sheets.

## Setup

1. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

2. Configure Google API access:
   
   You have two options for authentication:

   a) **Service Account** (recommended for automated scripts):
      - Go to the [Google Cloud Console](https://console.cloud.google.com/)
      - Create a new project or select an existing one
      - Enable the Google Drive API and Google Sheets API
      - Create a Service Account
      - Download the JSON key file and save it as `credentials.json` in the project root
      - For Google Sheets access, share the target spreadsheet with the service account email

   b) **OAuth2** (better for desktop apps with user interaction):
      - Go to the [Google Cloud Console](https://console.cloud.google.com/)
      - Create a new project or select an existing one
      - Enable the Google Drive API and Google Sheets API
      - Configure the OAuth consent screen
      - Create OAuth client ID credentials for a Desktop application
      - Download the JSON file and save it as `client_secret.json` in the project root

3. Configure environment variables for sensitive data:
   
   Copy the `.env.example` file to `.env` and fill in your details:
   ```
   cp .env.example .env
   # Then edit .env with your settings
   ```

   The `.env` file should contain:
   ```
   # Data folder path
   OPS_SIM_DATA_FOLDER=./data

   # Google Drive configuration
   OPS_SIM_GDRIVE_FOLDER_ID=your_folder_id_here
   OPS_SIM_USER_EMAIL=your.email@example.com

   # Google Sheets configuration
   OPS_SIM_SHEET_ID=your_sheet_id_here
   ```

   Make sure to keep your `.env` file private and never commit it to a public repository.

## Usage

Run the main script to process data and optionally sync to Google Drive/Sheets:

```python
# Without syncing to Google
python main.py

# With syncing to Google (using folder ID and user email from config.py)
python main.py --sync-to-cloud

# With specific Google Sheet ID (overrides the one in config.py)
python main.py --sync-to-cloud --sheet-id YOUR_SHEET_ID

# Override the folder ID from command line
python main.py --sync-to-cloud --folder-id YOUR_FOLDER_ID

# Share with a different email (overrides config.py)
python main.py --sync-to-cloud --user-email different.email@example.com

# Only update Google Sheets without uploading to Drive
python main.py --sync-to-cloud --no-drive-upload
```

All files will be stored in the Google Drive folder specified in your environment variables or by command line arguments and shared with the specified user email.

When you run the script with `--sync-to-cloud` for the first time, it will:
1. Create a new Google Sheet named "Master" (or whatever name is set in `config.py`)
2. Save the Sheet ID to your config store
3. Reuse the same Google Sheet on subsequent runs

If you want to create a new sheet instead of updating the existing one, you can run with a specific `--sheet-id` parameter.

Or import the functions in your own code:

```python
from main import run_all
import config

# Run with syncing enabled and specific sheet ID
sheet_id, file_id = run_all(
    sync_to_cloud=True,
    sheet_id="your_sheet_id",  # Optional, will create new if None
    upload_to_drive=True  # Set to False to skip Drive upload
)
```

## Security

This project requires access to Google APIs through service account credentials or OAuth. To protect your sensitive data:

1. Never commit `.env`, `credentials.json`, `client_secret.json`, or `token.pickle` to public repositories
2. Use environment variables instead of hardcoding sensitive values
3. The included `.gitignore` file is set up to exclude sensitive files from git

## Testing

To test the Google Drive and Sheets integration, run:

```python
# Run all tests
python test_gdrive_sync.py

# Override config values from command line
python test_gdrive_sync.py --folder-id YOUR_FOLDER_ID
python test_gdrive_sync.py --user-email different.email@example.com

# Skip uploading to Drive during sync test
python test_gdrive_sync.py --no-drive-upload

# Test only specific components
python test_gdrive_sync.py --test-sheets-only
python test_gdrive_sync.py --test-drive-only
python test_gdrive_sync.py --test-sync-only
```

## Functions

The project includes the following key functions:

- `append_to_master()`: Consolidates latest Excel data into a master file
- `DataAnalyzer.add_current_price()`: Adds current price data based on history
- `DataAnalyzer.add_capacity_allocation()`: Adds capacity allocation data based on history
- `sync_to_google()`: Syncs the master file to Google Drive and Google Sheets
- `update_google_sheet()`: Updates or creates a Google Sheet from an Excel file
- `upload_to_gdrive()`: Uploads a file to Google Drive

## Finding IDs

- **Google Sheet ID**: In the URL of your Google Sheet: `https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit`
- **Google Drive Folder ID**: In the URL when you open a folder: `https://drive.google.com/drive/folders/{FOLDER_ID}` 

## Configuration

The `config.py` file allows you to configure various aspects of the application. Most sensitive settings are now handled through environment variables for security. The main settings include:

```python
# Data folder path - override with OPS_SIM_DATA_FOLDER environment variable
DATA_FOLDER_PATH = Path(os.environ.get("OPS_SIM_DATA_FOLDER", "./data"))

# Google Drive configuration
GDRIVE_FOLDER_ID = os.environ.get("OPS_SIM_GDRIVE_FOLDER_ID", "")
USER_EMAIL = os.environ.get("OPS_SIM_USER_EMAIL", "")

# Google Sheets configuration
SHEET_NAME = "Master"
SHEET_ID = os.environ.get("OPS_SIM_SHEET_ID", "")
```