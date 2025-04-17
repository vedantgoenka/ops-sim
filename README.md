# Ops Simulation Data Sync

[![MIT License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

A Python tool for analyzing and syncing operational simulation data with Google Sheets.

## Overview

This project helps manage and analyze operational simulation data stored in Excel files. It provides functionality to:

- Consolidate data from multiple Excel files
- Analyze operational data (pricing, capacity allocation, batch sizes)
- Sync data to Google Sheets for easy sharing and collaboration
- Upload files to Google Drive

## Features

- **Data Consolidation**: Append data from multiple Excel files into a master file
- **Data Analysis**: Add derived data fields like current price and capacity allocation
- **Google Integration**: Sync data to Google Sheets and Drive
- **Environment-based Configuration**: Keep sensitive data secure using environment variables
- **Command-line Interface**: Run operations with command-line arguments

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/yourusername/ops-sim.git
   cd ops-sim
   ```

2. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

3. Set up your environment:
   ```
   cp .env.example .env
   # Edit .env with your settings
   ```

4. Set up Google API credentials:
   - Get a service account or OAuth credentials from Google Cloud Console
   - Save as `credentials.json` or `client_secret.json`
   - See [Setup](#setup) section for detailed instructions

## Setup

### Google API Setup

1. Create a project in [Google Cloud Console](https://console.cloud.google.com/)
2. Enable Google Drive API and Google Sheets API
3. Choose authentication method:
   
   **Option 1: Service Account (recommended for automated scripts)**
   - Create a service account in GCP Console
   - Download JSON key and save as `credentials.json`
   - Share Google Sheets with service account email
   
   **Option 2: OAuth (better for desktop apps)**
   - Configure OAuth consent screen
   - Create OAuth client ID for Desktop app
   - Download JSON and save as `client_secret.json`

### Environment Variables

Create a `.env` file with the following variables:

```
# Data folder path
OPS_SIM_DATA_FOLDER=./data

# Google Drive configuration
OPS_SIM_GDRIVE_FOLDER_ID=your_folder_id_here
OPS_SIM_USER_EMAIL=your.email@example.com

# Google Sheets configuration
OPS_SIM_SHEET_ID=your_sheet_id_here
```

## Usage

### Command Line

Run the main script with various options:

```bash
# Basic usage without Google sync
python main.py

# Sync to Google Sheets/Drive
python main.py --sync-to-cloud

# Specify Sheet ID
python main.py --sync-to-cloud --sheet-id YOUR_SHEET_ID

# Specify Google Drive folder
python main.py --sync-to-cloud --folder-id YOUR_FOLDER_ID

# Share with specific email
python main.py --sync-to-cloud --user-email user@example.com

# Upload to Google Drive
python main.py --sync-to-cloud --upload-to-drive
```

### In Your Code

```python
from main import run_all

# Run all operations with Google sync
sheet_id, file_id = run_all(
    sync_to_cloud=True,
    sheet_id="optional_sheet_id",
    upload_to_drive=True
)
```

## Project Structure

- `main.py` - Main entry point and command line interface
- `append.py` - Functions for appending data to the master file
- `analysis.py` - Data analysis functions
- `gdrive_sync.py` - Google Drive and Sheets integration
- `config.py` - Configuration settings
- `.env` - Environment variables (not included, create from .env.example)
- `credentials.json` - Google API credentials (not included)

## Security

- Never commit `.env`, `credentials.json`, or `token.pickle` to public repositories
- Use environment variables for sensitive configuration
- The included `.gitignore` is set up to exclude sensitive files

## Testing

```bash
# Run all tests
python test_gdrive_sync.py

# Test specific components
python test_gdrive_sync.py --test-sheets-only
python test_gdrive_sync.py --test-drive-only
```

## Finding Google IDs

- **Sheet ID**: From a Google Sheet URL: `https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit`
- **Folder ID**: From a Google Drive folder URL: `https://drive.google.com/drive/folders/{FOLDER_ID}`

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. 