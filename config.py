#!/usr/bin/env python3
"""
Configuration settings for the ops-sim project.
This module contains all configuration settings for data processing and Google Drive/Sheets sync.
"""

# Standard library imports
import os
import json
from pathlib import Path
from typing import Optional, Dict, Any

# Third-party imports
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Constants
# ---------

# Set default data folder path, but allow override via environment variable
DEFAULT_DATA_PATH = Path("./data")
DATA_FOLDER_PATH = Path(os.environ.get("OPS_SIM_DATA_FOLDER", DEFAULT_DATA_PATH))

# Default values for analysis
DEFAULT_PRICE = 180  # Default product price
DEFAULT_ALLOCATION = 50.0  # Default capacity allocation percentage
DEFAULT_INITIAL_BATCH_SIZE = 80  # Default initial batch size
DEFAULT_FINAL_BATCH_SIZE = 50  # Default final batch size

# Google Drive configuration - Use environment variables for sensitive data
GDRIVE_FOLDER_ID = os.environ.get("OPS_SIM_GDRIVE_FOLDER_ID", "")
USER_EMAIL = os.environ.get("OPS_SIM_USER_EMAIL", "")

# Google Sheets configuration
SHEET_NAME = "Master"
SHEET_ID = os.environ.get("OPS_SIM_SHEET_ID", "")

# File paths configuration
MASTER_FILE = DATA_FOLDER_PATH / "Master.xlsx"
CREDENTIALS_FILE = Path("credentials.json")
CLIENT_SECRET_FILE = Path("client_secret.json")  # Optional
TOKEN_PICKLE_FILE = Path("token.pickle")
CONFIG_STORE_FILE = Path("config_store.json")

# Print data path for debugging
print(f"Using data folder path: {DATA_FOLDER_PATH}")

def validate_paths() -> None:
    """Validate that required paths exist and are accessible.
    
    Raises:
        FileNotFoundError: If any required file or directory is missing
        PermissionError: If any required file or directory is not accessible
    """
    # Validate data folder
    if not DATA_FOLDER_PATH.exists():
        try:
            # Try to create the data folder if it doesn't exist
            DATA_FOLDER_PATH.mkdir(parents=True, exist_ok=True)
            print(f"Created data folder: {DATA_FOLDER_PATH}")
        except Exception as e:
            raise FileNotFoundError(f"Could not create data folder: {DATA_FOLDER_PATH}. Error: {str(e)}")
    
    if not os.access(DATA_FOLDER_PATH, os.R_OK | os.W_OK):
        raise PermissionError(f"Data folder not accessible: {DATA_FOLDER_PATH}")
    
    # Validate at least one authentication method exists when attempting to use Google APIs
    # Note: We only check this if we're actually going to use Google APIs
    if (GDRIVE_FOLDER_ID or SHEET_ID) and not (
            CREDENTIALS_FILE.exists() or CLIENT_SECRET_FILE.exists()):
        print(f"Warning: No authentication files found. You will need to create either "
              f"{CREDENTIALS_FILE} (service account) or {CLIENT_SECRET_FILE} (OAuth) "
              f"to use Google API functionality.")
    
    # If credentials.json exists, make sure it's readable
    if CREDENTIALS_FILE.exists() and not os.access(CREDENTIALS_FILE, os.R_OK):
        raise PermissionError(f"File not readable: {CREDENTIALS_FILE}")


def load_config_store() -> Dict[str, Any]:
    """Load persistent configuration from the config store file.
    
    Returns:
        Dict containing the stored configuration values.
        
    Raises:
        FileNotFoundError: If the config store file doesn't exist
        json.JSONDecodeError: If the config store file is not valid JSON
    """
    if not CONFIG_STORE_FILE.exists():
        return {}
    
    try:
        with open(CONFIG_STORE_FILE, 'r') as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        raise json.JSONDecodeError(f"Invalid JSON in config store file: {e}", e.doc, e.pos)


def save_config_store(config_data: Dict[str, Any]) -> None:
    """Save configuration to the config store file.
    
    Args:
        config_data: Dictionary containing configuration values to save.
        
    Raises:
        PermissionError: If the config store file is not writable
    """
    try:
        with open(CONFIG_STORE_FILE, 'w') as f:
            json.dump(config_data, f, indent=4)
    except PermissionError:
        raise PermissionError(f"Config store file not writable: {CONFIG_STORE_FILE}")


def get_sheet_id() -> Optional[str]:
    """Get the stored Google Sheet ID from the config store.
    
    Returns:
        The stored Google Sheet ID, or None if not found.
    """
    config_data = load_config_store()
    return config_data.get('SHEET_ID')


def save_sheet_id(sheet_id: str) -> None:
    """Save the Google Sheet ID to the config store.
    
    Args:
        sheet_id: The Google Sheet ID to save.
    """
    config_data = load_config_store()
    config_data['SHEET_ID'] = sheet_id
    save_config_store(config_data)


# Validate paths on module import
validate_paths()
