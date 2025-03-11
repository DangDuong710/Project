import os
import gspread
import logging
import sys
from fastapi import FastAPI, HTTPException, Query
from pydantic import BaseModel
from oauth2client.service_account import ServiceAccountCredentials
from typing import List, Optional
from fastapi.middleware.cors import CORSMiddleware
from datetime import datetime

# Configure logging with UTF-8 encoding
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("api_log.log", encoding='utf-8'),  # Added encoding='utf-8' here
        logging.StreamHandler(stream=sys.stdout)  # Ensure stdout is used instead of default stderr
    ]
)
logger = logging.getLogger("order_api")

app = FastAPI(title="Order Code Extraction API",
              description="API to extract order codes from PDF files and update to Google Sheets")

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Default configuration
DEFAULT_CONFIG = dict(SHEET_KEY="1kNMeY5JrRbvocmYktfWOoWJUamW6W8AQNATDgjgtGW0",
                      SHEET_NAME="Get_Orders_Code",
                      MAIN_FOLDER="D:\\Test_order_code_api",
                      CREDENTIALS_FILE="decent-trail-451507-d7-d59973874d84.json")


# Data models
class OrderCode(BaseModel):
    order_code: str
    seller: str


class ExtractionResult(BaseModel):
    date: str
    total_orders: int
    order_list: List[OrderCode]
    message: str


class UpdateConfig(BaseModel):
    sheet_key: Optional[str] = None
    sheet_name: Optional[str] = None
    folder_path: Optional[str] = None
    credentials_file: Optional[str] = None


# Function to extract order codes
def extract_order_codes(main_folder: str, target_date: str):
    logger.info(f"Starting order code extraction for date {target_date} from folder {main_folder}")
    order_data = set()

    # Check if the main folder exists
    if not os.path.exists(main_folder):
        logger.error(f"Main folder does not exist: {main_folder}")
        raise HTTPException(status_code=400, detail=f"Main folder does not exist: {main_folder}")

    # Iterate through all "Machine X" folders in the main folder
    logger.info(f"Found main folder, starting scan of subfolders")
    for machine_folder in os.listdir(main_folder):
        machine_path = os.path.join(main_folder, machine_folder)

        # Check if it is a folder (skip files)
        if os.path.isdir(machine_path):
            logger.info(f"Scanning machine folder: {machine_folder}")
            # Iterate through month folders (e.g., "2025_2")
            for month_folder in os.listdir(machine_path):
                month_path = os.path.join(machine_path, month_folder)

                # Check if the month folder is a directory
                if os.path.isdir(month_path):
                    logger.debug(f"Scanning month folder: {month_folder}")
                    # Try both formats: with and without leading zeros
                    # Example: try both "2025_3_1" and "2025_03_01"
                    date_formats = [target_date]

                    # Split and prepare alternative format
                    parts = target_date.split("_")
                    if len(parts) == 3:
                        year, month, day = parts
                        # Add format with leading zeros removed
                        alternative_format = f"{year}_{int(month)}_{int(day)}"
                        if alternative_format != target_date:
                            date_formats.append(alternative_format)

                    for date_format in date_formats:
                        date_folder_path = os.path.join(month_path, date_format)

                        # If the target date folder exists
                        if os.path.isdir(date_folder_path):
                            logger.info(f"✅ Found date folder: {date_folder_path}")

                            # Iterate through all subfolders in the date folder
                            for subfolder in os.listdir(date_folder_path):
                                subfolder_path = os.path.join(date_folder_path, subfolder)

                                if os.path.isdir(subfolder_path):
                                    logger.info(f"  🔍 Scanning subfolder: {subfolder_path}")

                                    # Scan all PDF files in the subfolder
                                    for file in os.listdir(subfolder_path):
                                        if file.lower().endswith(".pdf"):
                                            file_name = file.replace(".pdf", "").replace("-", "_")  # Replace '-' with '_'
                                            parts = file_name.split("_")  # Split into components

                                            # Check the number of components in the list
                                            if len(parts) <= 9:
                                                order_code = "UNKNOWN"
                                                seller = "UNKNOWN"
                                            else:
                                                if parts[5] == "1":
                                                    order_code = parts[2]
                                                else:
                                                    order_code = parts[0]

                                                # Get seller name at position [7]
                                                seller = parts[7] if len(parts) > 7 else "UNKNOWN"

                                            # If order_code is valid, add to the set (to remove duplicates)
                                            if order_code != "UNKNOWN":
                                                order_data.add((order_code, seller))
                                                logger.debug(
                                                    f"      📄 {file} → Order code: {order_code}, Seller: {seller}")

    logger.info(f"Extracted {len(order_data)} unique order codes")
    return order_data


# Get the list of all sheet names
def get_sheet_list(sheet_key: str, credentials_file: str):
    try:
        # Check credentials file
        if not os.path.exists(credentials_file):
            logger.error(f"Credentials file does not exist: {credentials_file}")
            raise HTTPException(status_code=400, detail=f"Credentials file does not exist: {credentials_file}")

        # Authenticate Google Sheets API
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
        client = gspread.authorize(creds)

        # Open Google Sheet and get all worksheet names
        spreadsheet = client.open_by_key(sheet_key)
        worksheets = spreadsheet.worksheets()

        sheet_list = [sheet.title for sheet in worksheets]
        logger.info(f"Retrieved {len(sheet_list)} sheets from Google Sheets: {sheet_list}")

        return sheet_list
    except Exception as e:
        logger.error(f"Error retrieving sheet list: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Failed to retrieve sheet list: {str(e)}")


# Function to update Google Sheet
def update_google_sheet(sheet_key: str, sheet_name: str, credentials_file: str, target_date: str,
                        order_data: set):
    logger.info(f"Starting update of Google Sheet '{sheet_name}' with {len(order_data)} order codes")
    try:
        # Check credentials file
        if not os.path.exists(credentials_file):
            logger.error(f"Credentials file does not exist: {credentials_file}")
            raise HTTPException(status_code=400, detail=f"Credentials file does not exist: {credentials_file}")

        # Authenticate Google Sheets API
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
        client = gspread.authorize(creds)

        # Open Google Sheet
        spreadsheet = client.open_by_key(sheet_key)

        # Get list of existing sheets
        sheet_list = [sheet.title for sheet in spreadsheet.worksheets()]
        logger.info(f"Existing sheets: {sheet_list}")

        # Check if sheet_name exists
        if sheet_name not in sheet_list:
            logger.warning(f"Sheet '{sheet_name}' does not exist, creating new one")
            spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=20)

        sheet = spreadsheet.worksheet(sheet_name)
        sheet.clear()

        # Write date to cell A1
        sheet.update("A1", [[target_date]])
        logger.info(f"Updated date {target_date} in cell A1")

        # Convert set to list and write to Google Sheet
        if order_data:
            logger.info(f"Updating {len(order_data)} order codes to sheet")
            sheet.update("A2", list(order_data))
            logger.info("Google Sheet updated successfully")

        return True
    except Exception as e:
        logger.error(f"Error updating Google Sheet: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Failed to update Google Sheet: {str(e)}")


@app.get("/", tags=["Information"])
def get_root():
    logger.info("Request to API homepage")
    return {"message": "Order Code Extraction API", "version": "1.0.0"}


@app.get("/sheets", tags=["Google Sheets"])
def get_all_sheets():
    """Retrieve the list of all sheets in Google Sheets"""
    logger.info("Request to retrieve all sheets")
    try:
        sheets = get_sheet_list(
            DEFAULT_CONFIG["SHEET_KEY"],
            DEFAULT_CONFIG["CREDENTIALS_FILE"]
        )
        logger.info(f"Returned list of {len(sheets)} sheets")
        return {
            "message": "Successfully retrieved sheet list",
            "total_sheets": len(sheets),
            "sheet_list": sheets,
            "current_sheet": DEFAULT_CONFIG["SHEET_NAME"]
        }
    except HTTPException as he:
        logger.error(f"HTTP error retrieving sheet list: {he.detail}")
        raise he
    except Exception as e:
        logger.error(f"Unknown error retrieving sheet list: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error retrieving sheet list: {str(e)}")


@app.get("/extract", response_model=ExtractionResult, tags=["Extraction"])
def extract_orders(
        date: str = Query(..., description="Date to scan in YYYY_M_D format (e.g., 2025_2_15 or 2025_02_15)"),
        update_sheet: bool = Query(True, description="Whether to update Google Sheet or not")):
    logger.info(f"Request to extract order codes for date: {date}, update sheet: {update_sheet}")
    try:
        main_folder = DEFAULT_CONFIG["MAIN_FOLDER"]
        start_time = datetime.now()

        # Extract order codes
        try:
            order_data = extract_order_codes(main_folder, date)
        except Exception as e:
            import traceback
            error_trace = traceback.format_exc()
            logger.error(f"Error extracting order codes: {str(e)}\n{error_trace}")
            raise HTTPException(status_code=500, detail=f"Error extracting order codes: {str(e)}\n{error_trace}")

        # Convert set to list of OrderCode objects
        order_list = [OrderCode(order_code=code, seller=seller) for code, seller in order_data]
        logger.info(f"Extracted {len(order_list)} order codes")

        # Update Google Sheet if requested
        sheet_updated = False
        sheet_error = None
        if update_sheet and order_data:
            try:
                logger.info("Starting Google Sheet update")
                sheet_updated = update_google_sheet(
                    DEFAULT_CONFIG["SHEET_KEY"],
                    DEFAULT_CONFIG["SHEET_NAME"],
                    DEFAULT_CONFIG["CREDENTIALS_FILE"],
                    date,
                    order_data
                )
            except Exception as e:
                sheet_error = str(e)
                logger.error(f"Error updating Google Sheet: {sheet_error}")

        # Calculate processing time
        processing_time = (datetime.now() - start_time).total_seconds()
        logger.info(f"Completed processing in {processing_time:.2f} seconds")

        # Prepare response message
        message = f"Successfully extracted {len(order_list)} order codes in {processing_time:.2f} seconds"
        if update_sheet:
            if sheet_updated:
                message += f" and updated Google Sheet '{DEFAULT_CONFIG['SHEET_NAME']}'"
            elif sheet_error:
                message += f", but failed to update Google Sheet: {sheet_error}"
            else:
                message += ", but did not update Google Sheet"

        result = ExtractionResult(
            date=date,
            total_orders=len(order_list),
            order_list=order_list,
            message=message
        )
        logger.info(f"Returning result: {message}")
        return result

    except HTTPException as he:
        logger.error(f"HTTP error: {he.detail}")
        raise he
    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        logger.error(f"Unknown error: {str(e)}\nDetails: {error_details}")
        raise HTTPException(status_code=500, detail=f"Error: {str(e)}\nDetails: {error_details}")


@app.post("/config", tags=["Configuration"])
def update_configuration(config: UpdateConfig):
    """Update API configuration parameters"""
    logger.info(f"Request to update configuration: {config}")
    try:
        updated = False
        changes = []

        if config.sheet_key:
            old_value = DEFAULT_CONFIG["SHEET_KEY"]
            DEFAULT_CONFIG["SHEET_KEY"] = config.sheet_key
            changes.append(f"Sheet key: {old_value} -> {config.sheet_key}")
            updated = True

        if config.sheet_name:
            old_value = DEFAULT_CONFIG["SHEET_NAME"]
            DEFAULT_CONFIG["SHEET_NAME"] = config.sheet_name
            changes.append(f"Sheet name: {old_value} -> {config.sheet_name}")
            updated = True

        if config.folder_path:
            # Skip path check if 'string' (for testing)
            if config.folder_path == "string" or os.path.isdir(config.folder_path):
                old_value = DEFAULT_CONFIG["MAIN_FOLDER"]
                DEFAULT_CONFIG["MAIN_FOLDER"] = config.folder_path
                changes.append(f"Main folder: {old_value} -> {config.folder_path}")
                updated = True
            else:
                logger.error(f"Invalid folder path: {config.folder_path}")
                raise HTTPException(status_code=400, detail="Invalid folder path")

        if config.credentials_file:
            # Skip file check if 'string' (for testing)
            if config.credentials_file == "string" or os.path.isfile(config.credentials_file):
                old_value = DEFAULT_CONFIG["CREDENTIALS_FILE"]
                DEFAULT_CONFIG["CREDENTIALS_FILE"] = config.credentials_file
                changes.append(f"Credentials file: {old_value} -> {config.credentials_file}")
                updated = True
            else:
                logger.error(f"Credentials file not found: {config.credentials_file}")
                raise HTTPException(status_code=400, detail="Credentials file not found")

        if updated:
            logger.info(f"Configuration updated successfully: {', '.join(changes)}")
            return {"message": "Configuration updated successfully", "config": DEFAULT_CONFIG, "changes": changes}
        else:
            logger.info("No configuration changes were made")
            return {"message": "No configuration changes were made", "config": DEFAULT_CONFIG}

    except HTTPException as he:
        logger.error(f"HTTP error updating configuration: {he.detail}")
        raise he
    except Exception as e:
        logger.error(f"Unknown error updating configuration: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/debug/folder", tags=["Debugging"])
def check_folder(date: str = Query(..., description="Date to check in YYYY_M_D or YYYY_MM_DD format")):
    logger.info(f"Request to check folder for date: {date}")
    main_folder = DEFAULT_CONFIG["MAIN_FOLDER"]
    result = {
        "main_folder_exists": os.path.exists(main_folder),
        "main_folder_path": main_folder,
        "date": date,
        "machine_folder_list": [],
    }

    if result["main_folder_exists"]:
        logger.info(f"Main folder exists: {main_folder}")
        for machine_folder in os.listdir(main_folder):
            machine_path = os.path.join(main_folder, machine_folder)
            if os.path.isdir(machine_path):
                logger.debug(f"Checking machine folder: {machine_folder}")
                machine_info = {
                    "name": machine_folder,
                    "path": machine_path,
                    "month_folder_list": []
                }

                for month_folder in os.listdir(machine_path):
                    month_path = os.path.join(machine_path, month_folder)
                    if os.path.isdir(month_path):
                        logger.debug(f"Checking month folder: {month_folder}")
                        # Try both date formats
                        parts = date.split("_")
                        date_formats = [date]

                        if len(parts) == 3:
                            year, month, day = parts
                            # Add format without leading zeros
                            alternative_format = f"{year}_{int(month)}_{int(day)}"
                            if alternative_format != date:
                                date_formats.append(alternative_format)

                        month_info = {
                            "name": month_folder,
                            "path": month_path,
                            "checked_date_formats": date_formats
                        }

                        for format in date_formats:
                            date_path = os.path.join(month_path, format)
                            exists = os.path.exists(date_path)
                            month_info[f"date_folder_{format}_exists"] = exists
                            month_info[f"date_folder_{format}_path"] = date_path

                            if exists:
                                logger.info(f"Found date folder: {date_path}")
                                month_info["contents"] = os.listdir(date_path)

                        machine_info["month_folder_list"].append(month_info)

                result["machine_folder_list"].append(machine_info)
    else:
        logger.warning(f"Main folder does not exist: {main_folder}")

    logger.info(f"Folder check result: {len(result['machine_folder_list'])} machines found")
    return result


if __name__ == "__main__":
    import uvicorn

    # Configure system to handle Unicode
    if sys.platform.startswith('win'):
        # On Windows, ensure console supports Unicode
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

    logger.info("Application is starting...")
    uvicorn.run(app, host="0.0.0.0", port=8000)