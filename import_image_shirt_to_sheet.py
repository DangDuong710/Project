import os
import pandas as pd
import json
import time
import io
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from googleapiclient.http import MediaFileUpload


class GoogleSheetImageImporter:
    def __init__(self):
        # Hardcoded values
        self.sheet_id = "1J8kR0qsVsU_v7yiJl-dd0Ti5nuLxczskFlBx0x3leC0"  # Replace with your actual Google Sheet ID
        self.sheet_name = "special_color_list"  # Replace with your actual sheet name
        self.source_folder = r"D:\FlashPOD_Productdetails\test\PNG"  # Replace with your actual image folder path
        self.credentials_path = r"D:\Project\decent-trail-451507-d7-d59973874d84.json"  # Replace with path to your service account JSON
        self.drive_folder_id = "1vYv5kTTKHURTwsaLEhvxWifwIgLnHZff"  # Replace with your Google Drive folder ID to store images

        # Initialize variables
        self.df = None
        self.sheets_service = None
        self.drive_service = None

    def log(self, message):
        """Print log messages with timestamp"""
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{timestamp}] {message}")

    def connect_to_services(self):
        """Connect to Google Sheet and Drive using service account credentials"""
        try:
            self.log("Đang kết nối với Google Services...")

            # Authenticate with Google APIs
            SCOPES = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]

            # Load service account credentials from JSON file
            creds = Credentials.from_service_account_file(
                self.credentials_path, scopes=SCOPES)

            # Create services to interact with Google APIs
            self.sheets_service = build('sheets', 'v4', credentials=creds)
            self.drive_service = build('drive', 'v3', credentials=creds)

            # Get data from Sheet
            sheet = self.sheets_service.spreadsheets()
            result = sheet.values().get(
                spreadsheetId=self.sheet_id,
                range=f"{self.sheet_name}!A:H"
            ).execute()
            values = result.get('values', [])

            if not values:
                self.log("Không tìm thấy dữ liệu trong sheet.")
                return False

            # Get header row
            header = values[0]

            # Process data rows - ensure all rows have the same length as header
            processed_rows = []
            for row in values[1:]:
                # If row has fewer columns than header, pad with empty strings
                if len(row) < len(header):
                    row_padded = row + [''] * (len(header) - len(row))
                    processed_rows.append(row_padded)
                else:
                    # If row has more columns than header, truncate
                    processed_rows.append(row[:len(header)])

            # Create DataFrame from processed data
            self.df = pd.DataFrame(processed_rows, columns=header)
            self.log(f"Đã tải dữ liệu với {len(self.df)} dòng")
            return True

        except Exception as e:
            error_msg = f"Lỗi kết nối: {str(e)}"
            self.log(error_msg)
            return False

    def upload_image_to_drive(self, image_path, file_name):
        """Upload image to Google Drive and return public URL"""
        try:
            # File metadata
            file_metadata = {
                'name': file_name,
                'parents': [self.drive_folder_id]
            }

            # Upload file
            media = MediaFileUpload(image_path,
                                    mimetype='image/png',
                                    resumable=True)

            file = self.drive_service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id,webViewLink'
            ).execute()

            # Make the file publicly accessible
            self.drive_service.permissions().create(
                fileId=file.get('id'),
                body={'type': 'anyone', 'role': 'reader'},
                fields='id'
            ).execute()

            # Get direct link for image
            file_id = file.get('id')
            direct_link = f"https://drive.google.com/uc?export=view&id={file_id}"

            self.log(f"Đã tải ảnh lên Google Drive: {file_name}")
            return direct_link

        except Exception as e:
            self.log(f"Lỗi khi tải ảnh lên Drive: {str(e)}")
            return None

    def find_and_import_images(self):
        """Find and import images to Google Sheet"""
        if not os.path.exists(self.source_folder):
            self.log(f"Thư mục nguồn không tồn tại: {self.source_folder}")
            return False

        try:
            self.log("Bắt đầu tìm và import ảnh...")

            # Check DataFrame structure
            required_columns = ['color_name', 'path']
            missing_columns = [col for col in required_columns if col not in self.df.columns]
            if missing_columns:
                error_msg = f"Thiếu các cột: {', '.join(missing_columns)}"
                self.log(error_msg)
                return False

            # Add 'note' column if not exists
            if 'note' not in self.df.columns:
                self.df['note'] = ''

            # Create mapping between DataFrame index and row number in Google Sheet
            # Add 2 because row 1 is the header and indexing starts from 0
            row_mapping = {i: i + 2 for i in range(len(self.df))}

            # Process each row
            source_base = self.source_folder

            success_count = 0
            error_count = 0
            updated_cells = []
            batch_size = 10  # Update sheet every 10 images

            for idx, row in self.df.iterrows():
                try:
                    color_name = row['color_name']
                    path_info = row['path']

                    # Get information from path
                    parts = path_info.split('\\')
                    if len(parts) >= 2:
                        product_type = parts[-2]  # SHIRT, HOODIE, SWEATSHIRT
                        style_code = parts[-1]  # G5000B, G5000, etc.

                        # Find images in source folder
                        search_dir = os.path.join(source_base, product_type, style_code)

                        # Check if directory exists
                        if not os.path.exists(search_dir):
                            alternative_search_dir = os.path.join(source_base, "PNG", product_type, style_code)
                            if os.path.exists(alternative_search_dir):
                                search_dir = alternative_search_dir
                            else:
                                self.log(f"Thư mục không tồn tại: {search_dir}")
                                error_count += 1
                                continue

                        # Find all files containing the color name
                        found_files = []
                        for file in os.listdir(search_dir):
                            if color_name.lower() in file.lower() and (
                                    file.endswith('.png') or file.endswith('.jpg') or file.endswith('.jpeg')):
                                found_files.append(file)

                        if not found_files:
                            self.log(f"Không tìm thấy ảnh cho màu {color_name} trong {search_dir}")
                            error_count += 1
                            continue

                        # Get the first file found
                        image_file = found_files[0]
                        image_path = os.path.join(search_dir, image_file)

                        # Upload image to Google Drive
                        image_url = self.upload_image_to_drive(image_path, image_file)

                        if not image_url:
                            self.log(f"Không thể tải ảnh lên cho màu {color_name}")
                            error_count += 1
                            continue

                        # Create IMAGE formula for Google Sheet
                        image_formula = f'=IMAGE("{image_url}"; 1)'  # Mode 1: Image scaled to fit cell

                        # Update image path to note column in DataFrame
                        self.df.at[idx, 'note'] = image_url

                        # Add to update list
                        sheet_row = row_mapping[idx]
                        updated_cells.append({
                            'range': f"{self.sheet_name}!H{sheet_row}",
                            'values': [[image_formula]]
                        })

                        self.log(f"Đã tìm và tải lên ảnh cho màu {color_name}: {image_file}")
                        success_count += 1

                        # Update every batch_size images
                        if len(updated_cells) >= batch_size:
                            self.update_sheet(updated_cells)
                            updated_cells = []
                    else:
                        self.log(f"Định dạng đường dẫn không đúng cho dòng {idx + 1}: {path_info}")
                        error_count += 1
                except Exception as e:
                    self.log(f"Lỗi xử lý dòng {idx + 1}: {str(e)}")
                    error_count += 1

            # Update remaining cells
            if updated_cells:
                self.update_sheet(updated_cells)

            # Summarize results
            self.log(f"\nKết quả:\n- Thành công: {success_count}\n- Lỗi: {error_count}\n- Tổng cộng: {len(self.df)}")

            return success_count > 0

        except Exception as e:
            error_msg = f"Lỗi: {str(e)}"
            self.log(error_msg)
            return False

    def update_sheet(self, cells):
        """Update multiple cells in Google Sheet"""
        try:
            batch_update_values_request_body = {
                'value_input_option': 'USER_ENTERED',
                'data': cells
            }

            self.sheets_service.spreadsheets().values().batchUpdate(
                spreadsheetId=self.sheet_id,
                body=batch_update_values_request_body
            ).execute()

            self.log(f"Đã cập nhật {len(cells)} ô trong Google Sheet")
            return True
        except Exception as e:
            self.log(f"Lỗi cập nhật Google Sheet: {str(e)}")
            return False


def main():
    """Main function to run the tool"""
    print("=== Google Sheet Image Importer Tool (với tải ảnh lên Google Drive) ===")

    # Initialize the tool
    importer = GoogleSheetImageImporter()

    # Connect to Google services
    print("\nĐang kết nối đến Google Services...")
    if not importer.connect_to_services():
        print("Không thể kết nối đến Google Services. Chương trình kết thúc.")
        return

    # Find and import images
    print("\nBắt đầu quá trình tìm và import ảnh...")
    if importer.find_and_import_images():
        print("\nQuá trình import ảnh hoàn tất!")
    else:
        print("\nQuá trình import ảnh thất bại hoặc có lỗi.")


if __name__ == "__main__":
    main()