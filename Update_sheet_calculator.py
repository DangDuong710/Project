import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import os
import json
import shutil
from googleapiclient.discovery import build
import datetime

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('luminous-lodge-321503-c17157d58b87.json', scope)
client = gspread.authorize(creds)

folder_path = r'D:\Fix Image'
spreadsheet_key = '1kNMeY5JrRbvocmYktfWOoWJUamW6W8AQNATDgjgtGW0'
spreadsheet = client.open_by_key(spreadsheet_key)

#-------------------------------Update_sheet_calculator----------------------------------#
def convert_to_json_compliant(value):
    if pd.isna(value):
        return ""
    if isinstance(value, float) and (value == float('inf') or value == float('-inf')):
        return str(value)
    try:
        json.dumps(value)
        return value
    except (OverflowError, ValueError):
        return str(value) if isinstance(value, float) else value

processed_folder = os.path.join(folder_path, "Updated")
if not os.path.exists(processed_folder):
    os.makedirs(processed_folder)

for filename in os.listdir(folder_path):
    if filename.endswith(('.xlsx', '.csv')):
        # Bỏ qua các tệp có tên bắt đầu bằng "POD_Check_Files_"
        if filename.startswith("POD_Check_Files_"):
            print(f"Bỏ qua tệp {filename} ")
            continue

        filepath = os.path.join(folder_path, filename)
        sheet_name = None

        try:
            df = None
            try:
                if filename.endswith('.xlsx'):
                    df = pd.read_excel(filepath)
                else:
                    df = pd.read_csv(filepath)
            except Exception as e:
                print(f"Lỗi khi đọc tệp {filename}: {e}")
                continue

            if df is not None:
                # Sheet name rules
                if "đơn pending" in filename.lower():
                    sheet_name = "Đơn PENDING"
                elif "product (product.template)" in filename.lower():
                    sheet_name = "On hand"
                elif "d7 = 7 days ago" in filename.lower():
                    sheet_name = "Data 7 days"
                else:
                    sheet_name = filename[:-5]
                if sheet_name == "On hand":
                    try:
                        df = df[['Name', 'Quantity On Hand']]  # Get data name and on hand only
                    except KeyError as e:  # Find column fail
                        print(f"Lỗi: Thiếu cột trong tệp {filename}: {e}")
                        continue  # skip file
                df = df.fillna('')
                df = df.apply(lambda series: series.map(convert_to_json_compliant))

                try:
                    worksheet = spreadsheet.worksheet(sheet_name)
                except gspread.exceptions.WorksheetNotFound:
                    worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=df.shape[0] + 100,
                                                          cols=df.shape[1] + 10)

                worksheet.update([df.columns.values.tolist()] + df.values.tolist(),
                                 value_input_option=gspread.utils.ValueInputOption.user_entered)
                new_filename = filename
                if os.path.exists(os.path.join(processed_folder, filename)):
                    now = datetime.datetime.now()
                    timestamp = now.strftime("%Y%m%d%H%M%S")
                    name, ext = os.path.splitext(filename)
                    new_filename = f"{name}_{timestamp}{ext}"
                shutil.move(filepath, os.path.join(processed_folder, new_filename))
                print(f"Đã di chuyển tệp {filename} vào thư mục Updated.")
        except Exception as e:
            if sheet_name is not None:
                print(f"Lỗi khi xử lý sheet {sheet_name}: {e}")
            else:
                print(f"Lỗi khi xử lý tệp {filename}: {e}")

print("Get data successfully!")

