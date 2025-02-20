import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import os
import json
import shutil
import datetime
import tkinter as tk
from tkinter import filedialog

# Thiết lập kết nối Google Sheets
import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('decent-trail-451507-d7-d59973874d84.json', scope)
client = gspread.authorize(creds)

print("✅ Kết nối Google Sheets thành công!")


# Hiển thị hộp thoại để chọn nhiều file
root = tk.Tk()
root.withdraw()  # Ẩn cửa sổ chính

file_paths = filedialog.askopenfilenames(
    title="Chọn các file Excel hoặc CSV",
    filetypes=[("Excel & CSV files", "*.xlsx;*.csv"), ("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
)

if not file_paths:
    print("Không có file nào được chọn. Thoát chương trình.")
    exit()

# Thư mục chứa file đã xử lý
processed_folder = os.path.join(os.path.dirname(file_paths[0]), "Updated")
if not os.path.exists(processed_folder):
    os.makedirs(processed_folder)


# Hàm chuyển đổi dữ liệu để tránh lỗi JSON
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


# Lặp qua từng file đã chọn
spreadsheet_key = '1kNMeY5JrRbvocmYktfWOoWJUamW6W8AQNATDgjgtGW0'
spreadsheet = client.open_by_key(spreadsheet_key)

for file_path in file_paths:
    filename = os.path.basename(file_path)
    sheet_name = None

    try:
        df = None
        try:
            if filename.endswith('.xlsx'):
                df = pd.read_excel(file_path)
            else:
                df = pd.read_csv(file_path)
        except Exception as e:
            print(f"Lỗi khi đọc tệp {filename}: {e}")
            continue

        if df is not None:
            # Xác định tên sheet
            if "đơn pending" in filename.lower():
                sheet_name = "Đơn PENDING"
            elif "product (product.template)" in filename.lower():
                sheet_name = "On hand"
            elif "d7 = 7 days ago" in filename.lower():
                sheet_name = "Data 7 days"
            else:
                sheet_name = filename[:-5]

            # Xử lý sheet "On hand"
            if sheet_name == "On hand":
                try:
                    df = df[['Name', 'Quantity On Hand']]
                except KeyError as e:
                    print(f"Lỗi: Thiếu cột trong tệp {filename}: {e}")
                    continue

            # Làm sạch dữ liệu
            df = df.fillna('')
            df = df.apply(lambda series: series.map(convert_to_json_compliant))

            # Cập nhật lên Google Sheet
            try:
                worksheet = spreadsheet.worksheet(sheet_name)
            except gspread.exceptions.WorksheetNotFound:
                worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=df.shape[0] + 100, cols=df.shape[1] + 10)

            worksheet.update([df.columns.values.tolist()] + df.values.tolist(),
                             value_input_option=gspread.utils.ValueInputOption.user_entered)

            # Di chuyển file đã xử lý
            new_filename = filename
            if os.path.exists(os.path.join(processed_folder, filename)):
                now = datetime.datetime.now()
                timestamp = now.strftime("%Y%m%d%H%M%S")
                name, ext = os.path.splitext(filename)
                new_filename = f"{name}_{timestamp}{ext}"

            shutil.move(file_path, os.path.join(processed_folder, new_filename))
            print(f"Đã di chuyển tệp {filename} vào thư mục Updated.")

    except Exception as e:
        print(f"Lỗi khi xử lý tệp {filename}: {e}")

print("Get data successfully!")
