import openpyxl
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os

# Nhập thông tin PART và ngày DATE
PART = int(input("Nhập PART (1-5): "))
DATE = "1/1/25"
FOLDER_PATH = r"D:\Fix Image"
print(f"PART: {PART} \nDATE: {DATE} \nFolder: {FOLDER_PATH}")


# Hàm kết nối Google Sheet bằng ID
def connect_to_google_sheet_by_id(json_keyfile, sheet_id, sheet_name):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile, scope)
    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(sheet_id)
    return spreadsheet

# Ánh xạ dữ liệu từ file Excel
def extract_data_from_excel(file_path):
    workbook = openpyxl.load_workbook(file_path, data_only=True)
    data = {}
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        if sheet_name.startswith("Machine"):
            printer_name = sheet_name.replace("Machine", "Printer")
        else:
            printer_name = sheet_name
        cell_value = sheet['C2'].value
        data[printer_name] = cell_value
    return data


# Cập nhật dữ liệu lên Google Sheet
def update_google_sheet(sheet, target_date, part, data):
    part_column_index = 2 + part
    all_rows = sheet.get_all_values()
    current_date = None
    processed_printers = set()

    for i, row in enumerate(all_rows[1:], start=2):
        sheet_date = row[0].strip() if row[0] else None
        printer_name = row[1].strip() if len(row) > 1 else None

        if sheet_date:
            current_date = sheet_date
        if not current_date or not printer_name:
            continue
        if current_date == target_date:
            if printer_name in processed_printers:
                print(f"Printer '{printer_name}' đã được cập nhật trước đó. Bỏ qua.")
                continue
            num_file = data.get(printer_name, 0)

            # Ghi dữ liệu vào Google Sheet
            try:
                sheet.update_cell(i, part_column_index, num_file)
                print(
                    f"Cập nhật Printer: '{printer_name}' | Dòng: {i} | Cột: {part_column_index} | Num_file: {num_file}")
                processed_printers.add(printer_name)
            except Exception as e:
                print(f"Lỗi khi cập nhật Printer '{printer_name}': {e}")


# Lấy tất cả các file Excel phù hợp trong thư mục
def get_excel_files_from_folder(folder_path):
    matched_files = []
    file_pattern = re.compile(r"POD_Check_Files_.*\.xlsx$")

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file_pattern.match(file):
                matched_files.append(os.path.join(root, file))
    return matched_files


# Chính xử lý
if __name__ == "__main__":
    json_keyfile = 'luminous-lodge-321503-c17157d58b87.json'
    sheet_id = '1kNMeY5JrRbvocmYktfWOoWJUamW6W8AQNATDgjgtGW0'
    sheet_name = 'January'
    spreadsheet = connect_to_google_sheet_by_id(json_keyfile, sheet_id, sheet_name)
    sheet = spreadsheet.worksheet(sheet_name)  # Truy cập sheet cụ thể
    excel_files = get_excel_files_from_folder(FOLDER_PATH)

    if not excel_files:
        print("Không tìm thấy file Excel nào phù hợp trong thư mục.")
    else:
        print(f"Đã tìm thấy {len(excel_files)} file Excel phù hợp:")
        for file in excel_files:
            print(f"- {file}")

    for file_path in excel_files:
        print(f"\nĐang xử lý file: {file_path}")
        extracted_data = extract_data_from_excel(file_path)
        print("Dữ liệu từ file Excel:")
        for printer, num_file in extracted_data.items():
            print(f"{printer}: {num_file}")
        update_google_sheet(sheet, DATE, PART, extracted_data)
