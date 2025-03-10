from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
# from fastapi.responses import JSONResponse
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import os
import json
import shutil
import datetime
from typing import List
import tempfile
import uvicorn

app = FastAPI(title="Google Sheets Data Uploader API",
              description="API để tải lên và cập nhật dữ liệu từ file Excel/CSV vào Google Sheets")

# Thiết lập kết nối Google Sheets
SCOPE = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
CREDENTIALS_FILE = 'decent-trail-451507-d7-d59973874d84.json'
SPREADSHEET_KEY = '1kNMeY5JrRbvocmYktfWOoWJUamW6W8AQNATDgjgtGW0'
PROCESSED_FOLDER = "uploaded_files"

# Đảm bảo thư mục processed tồn tại
if not os.path.exists(PROCESSED_FOLDER):
    os.makedirs(PROCESSED_FOLDER)


def get_gspread_client():
    """Khởi tạo và trả về client gspread đã được xác thực"""
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, SCOPE)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        raise Exception(f"Không thể kết nối với Google Sheets: {str(e)}")


def convert_to_json_compliant(value):
    """Chuyển đổi giá trị để tương thích với JSON"""
    if pd.isna(value):
        return ""
    if isinstance(value, float) and (value == float('inf') or value == float('-inf')):
        return str(value)
    try:
        json.dumps(value)
        return value
    except (OverflowError, ValueError):
        return str(value) if isinstance(value, float) else value


async def process_file(file_path: str, original_filename: str):
    """Xử lý file và cập nhật lên Google Sheets"""
    try:
        filename = original_filename.lower()
        client = get_gspread_client()
        spreadsheet = client.open_by_key(SPREADSHEET_KEY)

        # Đọc file dựa vào định dạng
        df = None
        if filename.endswith('.xlsx'):
            df = pd.read_excel(file_path)
        elif filename.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            raise HTTPException(status_code=400, detail="Định dạng file không được hỗ trợ. Chỉ chấp nhận .xlsx và .csv")

        # Xác định tên sheet
        sheet_name = None
        if "đơn pending" in filename:
            sheet_name = "Đơn PENDING"
        elif "product (product.template)" in filename:
            sheet_name = "On hand"
        elif "d7 = 7 days ago" in filename:
            sheet_name = "Data 7 days"
        else:
            # Lấy tên file không có phần mở rộng
            sheet_name = os.path.splitext(original_filename)[0]

        # Xử lý sheet "On hand"
        if sheet_name == "On hand":
            try:
                df = df[['Name', 'Quantity On Hand']]
            except KeyError as e:
                raise HTTPException(status_code=400, detail=f"Thiếu cột trong file: {str(e)}")

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

        # Lưu file đã xử lý
        now = datetime.datetime.now()
        timestamp = now.strftime("%Y%m%d%H%M%S")
        name, ext = os.path.splitext(original_filename)
        new_filename = f"{name}_{timestamp}{ext}"

        saved_path = os.path.join(PROCESSED_FOLDER, new_filename)
        shutil.copy(file_path, saved_path)

        return {"status": "success",
                "message": f"File {original_filename} đã được xử lý và cập nhật lên sheet {sheet_name}"}

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi khi xử lý file {original_filename}: {str(e)}")


@app.get("/")
async def root():
    """API root endpoint"""
    return {"message": "Google Sheets Data Uploader API", "version": "1.0"}


@app.post("/upload-single-file/")
async def upload_single_file(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    """
    Upload một file Excel/CSV để cập nhật lên Google Sheets
    """
    if not file.filename.lower().endswith(('.xlsx', '.csv')):
        raise HTTPException(status_code=400, detail="Chỉ chấp nhận file Excel (.xlsx) hoặc CSV (.csv)")

    # Tạo file tạm thời
    temp_file = tempfile.NamedTemporaryFile(delete=False)
    try:
        # Ghi dữ liệu từ file upload vào file tạm thời
        contents = await file.read()
        temp_file.write(contents)
        temp_file.close()

        # Xử lý file trong background
        background_tasks.add_task(process_file, temp_file.name, file.filename)

        return {"status": "processing", "message": f"File {file.filename} đang được xử lý"}

    except Exception as e:
        # Đảm bảo file tạm thời được xóa
        if os.path.exists(temp_file.name):
            os.unlink(temp_file.name)
        raise HTTPException(status_code=500, detail=f"Lỗi khi tải file lên: {str(e)}")


@app.post("/upload-multiple-files/")
async def upload_multiple_files(background_tasks: BackgroundTasks, files: List[UploadFile] = File(...)):
    """
    Upload nhiều file Excel/CSV để cập nhật lên Google Sheets
    """
    if not files:
        raise HTTPException(status_code=400, detail="Không có file nào được tải lên")

    results = []
    temp_files = []

    # Kiểm tra các file và tạo file tạm thời
    for file in files:
        if not file.filename.lower().endswith(('.xlsx', '.csv')):
            raise HTTPException(status_code=400,
                                detail=f"File {file.filename} không được hỗ trợ. Chỉ chấp nhận .xlsx và .csv")

        temp_file = tempfile.NamedTemporaryFile(delete=False)
        contents = await file.read()
        temp_file.write(contents)
        temp_file.close()
        temp_files.append((temp_file.name, file.filename))

    # Xử lý các file trong background
    for temp_path, original_filename in temp_files:
        background_tasks.add_task(process_file, temp_path, original_filename)
        results.append({"filename": original_filename, "status": "processing"})

    return {"files": results, "message": "Các file đang được xử lý"}


@app.get("/health-check/")
async def health_check():
    """Kiểm tra kết nối với Google Sheets"""
    try:
        client = get_gspread_client()
        spreadsheet = client.open_by_key(SPREADSHEET_KEY)
        worksheets = [ws.title for ws in spreadsheet.worksheets()]
        return {
            "status": "healthy",
            "google_sheets_connection": "successful",
            "spreadsheet_id": SPREADSHEET_KEY,
            "worksheets": worksheets
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi kết nối: {str(e)}")


if __name__ == "__main__":
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)