import fitz
import pytesseract as tess
import os
import time
import gspread
from google.oauth2.service_account import Credentials

# Cấu hình Tesseract
tess.pytesseract.tesseract_cmd = r"D:\Project\Tesseract-OCR\tesseract.exe"


def scan_pdf_for_text(pdf_path, search_texts=["easypost", "USPS MEDIA MAIL"]):
    """
    Quét PDF và tìm kiếm các text cụ thể
    """
    try:
        print(f"Đang xử lý: {os.path.basename(pdf_path)}")

        # Mở PDF
        pdf_document = fitz.open(pdf_path)
        all_text = ""

        # Xử lý từng trang
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]

            # Chuyển trang thành hình ảnh
            pix = page.get_pixmap()

            # Tạo tên file tạm cho hình ảnh
            temp_image_path = f"temp_page_{page_num}.png"
            pix.save(temp_image_path)

            try:
                # OCR nhận dạng text
                text = tess.image_to_string(temp_image_path)
                all_text += text + " "

                # Xóa file tạm
                os.remove(temp_image_path)

            except Exception as e:
                print(f"Lỗi OCR trang {page_num}: {e}")
                # Xóa file tạm nếu có lỗi
                if os.path.exists(temp_image_path):
                    os.remove(temp_image_path)
                continue

        pdf_document.close()

        # Kiểm tra có chứa text nào cần tìm không
        found_texts = []
        for search_text in search_texts:
            if search_text.lower() in all_text.lower():
                found_texts.append(search_text)

        if found_texts:
            return True, all_text, found_texts
        else:
            return False, all_text, []

    except Exception as e:
        print(f"Lỗi xử lý file {pdf_path}: {e}")
        return False, "", []


def setup_google_sheets(json_keyfile_path, sheet_url):
    """
    Kết nối với Google Sheets
    """
    try:
        # Định nghĩa scope
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]

        # Xác thực
        creds = Credentials.from_service_account_file(json_keyfile_path, scopes=scope)
        client = gspread.authorize(creds)

        # Mở sheet
        sheet = client.open_by_url(sheet_url).sheet1  # Hoặc .worksheet('Sheet1')

        return sheet
    except Exception as e:
        print(f"Lỗi kết nối Google Sheets: {e}")
        return None


def update_google_sheets_batch(sheet, results_dict):
    """
    Cập nhật Google Sheets theo batch để tối ưu hiệu suất
    """
    try:
        # Lấy tất cả dữ liệu từ cột B (Order Code)
        order_codes = sheet.col_values(2)  # Cột B = index 2

        # Chuẩn bị dữ liệu batch update
        batch_updates = []
        updated_count = 0

        # Duyệt qua từng order code (bắt đầu từ hàng 2 - bỏ header)
        for row_index, order_code in enumerate(order_codes[1:], start=2):
            if order_code and order_code.strip() in results_dict:
                cell_address = f'L{row_index}'
                batch_updates.append({
                    'range': cell_address,
                    'values': [[results_dict[order_code.strip()]]]
                })
                updated_count += 1
                print(f"✓ Chuẩn bị cập nhật {order_code}: {results_dict[order_code.strip()]}")

        # Thực hiện batch update
        if batch_updates:
            sheet.batch_update(batch_updates)
            print(f"\n✅ Đã cập nhật {updated_count} dòng trong Google Sheets!")
        else:
            print("\n⚠️ Không có dữ liệu nào cần cập nhật")

        return True

    except Exception as e:
        print(f"Lỗi cập nhật Google Sheets: {e}")
        return False


def get_order_code_from_filename(filename):
    """
    Lấy order code từ tên file PDF (bỏ extension .pdf)
    """
    return os.path.splitext(filename)[0]



def scan_folder_for_pdfs(folder_path, search_texts=["easypost", "USPS MEDIA MAIL"], sheet=None):
    """
    Quét tất cả PDF trong folder và tìm các text
    Cập nhật Google Sheets mỗi 10 file tìm thấy
    """
    found_files = []
    results_dict = {}  # Dictionary để lưu kết quả: {order_code: keywords_found}
    batch_results = {}  # Dictionary cho batch hiện tại
    found_count = 0

    # Kiểm tra folder có tồn tại không
    if not os.path.exists(folder_path):
        print(f"Folder không tồn tại: {folder_path}")
        return found_files, results_dict

    # Lấy danh sách tất cả file PDF
    pdf_files = []
    for file in os.listdir(folder_path):
        if file.lower().endswith('.pdf'):
            pdf_files.append(os.path.join(folder_path, file))

    print(f"Tìm thấy {len(pdf_files)} file PDF trong folder")
    print("-" * 50)

    # Xử lý từng file PDF
    for i, pdf_path in enumerate(pdf_files, 1):
        print(f"[{i}/{len(pdf_files)}] ", end="")

        found, text, found_texts = scan_pdf_for_text(pdf_path, search_texts)

        # Lấy order code từ tên file
        order_code = get_order_code_from_filename(os.path.basename(pdf_path))

        if found:
            keywords_str = ", ".join(found_texts)
            found_files.append({
                'file_path': pdf_path,
                'file_name': os.path.basename(pdf_path),
                'order_code': order_code,
                'text_content': text,
                'found_keywords': found_texts
            })
            results_dict[order_code] = keywords_str
            batch_results[order_code] = keywords_str
            found_count += 1
            print(f"✓ FOUND: {keywords_str}")

            # Cập nhật Google Sheets mỗi 10 file tìm thấy
            if found_count % 10 == 0 and sheet:
                print(f"\n🔄 Cập nhật Google Sheets - Batch {found_count // 10}...")
                if update_google_sheets_batch(sheet, batch_results):
                    print(f"✅ Đã cập nhật {len(batch_results)} mục vào Google Sheets")
                    batch_results = {}  # Reset batch
                else:
                    print("❌ Lỗi cập nhật batch")
                print("-" * 50)
        else:
            print("✗ Not found")

        # Thêm delay nhỏ để tránh quá tải
        time.sleep(0.1)

    # Cập nhật batch cuối cùng (nếu còn dư)
    if batch_results and sheet:
        print(f"\n🔄 Cập nhật batch cuối cùng ({len(batch_results)} mục)...")
        if update_google_sheets_batch(sheet, batch_results):
            print(f"✅ Đã cập nhật batch cuối thành công")
        else:
            print("❌ Lỗi cập nhật batch cuối")

    return found_files, results_dict


# MAIN EXECUTION
if __name__ == "__main__":
    folder_path = r"D:\FILE\New folder"

    # GOOGLE SHEETS CONFIG - THAY ĐỔI NHỮNG THÔNG TIN NÀY
    json_keyfile_path = r"D:\Project\decent-trail-451507-d7-d59973874d84.json"  # File JSON credentials
    sheet_url = "https://docs.google.com/spreadsheets/d/1_A6QdeYj_HCT23Y1yFMCHxBE6wMoYth5QL__ZZ_AuYU/edit"  # URL Google Sheet

    search_texts = ["easypost", "USPS MEDIA MAIL"]

    print(f"Bắt đầu quét folder: {folder_path}")
    print(f"Google Sheet URL: {sheet_url}")
    print(f"Tìm kiếm từ khóa: {search_texts}")
    print("=" * 60)

    # Kết nối Google Sheets
    print("Đang kết nối Google Sheets...")
    sheet = setup_google_sheets(json_keyfile_path, sheet_url)

    if not sheet:
        print("❌ Không thể kết nối Google Sheets. Kiểm tra lại cấu hình!")
        exit()

    print("✅ Kết nối Google Sheets thành công!")

    # Bắt đầu quét PDF
    start_time = time.time()
    results, results_dict = scan_folder_for_pdfs(folder_path, search_texts, sheet)
    end_time = time.time()

    # Hiển thị kết quả
    print("\n" + "=" * 60)
    print("KẾT QUẢ TỔNG HỢP:")
    print(f"Thời gian xử lý: {end_time - start_time:.2f} giây")
    print(f"Số file chứa từ khóa: {len(results)}")

    if results:
        print("\nDanh sách file chứa từ khóa:")
        for i, result in enumerate(results, 1):
            print(f"{i}. Order Code: {result['order_code']}")
            print(f"   File: {result['file_name']}")
            print(f"   Từ khóa tìm thấy: {', '.join(result['found_keywords'])}")
            print()

        print("✅ Tất cả kết quả đã được cập nhật vào Google Sheets!")

    else:
        print(f"\nKhông tìm thấy file nào chứa từ khóa: {search_texts}")

    print("\nHoàn thành!")