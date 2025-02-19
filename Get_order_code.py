import os
import gspread
from certifi import contents
from oauth2client.service_account import ServiceAccountCredentials

# Thông tin Google Sheet
SHEET_KEY = "1kNMeY5JrRbvocmYktfWOoWJUamW6W8AQNATDgjgtGW0"
SHEET_NAME = "Get_Orders_Code"

# Đường dẫn folder tổng và ngày cần quét
folder_tong = r"D:\New folder"
ngay_can_quet = "2025_2_15"

# Xác thực Google Sheets API
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("luminous-lodge-321503-c17157d58b87.json", scope)
client = gspread.authorize(creds)

# Mở Google Sheet
sheet = client.open_by_key(SHEET_KEY).worksheet(SHEET_NAME)
sheet.clear()
# Ghi ngày vào ô A1
sheet.update("A1", [[ngay_can_quet]])

# Tập hợp để lưu order code duy nhất
order_data = set()

# Duyệt qua tất cả folder "Machine X" trong folder tổng
for machine_folder in os.listdir(folder_tong):
    machine_path = os.path.join(folder_tong, machine_folder)

    # Kiểm tra nếu là folder (bỏ qua file)
    if os.path.isdir(machine_path):
        print(f"📂 Đang kiểm tra folder máy: {machine_path}")

        # Duyệt vào các folder tháng (VD: "2025_2")
        for month_folder in os.listdir(machine_path):
            month_path = os.path.join(machine_path, month_folder)

            # Kiểm tra nếu folder tháng là thư mục
            if os.path.isdir(month_path):
                ngay_folder_path = os.path.join(month_path, ngay_can_quet)

                # Nếu tồn tại folder ngày cần quét
                if os.path.isdir(ngay_folder_path):
                    print(f"✅ Tìm thấy folder ngày: {ngay_folder_path}")

                    # Duyệt qua tất cả folder con trong folder ngày
                    for subfolder in os.listdir(ngay_folder_path):
                        subfolder_path = os.path.join(ngay_folder_path, subfolder)

                        if os.path.isdir(subfolder_path):
                            print(f"  🔍 Đang quét folder con: {subfolder_path}")

                            # Quét tất cả file PDF trong folder con
                            for file in os.listdir(subfolder_path):
                                if file.lower().endswith(".pdf"):
                                    file_name = file.replace(".pdf", "").replace("-", "_")  # Thay '-' thành '_'
                                    parts = file_name.split("_")  # Chia thành các phần tử

                                    # Kiểm tra số phần tử trong danh sách
                                    if len(parts) <= 9:
                                        order_code = "UNKNOWN"
                                        seller = "UNKNOWN"
                                    else:
                                        if parts[5] == "1":
                                            order_code = parts[2]
                                        else:
                                            order_code = parts[0]

                                        # Lấy tên seller ở vị trí [7]
                                        seller = parts[7] if len(parts) > 7 else "UNKNOWN"

                                    # Nếu order_code hợp lệ, thêm vào tập hợp (để loại bỏ trùng lặp)
                                    if order_code != "UNKNOWN":
                                        order_data.add((order_code, seller))
                                        print(f"      📄 {file} → Order Code: {order_code}, Seller: {seller}")


# Chuyển tập hợp thành danh sách và ghi lên Google Sheet
if order_data:
    sheet.update("A2", list(order_data))

print("\n✅ Hoàn thành quét & cập nhật Google Sheet!")
