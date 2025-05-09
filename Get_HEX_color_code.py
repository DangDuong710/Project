import os
import cv2
import numpy as np
from PIL import Image
from sklearn.cluster import KMeans
import glob
import csv

# CÁC ĐƯỜNG DẪN CỐ ĐỊNH
DUONG_DAN_THU_MUC_ANH = r"D:\FlashPOD_Productdetails\PNG file\SHIRT\GILDAN"  # Thay đổi đường dẫn này theo máy tính của bạn
DUONG_DAN_FILE_CSV = r"D:\FlashPOD_Productdetails\PNG file\SHIRT\GILDAN\Hex_code.csv"  # Thay đổi đường dẫn này theo nhu cầu của bạn


def lay_mau_chinh(duong_dan_anh):
    """Trích xuất màu chính từ một ảnh, tập trung vào vùng trung tâm."""
    # Tải ảnh bằng PIL trước (xử lý tốt hơn với nhiều định dạng)
    pil_img = Image.open(duong_dan_anh)
    # Chuyển sang định dạng OpenCV
    img = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)

    # Lấy kích thước ảnh
    height, width = img.shape[:2]

    # Tính toán vùng trung tâm (crop khoảng 60% trung tâm của ảnh)
    center_x, center_y = width // 2, height // 2
    crop_width = int(width * 0.6)
    crop_height = int(height * 0.6)

    # Xác định tọa độ cắt
    x1 = max(0, center_x - crop_width // 2)
    y1 = max(0, center_y - crop_height // 2)
    x2 = min(width, center_x + crop_width // 2)
    y2 = min(height, center_y + crop_height // 2)

    # Cắt vùng trung tâm của ảnh
    center_img = img[y1:y2, x1:x2]

    # Thay đổi kích thước ảnh để tăng tốc xử lý
    resized_img = cv2.resize(center_img, (100, 100))

    # Lọc bỏ pixel quá sáng (có thể là nền trắng)
    # Chuyển sang HSV để dễ lọc màu
    hsv_img = cv2.cvtColor(resized_img, cv2.COLOR_BGR2HSV)

    # Lọc ra những pixel không phải màu trắng (dựa vào giá trị S và V)
    # Pixel có S thấp và V cao thường là màu trắng/xám
    mask = (hsv_img[:, :, 1] > 20) | (hsv_img[:, :, 2] < 220)

    # Nếu sau khi lọc, còn quá ít pixel, dùng ảnh gốc
    if np.sum(mask) < 1000:  # Nếu ít hơn 1000 pixel không phải màu trắng
        filtered_pixels = resized_img.reshape(-1, 3)
    else:
        # Chỉ lấy các pixel không phải màu trắng để phân tích
        filtered_pixels = resized_img.reshape(-1, 3)[mask.flatten()]


    if len(filtered_pixels) < 100:
        filtered_pixels = resized_img.reshape(-1, 3)
    kmeans = KMeans(n_clusters=5)
    kmeans.fit(filtered_pixels)

    # Lấy ra trung tâm các cụm
    centers = kmeans.cluster_centers_

    # Đếm số lượng pixel trong mỗi cụm
    counts = np.bincount(kmeans.labels_)

    # Sắp xếp các cụm theo số lượng pixel giảm dần
    sorted_indices = np.argsort(counts)[::-1]

    # Lấy màu của cụm lớn nhất
    dominant_color = centers[sorted_indices[0]]

    # Chuyển đổi BGR sang RGB
    dominant_color = dominant_color[::-1]

    return dominant_color


def rgb_to_hex(rgb):
    """Chuyển đổi một tuple RGB thành mã màu hex."""
    return '#{:02x}{:02x}{:02x}'.format(int(rgb[0]), int(rgb[1]), int(rgb[2]))


def trich_xuat_mau_tu_thu_muc(duong_dan_thu_muc):
    """Xử lý tất cả các ảnh trong một thư mục và trích xuất thông tin màu."""
    du_lieu_mau = {}

    # Lấy tất cả các tệp ảnh
    phan_mo_rong_anh = ['*.jpg', '*.jpeg', '*.png', '*.gif', '*.bmp']
    cac_tep_anh = []
    for ext in phan_mo_rong_anh:
        cac_tep_anh.extend(glob.glob(os.path.join(duong_dan_thu_muc, ext)))

    # Xử lý từng ảnh
    for duong_dan_anh in cac_tep_anh:
        # Trích xuất tên màu từ tên tệp
        ten_mau = os.path.splitext(os.path.basename(duong_dan_anh))[0]

        try:
            # Lấy màu chính
            mau_chinh = lay_mau_chinh(duong_dan_anh)

            # Chuyển đổi sang hex
            ma_mau_hex = rgb_to_hex(mau_chinh)

            # Lưu trong từ điển
            du_lieu_mau[ten_mau] = ma_mau_hex

            print(f"Đã xử lý: {ten_mau} - {ma_mau_hex}")

        except Exception as e:
            print(f"Lỗi khi xử lý {duong_dan_anh}: {e}")

    return du_lieu_mau


def luu_ket_qua_csv(du_lieu_mau, ten_file_xuat):
    """Lưu dữ liệu màu vào một tệp CSV."""
    with open(ten_file_xuat, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        # Viết tiêu đề
        writer.writerow(['Color', 'Color Code'])
        # Viết dữ liệu
        for ten_mau, ma_hex in du_lieu_mau.items():
            writer.writerow([ten_mau, ma_hex])

    print(f"Kết quả đã được lưu vào {ten_file_xuat}")


def xu_ly_mot_anh(duong_dan_anh):
    """Xử lý một ảnh để kiểm tra."""
    try:
        ten_mau = os.path.splitext(os.path.basename(duong_dan_anh))[0]
        mau_chinh = lay_mau_chinh(duong_dan_anh)
        ma_mau_hex = rgb_to_hex(mau_chinh)
        print(f"Tên màu: {ten_mau}")
        print(f"Mã màu HEX: {ma_mau_hex}")
    except Exception as e:
        print(f"Lỗi: {e}")


def main():
    print(f"Đang xử lý ảnh từ thư mục: {DUONG_DAN_THU_MUC_ANH}")
    print(f"Kết quả sẽ được lưu vào: {DUONG_DAN_FILE_CSV}")

    # Kiểm tra xem thư mục đầu vào có tồn tại không
    if not os.path.exists(DUONG_DAN_THU_MUC_ANH):
        print(f"Lỗi: Thư mục {DUONG_DAN_THU_MUC_ANH} không tồn tại.")
        return

    # Kiểm tra xem thư mục chứa file đầu ra có tồn tại không
    output_dir = os.path.dirname(DUONG_DAN_FILE_CSV)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Đã tạo thư mục {output_dir}")

    # Trích xuất màu
    du_lieu_mau = trich_xuat_mau_tu_thu_muc(DUONG_DAN_THU_MUC_ANH)

    # Lưu kết quả vào CSV
    luu_ket_qua_csv(du_lieu_mau, DUONG_DAN_FILE_CSV)


if __name__ == "__main__":
    # Bạn có thể bỏ comment dòng dưới đây và thêm đường dẫn để kiểm tra một ảnh cụ thể
    # xu_ly_mot_anh("đường_dẫn_đến_ảnh_cần_kiểm_tra.jpg")
    main()