import os
import cv2
import numpy as np
import pandas as pd
from collections import defaultdict
from sklearn.cluster import KMeans


def extract_dominant_color(image_path):
    """Trích xuất màu chủ đạo từ ảnh áo"""
    try:
        # Đọc ảnh
        image = cv2.imread(image_path)
        if image is None:
            return None, None

        # Chuyển từ BGR sang RGB
        image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)

        # Reshape ảnh thành mảng 1 chiều các pixel
        pixels = image.reshape(-1, 3)

        # Chỉ lấy các pixel không phải màu trắng (background)
        # Lọc ra các pixel không phải nền trắng (đủ xa từ 255,255,255)
        mask = np.sum(pixels - [255, 255, 255], axis=1) ** 2 > 3000
        foreground_pixels = pixels[mask]

        # Nếu ít pixel hơn, sử dụng toàn bộ ảnh
        if len(foreground_pixels) < 100:
            foreground_pixels = pixels

        # Sử dụng K-means để tìm 3 cụm màu phổ biến nhất
        kmeans = KMeans(n_clusters=3, n_init=10)
        kmeans.fit(foreground_pixels)

        # Lấy các màu trung tâm cùng số pixel trong mỗi cụm
        colors = kmeans.cluster_centers_.astype(int)
        counts = np.bincount(kmeans.labels_)

        # Chọn màu phổ biến nhất (không phải màu trắng hay đen)
        sorted_indices = np.argsort(counts)[::-1]
        for idx in sorted_indices:
            color = colors[idx]
            # Kiểm tra xem màu có phải là trắng hoặc đen không
            brightness = np.sum(color)
            if brightness < 700 and brightness > 60:  # Không quá trắng hoặc quá đen
                dominant_color = color
                break
        else:
            # Nếu không tìm thấy màu phù hợp, lấy màu phổ biến nhất
            dominant_color = colors[sorted_indices[0]]

        # Chuyển đổi từ RGB sang HEX
        hex_color = '#{:02x}{:02x}{:02x}'.format(
            dominant_color[0], dominant_color[1], dominant_color[2])

        return dominant_color, hex_color

    except Exception as e:
        print(f"Lỗi khi xử lý {image_path}: {e}")
        return None, None


def extract_subfolder_path(full_path, base_directory):
    """Trích xuất đường dẫn thư mục con từ đường dẫn đầy đủ
    Ví dụ: D:\FlashPOD_Productdetails\test\PNG\HOODIE\GILDAN\CARDINALRED.png -> HOODIE\GILDAN
    """
    # Xóa phần base_directory từ đường dẫn đầy đủ
    rel_path = os.path.relpath(full_path, base_directory)

    # Tách đường dẫn thành các thành phần
    path_parts = os.path.split(rel_path)

    # Lấy thư mục chứa file
    containing_folder = os.path.dirname(rel_path)

    # Tách thư mục chứa file theo dấu phân cách
    folder_parts = containing_folder.split(os.sep)

    # Lấy hai thư mục cuối cùng (nếu có)
    if len(folder_parts) >= 2:
        return os.path.join(folder_parts[-2], folder_parts[-1])
    elif len(folder_parts) == 1 and folder_parts[0]:
        return folder_parts[0]
    else:
        return ""


def find_colors_from_csv_list(directory, csv_file_path):
    """Tìm tất cả các ảnh có tên màu từ danh sách trong file CSV"""
    # Đọc danh sách màu từ file CSV
    try:
        df = pd.read_csv(csv_file_path)
        if 'color_name' in df.columns:
            color_list = df['color_name'].dropna().unique().tolist()
        else:
            # Nếu không có cột color_name, sử dụng cột đầu tiên
            first_col = df.columns[0]
            color_list = df[first_col].dropna().unique().tolist()
    except Exception as e:
        print(f"Lỗi khi đọc file CSV: {e}")
        return []

    print(f"Đã đọc được {len(color_list)} màu từ file CSV")

    # Tìm thông tin cho các màu
    color_data = []

    # Chuyển danh sách màu thành set để tìm kiếm nhanh hơn
    color_set = set(color.upper() for color in color_list)

    # Quét qua thư mục
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp')):
                # Lấy tên file không bao gồm phần mở rộng
                name_without_ext = os.path.splitext(file)[0].upper()

                # Kiểm tra nếu tên file nằm trong danh sách màu cần tìm
                if name_without_ext in color_set:
                    full_path = os.path.join(root, file)
                    rgb_color, hex_color = extract_dominant_color(full_path)

                    if rgb_color is not None:
                        # Lấy đường dẫn thư mục con
                        subfolder = extract_subfolder_path(full_path, directory)

                        color_data.append({
                            'color_name': name_without_ext,
                            'hex_code': hex_color,
                            'path': subfolder
                        })

    return color_data


def main():
    # Đường dẫn thư mục chứa ảnh
    image_directory = r"D:\FlashPOD_Productdetails\test\PNG"

    # Đường dẫn file CSV chứa danh sách màu cần tìm
    input_csv_file = r"C:\Users\Flashship\Desktop\duplicatecolor.csv"

    # Đường dẫn file kết quả
    output_file = r"D:\FlashPOD_Productdetails\test\PNG\color_list_from_input.csv"

    print("Đang tìm thông tin cho các màu từ danh sách...")
    results = find_colors_from_csv_list(image_directory, input_csv_file)

    if not results:
        print("Không tìm thấy thông tin cho các màu trong danh sách.")
        return

    # Lưu kết quả vào file CSV
    df = pd.DataFrame(results)
    df.to_csv(output_file, index=False)

    print(f"Đã tìm thấy {len(results)} kết quả.")
    print(f"Kết quả đã được lưu vào file: {os.path.abspath(output_file)}")

    # Hiển thị kết quả
    print("\nOut_put_list:")
    print(df)


if __name__ == "__main__":
    main()