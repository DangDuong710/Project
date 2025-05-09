import os
import cv2
import numpy as np
import pandas as pd
from collections import defaultdict
from sklearn.cluster import KMeans
import matplotlib.colors as mcolors


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


def compare_colors(color1, color2, threshold=30):
    """So sánh hai màu RGB, trả về True nếu tương tự nhau"""
    if color1 is None or color2 is None:
        return False

    # Tính khoảng cách Euclidean giữa hai màu
    distance = np.sqrt(np.sum((color1 - color2) ** 2))
    return distance < threshold


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


def find_duplicates_and_compare_colors(directory):
    """Tìm các file trùng tên và so sánh màu sắc"""
    # Nhóm các file theo tên (không bao gồm phần mở rộng)
    name_groups = defaultdict(list)

    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp')):
                # Lấy tên file không bao gồm phần mở rộng
                name_without_ext = os.path.splitext(file)[0]
                full_path = os.path.join(root, file)
                name_groups[name_without_ext].append(full_path)

    # Kết quả để xuất ra CSV
    results = []

    # Xử lý các nhóm có từ 2 file trở lên (trùng tên)
    for name, file_paths in name_groups.items():
        if len(file_paths) >= 2:
            colors = []
            for path in file_paths:
                rgb_color, hex_color = extract_dominant_color(path)
                if rgb_color is not None:
                    colors.append((path, rgb_color, hex_color))

            # So sánh các màu sắc trong nhóm
            if len(colors) >= 2:
                for i in range(len(colors)):
                    for j in range(i + 1, len(colors)):
                        path1, rgb1, hex1 = colors[i]
                        path2, rgb2, hex2 = colors[j]

                        if not compare_colors(rgb1, rgb2):
                            # Lấy đường dẫn thư mục con
                            subfolder1 = extract_subfolder_path(path1, directory)
                            subfolder2 = extract_subfolder_path(path2, directory)

                            # Nếu màu khác nhau, thêm vào kết quả
                            results.append({
                                'color_name': name,
                                'file1': os.path.basename(path1),
                                'hex1': hex1,
                                'file2': os.path.basename(path2),
                                'hex2': hex2,
                                'path1': subfolder1,  # Thay thế bằng đường dẫn thư mục con
                                'path2': subfolder2  # Thay thế bằng đường dẫn thư mục con
                            })

    return results


def main():
    # Nhập đường dẫn từ người dùng
    image_directory = r"D:\FlashPOD_Productdetails\test\PNG"
    output_path = r"D:\FlashPOD_Productdetails\test\PNG"

    # Xử lý đường dẫn output
    if not output_path:
        output_file = "tshirt_color_differences.csv"
    else:
        # Đảm bảo đường dẫn tồn tại
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
                print(f"Đã tạo thư mục: {output_dir}")
            except Exception as e:
                print(f"Không thể tạo thư mục {output_dir}: {e}")
                output_path = "tshirt_color_differences.csv"
                print(f"Sẽ lưu kết quả vào thư mục hiện tại: {output_path}")

        # Kiểm tra xem output_path có phải là thư mục không
        if os.path.isdir(output_path):
            output_file = os.path.join(output_path, "tshirt_color_differences.csv")
        else:
            output_file = output_path

    print("Đang quét và phân tích màu sắc...")
    results = find_duplicates_and_compare_colors(image_directory)

    if not results:
        print("Không tìm thấy ảnh nào có tên giống nhau nhưng màu khác nhau.")
        return

    # Lưu kết quả vào file CSV
    df = pd.DataFrame(results)
    df.to_csv(output_file, index=False)

    print(f"Đã tìm thấy {len(results)} cặp ảnh có tên giống nhau nhưng màu khác nhau.")
    print(f"Kết quả đã được lưu vào file: {os.path.abspath(output_file)}")


if __name__ == "__main__":
    main()