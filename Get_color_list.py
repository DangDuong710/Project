import os
import argparse
from PIL import Image
from collections import Counter


def get_most_common_color(image_path):
    """
    Trích xuất mã màu HEX phổ biến nhất từ ảnh
    """
    try:
        img = Image.open(image_path)
        # Chuyển sang RGB nếu cần
        if img.mode != 'RGB':
            img = img.convert('RGB')

        # Giảm kích thước ảnh để tăng tốc độ xử lý
        img = img.resize((100, 100))

        # Lấy tất cả các pixel
        pixels = list(img.getdata())

        # Đếm số lượng mỗi màu
        color_counter = Counter(pixels)

        # Tìm màu phổ biến nhất (loại bỏ màu trắng và đen nếu cần)
        most_common_color = color_counter.most_common(10)

        # Lấy màu phổ biến nhất không phải trắng hoặc đen
        for color in most_common_color:
            rgb = color[0]
            # Bỏ qua màu trắng và đen
            if not (rgb[0] > 240 and rgb[1] > 240 and rgb[2] > 240) and not (
                    rgb[0] < 15 and rgb[1] < 15 and rgb[2] < 15):
                hex_color = '#{:02x}{:02x}{:02x}'.format(rgb[0], rgb[1], rgb[2])
                return hex_color.upper()

        # Nếu không tìm thấy màu phù hợp, trả về màu phổ biến nhất
        rgb = most_common_color[0][0]
        hex_color = '#{:02x}{:02x}{:02x}'.format(rgb[0], rgb[1], rgb[2])
        return hex_color.upper()
    except Exception as e:
        print(f"Lỗi khi xử lý ảnh {image_path}: {e}")
        return "N/A"


def scan_folders(root_path):
    """
    Quét thư mục để tìm các file ảnh và trích xuất thông tin
    """
    results = []

    for type_folder in os.listdir(root_path):
        type_path = os.path.join(root_path, type_folder)
        if os.path.isdir(type_path):
            for material_folder in os.listdir(type_path):
                material_path = os.path.join(type_path, material_folder)
                if os.path.isdir(material_path):
                    # Đây là nơi chứa các file ảnh
                    for color_file in os.listdir(material_path):
                        if color_file.lower().endswith(('.png', '.jpg', '.jpeg')):
                            image_path = os.path.join(material_path, color_file)

                            # Lấy tên màu là tên file (không bao gồm phần mở rộng)
                            color_name = os.path.splitext(color_file)[0]

                            # Lấy loại và chất liệu từ đường dẫn
                            type_name = type_folder
                            material_name = material_folder

                            # Lấy mã hex
                            hex_code = get_most_common_color(image_path)

                            results.append({
                                'color_name': color_name,
                                'hex_code': hex_code,
                                'type': type_name,
                                'material': material_name,
                                'file_path': image_path
                            })

    return results


def main():
    parser = argparse.ArgumentParser(description='Trích xuất thông tin màu từ các file ảnh.')
    parser.add_argument('--path', default=r"D:\FlashPOD_Productdetails\PNG file",
                        help='Đường dẫn đến thư mục gốc (mặc định: D:\\FlashPOD_Productdetails\\PNG file)')
    parser.add_argument('--output', default='color_data.csv',
                        help='Tên file output (mặc định: color_data.csv)')

    args = parser.parse_args()

    # Đường dẫn thư mục gốc
    root_path = args.path

    # Tạo đường dẫn đầy đủ cho file output trong thư mục gốc
    output_file_path = os.path.join(root_path, args.output)

    print(f"Đang quét thư mục: {root_path}")
    results = scan_folders(root_path)

    # Xuất kết quả ra console
    print("\nKết quả:")
    print(f"{'Màu':<20} {'Mã HEX':<10} {'Type':<15} {'Material':<15}")
    print("-" * 60)

    for item in results:
        print(f"{item['color_name']:<20} {item['hex_code']:<10} {item['type']:<15} {item['material']:<15}")

    # Xuất kết quả ra file CSV trong thư mục gốc
    with open(output_file_path, 'w', encoding='utf-8') as f:
        f.write("color_name,hex_code,type,material,file_path\n")
        for item in results:
            f.write(f"{item['color_name']},{item['hex_code']},{item['type']},{item['material']},{item['file_path']}\n")

    print(f"\nĐã lưu kết quả vào file: {output_file_path}")
    print(f"Tổng số màu đã xử lý: {len(results)}")


if __name__ == "__main__":
    main()