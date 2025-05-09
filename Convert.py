from PIL import Image
import os


def convert_image():
    # Đường dẫn đầu vào và đầu ra cố định
    input_path = r"D:\FlashPOD_Productdetails\ORG\G5000"
    output_path = r"D:\FlashPOD_Productdetails\ORG\Fix G5000"

    # Định dạng đầu vào và đầu ra cố định
    input_format = "jpg"
    output_format = "png"

    # Tạo thư mục đầu ra nếu chưa tồn tại
    if not os.path.exists(output_path):
        os.makedirs(output_path)

    # Đếm số file đã chuyển đổi
    count = 0

    # Duyệt qua tất cả các file trong thư mục đầu vào
    for filename in os.listdir(input_path):
        if filename.lower().endswith(f".{input_format}"):
            try:
                # Đường dẫn đầy đủ của file đầu vào
                input_file = os.path.join(input_path, filename)

                # Tên file đầu ra (không có phần mở rộng)
                output_filename = os.path.splitext(filename)[0]

                # Đường dẫn đầy đủ của file đầu ra
                output_file = os.path.join(output_path, f"{output_filename}.{output_format}")

                # Mở và chuyển đổi ảnh
                img = Image.open(input_file)
                img.save(output_file)

                count += 1
                print(f"Đã chuyển đổi: {filename} -> {output_filename}.{output_format}")

            except Exception as e:
                print(f"Lỗi khi chuyển đổi {filename}: {e}")

    print(f"\nHoàn thành! Đã chuyển đổi {count} ảnh từ {input_format} sang {output_format}.")


if __name__ == "__main__":
    print("Bắt đầu chuyển đổi ảnh...")
    convert_image()
    print("Chương trình kết thúc.")