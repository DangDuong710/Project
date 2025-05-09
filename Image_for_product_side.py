from PIL import Image
import os
import glob
import sys

# Cấu hình cố định
INPUT_DIRECTORY = r"D:\FlashPOD_Productdetails\ORG\G5000"
OUTPUT_DIRECTORY = r"D:\FlashPOD_Productdetails\ORG\Fix G5000"
TARGET_WIDTH = 592
TARGET_HEIGHT = 592
TARGET_SIZE_KB = 100 # Kích thước mục tiêu (KB)
BACKGROUND_COLOR = (255, 255, 255)  # Màu trắng

# Định dạng file - có thể tùy chỉnh
INPUT_FORMATS = ["jpg", "jpeg", "png", "bmp", "webp", "tiff"]  # Các định dạng đầu vào được hỗ trợ
OUTPUT_FORMAT = "jpeg"  # Định dạng đầu ra (png, jpg, webp, etc.)


def ensure_directory_exists(directory_path):
    """Đảm bảo thư mục tồn tại, nếu không thì tạo mới."""
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)
        print(f"Đã tạo thư mục: {directory_path}")


def compress_image(img, output_path, output_format, target_kb=100, quality_start=85, min_quality=20):
    """
    Nén ảnh để đạt được kích thước gần với target_kb nhất có thể.

    Args:
        img: Đối tượng ảnh PIL
        output_path: Đường dẫn lưu ảnh
        output_format: Định dạng đầu ra (PNG, JPEG, etc.)
        target_kb: Kích thước mục tiêu tính bằng KB
        quality_start: Chất lượng khởi đầu (1-100)
        min_quality: Chất lượng tối thiểu chấp nhận được

    Returns:
        tuple: (quality_used, actual_size_kb)
    """
    # Chuẩn hóa định dạng đầu ra
    output_format = output_format.upper()

    # Xác định định dạng lưu (PIL sử dụng "JPEG" thay vì "JPG")
    save_format = "JPEG" if output_format == "JPG" else output_format

    # Kiểm tra xem định dạng có hỗ trợ nén chất lượng hay không
    supports_quality = output_format in ["JPEG", "JPG", "WEBP"]

    if not supports_quality:
        # Với các định dạng không hỗ trợ chất lượng (như PNG), sử dụng tối ưu hóa thay thế
        if output_format == "PNG":
            img.save(output_path, format=save_format, optimize=True, compress_level=9)
        else:
            img.save(output_path, format=save_format)
        current_size = os.path.getsize(output_path) / 1024
        return None, current_size

    # Xử lý cho các định dạng hỗ trợ chất lượng
    quality = quality_start
    target_bytes = target_kb * 1024

    # Thử lưu ảnh với chất lượng ban đầu để kiểm tra
    img.save(output_path, format=save_format, quality=quality, optimize=True)
    current_size = os.path.getsize(output_path)

    if current_size <= target_bytes:
        # Nếu kích thước đã nhỏ hơn mục tiêu, không cần nén thêm
        return quality, current_size / 1024

    # Tìm kiếm nhị phân để đạt được chất lượng phù hợp
    max_quality = quality_start
    min_quality = min_quality

    while max_quality - min_quality > 1:
        quality = (max_quality + min_quality) // 2
        img.save(output_path, format=save_format, quality=quality, optimize=True)
        current_size = os.path.getsize(output_path)

        if current_size > target_bytes:
            max_quality = quality
        else:
            min_quality = quality

    # Lưu với chất lượng cuối cùng
    quality = min_quality
    img.save(output_path, format=save_format, quality=quality, optimize=True)
    final_size = os.path.getsize(output_path)

    return quality, final_size / 1024


def convert_and_resize(input_path, output_path, output_format):
    """
    Chuyển đổi, resize và giảm dung lượng ảnh.

    Args:
        input_path: Đường dẫn đến file ảnh đầu vào
        output_path: Đường dẫn đến file ảnh đầu ra
        output_format: Định dạng file đầu ra
    """
    try:
        # Mở ảnh gốc
        img = Image.open(input_path)

        # Chuyển đổi sang RGB nếu cần thiết (cho các định dạng như RGBA, P, etc.)
        if img.mode != "RGB":
            img = img.convert("RGB")

        # Tính toán tỷ lệ để giữ nguyên hình dạng
        orig_width, orig_height = img.size
        ratio = min(TARGET_WIDTH / orig_width, TARGET_HEIGHT / orig_height)
        new_size = (int(orig_width * ratio), int(orig_height * ratio))

        # Resize ảnh giữ nguyên tỷ lệ
        img_resized = img.resize(new_size, Image.LANCZOS)

        # Tạo một ảnh trắng có kích thước đích
        new_img = Image.new("RGB", (TARGET_WIDTH, TARGET_HEIGHT), BACKGROUND_COLOR)

        # Paste ảnh đã resize vào giữa ảnh mới
        paste_position = ((TARGET_WIDTH - new_size[0]) // 2,
                          (TARGET_HEIGHT - new_size[1]) // 2)
        new_img.paste(img_resized, paste_position)

        # Nén và lưu ảnh
        quality, final_size = compress_image(new_img, output_path, output_format, target_kb=TARGET_SIZE_KB)

        quality_info = f", chất lượng={quality}" if quality is not None else ""
        print(f"Đã chuyển đổi và lưu ảnh: {os.path.basename(input_path)}")
        print(f"  → {os.path.basename(output_path)} ({final_size:.2f}KB{quality_info})")

    except Exception as e:
        print(f"Lỗi khi xử lý {input_path}: {e}")
        return None


def process_directory():
    """Xử lý tất cả ảnh theo định dạng đầu vào trong thư mục đầu vào."""
    # Đảm bảo thư mục đầu ra tồn tại
    ensure_directory_exists(OUTPUT_DIRECTORY)

    # Tìm tất cả file theo định dạng đầu vào
    input_files = []
    for format in INPUT_FORMATS:
        pattern = os.path.join(INPUT_DIRECTORY, f"*.{format.lower()}")
        input_files.extend(glob.glob(pattern))
        # Kiểm tra cả định dạng viết hoa
        if format.lower() != format.upper():
            pattern = os.path.join(INPUT_DIRECTORY, f"*.{format.upper()}")
            input_files.extend(glob.glob(pattern))

    if not input_files:
        print(f"Không tìm thấy file theo định dạng {', '.join(INPUT_FORMATS)} trong thư mục {INPUT_DIRECTORY}")
        return

    print(f"Tìm thấy {len(input_files)} file. Bắt đầu xử lý...")

    # Xử lý từng file
    for i, input_file in enumerate(input_files, 1):
        base_name = os.path.basename(input_file)
        name_without_ext = os.path.splitext(base_name)[0]
        output_path = os.path.join(OUTPUT_DIRECTORY, f"{name_without_ext}.{OUTPUT_FORMAT.lower()}")

        print(f"[{i}/{len(input_files)}] Đang xử lý {base_name}...")
        convert_and_resize(input_file, output_path, OUTPUT_FORMAT)

    print(f"\nHoàn thành! Đã xử lý {len(input_files)} ảnh.")
    print(f"Các ảnh đã được lưu tại: {OUTPUT_DIRECTORY}")


def display_settings():
    """Hiển thị cài đặt hiện tại của công cụ."""
    print("=== CÔNG CỤ CHUYỂN ĐỔI ẢNH ===")
    print(f"Thư mục đầu vào: {INPUT_DIRECTORY}")
    print(f"Thư mục đầu ra: {OUTPUT_DIRECTORY}")
    print(f"Định dạng đầu vào được hỗ trợ: {', '.join(INPUT_FORMATS)}")
    print(f"Định dạng đầu ra: {OUTPUT_FORMAT}")
    print(f"Kích thước đích: {TARGET_WIDTH}x{TARGET_HEIGHT} pixels")
    print(f"Dung lượng đích: {TARGET_SIZE_KB}KB")
    print("=" * 45)


if __name__ == "__main__":
    display_settings()
    process_directory()