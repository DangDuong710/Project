#!/usr/bin/env python3

import requests
import pandas as pd
from pathlib import Path
import time
import re
import concurrent.futures
from threading import Lock

# ===== CONFIG =====
DOWNLOAD_FOLDER = r"D:\FILE\New folder"
SHEET_URL = "https://docs.google.com/spreadsheets/d/1_A6QdeYj_HCT23Y1yFMCHxBE6wMoYth5QL__ZZ_AuYU/edit"
DELAY_BETWEEN_DOWNLOADS = 0.5
MAX_FILES = None
MAX_WORKERS = 5


# ==================

class FastGSheetDownloader:
    def __init__(self, download_folder):
        self.download_folder = Path(download_folder)
        self.download_folder.mkdir(exist_ok=True)
        self.lock = Lock()
        self.success_count = 0
        self.failed_count = 0

        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'application/pdf,application/octet-stream,*/*',
        }

    def get_csv_url(self, sheet_url):
        if '/d/' in sheet_url:
            sheet_id = sheet_url.split('/d/')[1].split('/')[0]
        else:
            sheet_id = sheet_url
        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"

    def read_sheet_data(self, sheet_url):
        try:
            csv_url = self.get_csv_url(sheet_url)
            print(f"Reading sheet data...")

            df = pd.read_csv(csv_url)
            data_pairs = []

            for _, row in df.iterrows():
                try:
                    order_code = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else None
                    link = str(row.iloc[10]).strip() if pd.notna(row.iloc[10]) else None

                    if (order_code and link and order_code != 'nan' and link != 'nan' and
                            'http' in link and ('.pdf' in link.lower() or 'mangoteeprints.com' in link)):
                        data_pairs.append((order_code, link))
                except:
                    continue

            print(f"Found {len(data_pairs)} valid pairs")
            return data_pairs

        except Exception as e:
            print(f"Error reading sheet: {e}")
            return []

    def download_single(self, order_code, url):
        try:
            safe_filename = re.sub(r'[<>:"/\\|?*]', '_', order_code)
            if not safe_filename.endswith('.pdf'):
                safe_filename += '.pdf'

            file_path = self.download_folder / safe_filename

            if file_path.exists():
                with self.lock:
                    self.success_count += 1
                return f"EXISTS: {safe_filename}"

            response = requests.get(url, headers=self.headers, timeout=30, stream=True)
            response.raise_for_status()

            with open(file_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            with self.lock:
                self.success_count += 1
            return f"SUCCESS: {safe_filename}"

        except Exception as e:
            with self.lock:
                self.failed_count += 1
            return f"FAILED: {order_code} - {str(e)[:50]}"

    def print_progress_bar(self, completed, total, success, failed, elapsed):
        progress = completed / total
        bar_length = 30
        filled_length = int(bar_length * progress)
        bar = '█' * filled_length + '░' * (bar_length - filled_length)

        speed = completed / elapsed if elapsed > 0 else 0
        eta = (total - completed) / speed if speed > 0 else 0

        print(f"\r[{bar}] {completed}/{total} ({progress:.1%}) | "
              f"✓{success} ✗{failed} | {speed:.1f}/s | ETA: {eta:.0f}s", end='', flush=True)

    def download_all(self, sheet_url, max_workers=5, delay=0.5, max_files=None):
        print("Fast PDF Downloader Started")
        print("-" * 50)

        data_pairs = self.read_sheet_data(sheet_url)

        if not data_pairs:
            print("No valid data found")
            return {"success": 0, "failed": 0, "total": 0}

        if max_files:
            data_pairs = data_pairs[:max_files]

        print(f"Files to download: {len(data_pairs)}")
        print(f"Workers: {max_workers}")
        print(f"Folder: {self.download_folder}")
        print("-" * 50)

        start_time = time.time()
        completed = 0

        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            print("Submitting tasks...")

            # Submit all tasks at once
            futures = [executor.submit(self.download_single, order_code, link)
                       for order_code, link in data_pairs]

            print("Starting downloads...")

            # Process completed futures
            for future in concurrent.futures.as_completed(futures):
                result = future.result()
                completed += 1
                elapsed = time.time() - start_time

                self.print_progress_bar(completed, len(data_pairs),
                                        self.success_count, self.failed_count, elapsed)

        elapsed_time = time.time() - start_time

        print(f"\n{'-' * 50}")
        print(f"COMPLETED in {elapsed_time:.1f}s")
        print(f"✓ Success: {self.success_count}")
        print(f"✗ Failed: {self.failed_count}")
        print(f"+ Total: {len(data_pairs)}")
        print(f"⚡ Speed: {len(data_pairs) / elapsed_time:.1f} files/sec")
        print("-" * 50)

        return {
            "success": self.success_count,
            "failed": self.failed_count,
            "total": len(data_pairs),
            "time": elapsed_time
        }


def main():
    downloader = FastGSheetDownloader(DOWNLOAD_FOLDER)
    downloader.download_all(
        sheet_url=SHEET_URL,
        max_workers=MAX_WORKERS,
        delay=DELAY_BETWEEN_DOWNLOADS,
        max_files=MAX_FILES
    )


if __name__ == "__main__":
    main()