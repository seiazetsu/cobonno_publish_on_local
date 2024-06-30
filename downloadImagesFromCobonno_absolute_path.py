import csv
import requests
import os
import uuid
from urllib.parse import urlparse, unquote
from PyQt5.QtWidgets import QApplication, QFileDialog

def download_image(url, save_path):
    response = requests.get(url)
    if response.status_code == 200:
        with open(save_path, 'wb') as file:
            file.write(response.content)
        print(f"Downloaded: {save_path}")
    else:
        print(f"Failed to download: {url}")

def process_csv(csv_file_path, download_folder):
    updated_rows = []
    with open(csv_file_path, 'r') as csv_file:
        csv_reader = csv.reader(csv_file)
        headers = next(csv_reader)
        updated_rows.append(headers)  # ヘッダーを追加
        for row in csv_reader:
            date, title, description, image_url = row
            parsed_url = urlparse(image_url)
            image_name = os.path.basename(unquote(parsed_url.path))  # パーセントエンコーディングをデコード
            unique_id = str(uuid.uuid4())  # ユニークなIDを生成
            file_extension = os.path.splitext(image_name)[1]
            unique_image_name = f"{os.path.splitext(image_name)[0]}_{unique_id}{file_extension}"
            save_path = os.path.join(download_folder, unique_image_name)
            download_image(image_url, save_path)
            row[3] = os.path.abspath(save_path)  # 画像URLを絶対パスに置き換え
            updated_rows.append(row)

    updated_csv_file_path = os.path.splitext(csv_file_path)[0] + '_updated.csv'
    with open(updated_csv_file_path, 'w', newline='') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerows(updated_rows)
    print(f"Updated CSV saved to: {updated_csv_file_path}")

    # 元のCSVファイルを削除
    os.remove(csv_file_path)
    print(f"Original CSV file deleted: {csv_file_path}")

    # 生成されたCSVファイルの内容を表示
    with open(updated_csv_file_path, 'r') as csv_file:
        csv_content = csv_file.read()
        print("Updated CSV Content:")
        print(csv_content)

# ファイルダイアログでCSVファイルのパスを選択
app = QApplication([])
csv_file_path, _ = QFileDialog.getOpenFileName(None, "Select CSV File", "", "CSV files (*.csv)")

# ダウンロードフォルダの指定
download_folder = 'images'

if not os.path.exists(download_folder):
    os.makedirs(download_folder)

if csv_file_path:
    process_csv(csv_file_path, download_folder)
else:
    print("No file selected.")
