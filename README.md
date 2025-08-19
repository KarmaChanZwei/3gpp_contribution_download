import os
import pandas as pd
import requests
from urllib.parse import urlparse
from openpyxl import load_workbook
import re
import time

current_dir = os.path.dirname(os.path.abspath(__file__))
excel_files = [f for f in os.listdir(current_dir) if f.endswith('.xlsx')]
WINDOWS_MAX_PATH_LENGTH = 255

def clean_filename(filename):
    filename = re.sub(r'[\\/*?:"<>|]', '_', filename)
    if len(filename) > WINDOWS_MAX_PATH_LENGTH:
        filename = filename[:WINDOWS_MAX_PATH_LENGTH]
    return filename

def download_file(url, save_path, max_retries=5, timeout=60):
    for attempt in range(max_retries):
        try:
            headers = {}
            if os.path.exists(save_path):
                existing_size = os.path.getsize(save_path)
                headers['Range'] = f'bytes={existing_size}-'
            else:
                existing_size = 0

            r = requests.get(url, headers=headers, timeout=timeout, stream=True, verify=False)
            if r.status_code in (200, 206):
                mode = 'ab' if existing_size > 0 else 'wb'
                with open(save_path, mode) as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
                return True
            else:
                print(f"下载失败 {url}, 状态码: {r.status_code}")
        except Exception as e:
            print(f"第 {attempt+1} 次下载失败: {e}")
            time.sleep(3)
    return False

if len(excel_files) > 0:
    excel_file = os.path.join(current_dir, excel_files[0])
    print(f"读取文件: {excel_file}")

    workbook = load_workbook(excel_file)
    sheet = workbook.active
    df = pd.read_excel(excel_file)

    total_files = len(df) - 1  # 除去标题行
    finished = 0

    for index, row in df.iterrows():
        if index == 0:  # 跳过标题行
            continue

        folder_name = clean_filename(str(row.iloc[10]))
        folder_path = os.path.join(current_dir, folder_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        cell = sheet.cell(row=index+2, column=1)
        hyperlink = cell.hyperlink.target if cell.hyperlink else None

        if hyperlink:
            url_path = urlparse(hyperlink).path
            zip_filename = os.path.basename(url_path)
            save_path = os.path.join(folder_path, zip_filename)

            success = download_file(hyperlink, save_path)
            if success:
                new_filename = clean_filename(f"{str(row.iloc[0])} {str(row.iloc[2])} {str(row.iloc[1])}.zip")
                new_path = os.path.join(folder_path, new_filename)
                os.rename(save_path, new_path)
                print(f"文件下载并重命名为: {new_path}")
            else:
                print(f"最终下载失败: {hyperlink}")
        else:
            print(f"超链接无效或不存在: {hyperlink}")

        # 更新整体进度
        finished += 1
        percent = (finished / total_files) * 100
        print(f"整体进度: {finished}/{total_files} ({percent:.2f}%)")

else:
    print("当前目录中没有找到Excel文件。")
