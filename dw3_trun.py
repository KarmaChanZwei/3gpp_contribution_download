import os
import pandas as pd
import requests
from urllib.parse import urlparse
from openpyxl import load_workbook
import re
import time
from tqdm import tqdm
import urllib3

# 关闭 urllib3 的 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 获取当前脚本所在目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# 查找目录中以 ".xlsx" 结尾的文件
excel_files = [f for f in os.listdir(current_dir) if f.endswith('.xlsx')]

WINDOWS_MAX_PATH_LENGTH = 255

def clean_filename(filename):
    filename = re.sub(r'[\\/*?:"<>|]', '_', filename)
    if len(filename) > WINDOWS_MAX_PATH_LENGTH:
        filename = filename[:WINDOWS_MAX_PATH_LENGTH]
    return filename

def download_file(url, save_path, max_retries=10, timeout=60):
    """
    下载文件，支持断点续传 + 进度条 + SSL fallback
    """
    for attempt in range(max_retries):
        try:
            headers = {}
            # 断点续传
            if os.path.exists(save_path):
                existing_size = os.path.getsize(save_path)
                headers['Range'] = f'bytes={existing_size}-'
            else:
                existing_size = 0

            # 先尝试 verify=True
            try_verify = True
            try:
                r = requests.get(url, headers=headers, timeout=timeout, stream=True, verify=True)
            except Exception:
                # 如果 SSL 验证失败，fallback 到 verify=False
                try_verify = False
                r = requests.get(url, headers=headers, timeout=timeout, stream=True, verify=False)

            if r.status_code in (200, 206):
                total_size = int(r.headers.get('content-length', 0)) + existing_size
                mode = 'ab' if existing_size > 0 else 'wb'

                # 下载进度条
                with open(save_path, mode) as f, tqdm(
                    total=total_size, unit='B', unit_scale=True,
                    desc=os.path.basename(save_path), initial=existing_size, leave=False
                ) as pbar:
                    for chunk in r.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
                            pbar.update(len(chunk))
                return True
            else:
                print(f"下载失败 {url}, 状态码: {r.status_code}, verify={try_verify}")
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

    # 外层 Excel 文件进度条
    for index, row in tqdm(df.iterrows(), total=len(df), desc="Excel文件处理进度", unit="file"):
        if index == 0:  # 跳过标题行
            continue

        folder_name = clean_filename(str(row.iloc[10]))
        folder_path = os.path.join(current_dir, folder_name)

        if len(folder_path) > WINDOWS_MAX_PATH_LENGTH:
            print(f"警告: 文件路径太长，已截断: {folder_path}")
            folder_path = folder_path[:WINDOWS_MAX_PATH_LENGTH]

        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        cell = sheet.cell(row=index+2, column=1)
        hyperlink = cell.hyperlink.target if cell.hyperlink else None

        if hyperlink:
            try:
                url_path = urlparse(hyperlink).path
                zip_filename = os.path.basename(url_path)
                save_path = os.path.join(folder_path, zip_filename)

                success = download_file(hyperlink, save_path)
                if success:
                    new_filename = clean_filename(f"{str(row.iloc[0])} {str(row.iloc[2])} {str(row.iloc[1])}.zip")
                    new_path = os.path.join(folder_path, new_filename)
                    os.rename(save_path, new_path)
                    print(f"\n文件下载并重命名为: {new_path}")
                else:
                    print(f"最终下载失败: {hyperlink}")
            except Exception as e:
                print(f"处理链接时出错: {hyperlink}, 错误: {str(e)}")
        else:
            print(f"超链接无效或不存在: {hyperlink}")
else:
    print("当前目录中没有找到Excel文件。")
