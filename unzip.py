import os
import shutil

def extract_zip_files():
    current_dir = os.getcwd()

    for root, dirs, files in os.walk(current_dir):
        for file in files:
            if file.endswith(".zip"):
                zip_path = os.path.join(root, file)
                
                # 解压缩文件到所在的目录中
                try:
                    shutil.unpack_archive(zip_path, root)
                    print(f"解压缩完成: {file} -> {root}")
                except Exception as e:
                    print(f"解压缩失败: {file}, 错误: {e}")

if __name__ == "__main__":
    extract_zip_files()
