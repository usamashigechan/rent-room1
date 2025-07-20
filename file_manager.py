# file_manager.py
import os
import shutil
from datetime import datetime

class FileManager:
    def __init__(self, base_folder):
        self.base_folder = base_folder
    
    def save_csv(self, df, filename):
        """CSVファイルの保存"""
        full_path = os.path.join(self.base_folder, filename)
        df.to_csv(full_path, index=False, encoding="utf-8-sig")
        return full_path
    
    def create_timestamped_folder(self, datestamp):
        """タイムスタンプ付きフォルダの作成と移動"""
        timestamp_folder = datetime.now().strftime("A%Y%m%d%H%M")
        dest_folder = os.path.join(self.base_folder, timestamp_folder)
        os.makedirs(dest_folder, exist_ok=True)
        
        for filename in os.listdir(self.base_folder):
            file_path = os.path.join(self.base_folder, filename)
            
            if os.path.isfile(file_path) and datestamp in filename:
                shutil.move(file_path, os.path.join(dest_folder, filename))
                print(f"移動: {filename} -> {dest_folder}")
        
        return dest_folder