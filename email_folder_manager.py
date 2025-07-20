# email_folder_manager.py
import os
import shutil
from datetime import datetime

class EmailFolderManager:
    def __init__(self, base_folder="C:\\111accommodationDB"):
        self.base_folder = base_folder
        
    def create_email_folder(self, email, application_id):
        """メールアドレス別フォルダの作成"""
        # メールアドレスをファイルシステム用に変換（@を_atに変更）
        safe_email = email.replace("@", "_at_").replace(".", "_")
        
        # フォルダ構造: C:\111accommodationDB\email\user_at_example_com\20241219_001
        email_folder = os.path.join(self.base_folder, "email", safe_email)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        app_folder = os.path.join(email_folder, f"{timestamp}_{application_id:03d}")
        
        os.makedirs(app_folder, exist_ok=True)
        print(f"メール別フォルダ作成: {app_folder}")
        
        return app_folder
    
    def move_results_to_email_folder(self, source_folder, email_folder):
        """処理結果をメール別フォルダに移動"""
        if not os.path.exists(source_folder):
            print(f"移動元フォルダが存在しません: {source_folder}")
            return False
        
        try:
            # フォルダ内のすべてのファイルを移動
            for filename in os.listdir(source_folder):
                source_file = os.path.join(source_folder, filename)
                dest_file = os.path.join(email_folder, filename)
                
                if os.path.isfile(source_file):
                    shutil.move(source_file, dest_file)
                    print(f"移動完了: {filename}")
            
            # 元の空フォルダを削除
            if os.path.exists(source_folder) and not os.listdir(source_folder):
                os.rmdir(source_folder)
                print(f"空フォルダ削除: {source_folder}")
            
            return True
        except Exception as e:
            print(f"ファイル移動エラー: {e}")
            return False
    
    def get_user_folders(self, email):
        """ユーザーのフォルダ一覧取得"""
        safe_email = email.replace("@", "_at_").replace(".", "_")
        email_folder = os.path.join(self.base_folder, "email", safe_email)
        
        if not os.path.exists(email_folder):
            return []
        
        folders = []
        for folder_name in os.listdir(email_folder):
            folder_path = os.path.join(email_folder, folder_name)
            if os.path.isdir(folder_path):
                folders.append({
                    'folder_name': folder_name,
                    'folder_path': folder_path,
                    'created_time': datetime.fromtimestamp(os.path.getctime(folder_path))
                })
        
        return sorted(folders, key=lambda x: x['created_time'], reverse=True)