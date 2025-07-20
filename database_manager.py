# database_manager.py
import sqlite3
import os
import json
from datetime import datetime
import pytz

class DatabaseManager:
    def __init__(self, db_folder="C:\\111accommodationDB"):
        self.db_folder = db_folder
        self.db_path = os.path.join(db_folder, "apply.db")
        self.timezone = pytz.timezone('Asia/Tokyo')
        self.init_database()
    
    def get_jst_timestamp(self):
        """日本時間のタイムスタンプを取得"""
        return datetime.now(self.timezone).strftime('%Y-%m-%d %H:%M:%S')
    
    def init_database(self):
        """データベースとテーブルの初期化"""
        os.makedirs(self.db_folder, exist_ok=True)
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # アプリケーション申し込みテーブル作成
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS applications (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email TEXT NOT NULL,
                page_count INTEGER NOT NULL,
                stations TEXT NOT NULL,
                urls TEXT NOT NULL,
                request_timestamp TEXT NOT NULL,
                processing_status TEXT DEFAULT 'processing',
                output_folder TEXT,
                total_scraped_items INTEGER DEFAULT 0,
                scraped_stations INTEGER DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
        ''')
        
        # 処理結果テーブル作成
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS processing_results (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                application_id INTEGER,
                station_name TEXT NOT NULL,
                scraped_count INTEGER NOT NULL,
                csv_files TEXT,
                ppt_file TEXT,
                processing_time REAL,
                created_at TEXT NOT NULL,
                FOREIGN KEY (application_id) REFERENCES applications (id)
            )
        ''')
        
        conn.commit()
        conn.close()
        print(f"データベース初期化完了: {self.db_path}")
    
    def save_application(self, email, page_count, stations, urls, request_timestamp):
        """申し込み情報をデータベースに保存"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # リストをJSONとして保存
        stations_json = json.dumps(stations, ensure_ascii=False)
        urls_json = json.dumps(urls, ensure_ascii=False)
        
        jst_time = self.get_jst_timestamp()
        
        cursor.execute('''
            INSERT INTO applications 
            (email, page_count, stations, urls, request_timestamp, processing_status, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (email, page_count, stations_json, urls_json, request_timestamp, 'processing', jst_time, jst_time))
        
        application_id = cursor.lastrowid
        conn.commit()
        conn.close()
        
        print(f"申し込み情報保存完了 - ID: {application_id}, Email: {email}, 時刻: {jst_time}")
        return application_id
    
    def update_application_result(self, application_id, output_folder, total_scraped, scraped_stations):
        """申し込み処理結果の更新"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        jst_time = self.get_jst_timestamp()
        
        cursor.execute('''
            UPDATE applications 
            SET processing_status = ?, output_folder = ?, total_scraped_items = ?, 
                scraped_stations = ?, updated_at = ?
            WHERE id = ?
        ''', ('completed', output_folder, total_scraped, scraped_stations, jst_time, application_id))
        
        conn.commit()
        conn.close()
        print(f"処理結果更新完了 - ID: {application_id}, 時刻: {jst_time}")
    
    def save_station_result(self, application_id, station_name, scraped_count, 
                          csv_files=None, ppt_file=None, processing_time=None):
        """駅別処理結果の保存"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        csv_files_json = json.dumps(csv_files, ensure_ascii=False) if csv_files else None
        jst_time = self.get_jst_timestamp()
        
        cursor.execute('''
            INSERT INTO processing_results 
            (application_id, station_name, scraped_count, csv_files, ppt_file, processing_time, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (application_id, station_name, scraped_count, csv_files_json, ppt_file, processing_time, jst_time))
        
        conn.commit()
        conn.close()
        print(f"駅別結果保存完了 - {station_name}: {scraped_count}件, 時刻: {jst_time}")
    
    def update_application_error(self, application_id, error_message):
        """エラー発生時の状態更新"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        jst_time = self.get_jst_timestamp()
        
        cursor.execute('''
            UPDATE applications 
            SET processing_status = ?, updated_at = ?
            WHERE id = ?
        ''', (f'error: {error_message}', jst_time, application_id))
        
        conn.commit()
        conn.close()
        print(f"エラー状態更新完了 - ID: {application_id}, 時刻: {jst_time}")
    
    def get_application_history(self, email=None, limit=50):
        """申し込み履歴の取得"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        if email:
            cursor.execute('''
                SELECT * FROM applications 
                WHERE email = ? 
                ORDER BY created_at DESC 
                LIMIT ?
            ''', (email, limit))
        else:
            cursor.execute('''
                SELECT * FROM applications 
                ORDER BY created_at DESC 
                LIMIT ?
            ''', (limit,))
        
        results = cursor.fetchall()
        conn.close()
        
        return results
    
    def get_processing_stats(self):
        """処理統計の取得"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 全体統計
        cursor.execute('''
            SELECT 
                COUNT(*) as total_applications,
                SUM(CASE WHEN processing_status = 'completed' THEN 1 ELSE 0 END) as completed,
                SUM(CASE WHEN processing_status = 'processing' THEN 1 ELSE 0 END) as processing,
                SUM(CASE WHEN processing_status LIKE 'error:%' THEN 1 ELSE 0 END) as errors,
                SUM(total_scraped_items) as total_items_scraped
            FROM applications
        ''')
        
        stats = cursor.fetchone()
        
        # 最新の処理時刻も取得
        cursor.execute('''
            SELECT MAX(created_at) as latest_created_at,
                   MAX(updated_at) as latest_updated_at
            FROM applications
        ''')
        
        times = cursor.fetchone()
        conn.close()
        
        return {
            'total_applications': stats[0],
            'completed': stats[1],
            'processing': stats[2],
            'errors': stats[3],
            'total_items_scraped': stats[4] or 0,
            'latest_created_at': times[0],
            'latest_updated_at': times[1],
            'timezone': 'Asia/Tokyo'
        }