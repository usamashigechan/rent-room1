# cloud_main.py - Cloud Run対応版
import os
import time
import pickle
import base64
import glob
import mimetypes
import re
import json
import tempfile
import shutil
import sqlite3
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.cloud import storage

from flask import Flask, request, jsonify, send_from_directory
from config import Config
from scraper import PropertyScraper
from data_processor import DataProcessor
from statistics_calc import StatisticsCalculator
from visualizer import Visualizer
from powerpoint_generator import PowerPointGenerator
from summary_generator import SummaryGenerator
from file_manager import FileManager
from station_processor import StationProcessor

# Cloud Storage データベース管理クラス
class CloudDatabaseManager:
    def __init__(self):
        self.bucket_name = os.environ.get('STORAGE_BUCKET', 'your-property-data-bucket')
        self.db_filename = "apply.db"
        self.local_db_path = "apply.db"
        self.client = storage.Client()
        self.bucket = self.client.bucket(self.bucket_name)
        
    def download_db(self):
        """Cloud StorageからDBダウンロード"""
        try:
            blob = self.bucket.blob(self.db_filename)
            if blob.exists():
                blob.download_to_filename(self.local_db_path)
                print(f"既存データベースダウンロード完了: {self.local_db_path}")
                return True
            else:
                print("新規データベースを作成します")
                self.create_new_db()
                return False
        except Exception as e:
            print(f"データベースダウンロードエラー: {e}")
            self.create_new_db()
            return False
    
    def upload_db(self):
        """Cloud StorageにDBアップロード"""
        try:
            if os.path.exists(self.local_db_path):
                blob = self.bucket.blob(self.db_filename)
                blob.upload_from_filename(self.local_db_path)
                print("データベースアップロード完了")
                return True
        except Exception as e:
            print(f"データベースアップロードエラー: {e}")
        return False
    
    def create_new_db(self):
        """新規データベース作成"""
        conn = sqlite3.connect(self.local_db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS applications (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email TEXT NOT NULL,
                page INTEGER NOT NULL,
                stations TEXT NOT NULL,
                urls TEXT NOT NULL,
                request_timestamp TEXT,
                processing_status TEXT DEFAULT 'processing',
                output_folder TEXT,
                total_scraped_items INTEGER DEFAULT 0,
                scraped_stations INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS station_results (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                application_id INTEGER,
                station_name TEXT,
                scraped_count INTEGER,
                csv_files TEXT,
                ppt_file TEXT,
                processing_time REAL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (application_id) REFERENCES applications (id)
            )
        ''')
        
        conn.commit()
        conn.close()
        print("新規データベース作成完了")
    
    def cleanup(self):
        """ローカルDBファイル削除"""
        if os.path.exists(self.local_db_path):
            try:
                os.remove(self.local_db_path)
                print("ローカルデータベースファイル削除")
            except:
                pass

# Cloud Run用データベース操作クラス
class CloudRunDatabaseManager:
    def __init__(self, db_path):
        self.db_path = db_path
    
    def save_application(self, email, page, stations, urls, timestamp):
        """申し込み情報保存"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # リストをJSON文字列に変換
        stations_json = json.dumps(stations, ensure_ascii=False)
        urls_json = json.dumps(urls, ensure_ascii=False)
        
        cursor.execute('''
            INSERT INTO applications 
            (email, page, stations, urls, request_timestamp)
            VALUES (?, ?, ?, ?, ?)
        ''', (email, page, stations_json, urls_json, timestamp))
        
        application_id = cursor.lastrowid
        conn.commit()
        conn.close()
        
        print(f"申し込み情報保存完了: ID={application_id}")
        return application_id
    
    def update_application_result(self, application_id, output_folder, total_scraped, scraped_stations):
        """申し込み結果更新"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            UPDATE applications 
            SET output_folder=?, total_scraped_items=?, scraped_stations=?,
                processing_status='completed', updated_at=CURRENT_TIMESTAMP
            WHERE id=?
        ''', (output_folder, total_scraped, scraped_stations, application_id))
        
        conn.commit()
        conn.close()
        print(f"申し込み結果更新完了: ID={application_id}")
    
    def update_application_error(self, application_id, error_message):
        """申し込みエラー更新"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            UPDATE applications 
            SET processing_status='error', output_folder=?,
                updated_at=CURRENT_TIMESTAMP
            WHERE id=?
        ''', (error_message, application_id))
        
        conn.commit()
        conn.close()
    
    def save_station_result(self, application_id, station, count, csv_files, ppt_file, processing_time):
        """駅別結果保存"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        csv_files_json = json.dumps(csv_files, ensure_ascii=False) if csv_files else None
        
        cursor.execute('''
            INSERT INTO station_results
            (application_id, station_name, scraped_count, csv_files, ppt_file, processing_time)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (application_id, station, count, csv_files_json, ppt_file, processing_time))
        
        conn.commit()
        conn.close()
    
    def get_processing_stats(self):
        """処理統計取得"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT 
                COUNT(*) as total_applications,
                COUNT(CASE WHEN processing_status = 'completed' THEN 1 END) as completed,
                COUNT(CASE WHEN processing_status = 'error' THEN 1 END) as errors,
                SUM(total_scraped_items) as total_items,
                AVG(total_scraped_items) as avg_items
            FROM applications
        ''')
        
        stats = cursor.fetchone()
        conn.close()
        
        return {
            "total_applications": stats[0] or 0,
            "completed_applications": stats[1] or 0,
            "error_applications": stats[2] or 0,
            "total_scraped_items": stats[3] or 0,
            "average_items_per_request": round(stats[4] or 0, 2)
        }
    
    def get_application_history(self, email=None, limit=50):
        """申し込み履歴取得"""
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
        
        rows = cursor.fetchall()
        conn.close()
        return rows

# Gmail送信クラス（Cloud Run対応）
class CloudRunGmailSender:
    def __init__(self):
        self.SCOPES = ['https://www.googleapis.com/auth/gmail.send']
        self.service = self.authenticate()
    
    def authenticate(self):
        """Cloud Run用認証"""
        creds = None
        
        # 環境変数からサービスアカウント情報取得を試行
        service_account_path = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS')
        if service_account_path and os.path.exists(service_account_path):
            try:
                from google.oauth2 import service_account
                credentials = service_account.Credentials.from_service_account_file(
                    service_account_path, scopes=self.SCOPES
                )
                service = build('gmail', 'v1', credentials=credentials)
                print("Gmail認証成功（サービスアカウント）")
                return service
            except Exception as e:
                print(f"サービスアカウント認証失敗: {e}")
        
        # フォールバック: 既存のOAuth2.0フロー
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not os.path.exists('credentials.json'):
                    print("ERROR: credentials.json が見つかりません")
                    return None
                
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', self.SCOPES)
                creds = flow.run_local_server(port=0)
            
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
        
        try:
            service = build('gmail', 'v1', credentials=creds)
            print("Gmail認証成功（OAuth2.0）")
            return service
        except Exception as e:
            print(f"Gmail認証エラー: {e}")
            return None
    
    def get_station_from_filename(self, filename):
        """ファイル名から駅名を抽出"""
        patterns = [
            r'1fData_(.+?)_\d{8}_\d{6}\.csv',
            r'1fData_(.+?)_\d{8}\.csv',
            r'1e_(.+?)_\d{8}_\d{6}_ptt1\.pptx',
            r'1e_(.+?)_\d{14}_ptt1\.pptx',
            r'1[a-z]_(.+?)_\d{8}_\d{6}',
            r'1[a-z]_(.+?)_\d{8,14}',
            r'1[a-z]_(.+?)\.',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, filename)
            if match:
                station_name = match.group(1)
                print(f"駅名抽出成功: {filename} → {station_name}")
                return station_name
        
        if '_' in filename:
            parts = filename.split('_')
            if len(parts) >= 2 and parts[0].startswith('1'):
                potential_station = parts[1]
                print(f"駅名抽出（フォールバック）: {filename} → {potential_station}")
                return potential_station
        
        print(f"駅名抽出失敗: {filename}")
        return filename
    
    def create_message(self, to_email, subject, body, attachments=None):
        """メッセージ作成"""
        message = MIMEMultipart()
        message['to'] = to_email
        message['subject'] = subject
        message.attach(MIMEText(body, 'plain', 'utf-8'))
        
        if attachments:
            for i, file_path in enumerate(attachments):
                if os.path.exists(file_path):
                    original_filename = os.path.basename(file_path)
                    print(f"添付ファイル {i+1} 処理中: {original_filename}")
                    
                    content_type, encoding = mimetypes.guess_type(file_path)
                    if content_type is None or encoding is not None:
                        content_type = 'application/octet-stream'
                    
                    main_type, sub_type = content_type.split('/', 1)
                    
                    with open(file_path, 'rb') as attachment_file:
                        if main_type == 'text':
                            part = MIMEText(attachment_file.read().decode('utf-8'), _subtype=sub_type)
                        else:
                            part = MIMEBase(main_type, sub_type)
                            part.set_payload(attachment_file.read())
                            encoders.encode_base64(part)
                    
                    try:
                        part.add_header(
                            'Content-Disposition',
                            'attachment',
                            filename=('utf-8', '', original_filename)
                        )
                    except:
                        try:
                            encoded_filename = Header(original_filename, 'utf-8').encode()
                            part.add_header('Content-Disposition', f'attachment; filename="{encoded_filename}"')
                        except:
                            safe_filename = re.sub(r'[^\w\-_\.]', '_', original_filename)
                            part.add_header('Content-Disposition', f'attachment; filename="{safe_filename}"')
                    
                    message.attach(part)
                    print(f"添付ファイル追加完了: {original_filename}")
                else:
                    print(f"添付ファイルが見つかりません: {file_path}")
        
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
        return {'raw': raw_message}
    
    def send_email(self, to_email, subject, body, attachments=None):
        """メール送信"""
        print(f"*** send_email 呼び出し ***")
        print(f"宛先: {to_email}, 件名: {subject}")
        
        if not self.service:
            print("Gmail認証が完了していません")
            return False
        
        try:
            message = self.create_message(to_email, subject, body, attachments)
            result = self.service.users().messages().send(userId='me', body=message).execute()
            print(f"*** メール送信成功: {result['id']} ***")
            return True
        except Exception as e:
            print(f"メール送信エラー: {e}")
            return False
    
    def send_reception_email(self, email, page, stations, urls):
        """受付確認メール"""
        subject = "受け付けました"
        stations_str = ', '.join(stations) if isinstance(stations, list) else str(stations)
        body = f"""{email}の方、{page}ページ、{stations_str}駅、{urls}の内容で受け付けました。しばらくしたら返信メールを送信されるのでお待ち下さい。遅い場合はエラーの可能性があります。"""
        return self.send_email(email, subject, body)
    
    def send_result_email(self, email, folder_path, stations_processed=None):
        """結果送信メール"""
        subject = "お問い合わせ結果"
        
        if stations_processed:
            stations_str = ', '.join(stations_processed)
            body = f"""お問い合わせいただいた{stations_str}駅の分析結果をお送りします。

添付ファイルの内容：
- 物件データ：スクレイピングした賃貸物件の詳細データ
- 分析レポート：統計分析とグラフを含むPowerPoint資料
- 統計グラフ：価格分布や間取り別統計のグラフ

※これらの結果は参考情報であり、必ず正しいとは限らないとお考え下さい。"""
        else:
            body = "添付ファイルの結果になりましたが、必ず正しいとは限らないとお考え下さい"
        
        attachments = []
        if os.path.exists(folder_path):
            pattern = os.path.join(folder_path, "1*")
            files = glob.glob(pattern)
            files.sort()
            
            for file_path in files:
                if os.path.isfile(file_path):
                    file_size = os.path.getsize(file_path)
                    filename = os.path.basename(file_path)
                    station = self.get_station_from_filename(filename)
                    print(f"添付対象ファイル: {filename} ({file_size} bytes) - 駅名: {station}")
                    attachments.append(file_path)
        
        print(f"結果メール送信: フォルダ={folder_path}, 添付={len(attachments)}個")
        return self.send_email(email, subject, body, attachments)

# Cloud Run用設定クラス
class CloudRunConfig(Config):
    def __init__(self):
        super().__init__()
        # 一時ディレクトリを使用
        self.FOLDER_PATH = tempfile.mkdtemp(prefix="property_data_")
        print(f"一時作業フォルダ作成: {self.FOLDER_PATH}")
    
    def cleanup(self):
        """一時ディレクトリ削除"""
        if hasattr(self, 'FOLDER_PATH') and os.path.exists(self.FOLDER_PATH):
            shutil.rmtree(self.FOLDER_PATH, ignore_errors=True)
            print(f"一時作業フォルダ削除: {self.FOLDER_PATH}")

# グローバル変数
cloud_db_manager = None
config = None
gmail_sender = None

# Flaskアプリ初期化
app = Flask(__name__)

def initialize_cloud_environment():
    """Cloud Run環境初期化"""
    global cloud_db_manager, config, gmail_sender
    
    try:
        # Cloud Storage データベース管理
        cloud_db_manager = CloudDatabaseManager()
        cloud_db_manager.download_db()
        
        # 設定初期化
        config = CloudRunConfig()
        
        # Gmail送信者初期化
        gmail_sender = CloudRunGmailSender()
        
        print("Cloud Run環境初期化完了")
        return True
    except Exception as e:
        print(f"Cloud Run環境初期化失敗: {e}")
        return False

def cleanup_cloud_environment():
    """Cloud Run環境クリーンアップ"""
    global cloud_db_manager, config
    
    try:
        # データベースアップロード
        if cloud_db_manager:
            cloud_db_manager.upload_db()
            cloud_db_manager.cleanup()
        
        # 一時ファイル削除
        if config:
            config.cleanup()
        
        print("Cloud Run環境クリーンアップ完了")
    except Exception as e:
        print(f"クリーンアップエラー: {e}")

# CORS設定
@app.after_request
def after_request(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type,Authorization'
    response.headers['Access-Control-Allow-Methods'] = 'GET,PUT,POST,DELETE,OPTIONS'
    response.headers['Access-Control-Max-Age'] = '86400'
    return response

# HTMLファイル配信
@app.route('/')
def index():
    return send_from_directory('.', 'keikyuuLine2.html')

@app.route('/script.js')
def script():
    return send_from_directory('.', 'script.js')

@app.route('/<path:filename>')
def serve_file(filename):
    return send_from_directory('.', filename)

@app.route('/test')
def test():
    return jsonify({"message": "Cloud Run server is running!", "status": "OK"})

@app.route('/health')
def health_check():
    """ヘルスチェック"""
    return jsonify({
        "status": "healthy",
        "platform": "Google Cloud Run",
        "gmail_ready": gmail_sender.service is not None if gmail_sender else False,
        "storage_ready": cloud_db_manager is not None
    })

# メイン処理API
@app.route('/receive', methods=['POST', 'OPTIONS'])
def receive_and_scrape_data():
    if request.method == 'OPTIONS':
        return jsonify({})
    
    start_time = time.time()
    application_id = None
    
    try:
        # Cloud Run環境初期化
        if not initialize_cloud_environment():
            return jsonify({"error": "環境初期化失敗"}), 500
        
        # リクエストデータ解析
        request_data = request.get_json()
        if not request_data:
            return jsonify({"error": "Invalid JSON or no data"}), 400

        email = request_data.get('email', '')
        num_pages = int(request_data.get('page', 3))
        stations = request_data.get('stations', ["三軒茶屋"])
        urls = request_data.get('urls', [
            "https://suumo.jp/jj/chintai/ichiran/FR301FC005/?ar=030&bs=040&ra=013&rn=0230&ek=023016720&cb=0.0&ct=9999999&mb=0&mt=9999999&et=9999999&cn=9999999&shkr1=03&shkr2=03&shkr3=03&shkr4=03&sngz=&po1=25&po2=99&pc=100&page="
        ])

        print(f"Cloud Run処理開始: {len(stations)}駅, {num_pages}ページ")

        # データベース管理
        db_manager = CloudRunDatabaseManager(cloud_db_manager.local_db_path)
        
        # 申し込み情報保存
        request_timestamp = Config.get_timestamp()
        application_id = db_manager.save_application(email, num_pages, stations, urls, request_timestamp)
        
        # 受付確認メール送信
        reception_email_sent = gmail_sender.send_reception_email(email, num_pages, stations, urls)
        
        # コンポーネント初期化
        os.makedirs(config.FOLDER_PATH, exist_ok=True)
        scraper = PropertyScraper()
        processor = DataProcessor()
        stats_calc = StatisticsCalculator()
        visualizer = Visualizer()
        ppt_gen = PowerPointGenerator()
        file_manager = FileManager(config.FOLDER_PATH)
        summary_gen = SummaryGenerator()
        
        station_processor = StationProcessor(scraper, processor, stats_calc, 
                                           visualizer, ppt_gen, file_manager)
        
        # 各駅処理
        total_scraped = 0
        all_station_data = []
        processed_stations = []
        
        for station, base_url in zip(stations, urls):
            try:
                print(f"=== {station}駅の処理開始 ===")
                station_start_time = time.time()
                result = station_processor.process_station(station, base_url, num_pages)
                processing_time = time.time() - station_start_time
                
                if result:
                    all_station_data.append(result)
                    total_scraped += result['count']
                    processed_stations.append(station)
                    
                    from datetime import datetime
                    datetime_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                    csv_files = [f"1fData_{station}_{datetime_str}.csv"]
                    ppt_file = f"1e_{station}_{datetime_str}_ptt1.pptx"
                    
                    db_manager.save_station_result(
                        application_id, station, result['count'], 
                        csv_files, ppt_file, processing_time
                    )
                    
                    print(f"=== {station}駅の処理完了（{result['count']}件） ===")
                
            except Exception as e:
                print(f"{station}: 処理中にエラー: {e}")
                db_manager.save_station_result(application_id, station, 0, None, None, None)
        
        # 総合まとめ作成
        if all_station_data:
            try:
                summary_gen.create_summary(config.FOLDER_PATH)
                print("総合まとめ作成完了")
            except Exception as e:
                print(f"総合まとめ作成エラー: {e}")
        
        # 結果メール送信
        result_email_sent = gmail_sender.send_result_email(email, config.FOLDER_PATH, processed_stations)
        
        # データベース結果更新
        db_manager.update_application_result(
            application_id, "cloud_run_processed", total_scraped, len(all_station_data)
        )
        
        processing_time = time.time() - start_time
        
        return jsonify({
            "message": "Cloud Run処理完了！",
            "application_id": application_id,
            "page": num_pages,
            "email": email,
            "stations": stations,
            "processed_stations": processed_stations,
            "scraped_stations": len(all_station_data),
            "total_scraped_items": total_scraped,
            "processing_time": round(processing_time, 2),
            "reception_email_sent": reception_email_sent,
            "result_email_sent": result_email_sent,
            "platform": "Google Cloud Run"
        })
        
    except Exception as e:
        error_message = str(e)
        print(f"処理エラー: {error_message}")
        
        if application_id and cloud_db_manager:
            db_manager = CloudRunDatabaseManager(cloud_db_manager.local_db_path)
            db_manager.update_application_error(application_id, error_message)
        
        return jsonify({
            "error": f"処理中にエラーが発生しました: {error_message}",
            "application_id": application_id,
            "platform": "Google Cloud Run"
        }), 500
    
    finally:
        # 必ずクリーンアップ実行
        cleanup_cloud_environment()

# 統計・履歴API
@app.route('/stats')
def get_stats():
    """統計情報取得"""
    try:
        temp_db_manager = CloudDatabaseManager()
        temp_db_manager.download_db()
        
        db_manager = CloudRunDatabaseManager(temp_db_manager.local_db_path)
        stats = db_manager.get_processing_stats()
        
        temp_db_manager.cleanup()
        return jsonify(stats)
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/history')
def get_history():
    """履歴取得"""
    try:
        email = request.args.get('email')
        limit = int(request.args.get('limit', 50))
        
        temp_db_manager = CloudDatabaseManager()
        temp_db_manager.download_db()
        
        db_manager = CloudRunDatabaseManager(temp_db_manager.local_db_path)
        history = db_manager.get_application_history(email, limit)
        
        temp_db_manager.cleanup()
        
        return jsonify([{
            'id': row[0],
            'email': row[1],
            'page': row[2],
            'stations': row[3],
            'urls': row[4],
            'request_timestamp': row[5],
            'processing_status': row[6],
            'output_folder': row[7],
            'total_scraped_items': row[8],
            'scraped_stations': row[9],
            'created_at': row[10],
            'updated_at': row[11]
        } for row in history])
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    print("=" * 60)
    print("Google Cloud Run 対応サーバー起動中...")
    print("=" * 60)
    print("1. Cloud Storage バケット設定: STORAGE_BUCKET環境変数")
    print("2. Gmail認証: GOOGLE_APPLICATION_CREDENTIALS環境変数")
    print("3. データベース: Cloud Storage上のapply.db")
    print("4. 一時ファイル: 処理後自動削除")
    print("5. 同時実行制御: Cloud Run設定で制御")
    print("=" * 60)
    
    # Cloud Run用のポート設定
    port = int(os.environ.get('PORT', 8080))
    app.run(debug=False, host='0.0.0.0', port=port)