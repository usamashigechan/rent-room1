# main_flask_app.py 
import os
os.environ['MPLBACKEND'] = 'Agg'
import threading
import time
import pickle
import base64
import glob
import mimetypes
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

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
from database_manager import DatabaseManager

# Gmail送信クラス
class GmailSender:
    def __init__(self):
        self.SCOPES = ['https://www.googleapis.com/auth/gmail.send']
        self.service = self.authenticate()
    
    def authenticate(self):
        """OAuth2.0認証"""
        creds = None
        
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
            print("Gmail認証成功")
            return service
        except Exception as e:
            print(f"Gmail認証エラー: {e}")
            return None
    
    def get_station_from_filename(self, filename):
        """ファイル名から駅名を抽出"""
        # ファイル名パターンの例: 
        # 1fData_三軒茶屋_20250719_143022.csv
        # 1e_三軒茶屋_20250719_143022_ptt1.pptx
        patterns = [
            r'1fData_(.+?)_\d{8}_\d{6}\.csv',           # 新形式CSVファイル
            r'1fData_(.+?)_\d{8}\.csv',                 # 旧形式CSVファイル
            r'1e_(.+?)_\d{8}_\d{6}_ptt1\.pptx',         # 新形式PowerPointファイル
            r'1e_(.+?)_\d{14}_ptt1\.pptx',              # 旧形式PowerPointファイル
            r'1[a-z]_(.+?)_\d{8}_\d{6}',               # その他の新形式ファイル
            r'1[a-z]_(.+?)_\d{8,14}',                  # その他の旧形式ファイル
            r'1[a-z]_(.+?)\.', 	                        # 汎用パターン
        ]
        
        for pattern in patterns:
            match = re.search(pattern, filename)
            if match:
                station_name = match.group(1)
                print(f"駅名抽出成功: {filename} → {station_name}")
                return station_name
        
        # パターンにマッチしない場合はファイル名から手動抽出を試行
        if '_' in filename:
            parts = filename.split('_')
            if len(parts) >= 2 and parts[0].startswith('1'):
                potential_station = parts[1]
                print(f"駅名抽出（フォールバック）: {filename} → {potential_station}")
                return potential_station
        
        print(f"駅名抽出失敗: {filename}")
        return filename
    
    def create_readable_filename(self, original_filename, station_name=None):
        """読みやすいファイル名を作成"""
        # 拡張子を取得
        base_name, ext = os.path.splitext(original_filename)
        
        # ファイル名から駅名を抽出（引数で指定されていない場合）
        if not station_name:
            station_name = self.get_station_from_filename(original_filename)
        
        # ファイルタイプを判定
        if original_filename.startswith('1fData_'):
            file_type = "物件データ"
        elif original_filename.startswith('1e_') and 'ptt1' in original_filename:
            file_type = "分析レポート"
        elif original_filename.startswith('1g_'):
            file_type = "統計グラフ"
        elif original_filename.startswith('1h_'):
            file_type = "まとめ資料"
        else:
            file_type = "結果ファイル"
        
        # 新しいファイル名を作成（駅名を確実に含める）
        if station_name and station_name != original_filename and station_name.strip():
            readable_name = f"{station_name}_{file_type}{ext}"
        else:
            # 駅名が抽出できない場合は元のファイル名から時刻部分を除去
            if '_' in base_name:
                parts = base_name.split('_')
                if len(parts) >= 2:
                    extracted_station = parts[1]  # 1fData_駅名_datetime の駅名部分
                    readable_name = f"{extracted_station}_{file_type}{ext}"
                else:
                    readable_name = f"{file_type}{ext}"
            else:
                readable_name = f"{file_type}{ext}"
        
        print(f"ファイル名変換:")
        print(f"  元ファイル名: {original_filename}")
        print(f"  抽出駅名: {station_name}")
        print(f"  変換後: {readable_name}")
        return readable_name
    
    def create_message(self, to_email, subject, body, attachments=None):
        """メッセージの作成（元のファイル名をそのまま使用版）"""
        message = MIMEMultipart()
        message['to'] = to_email
        message['subject'] = subject
        message.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # 添付ファイルの追加
        if attachments:
            for i, file_path in enumerate(attachments):
                if os.path.exists(file_path):
                    original_filename = os.path.basename(file_path)
                    print(f"添付ファイル {i+1} 処理中:")
                    print(f"  ファイル名: {original_filename}")
                    
                    # MIMEタイプの判定
                    content_type, encoding = mimetypes.guess_type(file_path)
                    if content_type is None or encoding is not None:
                        content_type = 'application/octet-stream'
                    
                    main_type, sub_type = content_type.split('/', 1)
                    
                    # ファイルの読み込み
                    with open(file_path, 'rb') as attachment_file:
                        if main_type == 'text':
                            part = MIMEText(attachment_file.read().decode('utf-8'), _subtype=sub_type)
                        else:
                            part = MIMEBase(main_type, sub_type)
                            part.set_payload(attachment_file.read())
                            encoders.encode_base64(part)
                    
                    # 元のファイル名をそのまま使用（複数の方法で確実に）
                    try:
                        # 方法1: RFC 2231形式（推奨）
                        part.add_header(
                            'Content-Disposition',
                            'attachment',
                            filename=('utf-8', '', original_filename)
                        )
                    except:
                        try:
                            # 方法2: 標準的な方法
                            encoded_filename = Header(original_filename, 'utf-8').encode()
                            part.add_header('Content-Disposition', f'attachment; filename="{encoded_filename}"')
                        except:
                            # 方法3: フォールバック（ASCIIセーフ）
                            safe_filename = re.sub(r'[^\w\-_\.]', '_', original_filename)
                            part.add_header('Content-Disposition', f'attachment; filename="{safe_filename}"')
                    
                    message.attach(part)
                    print(f"  添付ファイル追加完了: {original_filename}")
                    
                else:
                    print(f"添付ファイルが見つかりません: {file_path}")
        
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
        return {'raw': raw_message}
    
    def send_email(self, to_email, subject, body, attachments=None):
        """メール送信"""
        print(f"*** send_email 呼び出し ***")
        print(f"  宛先: {to_email}")
        print(f"  件名: {subject}")
        print(f"  添付ファイル数: {len(attachments) if attachments else 0}")
        
        if not self.service:
            print("Gmail認証が完了していません")
            return False
        
        try:
            print("メッセージ作成中...")
            message = self.create_message(to_email, subject, body, attachments)
            
            print("Gmail API でメール送信中...")
            result = self.service.users().messages().send(userId='me', body=message).execute()
            
            print(f"*** メール送信成功 ***")
            print(f"  メッセージID: {result['id']}")
            print(f"  送信先: {to_email}")
            print(f"  件名: {subject}")
            return True
            
        except HttpError as error:
            print(f"Gmail API エラー: {error}")
            return False
        except Exception as e:
            print(f"メール送信エラー: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def send_reception_email(self, email, page, stations, urls):
        """受付確認メールの送信"""
        subject = "受け付けました"
        stations_str = ', '.join(stations) if isinstance(stations, list) else str(stations)
        body = f"""{email}の方、{page}ページ、{stations_str}駅、{urls}の内容で受け付けました。しばらくしたら返信メールを送信されるのでお待ち下さい。遅い場合はエラーの可能性があります。"""
        
        return self.send_email(email, subject, body)
    
    def send_result_email(self, email, folder_path, stations_processed=None):
        """結果送信メールの送信（指定フォルダの1から始まるファイルを添付）"""
        subject = "お問い合わせ結果"
        
        # 処理した駅名を含む本文作成
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
        
        # 指定フォルダ内の1から始まるファイルを取得
        attachments = []
        
        if os.path.exists(folder_path):
            pattern = os.path.join(folder_path, "1*")
            files = glob.glob(pattern)
            
            # ファイルをソートして順序を安定化
            files.sort()
            
            for file_path in files:
                if os.path.isfile(file_path):
                    file_size = os.path.getsize(file_path)
                    filename = os.path.basename(file_path)
                    station = self.get_station_from_filename(filename)
                    print(f"添付対象ファイル: {filename} ({file_size} bytes) - 駅名: {station}")
                    attachments.append(file_path)
        
        print(f"結果メール送信:")
        print(f"  フォルダ: {folder_path}")
        print(f"  添付ファイル数: {len(attachments)}")
        print(f"  処理した駅: {stations_processed}")
        
        return self.send_email(email, subject, body, attachments)

# Flaskアプリの設定
app = Flask(__name__)

# Gmail送信クラスのインスタンス作成
gmail_sender = GmailSender()

# CORS設定
@app.after_request
def after_request(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type,Authorization'
    response.headers['Access-Control-Allow-Methods'] = 'GET,PUT,POST,DELETE,OPTIONS'
    response.headers['Access-Control-Max-Age'] = '86400'
    return response

# HTMLファイルを配信するルート
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
    return jsonify({"message": "Flask server is running!", "status": "OK"})

# Gmail認証状況確認
@app.route('/gmail-status')
def gmail_status():
    has_credentials = os.path.exists('credentials.json')
    has_token = os.path.exists('token.pickle')
    gmail_ready = gmail_sender.service is not None
    
    return jsonify({
        "credentials_json": has_credentials,
        "token_pickle": has_token,
        "gmail_authenticated": gmail_ready,
        "ready_to_send": gmail_ready
    })

# メールテスト用API
@app.route('/test-email')
def test_email():
    test_email_addr = request.args.get('email')
    if not test_email_addr:
        return jsonify({"error": "email parameter required"}), 400
    
    success = gmail_sender.send_email(
        test_email_addr,
        "Gmail送信テスト",
        "これはGmail送信のテストメールです。正常に動作しています。"
    )
    
    return jsonify({
        "message": "メールテスト完了",
        "success": success,
        "test_email": test_email_addr
    })

# 統計情報API
@app.route('/stats')
def get_stats():
    db_manager = DatabaseManager()
    stats = db_manager.get_processing_stats()
    return jsonify(stats)

# 履歴取得API
@app.route('/history')
def get_history():
    email = request.args.get('email')
    limit = int(request.args.get('limit', 50))
    
    db_manager = DatabaseManager()
    history = db_manager.get_application_history(email, limit)
    
    return jsonify([{
        'id': row[0],
        'email': row[1],
        'page_count': row[2],
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

# メイン処理API（スクレイピング・メール送信統合版）
@app.route('/receive', methods=['POST', 'OPTIONS'])
def receive_and_scrape_data():
    print(f"Request method: {request.method}")
    
    if request.method == 'OPTIONS':
        print("Handling OPTIONS request")
        return jsonify({})
    
    try:
        request_data = request.get_json()
        print(f"Received data: {request_data}")
    except Exception as e:
        print(f"JSON parse error: {e}")
        return jsonify({"error": f"JSON parse error: {str(e)}"}), 400
    
    if not request_data:
        return jsonify({"error": "Invalid JSON or no data"}), 400

    # リクエストデータの取得
    email = request_data.get('email', '')
    num_pages = int(request_data.get('page', 3))
    stations = request_data.get('stations', ["三軒茶屋"])
    urls = request_data.get('urls', [
        "https://suumo.jp/jj/chintai/ichiran/FR301FC005/?ar=030&bs=040&ra=013&rn=0230&ek=023016720&cb=0.0&ct=9999999&mb=0&mt=9999999&et=9999999&cn=9999999&shkr1=03&shkr2=03&shkr3=03&shkr4=03&sngz=&po1=25&po2=99&pc=100&page="
    ])

    print(f"スクレイピング開始:")
    print(f"  Page数: {num_pages}")
    print(f"  駅リスト: {stations}")
    print(f"  Email: {email}")

    # データベースマネージャーの初期化
    db_manager = DatabaseManager()
    
    # 申し込み情報をデータベースに保存
    request_timestamp = Config.get_timestamp()
    application_id = db_manager.save_application(email, num_pages, stations, urls, request_timestamp)
    
    # 受付確認メールの送信（JSデータ受け取り直後、添付なし）
    print("受付確認メール送信中（添付なし）...")
    reception_email_sent = gmail_sender.send_reception_email(email, num_pages, stations, urls)
    
    try:
        # コンポーネント初期化
        config = Config()
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
        
        total_scraped = 0
        all_station_data = []
        processed_stations = []
        
        # 各駅の処理
        for station, base_url in zip(stations, urls):
            try:
                print(f"\n=== {station}駅の処理開始 ===")
                start_time = time.time()
                result = station_processor.process_station(station, base_url, num_pages)
                processing_time = time.time() - start_time
                
                if result:
                    all_station_data.append(result)
                    total_scraped += result['count']
                    processed_stations.append(station)
                    
                    # 駅別結果をデータベースに保存（datetime付きファイル名）
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
                print(f"{station}: 処理中にエラーが発生しました: {e}")
                db_manager.save_station_result(
                    application_id, station, 0, None, None, None
                )
        
        # 総合まとめ作成
        if all_station_data:
            print("\n=== 総合まとめ作成開始 ===")
            try:
                summary_gen.create_summary(config.FOLDER_PATH)
                print("=== 総合まとめ作成完了 ===")
            except Exception as e:
                print(f"総合まとめ作成エラー: {e}")
        
        # 結果送信メール（処理完了後、ファイル移動前）
        print("\n=== 結果送信メール送信開始 ===")
        print(f"Gmail認証状況: {gmail_sender.service is not None}")
        print(f"処理した駅: {processed_stations}")
        print(f"出力フォルダ: {config.FOLDER_PATH}")
        
        # フォルダ内容の詳細確認
        if os.path.exists(config.FOLDER_PATH):
            all_files = os.listdir(config.FOLDER_PATH)
            print(f"フォルダ内の全ファイル({len(all_files)}個):")
            for f in all_files:
                file_path = os.path.join(config.FOLDER_PATH, f)
                if os.path.isfile(file_path):
                    file_size = os.path.getsize(file_path)
                    if f.startswith(('1c', '1e', '1f')):
                        print(f"  ★送信対象: {f} ({file_size} bytes)")
                    else:
                        print(f"  ×除外: {f} ({file_size} bytes)")
                else:
                    print(f"  (フォルダ): {f}")
        else:
            print(f"ERROR: 出力フォルダが存在しません: {config.FOLDER_PATH}")
        
        try:
            print("結果メール送信実行中...")
            result_email_sent = gmail_sender.send_result_email(email, config.FOLDER_PATH, processed_stations)
            print(f"結果メール送信結果: {result_email_sent}")
            
            if result_email_sent:
                print("SUCCESS: 結果メール送信成功")
            else:
                print("ERROR: 結果メール送信失敗")
                
        except Exception as e:
            print(f"ERROR: 結果メール送信中に例外発生: {e}")
            import traceback
            traceback.print_exc()
            result_email_sent = False
        
        print("=== 結果送信メール送信完了 ===\n")
        
        # ファイル移動（A+timestampフォルダに）
        print("\n=== ファイル移動開始 ===")
        datestamp = config.get_datestamp()
        dest_folder = file_manager.create_timestamped_folder(datestamp)
        
        print(f"ファイルの移動が完了しました。移動先: {dest_folder}")
        
        # データベースの処理結果を更新
        db_manager.update_application_result(
            application_id, dest_folder, total_scraped, len(all_station_data)
        )
        
        return jsonify({
            "message": "全処理完了！スクレイピング、統計処理、グラフ作成、PowerPoint作成、メール送信、ファイル移動が完了しました。",
            "application_id": application_id,
            "page": num_pages,
            "email": email,
            "stations": stations,
            "processed_stations": processed_stations,
            "scraped_stations": len(all_station_data),
            "total_scraped_items": total_scraped,
            "output_folder": dest_folder,
            "database_saved": True,
            "database_path": "C:\\111accommodationDB\\apply.db",
            "reception_email_sent": reception_email_sent,
            "result_email_sent": result_email_sent,
            "email_method": "Gmail API (OAuth2.0)"
        })
        
    except Exception as e:
        # エラー発生時のデータベース更新
        error_message = str(e)
        db_manager.update_application_error(application_id, error_message)
        
        return jsonify({
            "error": f"処理中にエラーが発生しました: {error_message}",
            "application_id": application_id,
            "database_saved": True,
            "reception_email_sent": reception_email_sent if 'reception_email_sent' in locals() else False,
            "email_method": "Gmail API (OAuth2.0)"
        }), 500

if __name__ == '__main__':
    print("=" * 60)
    print("完全統合サーバー起動中（地名入り添付ファイル名対応版）...")
    print("=" * 60)
    print("1. HTMLファイル 'keikyuuLine2.html' をこのPythonファイルと同じフォルダに置いてください")
    print("2. ブラウザで http://localhost:5000 にアクセスしてください")
    print("3. テスト用URL: http://localhost:5000/test")
    print("4. Gmail認証状況: http://localhost:5000/gmail-status")
    print("5. メールテスト: http://localhost:5000/test-email?email=test@example.com")
    print("6. 統計情報URL: http://localhost:5000/stats")
    print("7. 履歴情報URL: http://localhost:5000/history")
    print("8. 機能: スクレイピング → 統計処理 → グラフ作成 → PowerPoint作成 → データベース保存 → Gmail送信")
    print("9. データベース: C:\\111accommodationDB\\apply.db")
    print("10. 結果フォルダ: C:\\1111accommodation\\A{timestamp}")
    print("11. 新機能: 添付ファイル名に駅名を含める（例：三軒茶屋_物件データ.csv）")
    print("=" * 60)
    
    # Gmail認証状況チェック
    if not os.path.exists('credentials.json'):
        print("WARNING: credentials.json が見つかりません")
        print("Google Cloud Console からダウンロードしてください")
    
    if gmail_sender.service:
        print("OK: Gmail認証完了 - メール送信準備完了")
    else:
        print("WARNING: Gmail認証が必要です - 初回実行時にブラウザで認証してください")
    
    print("=" * 60)
    
    app.run(debug=True, host='0.0.0.0', port=5000)