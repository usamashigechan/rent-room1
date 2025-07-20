from gmail_sender import GmailSender

def test_send():
    # Gmail送信テスト
    sender = GmailSender()
    
    # メール送信
    success = sender.send_email(
        to_email="uji.shigetaka1@gmail.com",     # 宛先
        subject="テストメール",           # 件名
        body="これはテストメールです。"    # 本文
    )
    
    if success:
        print("テスト成功!")
    else:
        print("テスト失敗...")

if __name__ == "__main__":
    test_send()