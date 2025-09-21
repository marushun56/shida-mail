import os
from glob import glob
import pandas as pd
import win32com.client as win32
# import openpyxl # Excelファイルを扱うために内部的に使用されます
import re

# --- 設定項目 ---

# 0. 送信元メールアドレス（Outlookに追加済みのアカウント）
SENDER_EMAIL = ''

# 1. 送信用のフォルダ名（このスクリプトと同じ場所に作成）
# フォルダ内にある Excel ファイルをすべて添付します（.xlsx/.xlsm/.xls）
ATTACH_DIR = 'to_send'

# 2. 宛先が書かれたCSVファイル名 (同じフォルダにあるファイル名を指定)
# CSV形式: 送信先名,送信先メールアドレス,cc先名,cc先メールアドレス
CSV_FILENAME = 'mail_list.csv'

# 3. メール本文のファイル名 (同じフォルダにあるテキストファイル名を指定)
MAIL_BODY_FILE = 'mail_body.txt'

# 4. メールの件名
MAIL_SUBJECT = '【更新のご連絡】加盟店一覧'
# --- 設定はここまで ---


def is_valid_email(addr: str) -> bool:
    if not addr:
        return False
    addr = addr.strip()
    # 簡易チェック（必要なら強化してください）
    return re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", addr) is not None

def main():
    try:
        # スクリプトがあるフォルダの絶対パスを取得
        current_dir = os.path.dirname(os.path.abspath(__file__))
        
        # ファイル/フォルダのフルパスを作成
        attach_dir = os.path.join(current_dir, ATTACH_DIR)
        csv_path = os.path.join(current_dir, CSV_FILENAME)
        mail_body_path = os.path.join(current_dir, MAIL_BODY_FILE)

        # 添付フォルダの存在チェックとExcelファイル収集
        if not os.path.isdir(attach_dir):
            print(f"エラー: 添付フォルダが見つかりません: {attach_dir}")
            print("フォルダを作成し、送信したいExcelファイルを入れてください。")
            return
        # 対象拡張子のファイル一覧
        excel_files = []
        for pattern in ("*.xlsx", "*.xlsm", "*.xls"):
            excel_files.extend(glob(os.path.join(attach_dir, pattern)))
        if not excel_files:
            print(f"エラー: 添付フォルダにExcelファイルがありません: {attach_dir}")
            return
        else:
            print(f"添付対象ファイル: {len(excel_files)}件")
        if not os.path.exists(csv_path):
            print(f"エラー: CSVファイルが見つかりません: {csv_path}")
            return
        if not os.path.exists(mail_body_path):
            print(f"エラー: メール本文ファイルが見つかりません: {mail_body_path}")
            return

        # メール本文をファイルから読み込み
        with open(mail_body_path, 'r', encoding='utf-8') as f:
            mail_body_template = f.read().strip()

        # CSVファイルを読み込む（4列対応: 送信先名,送信先メール,cc先名,cc先メール）
        df = pd.read_csv(
            csv_path,
            header=None,            # まずはヘッダーなしで読む（先頭行がヘッダーでも後で自動スキップ）
            names=['to_name', 'to_email', 'cc_name', 'cc_email'],
            dtype=str,
            keep_default_na=False,  # 空文字はNaNにしない
            encoding='utf-8-sig',   # BOM付きUTF-8にも対応
            usecols=[0, 1, 2, 3]    # 4列を読み取り
        )
        # 前後空白を除去
        df['to_name'] = df['to_name'].str.strip()
        df['to_email'] = df['to_email'].str.strip().str.rstrip(';,，、')
        df['cc_name'] = df['cc_name'].str.strip()
        df['cc_email'] = df['cc_email'].str.strip().str.rstrip(';,，、')

        # 先頭行がヘッダーぽい場合は落とす
        if len(df) > 0 and df.iloc[0]['to_name'].lower() in ('name', 'to_name', '送信先名'):
            df = df.iloc[1:, :]

        # 無効なメールを除外（送信先は必須、CC先は任意）
        df = df[df['to_email'].apply(is_valid_email)]
        if df.empty:
            print("エラー: 有効な送信先メールアドレスがCSVにありません。")
            return

        # Outlookアプリケーションを起動
        outlook = win32.Dispatch('outlook.application')
        
        print("メールの送信を開始します...")

        # CSVの各行に対してメールを作成・送信
        for index, row in df.iterrows():
            recipient_name = (row['to_name'] or '').strip()
            recipient_email = (row['to_email'] or '').strip()
            cc_name = (row['cc_name'] or '').strip()
            cc_email = (row['cc_email'] or '').strip()

            mail = outlook.CreateItem(0)

            # 送信元アカウントを指定（Outlookに該当アカウントが追加されている必要があります）
            try:
                accounts = outlook.Session.Accounts
                for account in accounts:
                    if str(account.SmtpAddress).lower() == SENDER_EMAIL.lower():
                        mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))  # SendUsingAccount
                        break
                else:
                    print(f"警告: 送信元アカウント {SENDER_EMAIL} がOutlookに見つかりません。既定のアカウントから送信します。")
            except Exception as e:
                print(f"送信元アカウントの設定に失敗: {e}（既定のアカウントを使用）")

            # 宛先/CC をRecipients経由で追加し、解決できるか確認
            recips = mail.Recipients
            to_r = recips.Add(recipient_email)
            to_r.Type = 1  # To

            # CC（行ごとのCCメールがあれば追加）
            if cc_email and is_valid_email(cc_email):
                cc_r = recips.Add(cc_email)
                cc_r.Type = 2  # CC

            # 宛先解決
            if not recips.ResolveAll():
                cc_info = f" CC: {cc_email}" if cc_email else ""
                print(f"エラー: 宛先を解決できませんでした -> To: {recipient_email}{cc_info}")
                continue

            mail.Subject = MAIL_SUBJECT
            mail.Body = f"{recipient_name}様\n\n{mail_body_template}"
            # フォルダ内のExcelファイルをすべて添付
            for f in excel_files:
                mail.Attachments.Add(f)

            try:
                mail.Send()
                cc_info = f" (CC: {cc_name} <{cc_email}>)" if cc_email else ""
                print(f"送信完了: {recipient_name}様 ({recipient_email}){cc_info}")
            except Exception as e:
                print(f"送信失敗: {recipient_name}様 ({recipient_email}) -> {e}")

        print("すべてのメールの送信が完了しました。")

    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == '__main__':
    main()
