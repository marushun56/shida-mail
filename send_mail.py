import os
from glob import glob
import pandas as pd
import win32com.client as win32
# import openpyxl # Excelファイルを扱うために内部的に使用されます
import re

# --- 設定項目 ---

# 0. 送信元メールアドレス（Outlookに追加済みのアカウント）
SENDER_EMAIL = 'maruyama.shun23@gmail.com'

# 1. 送信用のフォルダ名（このスクリプトと同じ場所に作成）
# フォルダ内にある Excel ファイルをすべて添付します（.xlsx/.xlsm/.xls）
ATTACH_DIR = 'to_send'

# 2. 宛先が書かれたCSVファイル名 (同じフォルダにあるファイル名を指定)
CSV_FILENAME = 'mail_list.csv'

# 3. CCに追加するメールアドレス（複数ある場合はセミコロン;で区切る）
CC_ADDRESSES = 'cc1@example.com; cc2@example.com'

# 4. メールの件名
MAIL_SUBJECT = '【更新のご連絡】加盟店一覧'

# 5. メールの本文 ('''で囲むと改行もそのまま反映されます)
MAIL_BODY = '''
各位

いつもお世話になっております。

加盟店一覧を更新いたしましたので、ご確認をお願いいたします。
添付ファイルをご確認ください。

何卒よろしくお願い申し上げます。
'''
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

        # CSVファイルを読み込む（ヘッダー有無を吸収、空白や無効行を除去）
        df = pd.read_csv(
            csv_path,
            header=None,            # まずはヘッダーなしで読む（先頭行がヘッダーでも後で自動スキップ）
            names=['name', 'email'],
            dtype=str,
            keep_default_na=False,  # 空文字はNaNにしない
            encoding='utf-8-sig',   # BOM付きUTF-8にも対応
            usecols=[0, 1]          # 末尾の余分な空列（末尾カンマ）を無視
        )
        # 前後空白を除去
        df['name'] = df['name'].str.strip()
        # 空白に加えて末尾のカンマやセミコロンなども削除
        df['email'] = df['email'].str.strip().str.rstrip(';,，、')

        # 先頭行がヘッダーぽい場合は落とす
        if len(df) > 0 and df.iloc[0]['name'].lower() == 'name' and df.iloc[0]['email'].lower() == 'email':
            df = df.iloc[1:, :]

        # 無効なメールを除外
        df = df[df['email'].apply(is_valid_email)]
        if df.empty:
            print("エラー: 有効なメールアドレスがCSVにありません。")
            return

        # Outlookアプリケーションを起動
        outlook = win32.Dispatch('outlook.application')
        
        print("メールの送信を開始します...")

        # CSVの各行に対してメールを作成・送信
        for index, row in df.iterrows():
            recipient_name = (row['name'] or '').strip()
            recipient_email = (row['email'] or '').strip()

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

            # CC（空ならスキップ）
            cc_list = [a.strip() for a in CC_ADDRESSES.split(';') if a.strip()] if CC_ADDRESSES else []
            for cc in cc_list:
                if is_valid_email(cc):
                    r = recips.Add(cc)
                    r.Type = 2  # CC

            # 宛先解決
            if not recips.ResolveAll():
                print(f"エラー: 宛先を解決できませんでした -> To: {recipient_email}, CC: {', '.join(cc_list)}")
                continue

            mail.Subject = MAIL_SUBJECT
            mail.Body = f"{recipient_name}様\n\n{MAIL_BODY}"
            # フォルダ内のExcelファイルをすべて添付
            for f in excel_files:
                mail.Attachments.Add(f)

            try:
                mail.Send()
                print(f"送信完了: {recipient_name}様 ({recipient_email})")
            except Exception as e:
                print(f"送信失敗: {recipient_name}様 ({recipient_email}) -> {e}")

        print("すべてのメールの送信が完了しました。")

    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == '__main__':
    main()
