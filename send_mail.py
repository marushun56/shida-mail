import os
import pandas as pd
import win32com.client as win32
import openpyxl # Excelファイルを扱うために内部的に使用されます

# --- 設定項目 ---

# 1. 送信するExcelファイル名 (同じフォルダにあるファイル名を指定)
EXCEL_FILENAME = '加盟店一覧.xlsx'

# 2. 宛先が書かれたCSVファイル名 (同じフォルダにあるファイル名を指定)
CSV_FILENAME = 'mail_list.csv'

# 3. CCに追加するメールアドレス（複数ある場合はカンマで区切る）
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


def main():
    try:
        # スクリプトがあるフォルダの絶対パスを取得
        current_dir = os.path.dirname(os.path.abspath(__file__))
        
        # ファイルのフルパスを作成
        excel_path = os.path.join(current_dir, EXCEL_FILENAME)
        csv_path = os.path.join(current_dir, CSV_FILENAME)

        # ファイルの存在チェック
        if not os.path.exists(excel_path):
            print(f"エラー: Excelファイルが見つかりません: {excel_path}")
            return
        if not os.path.exists(csv_path):
            print(f"エラー: CSVファイルが見つかりません: {csv_path}")
            return

        # CSVファイルを読み込む（ヘッダーなしと仮定）
        df = pd.read_csv(csv_path, header=None, names=['name', 'email'])
        
        # Outlookアプリケーションを起動
        outlook = win32.Dispatch('outlook.application')
        
        print("メールの送信を開始します...")

        # CSVの各行に対してメールを作成・送信
        for index, row in df.iterrows():
            recipient_name = row['name']
            recipient_email = row['email']
            
            # メールアイテムを作成
            mail = outlook.CreateItem(0)
            
            # 宛先、CC、件名、本文を設定
            mail.To = recipient_email
            mail.CC = CC_ADDRESSES
            mail.Subject = MAIL_SUBJECT
            
            # 本文の最初に宛名を追加
            mail.Body = f"{recipient_name}様\n\n{MAIL_BODY}"
            
            # 添付ファイルを追加
            mail.Attachments.Add(excel_path)
            
            # メールを送信
            mail.Send()
            
            print(f"送信完了: {recipient_name}様 ({recipient_email})")

        print("すべてのメールの送信が完了しました。")

    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == '__main__':
    main()
