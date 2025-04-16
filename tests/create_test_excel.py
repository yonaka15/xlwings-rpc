#!/usr/bin/env python
"""
テスト用のExcelファイルを作成するスクリプト
"""
import os
import xlwings as xw

def create_test_excel():
    # ファイルパスの設定
    file_path = os.path.join(os.path.dirname(__file__), 'test_file.xlsx')
    
    # 既存のファイルがあれば削除
    if os.path.exists(file_path):
        os.remove(file_path)
    
    # 新しいExcelアプリケーションとブックの作成
    app = xw.App(visible=False)
    wb = app.books.add()
    
    # シートの取得
    sheet = wb.sheets[0]
    sheet.name = 'Sheet1'
    
    # サンプルデータの設定
    # ヘッダー
    headers = ['科目', '大項目', '中項目', '小項目', '摘要', '単価', '数量', '単位', '金額', '税']
    sheet.range('A1').value = headers
    
    # サンプルデータ
    data = [
        ['旅費交通費', '出張費', '交通費', '新幹線', '東京-大阪往復', 14000, 1, '回', 14000, 0.1],
        ['旅費交通費', '出張費', '宿泊費', 'ホテル', '大阪出張宿泊', 8000, 2, '泊', 16000, 0.1],
        ['消耗品費', '事務用品', 'パソコン周辺機器', 'マウス', 'ワイヤレスマウス', 3000, 2, '個', 6000, 0.1],
        ['消耗品費', '事務用品', '文具', 'ペン', 'ボールペン10本セット', 500, 3, 'セット', 1500, 0.1],
        ['通信費', '電話', '携帯電話', '通話料', '4月分', 5000, 1, '月', 5000, 0.1]
    ]
    sheet.range('A2').value = data
    
    # 列幅の調整
    for i, width in enumerate([15, 15, 15, 15, 20, 10, 10, 10, 10, 10]):
        sheet.columns[i].width = width
    
    # ファイル保存
    wb.save(file_path)
    wb.close()
    app.quit()
    
    print(f"テストファイルが作成されました: {file_path}")
    return file_path

if __name__ == '__main__':
    create_test_excel()
