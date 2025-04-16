"""
xlwings-rpcのチャートタイプマッピング機能をテストするスクリプト
"""
import json
import sys
import requests
import time

# JSON-RPCリクエストの作成
def create_request(method, params=None, request_id=1):
    return {
        "jsonrpc": "2.0",
        "method": method,
        "params": params,
        "id": request_id
    }

# リクエストの送信
def send_request(url, method, params=None, request_id=1):
    payload = create_request(method, params, request_id)
    response = requests.post(url, json=payload)
    return response.json()

def test_chart_types():
    """
    異なるチャートタイプのテスト
    """
    rpc_url = "http://127.0.0.1:8000/rpc"
    
    print(f"プラットフォーム: {sys.platform}")
    
    # 1. アプリケーションの作成
    app_response = send_request(rpc_url, "app.create", {"visible": True, "add_book": True}, 1)
    if "error" in app_response:
        print(f"Error creating app: {app_response['error']}")
        return
    
    app_pid = app_response["result"]["id"]
    print(f"アプリケーション作成: PID={app_pid}")
    
    # 2. テスト用データの設定
    time.sleep(1)  # アプリケーションが起動するのを待つ
    book_response = send_request(rpc_url, "book.list", {"pid": app_pid}, 2)
    
    if "error" in book_response:
        print(f"Error getting book list: {book_response['error']}")
        return
        
    book_name = book_response["result"][0]["name"]
    
    sheet_response = send_request(rpc_url, "sheet.list", {"book": book_name, "pid": app_pid}, 3)
    
    if "error" in sheet_response:
        print(f"Error getting sheet list: {sheet_response['error']}")
        return
        
    sheet_name = sheet_response["result"][0]["name"]
    
    # テストデータの作成
    data_response = send_request(rpc_url, "range.set_value", {
        "book": book_name,
        "sheet": sheet_name,
        "address": "A1:B5",
        "value": [
            ["カテゴリ", "値"],
            ["A", 10],
            ["B", 25],
            ["C", 15],
            ["D", 30]
        ],
        "pid": app_pid
    }, 4)
    
    if "error" in data_response:
        print(f"Error setting data: {data_response['error']}")
        return
        
    print(f"テストデータの設定: {book_name}, {sheet_name}")
    
    # 3. チャートの作成
    chart_response = send_request(rpc_url, "chart.add", {
        "book": book_name,
        "sheet": sheet_name,
        "left": 100,
        "top": 100,
        "width": 400,
        "height": 300,
        "pid": app_pid
    }, 5)
    
    if "error" in chart_response:
        print(f"Error creating chart: {chart_response['error']}")
        return
    
    chart_name = chart_response["result"]["name"]
    print(f"チャート作成: {chart_name}")
    
    # 4. データソースの設定
    source_response = send_request(rpc_url, "chart.set_source_data", {
        "book": book_name,
        "sheet": sheet_name,
        "chart": chart_name,
        "range": "A1:B5",
        "pid": app_pid
    }, 6)
    
    if "error" in source_response:
        print(f"Error setting source data: {source_response['error']}")
        return
    
    print("データソース設定: 成功")
    
    # 5. 各チャートタイプをテスト
    chart_types = [
        "column",         # Windowsスタイル
        "column_clustered", # MacOSスタイル
        "bar",
        "line",
        "pie",
        "area",
        "scatter",
        "doughnut"
    ]
    
    results = {}
    
    for i, chart_type in enumerate(chart_types):
        print(f"\nテスト: chart_type='{chart_type}'")
        
        type_response = send_request(rpc_url, "chart.set_chart_type", {
            "book": book_name,
            "sheet": sheet_name,
            "chart": chart_name,
            "type": chart_type,
            "pid": app_pid
        }, 7 + i)
        
        if "error" in type_response:
            print(f"  エラー: {type_response['error']['message']}")
            results[chart_type] = {"success": False, "error": type_response['error']['message']}
        else:
            result_type = type_response["result"].get("chart_type", "不明")
            print(f"  成功: タイプ設定={result_type}")
            results[chart_type] = {"success": True, "result_type": result_type}
    
    # 6. 結果の出力
    print("\n=== テスト結果のサマリー ===")
    for chart_type, result in results.items():
        status = "✅ 成功" if result["success"] else "❌ 失敗"
        details = f"→ {result['result_type']}" if result["success"] else f"→ {result['error']}"
        print(f"{chart_type.ljust(20)}: {status} {details}")
    
    # 7. 終了
    print("\nアプリケーションを終了します...")
    quit_response = send_request(rpc_url, "app.quit", {"pid": app_pid, "save_changes": False}, 20)
    
    if "error" in quit_response:
        print(f"Error quitting app: {quit_response['error']}")
    else:
        print("アプリケーション終了: 成功")
    
    print("テスト完了")

if __name__ == "__main__":
    test_chart_types()
