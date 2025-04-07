# xlwings-rpc

xlwings-rpcは、[xlwings](https://www.xlwings.org/)の機能をJSON-RPC 2.0プロトコルを通じて提供するサーバーです。これにより、ローカルまたはリモートのクライアントからExcelワークブックを操作することが可能になります。

## 特徴

- **JSON-RPC 2.0**: シンプルで軽量なRPCプロトコルの採用
- **xlwings統合**: xlwingsの強力な機能をリモートから利用可能
- **バッチ処理**: 複数の操作をバッチで処理可能
- **型安全**: Pydanticによる入力検証と型チェック
- **拡張性**: 簡単に拡張可能なモジュラー設計
- **非同期サポート**: 非同期処理により効率的なリクエスト処理

## インストール

```bash
pip install xlwings-rpc
```

または、ソースコードから直接インストール:

```bash
git clone https://github.com/yourusername/xlwings-rpc.git
cd xlwings-rpc
pip install -e .
```

## 使い方

### サーバーの起動

```bash
# デフォルト設定（localhost:8000）で起動
python -m xlwings_rpc

# ポートとホストを指定して起動
python -m xlwings_rpc --host 0.0.0.0 --port 5000
```

### Python クライアントからの使用例

```python
import json
import requests

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

# 使用例
rpc_url = "http://localhost:8000/rpc"

# アプリケーションの一覧を取得
apps = send_request(rpc_url, "app.list")
print(f"利用可能なExcelアプリケーション: {apps}")

# ワークブックを開く
book = send_request(rpc_url, "book.open", {"path": "/path/to/your/workbook.xlsx"})
print(f"開いたワークブック: {book}")

# シートの一覧を取得
sheets = send_request(rpc_url, "sheet.list", {"book": book["result"]["name"]})
print(f"シート一覧: {sheets}")

# A1:B5範囲の値を取得
values = send_request(rpc_url, "range.get_value", {
    "book": book["result"]["name"],
    "sheet": "Sheet1",
    "address": "A1:B5"
})
print(f"セル値: {values}")
```

### cURL からの使用例

基本的な cURL コマンド形式:

```bash
curl -X POST http://localhost:8000/rpc \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc": "2.0", "method": "メソッド名", "params": {パラメータ}, "id": 1}'
```

利用可能なExcelアプリケーションを取得:

```bash
curl -X POST http://localhost:8000/rpc \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc": "2.0", "method": "app.list", "id": 1}'
```

新しいExcelアプリケーションを作成:

```bash
curl -X POST http://localhost:8000/rpc \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc": "2.0", "method": "app.create", "params": {"visible": true, "add_book": true}, "id": 2}'
```

ワークブックを開く:

```bash
curl -X POST http://localhost:8000/rpc \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc": "2.0", "method": "book.open", "params": {"path": "/path/to/your/workbook.xlsx"}, "id": 3}'
```

シートの一覧を取得:

```bash
curl -X POST http://localhost:8000/rpc \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc": "2.0", "method": "sheet.list", "params": {"book": "Book1.xlsx"}, "id": 4}'
```

セル範囲の値を取得:

```bash
curl -X POST http://localhost:8000/rpc \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc": "2.0", "method": "range.get_value", "params": {"book": "Book1.xlsx", "sheet": "Sheet1", "address": "A1:B5"}, "id": 5}'
```

セル範囲に値を設定:

```bash
curl -X POST http://localhost:8000/rpc \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc": "2.0", "method": "range.set_value", "params": {"book": "Book1.xlsx", "sheet": "Sheet1", "address": "A1:B2", "value": [[1, 2], [3, 4]]}, "id": 6}'
```

## API 概要

xlwings-rpcは以下のカテゴリでメソッドを提供します:

- **app**: アプリケーションレベルの操作
  - `app.list`: 利用可能なExcelアプリケーションの一覧を取得
  - `app.create`: 新しいExcelアプリケーションを起動
  - `app.quit`: アプリケーションを終了

- **book**: ワークブックレベルの操作
  - `book.list`: 開いているワークブックの一覧を取得
  - `book.open`: ワークブックを開く
  - `book.close`: ワークブックを閉じる
  - `book.save`: ワークブックを保存

- **sheet**: シートレベルの操作
  - `sheet.list`: シートの一覧を取得
  - `sheet.add`: 新しいシートを追加
  - `sheet.delete`: シートを削除

- **range**: レンジレベルの操作
  - `range.get`: セル範囲の値を取得
  - `range.set_value`: セル範囲に値を設定
  - `range.clear`: セル範囲をクリア

## 依存関係

- Python 3.11以上
- xlwings 0.33.11以上
- FastAPI
- Uvicorn
- Pydantic

## ライセンス

このプロジェクトはMITライセンスの下で公開されています。
