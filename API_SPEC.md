# xlwings-rpc API 仕様書

## 概要

xlwings-rpcは、xlwingsの機能をJSON-RPC 2.0プロトコルを通じて提供するサーバーです。このAPIを使用することで、ローカルまたはリモートのクライアントからExcelワークブックを操作することができます。

## 基本情報

- **エンドポイント**: `http://<ホスト>:<ポート>/rpc`
- **メソッド**: POST
- **コンテンツタイプ**: `application/json`
- **プロトコル**: JSON-RPC 2.0

## リクエスト形式

```json
{
  "jsonrpc": "2.0",
  "method": "メソッド名",
  "params": {
    // オプションのパラメータ
  },
  "id": リクエストID
}
```

- `jsonrpc`: 常に "2.0" を指定（必須）
- `method`: 実行するメソッド名（必須）
- `params`: メソッドに渡すパラメータ（オプション、メソッドによって異なる）
- `id`: クライアント側で設定する一意のリクエストID（必須、レスポンスと関連付けるため）

## レスポンス形式

**成功時:**

```json
{
  "jsonrpc": "2.0",
  "result": {
    // メソッドの結果
  },
  "id": リクエストID
}
```

**エラー時:**

```json
{
  "jsonrpc": "2.0",
  "error": {
    "code": エラーコード,
    "message": "エラーメッセージ",
    "data": {
      // オプションの追加情報
    }
  },
  "id": リクエストID
}
```

## エラーコード

| コード | 説明 |
|--------|------|
| -32700 | パースエラー（無効なJSONが送信された） |
| -32600 | 無効なリクエスト（JSONは正しいが、無効なリクエスト） |
| -32601 | メソッドが見つからない |
| -32602 | 無効なパラメータ |
| -32603 | 内部エラー |
| -32000 | Excelアプリケーションが見つからない |
| -32001 | ワークブックが見つからない |
| -32002 | シートが見つからない |
| -32003 | レンジエラー |
| -32004 | Excelエラー |
| -32005 | 権限エラー |
| -32006 | タイムアウトエラー |

## バッチリクエスト

複数のメソッドを一度に呼び出す場合は、リクエストオブジェクトの配列を送信します。

```json
[
  {
    "jsonrpc": "2.0",
    "method": "メソッド名1",
    "params": {},
    "id": 1
  },
  {
    "jsonrpc": "2.0",
    "method": "メソッド名2",
    "params": {},
    "id": 2
  }
]
```

レスポンスも配列形式で返されます。

## 利用可能なメソッド

### アプリケーション操作

#### app.list

すべての実行中のExcelアプリケーションを取得します。

**パラメータ**: なし

**戻り値**:
```json
[
  {
    "id": アプリケーションのPID,
    "version": "Excel バージョン",
    "visible": true/false,
    "calculation": "計算モード",
    "screen_updating": true/false,
    "display_alerts": true/false
  },
  ...
]
```

#### app.get

指定されたPIDまたはアクティブなExcelアプリケーションを取得します。

**パラメータ**:
```json
{
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
{
  "id": アプリケーションのPID,
  "version": "Excel バージョン",
  "visible": true/false,
  "calculation": "計算モード",
  "screen_updating": true/false,
  "display_alerts": true/false
}
```

#### app.create

新しいExcelアプリケーションを作成します。

**パラメータ**:
```json
{
  "visible": true/false, // デフォルト: true
  "add_book": true/false // デフォルト: true
}
```

**戻り値**:
```json
{
  "id": 新しいアプリケーションのPID,
  "version": "Excel バージョン",
  "visible": true/false,
  "calculation": "計算モード",
  "screen_updating": true/false,
  "display_alerts": true/false
}
```

#### app.quit

Excelアプリケーションを終了します。

**パラメータ**:
```json
{
  "pid": アプリケーションのPID,
  "save_changes": true/false // デフォルト: true
}
```

**注意**: `save_changes` パラメータは、内部的にブックを保存するために使用されますが、最新の xlwings バージョンでは `quit()` メソッド自体は引数を受け付けません。このパラメータを `true` に設定すると、アプリケーションの終了前に開いているすべてのブックが自動的に保存されます。

**戻り値**:
```json
true
```

#### app.set_calculation

計算モードを設定します。

**パラメータ**:
```json
{
  "pid": アプリケーションのPID,
  "mode": "automatic" | "manual" | "semiautomatic"
}
```

**戻り値**:
```json
{
  "id": アプリケーションのPID,
  "version": "Excel バージョン",
  "visible": true/false,
  "calculation": "計算モード",
  "screen_updating": true/false,
  "display_alerts": true/false
}
```

#### app.get_calculation

現在の計算モードを取得します。

**パラメータ**:
```json
{
  "pid": アプリケーションのPID
}
```

**戻り値**:
```json
"automatic" | "manual" | "semiautomatic"
```

#### app.get_books

指定されたアプリケーションで開いているワークブックを取得します。

**パラメータ**:
```json
{
  "pid": アプリケーションのPID
}
```

**戻り値**:
```json
[
  {
    "name": "ワークブック名",
    "fullname": "ワークブックのフルパス",
    "path": "ワークブックのパス",
    "app_id": アプリケーションのPID,
    "sheets": ["シート名1", "シート名2", ...]
  },
  ...
]
```

### ワークブック操作

#### book.list

すべての開いているワークブックを取得します。

**パラメータ**:
```json
{
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
[
  {
    "name": "ワークブック名",
    "fullname": "ワークブックのフルパス",
    "path": "ワークブックのパス",
    "app_id": アプリケーションのPID,
    "sheets": ["シート名1", "シート名2", ...]
  },
  ...
]
```

#### book.get

指定されたワークブックを取得します。

**パラメータ**:
```json
{
  "name": "ワークブック名",
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
{
  "name": "ワークブック名",
  "fullname": "ワークブックのフルパス",
  "path": "ワークブックのパス",
  "app_id": アプリケーションのPID,
  "sheets": ["シート名1", "シート名2", ...]
}
```

#### book.open

ワークブックを開きます。

**パラメータ**:
```json
{
  "path": "ワークブックのパス",
  "pid": アプリケーションのPID (オプション),
  "read_only": true/false (オプション),
  "password": "パスワード" (オプション)
}
```

**戻り値**:
```json
{
  "name": "ワークブック名",
  "fullname": "ワークブックのフルパス",
  "path": "ワークブックのパス",
  "app_id": アプリケーションのPID,
  "sheets": ["シート名1", "シート名2", ...]
}
```

#### book.create

新しいワークブックを作成します。

**パラメータ**:
```json
{
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
{
  "name": "ワークブック名",
  "fullname": "ワークブックのフルパス",
  "path": "ワークブックのパス",
  "app_id": アプリケーションのPID,
  "sheets": ["シート名1", "シート名2", ...]
}
```

#### book.close

ワークブックを閉じます。

**パラメータ**:
```json
{
  "name": "ワークブック名",
  "pid": アプリケーションのPID (オプション),
  "save": true/false (オプション, デフォルト: true)
}
```

**戻り値**:
```json
true
```

#### book.save

ワークブックを保存します。

**パラメータ**:
```json
{
  "name": "ワークブック名",
  "pid": アプリケーションのPID (オプション),
  "path": "保存先パス" (オプション)
}
```

**戻り値**:
```json
{
  "name": "ワークブック名",
  "fullname": "ワークブックのフルパス",
  "path": "ワークブックのパス",
  "app_id": アプリケーションのPID,
  "sheets": ["シート名1", "シート名2", ...]
}
```

#### book.get_sheets

ワークブック内のすべてのシートを取得します。

**パラメータ**:
```json
{
  "name": "ワークブック名",
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
[
  {
    "name": "シート名",
    "book_name": "ワークブック名",
    "index": シートのインデックス,
    "used_range": "使用範囲のアドレス"
  },
  ...
]
```

### シート操作

#### sheet.list

ワークブック内のすべてのシートを取得します。

**パラメータ**:
```json
{
  "book": "ワークブック名",
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
[
  {
    "name": "シート名",
    "book_name": "ワークブック名",
    "index": シートのインデックス,
    "used_range": "使用範囲のアドレス"
  },
  ...
]
```

#### sheet.get

特定のシートを取得します。

**パラメータ**:
```json
{
  "book": "ワークブック名",
  "name": "シート名",
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
{
  "name": "シート名",
  "book_name": "ワークブック名",
  "index": シートのインデックス,
  "used_range": "使用範囲のアドレス"
}
```

#### sheet.add

新しいシートを追加します。

**パラメータ**:
```json
{
  "book": "ワークブック名",
  "name": "シート名" (オプション),
  "before": "既存のシート名" (オプション),
  "after": "既存のシート名" (オプション),
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
{
  "name": "シート名",
  "book_name": "ワークブック名",
  "index": シートのインデックス,
  "used_range": "使用範囲のアドレス"
}
```

#### sheet.delete

シートを削除します。

**パラメータ**:
```json
{
  "book": "ワークブック名",
  "name": "シート名",
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
true
```

#### sheet.rename

シートの名前を変更します。

**パラメータ**:
```json
{
  "book": "ワークブック名",
  "name": "現在のシート名",
  "new_name": "新しいシート名",
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
{
  "name": "新しいシート名",
  "book_name": "ワークブック名",
  "index": シートのインデックス,
  "used_range": "使用範囲のアドレス"
}
```

#### sheet.clear

シートの内容をクリアします。

**パラメータ**:
```json
{
  "book": "ワークブック名",
  "name": "シート名",
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
true
```

#### sheet.get_used_range

シートの使用範囲を取得します。

**パラメータ**:
```json
{
  "book": "ワークブック名",
  "name": "シート名",
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
{
  "address": "範囲のアドレス",
  "sheet_name": "シート名",
  "book_name": "ワークブック名",
  "value": [["セル値"]],
  "formula": [["セル数式"]],
  "shape": [行数, 列数],
  "row": 開始行,
  "column": 開始列,
  "row_height": 行の高さ,
  "column_width": 列の幅
}
```

#### sheet.activate

シートをアクティブにします。

**パラメータ**:
```json
{
  "book": "ワークブック名",
  "name": "シート名",
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
{
  "name": "シート名",
  "book_name": "ワークブック名",
  "index": シートのインデックス,
  "used_range": "使用範囲のアドレス"
}
```

### レンジ操作

#### range.get

特定のセル範囲を取得します。

**パラメータ**:
```json
{
  "book": "ワークブック名",
  "sheet": "シート名",
  "address": "セル範囲のアドレス (例: 'A1:B10')",
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
{
  "address": "範囲のアドレス",
  "sheet_name": "シート名",
  "book_name": "ワークブック名",
  "value": [["セル値"]],
  "formula": [["セル数式"]],
  "shape": [行数, 列数],
  "row": 開始行,
  "column": 開始列,
  "row_height": 行の高さ,
  "column_width": 列の幅
}
```

#### range.get_value

セル範囲の値を取得します。

**パラメータ**:
```json
{
  "book": "ワークブック名",
  "sheet": "シート名",
  "address": "セル範囲のアドレス",
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
[
  ["セル値1", "セル値2", ...],
  ...
]
```

#### range.set_value

セル範囲に値を設定します。

**パラメータ**:
```json
{
  "book": "ワークブック名",
  "sheet": "シート名",
  "address": "セル範囲のアドレス",
  "value": 単一の値またはセル値の2次元配列,
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
true
```

#### range.get_formula

セル範囲の数式を取得します。

**パラメータ**:
```json
{
  "book": "ワークブック名",
  "sheet": "シート名",
  "address": "セル範囲のアドレス",
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
[
  ["数式1", "数式2", ...],
  ...
]
```

#### range.set_formula

セル範囲に数式を設定します。

**パラメータ**:
```json
{
  "book": "ワークブック名",
  "sheet": "シート名",
  "address": "セル範囲のアドレス",
  "formula": 単一の数式または数式の2次元配列,
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
true
```

#### range.clear

セル範囲をクリアします。

**パラメータ**:
```json
{
  "book": "ワークブック名",
  "sheet": "シート名",
  "address": "セル範囲のアドレス",
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
true
```

#### range.get_as_dataframe

セル範囲をpandas DataFrameとして取得します。

**パラメータ**:
```json
{
  "book": "ワークブック名",
  "sheet": "シート名",
  "address": "セル範囲のアドレス",
  "header": true/false (オプション, デフォルト: true),
  "index": true/false (オプション, デフォルト: false),
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
{
  "type": "dataframe",
  "index": ["インデックス値"],
  "columns": ["列名"],
  "data": [
    ["セル値1", "セル値2", ...],
    ...
  ]
}
```

#### range.set_dataframe

pandas DataFrameをセル範囲に設定します。

**パラメータ**:
```json
{
  "book": "ワークブック名",
  "sheet": "シート名",
  "address": "セル範囲のアドレス",
  "dataframe": {
    "type": "dataframe",
    "index": ["インデックス値"],
    "columns": ["列名"],
    "data": [
      ["セル値1", "セル値2", ...],
      ...
    ]
  },
  "header": true/false (オプション, デフォルト: true),
  "index": true/false (オプション, デフォルト: false),
  "pid": アプリケーションのPID (オプション)
}
```

**戻り値**:
```json
true
```

## 使用例

### cURL からの使用例

アプリケーションの一覧を取得:

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

## エラーハンドリング

エラーが発生した場合、JSONレスポンスにはエラーオブジェクトが含まれます。各エラーには、コード、メッセージ、および追加情報が含まれることがあります。

```json
{
  "jsonrpc": "2.0",
  "error": {
    "code": -32603,
    "message": "Internal error: エラーメッセージ",
    "data": {
      "traceback": "スタックトレース情報..."
    }
  },
  "id": 1
}
```

エラーが発生した場合は、エラーコードとメッセージを確認して、適切な対処を行ってください。

## 注意事項

1. MacOSでは一部のxlwingsの機能に制限がある場合があります。特に`calculation`プロパティなど、一部のプロパティにアクセスする際に`k.missing_value`エラーが発生する可能性があります。

2. サーバーがエラーを適切に処理するように設計されていますが、一部のケースではクライアント側での追加のエラーハンドリングが必要になる場合があります。

3. 大量のデータを扱う場合は、レスポンスサイズが大きくなる可能性があるため、適切な範囲を指定してください。

## ヘルスチェック

サーバーの状態を確認するには、`/health`エンドポイントを使用できます。

```bash
curl http://localhost:8000/health
```

正常な場合は、以下のJSONレスポンスが返されます。

```json
{
  "status": "ok"
}
```