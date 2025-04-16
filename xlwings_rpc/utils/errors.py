"""
エラーハンドリングユーティリティ

JSON-RPC 2.0のエラーオブジェクトの生成と、
xlwingsの例外からJSON-RPCエラーへの変換を提供します。
"""
from typing import Any, Dict, Optional, Type, Union
import traceback
import xlwings as xw

# JSON-RPC 2.0標準エラーコード
PARSE_ERROR = -32700
INVALID_REQUEST = -32600
METHOD_NOT_FOUND = -32601
INVALID_PARAMS = -32602
INTERNAL_ERROR = -32603

# xlwings-rpc特有のエラーコード (サーバーエラー範囲: -32000 to -32099)
EXCEL_NOT_FOUND = -32000
WORKBOOK_NOT_FOUND = -32001
SHEET_NOT_FOUND = -32002
RANGE_ERROR = -32003
EXCEL_ERROR = -32004
PERMISSION_ERROR = -32005
TIMEOUT_ERROR = -32006
CHART_NOT_FOUND = -32007
CHART_TYPE_ERROR = -32008

# エラーメッセージ
ERROR_MESSAGES = {
    PARSE_ERROR: "Parse error",
    INVALID_REQUEST: "Invalid Request",
    METHOD_NOT_FOUND: "Method not found",
    INVALID_PARAMS: "Invalid params",
    INTERNAL_ERROR: "Internal error",
    EXCEL_NOT_FOUND: "Excel application not found",
    WORKBOOK_NOT_FOUND: "Workbook not found",
    SHEET_NOT_FOUND: "Sheet not found",
    RANGE_ERROR: "Range error",
    EXCEL_ERROR: "Excel error",
    PERMISSION_ERROR: "Permission denied",
    TIMEOUT_ERROR: "Operation timed out",
    CHART_NOT_FOUND: "Chart not found",
    CHART_TYPE_ERROR: "Invalid chart type",
}


def create_error_response(
    code: int, 
    message: Optional[str] = None, 
    data: Any = None, 
    id: Optional[Union[str, int]] = None
) -> Dict[str, Any]:
    """
    JSON-RPC 2.0エラーレスポンスを生成します。

    Args:
        code: エラーコード
        message: エラーメッセージ (Noneの場合は標準メッセージを使用)
        data: 追加のエラーデータ (オプション)
        id: リクエストID (Noneの場合はIDなし)

    Returns:
        JSON-RPC 2.0形式のエラーレスポンス
    """
    if message is None:
        message = ERROR_MESSAGES.get(code, "Unknown error")

    error = {
        "code": code,
        "message": message
    }

    if data is not None:
        error["data"] = data

    response = {
        "jsonrpc": "2.0",
        "error": error,
        "id": id
    }

    return response


def handle_exception(
    exception: Exception, 
    id: Optional[Union[str, int]] = None, 
    include_traceback: bool = False
) -> Dict[str, Any]:
    """
    Pythonの例外をJSON-RPCエラーレスポンスに変換します。

    Args:
        exception: 発生した例外
        id: リクエストID
        include_traceback: トレースバックを含めるかどうか

    Returns:
        JSON-RPC 2.0形式のエラーレスポンス
    """
    # xlwings固有の例外を処理
    if isinstance(exception, xw.XlwingsError):
        code = EXCEL_ERROR
        message = str(exception)
    elif isinstance(exception, ConnectionError):
        code = EXCEL_NOT_FOUND
        message = "Failed to connect to Excel"
    elif isinstance(exception, FileNotFoundError):
        code = WORKBOOK_NOT_FOUND
        message = f"File not found: {str(exception)}"
    elif isinstance(exception, ValueError):
        # ValueErrorの内容に応じてエラーコードを決定
        message = str(exception)
        if "Chart" in message and "not found" in message:
            code = CHART_NOT_FOUND
        elif "chart type" in message:
            code = CHART_TYPE_ERROR
        else:
            code = INVALID_PARAMS
    elif isinstance(exception, PermissionError):
        code = PERMISSION_ERROR
        message = str(exception)
    elif isinstance(exception, TimeoutError):
        code = TIMEOUT_ERROR
        message = "Operation timed out"
    else:
        # その他の例外はInternal Errorとして処理
        code = INTERNAL_ERROR
        message = f"Internal error: {str(exception)}"

    # 追加データにトレースバックを含める
    data = None
    if include_traceback:
        data = {
            "traceback": traceback.format_exc()
        }

    return create_error_response(code, message, data, id)
