"""
JSON-RPCサーバーの実装

FastAPIを使用したJSON-RPC 2.0サーバーの実装を提供します。
"""
from typing import Dict, Any, List, Union, Optional
import asyncio
import logging
from fastapi import FastAPI, Request, Response, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import uvicorn

from xlwings_rpc.utils.errors import (
    PARSE_ERROR, INVALID_REQUEST, METHOD_NOT_FOUND,
    create_error_response, handle_exception
)
from xlwings_rpc.methods.app import AppMethods
from xlwings_rpc.methods.book import BookMethods
from xlwings_rpc.methods.sheet import SheetMethods
from xlwings_rpc.methods.range import RangeMethods


# ロガーの設定
logger = logging.getLogger(__name__)

# FastAPIアプリケーションの作成
app = FastAPI(
    title="xlwings-rpc",
    description="JSON-RPC 2.0 API for xlwings",
    version="0.1.0"
)

# CORS設定
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# メソッドディスパッチャー
method_dispatcher = {
    # App メソッド
    "app.list": AppMethods.list,
    "app.get": AppMethods.get,
    "app.create": AppMethods.create,
    "app.quit": AppMethods.quit,
    "app.set_calculation": AppMethods.set_calculation,
    "app.get_calculation": AppMethods.get_calculation,
    "app.get_books": AppMethods.get_books,
    
    # Book メソッド
    "book.list": BookMethods.list,
    "book.get": BookMethods.get,
    "book.open": BookMethods.open,
    "book.create": BookMethods.create,
    "book.close": BookMethods.close,
    "book.save": BookMethods.save,
    "book.get_sheets": BookMethods.get_sheets,
    
    # Sheet メソッド
    "sheet.list": SheetMethods.list,
    "sheet.get": SheetMethods.get,
    "sheet.add": SheetMethods.add,
    "sheet.delete": SheetMethods.delete,
    "sheet.rename": SheetMethods.rename,
    "sheet.clear": SheetMethods.clear,
    "sheet.get_used_range": SheetMethods.get_used_range,
    "sheet.activate": SheetMethods.activate,
    
    # Range メソッド
    "range.get": RangeMethods.get,
    "range.get_value": RangeMethods.get_value,
    "range.set_value": RangeMethods.set_value,
    "range.get_formula": RangeMethods.get_formula,
    "range.set_formula": RangeMethods.set_formula,
    "range.clear": RangeMethods.clear,
    "range.get_as_dataframe": RangeMethods.get_as_dataframe,
    "range.set_dataframe": RangeMethods.set_dataframe,
}


async def process_request(request_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    単一のJSON-RPCリクエストを処理します。
    
    Args:
        request_data: JSON-RPCリクエストオブジェクト
    
    Returns:
        JSON-RPCレスポンスオブジェクト
    """
    # リクエストの検証
    if not isinstance(request_data, dict):
        return create_error_response(INVALID_REQUEST, id=None)
    
    # 必須フィールドの確認
    if "jsonrpc" not in request_data or request_data["jsonrpc"] != "2.0" or "method" not in request_data:
        return create_error_response(INVALID_REQUEST, id=request_data.get("id"))
    
    # IDの取得（通知の場合はNone）
    request_id = request_data.get("id")
    method = request_data["method"]
    params = request_data.get("params", {})
    
    # メソッドの存在確認
    if method not in method_dispatcher:
        return create_error_response(METHOD_NOT_FOUND, id=request_id)
    
    # メソッドの実行
    try:
        handler = method_dispatcher[method]
        result = await handler(params) if params else await handler()
        
        # 通知の場合はレスポンスを返さない
        if request_id is None:
            return None
        
        # 正常レスポンスの作成
        return {
            "jsonrpc": "2.0",
            "result": result,
            "id": request_id
        }
    except Exception as e:
        # エラーをJSON-RPC形式に変換
        logger.exception(f"Error processing method {method}: {str(e)}")
        return handle_exception(e, request_id, include_traceback=True)


async def process_batch_request(batch_request: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    バッチリクエストを処理します。
    
    Args:
        batch_request: JSON-RPCリクエストオブジェクトのリスト
    
    Returns:
        JSON-RPCレスポンスオブジェクトのリスト
    """
    # 並列処理のためのタスク作成
    tasks = [process_request(req) for req in batch_request]
    
    # 全タスクを実行
    responses = await asyncio.gather(*tasks)
    
    # Noneの応答（通知）を除去
    return [r for r in responses if r is not None]


@app.post("/rpc")
async def handle_rpc(request: Request) -> Response:
    """
    JSON-RPC 2.0リクエストを処理するエンドポイント
    
    Args:
        request: FastAPIリクエストオブジェクト
    
    Returns:
        JSON-RPCレスポンス
    """
    try:
        # リクエストボディのパース
        request_data = await request.json()
        
        # リクエストの型に応じた処理
        if isinstance(request_data, list):
            # バッチリクエスト
            if not request_data:
                # 空配列はエラー
                response_data = create_error_response(INVALID_REQUEST, id=None)
            else:
                response_data = await process_batch_request(request_data)
                # レスポンスが空の場合は何も返さない
                if not response_data:
                    return Response(status_code=204)
        else:
            # 単一リクエスト
            response_data = await process_request(request_data)
            # 通知の場合は何も返さない
            if response_data is None:
                return Response(status_code=204)
        
        # レスポンスの返却
        return Response(
            content=str(response_data).replace("'", '"'),
            media_type="application/json"
        )
    except Exception as e:
        # JSONパースエラーなど
        logger.exception(f"Error processing RPC request: {str(e)}")
        response_data = create_error_response(PARSE_ERROR, id=None)
        return Response(
            content=str(response_data).replace("'", '"'),
            media_type="application/json"
        )


@app.get("/health")
async def health_check() -> Dict[str, str]:
    """
    ヘルスチェックエンドポイント
    
    Returns:
        ステータス情報
    """
    return {"status": "ok"}


def start_server(host: str = "127.0.0.1", port: int = 8000):
    """
    サーバーを起動します。
    
    Args:
        host: ホストアドレス
        port: ポート番号
    """
    uvicorn.run(app, host=host, port=port)


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="xlwings-rpc server")
    parser.add_argument("--host", default="127.0.0.1", help="Host address to bind")
    parser.add_argument("--port", type=int, default=8000, help="Port to bind")
    
    args = parser.parse_args()
    start_server(args.host, args.port)
