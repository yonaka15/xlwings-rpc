"""
JSON-RPC 2.0 基本スキーマ

JSON-RPC 2.0プロトコルのリクエストとレスポンスのスキーマを定義します。
"""
from typing import Any, Dict, List, Optional, Union
from pydantic import BaseModel, Field


class JsonRpcRequest(BaseModel):
    """JSON-RPC 2.0リクエスト"""
    jsonrpc: str = Field("2.0", const=True)
    method: str
    params: Optional[Union[Dict[str, Any], List[Any]]] = None
    id: Optional[Union[str, int]] = None


class JsonRpcNotification(BaseModel):
    """JSON-RPC 2.0通知（ID無しのリクエスト）"""
    jsonrpc: str = Field("2.0", const=True)
    method: str
    params: Optional[Union[Dict[str, Any], List[Any]]] = None


class JsonRpcError(BaseModel):
    """JSON-RPC 2.0エラー"""
    code: int
    message: str
    data: Optional[Any] = None


class JsonRpcResponse(BaseModel):
    """JSON-RPC 2.0レスポンス"""
    jsonrpc: str = Field("2.0", const=True)
    result: Optional[Any] = None
    error: Optional[JsonRpcError] = None
    id: Optional[Union[str, int]] = None
