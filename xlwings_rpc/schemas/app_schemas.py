"""
アプリケーション関連のスキーマ

JSON-RPCのApp関連メソッドのリクエスト・レスポンススキーマを定義します。
"""
from typing import List, Optional, Dict, Any
from pydantic import BaseModel, Field


class AppInfo(BaseModel):
    """アプリケーション情報"""
    id: int
    version: str
    visible: bool
    calculation: str
    screen_updating: bool
    display_alerts: bool


class AppList(BaseModel):
    """アプリケーション一覧"""
    apps: List[AppInfo]


class AppGetRequest(BaseModel):
    """app.getリクエストパラメータ"""
    pid: Optional[int] = None


class AppCreateRequest(BaseModel):
    """app.createリクエストパラメータ"""
    visible: bool = True
    add_book: bool = True


class AppQuitRequest(BaseModel):
    """app.quitリクエストパラメータ"""
    pid: int
    save_changes: bool = True


class AppSetCalculationRequest(BaseModel):
    """app.set_calculationリクエストパラメータ"""
    pid: int
    mode: str = Field(..., description="計算モード ('automatic', 'manual', 'semiautomatic')")


class AppGetCalculationRequest(BaseModel):
    """app.get_calculationリクエストパラメータ"""
    pid: int


class AppGetBooksRequest(BaseModel):
    """app.get_booksリクエストパラメータ"""
    pid: int
