"""
データ変換ユーティリティ

xlwingsのオブジェクトとJSON-RPC間でのデータ変換を行います。
"""
from typing import Any, Dict, List, Union, Optional
import datetime
import logging
import xlwings as xw
import numpy as np
import pandas as pd


# ロガーの設定
logger = logging.getLogger(__name__)


def to_serializable(obj: Any) -> Any:
    """
    オブジェクトをJSONシリアライズ可能な形式に変換します。

    Args:
        obj: 変換するオブジェクト

    Returns:
        JSONシリアライズ可能なオブジェクト
    """
    # 基本型はそのまま返す
    if obj is None or isinstance(obj, (bool, int, float, str)):
        return obj
    
    # 日付型の変換
    if isinstance(obj, datetime.datetime):
        return obj.isoformat()
    if isinstance(obj, datetime.date):
        return obj.isoformat()
    
    # リストの変換（再帰的に変換）
    if isinstance(obj, (list, tuple)):
        return [to_serializable(item) for item in obj]
    
    # 辞書の変換（再帰的に変換）
    if isinstance(obj, dict):
        return {k: to_serializable(v) for k, v in obj.items()}
    
    # NumPy配列の変換
    if isinstance(obj, np.ndarray):
        return to_serializable(obj.tolist())
    
    # Pandas DataFrameの変換
    if isinstance(obj, pd.DataFrame):
        return {
            "type": "dataframe",
            "index": to_serializable(obj.index.tolist()),
            "columns": to_serializable(obj.columns.tolist()),
            "data": to_serializable(obj.values.tolist())
        }
    
    # Pandas Seriesの変換
    if isinstance(obj, pd.Series):
        return {
            "type": "series",
            "index": to_serializable(obj.index.tolist()),
            "data": to_serializable(obj.values.tolist())
        }
    
    # xlwings App オブジェクトの変換
    if isinstance(obj, xw.App):
        app_data = {"id": obj.pid}
        
        # 各プロパティを個別に try-except で囲む
        try:
            app_data["version"] = obj.version
        except Exception as e:
            app_data["version"] = "unknown"
            logger.warning(f"Error getting app version: {str(e)}")
        
        try:
            app_data["visible"] = obj.visible
        except Exception as e:
            app_data["visible"] = None
            logger.warning(f"Error getting app visibility: {str(e)}")
        
        # MacOS の k.missing_value エラーに対応
        try:
            app_data["calculation"] = str(obj.calculation)
        except KeyError as e:
            # MacOS では計算モードが取得できない場合がある
            if "k.missing_value" in str(e):
                app_data["calculation"] = "unknown"
                logger.warning("MacOS specific error: Unable to get calculation mode due to k.missing_value")
            else:
                raise
        except Exception as e:
            app_data["calculation"] = "unknown"
            logger.warning(f"Error getting app calculation mode: {str(e)}")
        
        try:
            app_data["screen_updating"] = obj.screen_updating
        except Exception as e:
            app_data["screen_updating"] = None
            logger.warning(f"Error getting app screen_updating: {str(e)}")
        
        try:
            app_data["display_alerts"] = obj.display_alerts
        except Exception as e:
            app_data["display_alerts"] = None
            logger.warning(f"Error getting app display_alerts: {str(e)}")
        
        return app_data
    
    # xlwings Book オブジェクトの変換
    if isinstance(obj, xw.Book):
        book_data = {}
        
        try:
            book_data["name"] = obj.name
        except Exception as e:
            book_data["name"] = "unknown"
            logger.warning(f"Error getting book name: {str(e)}")
        
        try:
            book_data["fullname"] = obj.fullname
        except Exception as e:
            book_data["fullname"] = None
            logger.warning(f"Error getting book fullname: {str(e)}")
        
        try:
            book_data["path"] = obj.fullname
        except Exception as e:
            book_data["path"] = None
            logger.warning(f"Error getting book path: {str(e)}")
        
        try:
            book_data["app_id"] = obj.app.pid if obj.app else None
        except Exception as e:
            book_data["app_id"] = None
            logger.warning(f"Error getting book app_id: {str(e)}")
        
        try:
            book_data["sheets"] = [sheet.name for sheet in obj.sheets]
        except Exception as e:
            book_data["sheets"] = []
            logger.warning(f"Error getting book sheets: {str(e)}")
        
        return book_data
    
    # xlwings Sheet オブジェクトの変換
    if isinstance(obj, xw.Sheet):
        sheet_data = {}
        
        try:
            sheet_data["name"] = obj.name
        except Exception as e:
            sheet_data["name"] = "unknown"
            logger.warning(f"Error getting sheet name: {str(e)}")
        
        try:
            sheet_data["book_name"] = obj.book.name
        except Exception as e:
            sheet_data["book_name"] = None
            logger.warning(f"Error getting sheet book_name: {str(e)}")
        
        try:
            sheet_data["index"] = obj.index
        except Exception as e:
            sheet_data["index"] = None
            logger.warning(f"Error getting sheet index: {str(e)}")
        
        try:
            sheet_data["used_range"] = str(obj.used_range.address)
        except Exception as e:
            sheet_data["used_range"] = None
            logger.warning(f"Error getting sheet used_range: {str(e)}")
        
        return sheet_data
    
    # xlwings Range オブジェクトの変換
    if isinstance(obj, xw.Range):
        range_data = {}
        
        try:
            range_data["address"] = obj.address
        except Exception as e:
            range_data["address"] = "unknown"
            logger.warning(f"Error getting range address: {str(e)}")
        
        try:
            range_data["sheet_name"] = obj.sheet.name
        except Exception as e:
            range_data["sheet_name"] = None
            logger.warning(f"Error getting range sheet_name: {str(e)}")
        
        try:
            range_data["book_name"] = obj.sheet.book.name
        except Exception as e:
            range_data["book_name"] = None
            logger.warning(f"Error getting range book_name: {str(e)}")
        
        try:
            range_data["value"] = to_serializable(obj.value)
        except Exception as e:
            range_data["value"] = None
            logger.warning(f"Error getting range value: {str(e)}")
        
        try:
            range_data["formula"] = to_serializable(obj.formula)
        except Exception as e:
            range_data["formula"] = None
            logger.warning(f"Error getting range formula: {str(e)}")
        
        try:
            range_data["shape"] = obj.shape
        except Exception as e:
            range_data["shape"] = None
            logger.warning(f"Error getting range shape: {str(e)}")
        
        try:
            range_data["row"] = obj.row
        except Exception as e:
            range_data["row"] = None
            logger.warning(f"Error getting range row: {str(e)}")
        
        try:
            range_data["column"] = obj.column
        except Exception as e:
            range_data["column"] = None
            logger.warning(f"Error getting range column: {str(e)}")
        
        try:
            range_data["row_height"] = obj.row_height
        except Exception as e:
            range_data["row_height"] = None
            logger.warning(f"Error getting range row_height: {str(e)}")
        
        try:
            range_data["column_width"] = obj.column_width
        except Exception as e:
            range_data["column_width"] = None
            logger.warning(f"Error getting range column_width: {str(e)}")
        
        return range_data
    
    # xlwings Chart オブジェクトの変換
    if isinstance(obj, xw.Chart):
        chart_data = {}
        
        try:
            chart_data["name"] = obj.name
        except Exception as e:
            chart_data["name"] = "unknown"
            logger.warning(f"Error getting chart name: {str(e)}")
        
        try:
            chart_data["chart_type"] = obj.chart_type
        except Exception as e:
            chart_data["chart_type"] = "unknown"
            logger.warning(f"Error getting chart type: {str(e)}")
        
        try:
            # シートとブックの情報
            if obj.parent:
                chart_data["sheet_name"] = obj.parent.name
                if obj.parent.book:
                    chart_data["book_name"] = obj.parent.book.name
        except Exception as e:
            chart_data["sheet_name"] = None
            chart_data["book_name"] = None
            logger.warning(f"Error getting chart parent info: {str(e)}")
        
        # 位置と大きさの情報
        try:
            chart_data["left"] = obj.left
        except Exception as e:
            chart_data["left"] = None
            logger.warning(f"Error getting chart left position: {str(e)}")
        
        try:
            chart_data["top"] = obj.top
        except Exception as e:
            chart_data["top"] = None
            logger.warning(f"Error getting chart top position: {str(e)}")
        
        try:
            chart_data["width"] = obj.width
        except Exception as e:
            chart_data["width"] = None
            logger.warning(f"Error getting chart width: {str(e)}")
        
        try:
            chart_data["height"] = obj.height
        except Exception as e:
            chart_data["height"] = None
            logger.warning(f"Error getting chart height: {str(e)}")
        
        return chart_data
    
    # その他のオブジェクトは文字列に変換
    return str(obj)


def from_json_value(value: Any) -> Any:
    """
    JSON値をPythonオブジェクトに変換します。
    特に、特殊なフォーマット（DataFrameなど）を元の型に復元します。

    Args:
        value: 変換するJSON値

    Returns:
        変換されたPythonオブジェクト
    """
    # 基本型はそのまま返す
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    
    # リストの変換（再帰的に変換）
    if isinstance(value, list):
        return [from_json_value(item) for item in value]
    
    # 辞書の変換
    if isinstance(value, dict):
        # DataFrameの復元
        if value.get("type") == "dataframe" and all(k in value for k in ["data", "columns", "index"]):
            return pd.DataFrame(
                data=value["data"],
                index=value["index"],
                columns=value["columns"]
            )
        
        # Seriesの復元
        if value.get("type") == "series" and all(k in value for k in ["data", "index"]):
            return pd.Series(
                data=value["data"],
                index=value["index"]
            )
        
        # 通常の辞書は再帰的に変換
        return {k: from_json_value(v) for k, v in value.items()}
    
    # その他の型はそのまま返す
    return value
