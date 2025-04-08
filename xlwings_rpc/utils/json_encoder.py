"""
カスタムJSONエンコーダー

xlwings-rpc用のカスタムJSONエンコーダーを提供します。
Python標準のJSONEncoderを拡張し、xlwingsの特殊なオブジェクトを適切にシリアライズします。
"""
import json
from typing import Any
import xlwings as xw
import numpy as np
import pandas as pd
import datetime
import logging
from xlwings_rpc.utils.converters import to_serializable

# ロガーの設定
logger = logging.getLogger(__name__)


class XlwingsRpcJSONEncoder(json.JSONEncoder):
    """
    xlwings-rpc用のカスタムJSONエンコーダー
    
    xlwingsの特殊なオブジェクトや、Python固有の表現を
    標準的なJSON形式に変換します。
    """
    
    def default(self, obj: Any) -> Any:
        """
        オブジェクトをJSONシリアライズ可能な形式に変換します。
        
        Args:
            obj: 変換するオブジェクト
            
        Returns:
            JSONシリアライズ可能なオブジェクト
        """
        # 既存のto_serializable関数を利用
        result = to_serializable(obj)
        
        # to_serializable関数で変換した結果がオブジェクトと同一の場合
        # （変換されなかった場合）は追加の変換処理を試みる
        if result is obj:
            # 日付型の変換
            if isinstance(obj, datetime.datetime):
                return obj.isoformat()
            if isinstance(obj, datetime.date):
                return obj.isoformat()
            
            # xlwings特有のVersionNumberクラスの処理
            if hasattr(obj, '__class__') and obj.__class__.__name__ == 'VersionNumber':
                return str(obj)
            
            # NumPy型の変換
            if isinstance(obj, (np.integer, np.int64, np.int32)):
                return int(obj)
            if isinstance(obj, (np.floating, np.float64, np.float32)):
                return float(obj)
            if isinstance(obj, np.bool_):
                return bool(obj)
            
            # 複合型の場合、通常のJSONEncoderのdefault処理を試みる
            try:
                return super().default(obj)
            except TypeError:
                # 最終手段として文字列化
                return str(obj)
        
        return result


def json_dumps(obj: Any) -> str:
    """
    オブジェクトをJSON文字列に変換します。
    
    Args:
        obj: 変換するオブジェクト
        
    Returns:
        JSON文字列
    """
    try:
        return json.dumps(obj, cls=XlwingsRpcJSONEncoder)
    except Exception as e:
        logger.error(f"JSONシリアライズエラー: {str(e)}")
        # 最終手段として、通常のJSON変換を試みる
        return json.dumps({"error": f"シリアライズエラー: {str(e)}"})
