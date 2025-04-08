"""
Excel Rangeアダプター

xlwingsのRangeオブジェクトとAPI間のインターフェースを提供します。
"""
from typing import Dict, List, Optional, Any, Union, Tuple
import logging
import xlwings as xw
import pandas as pd
import numpy as np
from xlwings_rpc.utils.converters import to_serializable

# ロガーの設定
logger = logging.getLogger(__name__)


class RangeAdapter:
    """
    xlwingsのRangeオブジェクトに対するアダプタークラス
    """
    
    @staticmethod
    def get_range(
        book_identifier: str, 
        sheet_identifier: Union[str, int], 
        address: str,
        pid: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        指定された範囲を取得します。

        Args:
            book_identifier: ワークブック名かフルパス
            sheet_identifier: シート名かインデックス
            address: セル範囲のアドレス (例: "A1", "A1:B5")
            pid: ExcelアプリケーションのプロセスID (オプション)

        Returns:
            範囲情報

        Raises:
            ValueError: ワークブック、シート、範囲が見つからない場合
        """
        try:
            if pid is not None:
                # 最新のxlwingsのAPIでは、appsコレクションから直接アクセスする
                try:
                    app = xw.apps[pid]
                except KeyError:
                    # PIDが見つからない場合
                    raise ValueError(f"No Excel application found with PID {pid}")
                
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            sheet = book.sheets[sheet_identifier]
            range_obj = sheet.range(address)
            return to_serializable(range_obj)
        except Exception as e:
            raise ValueError(f"Failed to get range '{address}' from sheet '{sheet_identifier}' in workbook '{book_identifier}': {str(e)}")
    
    @staticmethod
    def get_range_value(
        book_identifier: str, 
        sheet_identifier: Union[str, int], 
        address: str,
        pid: Optional[int] = None
    ) -> Any:
        """
        指定された範囲の値を取得します。

        Args:
            book_identifier: ワークブック名かフルパス
            sheet_identifier: シート名かインデックス
            address: セル範囲のアドレス
            pid: ExcelアプリケーションのプロセスID (オプション)

        Returns:
            範囲の値

        Raises:
            ValueError: ワークブック、シート、範囲が見つからない場合
        """
        try:
            if pid is not None:
                # 最新のxlwingsのAPIでは、appsコレクションから直接アクセスする
                try:
                    app = xw.apps[pid]
                except KeyError:
                    # PIDが見つからない場合
                    raise ValueError(f"No Excel application found with PID {pid}")
                
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            sheet = book.sheets[sheet_identifier]
            range_obj = sheet.range(address)
            return to_serializable(range_obj.value)
        except Exception as e:
            raise ValueError(f"Failed to get value of range '{address}' from sheet '{sheet_identifier}' in workbook '{book_identifier}': {str(e)}")
    
    @staticmethod
    def set_range_value(
        book_identifier: str, 
        sheet_identifier: Union[str, int], 
        address: str,
        value: Any,
        pid: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        指定された範囲に値を設定します。

        Args:
            book_identifier: ワークブック名かフルパス
            sheet_identifier: シート名かインデックス
            address: セル範囲のアドレス
            value: 設定する値
            pid: ExcelアプリケーションのプロセスID (オプション)

        Returns:
            更新された範囲情報

        Raises:
            ValueError: ワークブック、シート、範囲が見つからない場合
        """
        try:
            if pid is not None:
                # 最新のxlwingsのAPIでは、appsコレクションから直接アクセスする
                try:
                    app = xw.apps[pid]
                except KeyError:
                    # PIDが見つからない場合
                    raise ValueError(f"No Excel application found with PID {pid}")
                
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            sheet = book.sheets[sheet_identifier]
            range_obj = sheet.range(address)
            range_obj.value = value
            return to_serializable(range_obj)
        except Exception as e:
            raise ValueError(f"Failed to set value of range '{address}' in sheet '{sheet_identifier}' of workbook '{book_identifier}': {str(e)}")
    
    @staticmethod
    def get_range_formula(
        book_identifier: str, 
        sheet_identifier: Union[str, int], 
        address: str,
        pid: Optional[int] = None
    ) -> Any:
        """
        指定された範囲の数式を取得します。

        Args:
            book_identifier: ワークブック名かフルパス
            sheet_identifier: シート名かインデックス
            address: セル範囲のアドレス
            pid: ExcelアプリケーションのプロセスID (オプション)

        Returns:
            範囲の数式

        Raises:
            ValueError: ワークブック、シート、範囲が見つからない場合
        """
        try:
            if pid is not None:
                # 最新のxlwingsのAPIでは、appsコレクションから直接アクセスする
                try:
                    app = xw.apps[pid]
                except KeyError:
                    # PIDが見つからない場合
                    raise ValueError(f"No Excel application found with PID {pid}")
                
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            sheet = book.sheets[sheet_identifier]
            range_obj = sheet.range(address)
            return to_serializable(range_obj.formula)
        except Exception as e:
            raise ValueError(f"Failed to get formula of range '{address}' from sheet '{sheet_identifier}' in workbook '{book_identifier}': {str(e)}")
    
    @staticmethod
    def set_range_formula(
        book_identifier: str, 
        sheet_identifier: Union[str, int], 
        address: str,
        formula: Any,
        pid: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        指定された範囲に数式を設定します。

        Args:
            book_identifier: ワークブック名かフルパス
            sheet_identifier: シート名かインデックス
            address: セル範囲のアドレス
            formula: 設定する数式
            pid: ExcelアプリケーションのプロセスID (オプション)

        Returns:
            更新された範囲情報

        Raises:
            ValueError: ワークブック、シート、範囲が見つからない場合
        """
        try:
            if pid is not None:
                # 最新のxlwingsのAPIでは、appsコレクションから直接アクセスする
                try:
                    app = xw.apps[pid]
                except KeyError:
                    # PIDが見つからない場合
                    raise ValueError(f"No Excel application found with PID {pid}")
                
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            sheet = book.sheets[sheet_identifier]
            range_obj = sheet.range(address)
            range_obj.formula = formula
            return to_serializable(range_obj)
        except Exception as e:
            raise ValueError(f"Failed to set formula of range '{address}' in sheet '{sheet_identifier}' of workbook '{book_identifier}': {str(e)}")
    
    @staticmethod
    def clear_range(
        book_identifier: str, 
        sheet_identifier: Union[str, int], 
        address: str,
        pid: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        指定された範囲をクリアします。

        Args:
            book_identifier: ワークブック名かフルパス
            sheet_identifier: シート名かインデックス
            address: セル範囲のアドレス
            pid: ExcelアプリケーションのプロセスID (オプション)

        Returns:
            クリアされた範囲情報

        Raises:
            ValueError: ワークブック、シート、範囲が見つからない場合
        """
        try:
            if pid is not None:
                # 最新のxlwingsのAPIでは、appsコレクションから直接アクセスする
                try:
                    app = xw.apps[pid]
                except KeyError:
                    # PIDが見つからない場合
                    raise ValueError(f"No Excel application found with PID {pid}")
                
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            sheet = book.sheets[sheet_identifier]
            range_obj = sheet.range(address)
            range_obj.clear()
            return to_serializable(range_obj)
        except Exception as e:
            raise ValueError(f"Failed to clear range '{address}' in sheet '{sheet_identifier}' of workbook '{book_identifier}': {str(e)}")
    
    @staticmethod
    def get_range_as_dataframe(
        book_identifier: str, 
        sheet_identifier: Union[str, int], 
        address: str,
        header: bool = True,
        index: bool = False,
        pid: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        指定された範囲をDataFrameとして取得します。

        Args:
            book_identifier: ワークブック名かフルパス
            sheet_identifier: シート名かインデックス
            address: セル範囲のアドレス
            header: 最初の行をヘッダーとして使用するかどうか
            index: 最初の列をインデックスとして使用するかどうか
            pid: ExcelアプリケーションのプロセスID (オプション)

        Returns:
            DataFrame情報

        Raises:
            ValueError: ワークブック、シート、範囲が見つからない場合
        """
        try:
            if pid is not None:
                # 最新のxlwingsのAPIでは、appsコレクションから直接アクセスする
                try:
                    app = xw.apps[pid]
                except KeyError:
                    # PIDが見つからない場合
                    raise ValueError(f"No Excel application found with PID {pid}")
                
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            sheet = book.sheets[sheet_identifier]
            range_obj = sheet.range(address)
            df = range_obj.options(pd.DataFrame, header=header, index=index).value
            return to_serializable(df)
        except Exception as e:
            raise ValueError(f"Failed to get range '{address}' as DataFrame from sheet '{sheet_identifier}' in workbook '{book_identifier}': {str(e)}")
    
    @staticmethod
    def set_dataframe(
        book_identifier: str, 
        sheet_identifier: Union[str, int], 
        address: str,
        dataframe: Dict[str, Any],
        header: bool = True,
        index: bool = False,
        pid: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        指定された範囲にDataFrameを設定します。

        Args:
            book_identifier: ワークブック名かフルパス
            sheet_identifier: シート名かインデックス
            address: セル範囲のアドレス
            dataframe: 設定するDataFrame (シリアライズされた形式)
            header: ヘッダーを含めるかどうか
            index: インデックスを含めるかどうか
            pid: ExcelアプリケーションのプロセスID (オプション)

        Returns:
            更新された範囲情報

        Raises:
            ValueError: ワークブック、シート、範囲が見つからない場合
        """
        try:
            if pid is not None:
                # 最新のxlwingsのAPIでは、appsコレクションから直接アクセスする
                try:
                    app = xw.apps[pid]
                except KeyError:
                    # PIDが見つからない場合
                    raise ValueError(f"No Excel application found with PID {pid}")
                
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            sheet = book.sheets[sheet_identifier]
            range_obj = sheet.range(address)
            
            # シリアライズされたDataFrameを復元
            df = pd.DataFrame(
                data=dataframe["data"],
                index=dataframe["index"],
                columns=dataframe["columns"]
            )
            
            range_obj.options(pd.DataFrame, header=header, index=index).value = df
            return to_serializable(range_obj)
        except Exception as e:
            raise ValueError(f"Failed to set DataFrame to range '{address}' in sheet '{sheet_identifier}' of workbook '{book_identifier}': {str(e)}")
