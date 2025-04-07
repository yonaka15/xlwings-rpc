"""
レンジ関連のRPCメソッド

Excel Rangeに関連するJSON-RPCメソッドを実装します。
"""
from typing import Dict, List, Optional, Any, Union
from xlwings_rpc.adapters.range_adapter import RangeAdapter


class RangeMethods:
    """
    range.* 名前空間のRPCメソッド実装
    """
    
    @staticmethod
    async def get(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        range.get: 指定された範囲を取得します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - address (str): セル範囲のアドレス (例: "A1", "A1:B5")
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            範囲情報
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        address = params["address"]
        pid = params.get("pid")
        return RangeAdapter.get_range(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            address=address,
            pid=pid
        )
    
    @staticmethod
    async def get_value(params: Dict[str, Any]) -> Any:
        """
        range.get_value: 指定された範囲の値を取得します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - address (str): セル範囲のアドレス
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            範囲の値
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        address = params["address"]
        pid = params.get("pid")
        return RangeAdapter.get_range_value(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            address=address,
            pid=pid
        )
    
    @staticmethod
    async def set_value(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        range.set_value: 指定された範囲に値を設定します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - address (str): セル範囲のアドレス
                - value (Any): 設定する値
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            更新された範囲情報
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        address = params["address"]
        value = params["value"]
        pid = params.get("pid")
        return RangeAdapter.set_range_value(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            address=address,
            value=value,
            pid=pid
        )
    
    @staticmethod
    async def get_formula(params: Dict[str, Any]) -> Any:
        """
        range.get_formula: 指定された範囲の数式を取得します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - address (str): セル範囲のアドレス
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            範囲の数式
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        address = params["address"]
        pid = params.get("pid")
        return RangeAdapter.get_range_formula(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            address=address,
            pid=pid
        )
    
    @staticmethod
    async def set_formula(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        range.set_formula: 指定された範囲に数式を設定します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - address (str): セル範囲のアドレス
                - formula (Any): 設定する数式
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            更新された範囲情報
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        address = params["address"]
        formula = params["formula"]
        pid = params.get("pid")
        return RangeAdapter.set_range_formula(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            address=address,
            formula=formula,
            pid=pid
        )
    
    @staticmethod
    async def clear(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        range.clear: 指定された範囲をクリアします。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - address (str): セル範囲のアドレス
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            クリアされた範囲情報
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        address = params["address"]
        pid = params.get("pid")
        return RangeAdapter.clear_range(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            address=address,
            pid=pid
        )
    
    @staticmethod
    async def get_as_dataframe(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        range.get_as_dataframe: 指定された範囲をDataFrameとして取得します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - address (str): セル範囲のアドレス
                - header (Optional[bool]): 最初の行をヘッダーとして使用するかどうか
                - index (Optional[bool]): 最初の列をインデックスとして使用するかどうか
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            DataFrame情報
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        address = params["address"]
        header = params.get("header", True)
        index = params.get("index", False)
        pid = params.get("pid")
        return RangeAdapter.get_range_as_dataframe(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            address=address,
            header=header,
            index=index,
            pid=pid
        )
    
    @staticmethod
    async def set_dataframe(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        range.set_dataframe: 指定された範囲にDataFrameを設定します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - address (str): セル範囲のアドレス
                - dataframe (Dict[str, Any]): 設定するDataFrame (シリアライズされた形式)
                - header (Optional[bool]): ヘッダーを含めるかどうか
                - index (Optional[bool]): インデックスを含めるかどうか
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            更新された範囲情報
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        address = params["address"]
        dataframe = params["dataframe"]
        header = params.get("header", True)
        index = params.get("index", False)
        pid = params.get("pid")
        return RangeAdapter.set_dataframe(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            address=address,
            dataframe=dataframe,
            header=header,
            index=index,
            pid=pid
        )
