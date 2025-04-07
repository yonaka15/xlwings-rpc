"""
シート関連のRPCメソッド

Excel Sheetに関連するJSON-RPCメソッドを実装します。
"""
from typing import Dict, List, Optional, Any, Union
from xlwings_rpc.adapters.sheet_adapter import SheetAdapter


class SheetMethods:
    """
    sheet.* 名前空間のRPCメソッド実装
    """
    
    @staticmethod
    async def list(params: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        sheet.list: ワークブック内のすべてのシートを取得します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            シート情報のリスト
        """
        book_identifier = params["book"]
        pid = params.get("pid")
        return SheetAdapter.get_sheets(book_identifier=book_identifier, pid=pid)
    
    @staticmethod
    async def get(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        sheet.get: 特定のシートを取得します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            シート情報
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        pid = params.get("pid")
        return SheetAdapter.get_sheet(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            pid=pid
        )
    
    @staticmethod
    async def add(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        sheet.add: 新しいシートを追加します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - name (Optional[str]): 新しいシート名
                - before (Optional[Union[str, int]]): この前に追加
                - after (Optional[Union[str, int]]): この後に追加
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            新しいシート情報
        """
        book_identifier = params["book"]
        name = params.get("name")
        before = params.get("before")
        after = params.get("after")
        pid = params.get("pid")
        return SheetAdapter.add_sheet(
            book_identifier=book_identifier,
            name=name,
            before=before,
            after=after,
            pid=pid
        )
    
    @staticmethod
    async def delete(params: Dict[str, Any]) -> bool:
        """
        sheet.delete: シートを削除します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            成功した場合はTrue
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        pid = params.get("pid")
        return SheetAdapter.delete_sheet(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            pid=pid
        )
    
    @staticmethod
    async def rename(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        sheet.rename: シート名を変更します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - new_name (str): 新しいシート名
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            更新されたシート情報
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        new_name = params["new_name"]
        pid = params.get("pid")
        return SheetAdapter.rename_sheet(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            new_name=new_name,
            pid=pid
        )
    
    @staticmethod
    async def clear(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        sheet.clear: シートの内容をクリアします。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            更新されたシート情報
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        pid = params.get("pid")
        return SheetAdapter.clear_sheet(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            pid=pid
        )
    
    @staticmethod
    async def get_used_range(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        sheet.get_used_range: シートの使用範囲を取得します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            使用範囲情報
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        pid = params.get("pid")
        return SheetAdapter.get_used_range(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            pid=pid
        )
    
    @staticmethod
    async def activate(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        sheet.activate: シートをアクティブにします。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            アクティブにしたシート情報
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        pid = params.get("pid")
        return SheetAdapter.activate_sheet(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            pid=pid
        )
