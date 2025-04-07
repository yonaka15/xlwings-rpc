"""
ワークブック関連のRPCメソッド

Excel Workbookに関連するJSON-RPCメソッドを実装します。
"""
from typing import Dict, List, Optional, Any, Union
from xlwings_rpc.adapters.book_adapter import BookAdapter


class BookMethods:
    """
    book.* 名前空間のRPCメソッド実装
    """
    
    @staticmethod
    async def list(params: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
        """
        book.list: 開いているワークブックを取得します。

        Args:
            params: パラメータオブジェクト (オプション)
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            ワークブック情報のリスト
        """
        pid = params.get("pid") if params else None
        return BookAdapter.get_books(pid=pid)
    
    @staticmethod
    async def get(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        book.get: 特定のワークブックを取得します。

        Args:
            params: パラメータオブジェクト
                - name (str): ワークブック名かフルパス
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            ワークブック情報
        """
        book_identifier = params["name"]
        pid = params.get("pid")
        return BookAdapter.get_book(book_identifier=book_identifier, pid=pid)
    
    @staticmethod
    async def open(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        book.open: ワークブックを開きます。

        Args:
            params: パラメータオブジェクト
                - path (str): ワークブックのパス
                - pid (Optional[int]): ExcelアプリケーションのプロセスID
                - read_only (Optional[bool]): 読み取り専用で開くかどうか
                - password (Optional[str]): パスワード

        Returns:
            開いたワークブック情報
        """
        path = params["path"]
        pid = params.get("pid")
        read_only = params.get("read_only", False)
        password = params.get("password")
        return BookAdapter.open_book(
            path=path, pid=pid, read_only=read_only, password=password
        )
    
    @staticmethod
    async def create(params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        book.create: 新しいワークブックを作成します。

        Args:
            params: パラメータオブジェクト (オプション)
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            新しいワークブック情報
        """
        pid = params.get("pid") if params else None
        return BookAdapter.create_book(pid=pid)
    
    @staticmethod
    async def close(params: Dict[str, Any]) -> bool:
        """
        book.close: ワークブックを閉じます。

        Args:
            params: パラメータオブジェクト
                - name (str): ワークブック名かフルパス
                - pid (Optional[int]): ExcelアプリケーションのプロセスID
                - save (Optional[bool]): 変更を保存するかどうか
                - path (Optional[str]): 保存先パス

        Returns:
            成功した場合はTrue
        """
        book_identifier = params["name"]
        pid = params.get("pid")
        save = params.get("save", True)
        path = params.get("path")
        return BookAdapter.close_book(
            book_identifier=book_identifier, pid=pid, save=save, path=path
        )
    
    @staticmethod
    async def save(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        book.save: ワークブックを保存します。

        Args:
            params: パラメータオブジェクト
                - name (str): ワークブック名かフルパス
                - pid (Optional[int]): ExcelアプリケーションのプロセスID
                - path (Optional[str]): 保存先パス

        Returns:
            保存したワークブック情報
        """
        book_identifier = params["name"]
        pid = params.get("pid")
        path = params.get("path")
        return BookAdapter.save_book(
            book_identifier=book_identifier, pid=pid, path=path
        )
    
    @staticmethod
    async def get_sheets(params: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        book.get_sheets: ワークブック内のシートを取得します。

        Args:
            params: パラメータオブジェクト
                - name (str): ワークブック名かフルパス
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            シート情報のリスト
        """
        book_identifier = params["name"]
        pid = params.get("pid")
        return BookAdapter.get_book_sheets(
            book_identifier=book_identifier, pid=pid
        )
