"""
アプリケーション関連のRPCメソッド

Excel Applicationに関連するJSON-RPCメソッドを実装します。
"""
from typing import Dict, List, Optional, Any, Union
from xlwings_rpc.adapters.app_adapter import AppAdapter


class AppMethods:
    """
    app.* 名前空間のRPCメソッド実装
    """
    
    @staticmethod
    async def list() -> List[Dict[str, Any]]:
        """
        app.list: すべての実行中のExcelアプリケーションを取得します。

        Returns:
            アプリケーション情報のリスト
        """
        return AppAdapter.get_apps()
    
    @staticmethod
    async def get(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        app.get: 指定されたPIDまたはアクティブなExcelアプリケーションを取得します。

        Args:
            params: パラメータオブジェクト
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            アプリケーション情報
        """
        pid = params.get("pid")
        return AppAdapter.get_app(pid)
    
    @staticmethod
    async def create(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        app.create: 新しいExcelアプリケーションを作成します。

        Args:
            params: パラメータオブジェクト
                - visible (Optional[bool]): アプリケーションを表示するかどうか
                - add_book (Optional[bool]): 新しいブックを追加するかどうか

        Returns:
            新しいアプリケーション情報
        """
        visible = params.get("visible", True)
        add_book = params.get("add_book", True)
        return AppAdapter.create_app(visible=visible, add_book=add_book)
    
    @staticmethod
    async def quit(params: Dict[str, Any]) -> bool:
        """
        app.quit: Excelアプリケーションを終了します。

        Args:
            params: パラメータオブジェクト
                - pid (int): ExcelアプリケーションのプロセスID
                - save_changes (Optional[bool]): 変更を保存するかどうか

        Returns:
            成功した場合はTrue
        """
        pid = params["pid"]
        save_changes = params.get("save_changes", True)
        return AppAdapter.quit_app(pid=pid, save_changes=save_changes)
    
    @staticmethod
    async def set_calculation(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        app.set_calculation: 計算モードを設定します。

        Args:
            params: パラメータオブジェクト
                - pid (int): ExcelアプリケーションのプロセスID
                - mode (str): 計算モード ('automatic', 'manual', 'semiautomatic')

        Returns:
            更新されたアプリケーション情報
        """
        pid = params["pid"]
        mode = params["mode"]
        return AppAdapter.set_calculation(pid=pid, calculation_mode=mode)
    
    @staticmethod
    async def get_calculation(params: Dict[str, Any]) -> str:
        """
        app.get_calculation: 現在の計算モードを取得します。

        Args:
            params: パラメータオブジェクト
                - pid (int): ExcelアプリケーションのプロセスID

        Returns:
            計算モード
        """
        pid = params["pid"]
        return AppAdapter.get_calculation(pid=pid)
    
    @staticmethod
    async def get_books(params: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        app.get_books: 指定されたアプリケーションで開いているワークブックを取得します。

        Args:
            params: パラメータオブジェクト
                - pid (int): ExcelアプリケーションのプロセスID

        Returns:
            ワークブック情報のリスト
        """
        pid = params["pid"]
        return AppAdapter.get_app_books(pid=pid)
