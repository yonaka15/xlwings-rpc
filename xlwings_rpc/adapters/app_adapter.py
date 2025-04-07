"""
Excel Applicationアダプター

xlwingsのAppオブジェクトとAPI間のインターフェースを提供します。
"""
from typing import Dict, List, Optional, Any, Union
import xlwings as xw
from xlwings_rpc.utils.converters import to_serializable


class AppAdapter:
    """
    xlwingsのAppオブジェクトに対するアダプタークラス
    """
    
    @staticmethod
    def get_apps() -> List[Dict[str, Any]]:
        """
        すべての実行中のExcelアプリケーションを取得します。

        Returns:
            アプリケーション情報のリスト
        """
        apps = []
        for app in xw.apps:
            apps.append(to_serializable(app))
        return apps
    
    @staticmethod
    def get_app(pid: Optional[int] = None) -> Dict[str, Any]:
        """
        指定されたPIDまたはアクティブなExcelアプリケーションを取得します。

        Args:
            pid: ExcelアプリケーションのプロセスID (デフォルト: None)

        Returns:
            アプリケーション情報
        
        Raises:
            ValueError: 指定されたPIDのアプリケーションが見つからない場合
        """
        try:
            if pid is not None:
                app = xw.App(pid=pid)
            else:
                # アクティブなアプリケーションを取得、なければ新規作成
                try:
                    app = xw.apps.active
                    if app is None:
                        raise AttributeError("No active app")
                except (AttributeError, IndexError):
                    app = xw.App(visible=False)
            
            return to_serializable(app)
        except Exception as e:
            raise ValueError(f"Failed to get Excel application: {str(e)}")
    
    @staticmethod
    def create_app(visible: bool = True, add_book: bool = True) -> Dict[str, Any]:
        """
        新しいExcelアプリケーションを作成します。

        Args:
            visible: アプリケーションを表示するかどうか (デフォルト: True)
            add_book: 新しいブックを追加するかどうか (デフォルト: True)

        Returns:
            新しいアプリケーション情報
        """
        app = xw.App(visible=visible, add_book=add_book)
        return to_serializable(app)
    
    @staticmethod
    def quit_app(pid: int, save_changes: bool = True) -> bool:
        """
        Excelアプリケーションを終了します。

        Args:
            pid: ExcelアプリケーションのプロセスID
            save_changes: 変更を保存するかどうか (デフォルト: True)

        Returns:
            成功した場合はTrue

        Raises:
            ValueError: 指定されたPIDのアプリケーションが見つからない場合
        """
        try:
            app = xw.App(pid=pid)
            app.quit(save_changes)
            return True
        except Exception as e:
            raise ValueError(f"Failed to quit Excel application: {str(e)}")
    
    @staticmethod
    def set_calculation(pid: int, calculation_mode: str) -> Dict[str, Any]:
        """
        計算モードを設定します。

        Args:
            pid: ExcelアプリケーションのプロセスID
            calculation_mode: 計算モード ('automatic', 'manual', 'semiautomatic')

        Returns:
            更新されたアプリケーション情報

        Raises:
            ValueError: 無効な計算モードまたはPIDが指定された場合
        """
        valid_modes = {'automatic', 'manual', 'semiautomatic'}
        if calculation_mode.lower() not in valid_modes:
            raise ValueError(f"Invalid calculation mode. Valid values are: {', '.join(valid_modes)}")
        
        try:
            app = xw.App(pid=pid)
            app.calculation = calculation_mode.lower()
            return to_serializable(app)
        except Exception as e:
            raise ValueError(f"Failed to set calculation mode: {str(e)}")
    
    @staticmethod
    def get_calculation(pid: int) -> str:
        """
        現在の計算モードを取得します。

        Args:
            pid: ExcelアプリケーションのプロセスID

        Returns:
            計算モード

        Raises:
            ValueError: 指定されたPIDのアプリケーションが見つからない場合
        """
        try:
            app = xw.App(pid=pid)
            return str(app.calculation)
        except Exception as e:
            raise ValueError(f"Failed to get calculation mode: {str(e)}")
    
    @staticmethod
    def get_app_books(pid: int) -> List[Dict[str, Any]]:
        """
        指定されたアプリケーションで開いているワークブックを取得します。

        Args:
            pid: ExcelアプリケーションのプロセスID

        Returns:
            ワークブック情報のリスト

        Raises:
            ValueError: 指定されたPIDのアプリケーションが見つからない場合
        """
        try:
            app = xw.App(pid=pid)
            return [to_serializable(book) for book in app.books]
        except Exception as e:
            raise ValueError(f"Failed to get workbooks: {str(e)}")
