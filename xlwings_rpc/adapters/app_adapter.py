"""
Excel Applicationアダプター

xlwingsのAppオブジェクトとAPI間のインターフェースを提供します。
"""
from typing import Dict, List, Optional, Any, Union
import xlwings as xw
import logging
from xlwings_rpc.utils.converters import to_serializable

# ロガーの設定
logger = logging.getLogger(__name__)


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
                logger.debug(f"Attempting to get Excel app with PID: {pid}")
                try:
                    # 最新のxlwingsのAPIでは、appsコレクションから直接アクセスする
                    app = xw.apps[pid]
                except KeyError:
                    # PIDが見つからない場合
                    raise ValueError(f"No Excel application found with PID {pid}")
                except Exception as e:
                    logger.exception(f"Error accessing Excel app with PID {pid}: {str(e)}")
                    raise
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
            # 最新のxlwingsのAPIでは、appsコレクションから直接アクセスする
            try:
                app = xw.apps[pid]
            except KeyError:
                # PIDが見つからない場合
                raise ValueError(f"No Excel application found with PID {pid}")

            # 変更を保存する場合は、quit()の前に明示的に保存
            if save_changes:
                try:
                    # 開いているブックをすべて保存
                    for book in app.books:
                        if book.path:  # パスがある（保存済みのブック）の場合
                            book.save()
                except Exception as e:
                    logger.warning(f"Failed to save books before quitting: {str(e)}")
            
            # 公式ドキュメントによると、quit()は引数を取らない
            app.quit()  # 引数なしで呼び出し
            return True
        except Exception as e:
            # 終了に失敗した場合、killメソッドを試す
            logger.warning(f"Failed to quit Excel application: {str(e)}. Trying kill() method...")
            try:
                app.kill()
                return True
            except Exception as e2:
                raise ValueError(f"Failed to quit Excel application: {str(e)}. Kill attempt also failed: {str(e2)}")
    
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
            try:
                # 最新のxlwingsのAPIでは、appsコレクションから直接アクセスする
                app = xw.apps[pid]
            except KeyError:
                # PIDが見つからない場合
                raise ValueError(f"No Excel application found with PID {pid}")
            
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
            try:
                # 最新のxlwingsのAPIでは、appsコレクションから直接アクセスする
                app = xw.apps[pid]
            except KeyError:
                # PIDが見つからない場合
                raise ValueError(f"No Excel application found with PID {pid}")
            
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
            try:
                # 最新のxlwingsのAPIでは、appsコレクションから直接アクセスする
                app = xw.apps[pid]
            except KeyError:
                # PIDが見つからない場合
                raise ValueError(f"No Excel application found with PID {pid}")
            
            return [to_serializable(book) for book in app.books]
        except Exception as e:
            raise ValueError(f"Failed to get workbooks: {str(e)}")
