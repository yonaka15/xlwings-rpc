"""
Excel Sheetアダプター

xlwingsのSheetオブジェクトとAPI間のインターフェースを提供します。
"""
from typing import Dict, List, Optional, Any, Union
import logging
import xlwings as xw
from xlwings_rpc.utils.converters import to_serializable

# ロガーの設定
logger = logging.getLogger(__name__)


class SheetAdapter:
    """
    xlwingsのSheetオブジェクトに対するアダプタークラス
    """
    
    @staticmethod
    def get_sheets(book_identifier: str, pid: Optional[int] = None) -> List[Dict[str, Any]]:
        """
        ワークブック内のすべてのシートを取得します。

        Args:
            book_identifier: ワークブック名かフルパス
            pid: ExcelアプリケーションのプロセスID (オプション)

        Returns:
            シート情報のリスト
        
        Raises:
            ValueError: ワークブックが見つからない場合
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
            
            return [to_serializable(sheet) for sheet in book.sheets]
        except Exception as e:
            raise ValueError(f"Failed to get sheets for workbook '{book_identifier}': {str(e)}")
    
    @staticmethod
    def get_sheet(
        book_identifier: str, 
        sheet_identifier: Union[str, int], 
        pid: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        特定のシートを取得します。

        Args:
            book_identifier: ワークブック名かフルパス
            sheet_identifier: シート名かインデックス
            pid: ExcelアプリケーションのプロセスID (オプション)

        Returns:
            シート情報

        Raises:
            ValueError: ワークブックやシートが見つからない場合
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
            return to_serializable(sheet)
        except Exception as e:
            raise ValueError(f"Failed to get sheet '{sheet_identifier}' from workbook '{book_identifier}': {str(e)}")
    
    @staticmethod
    def add_sheet(
        book_identifier: str, 
        name: Optional[str] = None, 
        before: Optional[Union[str, int]] = None, 
        after: Optional[Union[str, int]] = None,
        pid: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        新しいシートを追加します。

        Args:
            book_identifier: ワークブック名かフルパス
            name: 新しいシート名 (オプション)
            before: この前に追加 (オプション)
            after: この後に追加 (オプション)
            pid: ExcelアプリケーションのプロセスID (オプション)

        Returns:
            新しいシート情報

        Raises:
            ValueError: ワークブックが見つからないか、シート追加に失敗した場合
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
            
            if before is not None:
                sheet = book.sheets.add(name=name, before=book.sheets[before])
            elif after is not None:
                sheet = book.sheets.add(name=name, after=book.sheets[after])
            else:
                sheet = book.sheets.add(name=name)
            
            return to_serializable(sheet)
        except Exception as e:
            raise ValueError(f"Failed to add sheet to workbook '{book_identifier}': {str(e)}")
    
    @staticmethod
    def delete_sheet(
        book_identifier: str, 
        sheet_identifier: Union[str, int], 
        pid: Optional[int] = None
    ) -> bool:
        """
        シートを削除します。

        Args:
            book_identifier: ワークブック名かフルパス
            sheet_identifier: シート名かインデックス
            pid: ExcelアプリケーションのプロセスID (オプション)

        Returns:
            成功した場合はTrue

        Raises:
            ValueError: ワークブックやシートが見つからない場合
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
            sheet.delete()
            return True
        except Exception as e:
            raise ValueError(f"Failed to delete sheet '{sheet_identifier}' from workbook '{book_identifier}': {str(e)}")
    
    @staticmethod
    def rename_sheet(
        book_identifier: str, 
        sheet_identifier: Union[str, int], 
        new_name: str,
        pid: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        シート名を変更します。

        Args:
            book_identifier: ワークブック名かフルパス
            sheet_identifier: シート名かインデックス
            new_name: 新しいシート名
            pid: ExcelアプリケーションのプロセスID (オプション)

        Returns:
            更新されたシート情報

        Raises:
            ValueError: ワークブックやシートが見つからない場合
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
            sheet.name = new_name
            return to_serializable(sheet)
        except Exception as e:
            raise ValueError(f"Failed to rename sheet '{sheet_identifier}' in workbook '{book_identifier}': {str(e)}")
    
    @staticmethod
    def clear_sheet(
        book_identifier: str, 
        sheet_identifier: Union[str, int], 
        pid: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        シートの内容をクリアします。

        Args:
            book_identifier: ワークブック名かフルパス
            sheet_identifier: シート名かインデックス
            pid: ExcelアプリケーションのプロセスID (オプション)

        Returns:
            更新されたシート情報

        Raises:
            ValueError: ワークブックやシートが見つからない場合
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
            sheet.clear()
            return to_serializable(sheet)
        except Exception as e:
            raise ValueError(f"Failed to clear sheet '{sheet_identifier}' in workbook '{book_identifier}': {str(e)}")
    
    @staticmethod
    def get_used_range(
        book_identifier: str, 
        sheet_identifier: Union[str, int], 
        pid: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        シートの使用範囲を取得します。

        Args:
            book_identifier: ワークブック名かフルパス
            sheet_identifier: シート名かインデックス
            pid: ExcelアプリケーションのプロセスID (オプション)

        Returns:
            使用範囲情報

        Raises:
            ValueError: ワークブックやシートが見つからない場合
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
            used_range = sheet.used_range
            return to_serializable(used_range)
        except Exception as e:
            raise ValueError(f"Failed to get used range for sheet '{sheet_identifier}' in workbook '{book_identifier}': {str(e)}")
    
    @staticmethod
    def activate_sheet(
        book_identifier: str, 
        sheet_identifier: Union[str, int], 
        pid: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        シートをアクティブにします。

        Args:
            book_identifier: ワークブック名かフルパス
            sheet_identifier: シート名かインデックス
            pid: ExcelアプリケーションのプロセスID (オプション)

        Returns:
            アクティブにしたシート情報

        Raises:
            ValueError: ワークブックやシートが見つからない場合
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
            sheet.activate()
            return to_serializable(sheet)
        except Exception as e:
            raise ValueError(f"Failed to activate sheet '{sheet_identifier}' in workbook '{book_identifier}': {str(e)}")
