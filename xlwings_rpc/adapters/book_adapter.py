"""
Excel Workbookアダプター

xlwingsのBookオブジェクトとAPI間のインターフェースを提供します。
"""
from typing import Dict, List, Optional, Any, Union
import os
import xlwings as xw
from xlwings_rpc.utils.converters import to_serializable


class BookAdapter:
    """
    xlwingsのBookオブジェクトに対するアダプタークラス
    """
    
    @staticmethod
    def get_books(pid: Optional[int] = None) -> List[Dict[str, Any]]:
        """
        開いているワークブックを取得します。

        Args:
            pid: ExcelアプリケーションのプロセスID (Noneの場合はすべてのアプリケーション)

        Returns:
            ワークブック情報のリスト
        """
        books = []
        if pid is not None:
            try:
                app = xw.App(pid=pid)
                for book in app.books:
                    books.append(to_serializable(book))
            except Exception as e:
                raise ValueError(f"Failed to get books for Excel application (PID {pid}): {str(e)}")
        else:
            for app in xw.apps:
                for book in app.books:
                    books.append(to_serializable(book))
        
        return books
    
    @staticmethod
    def get_book(book_identifier: str, pid: Optional[int] = None) -> Dict[str, Any]:
        """
        特定のワークブックを取得します。

        Args:
            book_identifier: ワークブック名かフルパス
            pid: ExcelアプリケーションのプロセスID (オプション)

        Returns:
            ワークブック情報

        Raises:
            ValueError: ワークブックが見つからない場合
        """
        try:
            if pid is not None:
                app = xw.App(pid=pid)
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            return to_serializable(book)
        except Exception as e:
            raise ValueError(f"Failed to get workbook '{book_identifier}': {str(e)}")
    
    @staticmethod
    def open_book(
        path: str, 
        pid: Optional[int] = None, 
        read_only: bool = False, 
        password: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        ワークブックを開きます。

        Args:
            path: ワークブックのパス
            pid: ExcelアプリケーションのプロセスID (オプション)
            read_only: 読み取り専用で開くかどうか
            password: パスワード (オプション)

        Returns:
            開いたワークブック情報

        Raises:
            ValueError: ファイルが見つからないか開けない場合
        """
        if not os.path.exists(path):
            raise ValueError(f"File not found: {path}")
        
        try:
            if pid is not None:
                app = xw.App(pid=pid)
                book = app.books.open(path, read_only=read_only, password=password)
            else:
                # アクティブなアプリケーションを使用するか、新しいアプリケーションを作成
                try:
                    app = xw.apps.active
                    if app is None:
                        raise AttributeError("No active app")
                except (AttributeError, IndexError):
                    app = xw.App(visible=False)
                
                book = app.books.open(path, read_only=read_only, password=password)
            
            return to_serializable(book)
        except Exception as e:
            raise ValueError(f"Failed to open workbook '{path}': {str(e)}")
    
    @staticmethod
    def create_book(pid: Optional[int] = None) -> Dict[str, Any]:
        """
        新しいワークブックを作成します。

        Args:
            pid: ExcelアプリケーションのプロセスID (オプション)

        Returns:
            新しいワークブック情報
        """
        try:
            if pid is not None:
                app = xw.App(pid=pid)
                book = app.books.add()
            else:
                # アクティブなアプリケーションを使用するか、新しいアプリケーションを作成
                try:
                    app = xw.apps.active
                    if app is None:
                        raise AttributeError("No active app")
                except (AttributeError, IndexError):
                    app = xw.App(visible=False)
                
                book = app.books.add()
            
            return to_serializable(book)
        except Exception as e:
            raise ValueError(f"Failed to create workbook: {str(e)}")
    
    @staticmethod
    def close_book(
        book_identifier: str, 
        pid: Optional[int] = None, 
        save: bool = True, 
        path: Optional[str] = None
    ) -> bool:
        """
        ワークブックを閉じます。

        Args:
            book_identifier: ワークブック名かフルパス
            pid: ExcelアプリケーションのプロセスID (オプション)
            save: 変更を保存するかどうか
            path: 保存先パス (オプション)

        Returns:
            成功した場合はTrue

        Raises:
            ValueError: ワークブックが見つからないか閉じられない場合
        """
        try:
            if pid is not None:
                app = xw.App(pid=pid)
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            if save and path:
                book.save(path=path)
            book.close(save=save)
            return True
        except Exception as e:
            raise ValueError(f"Failed to close workbook '{book_identifier}': {str(e)}")
    
    @staticmethod
    def save_book(
        book_identifier: str, 
        pid: Optional[int] = None, 
        path: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        ワークブックを保存します。

        Args:
            book_identifier: ワークブック名かフルパス
            pid: ExcelアプリケーションのプロセスID (オプション)
            path: 保存先パス (オプション)

        Returns:
            保存したワークブック情報

        Raises:
            ValueError: ワークブックが見つからないか保存できない場合
        """
        try:
            if pid is not None:
                app = xw.App(pid=pid)
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            if path:
                book.save(path=path)
            else:
                book.save()
            
            return to_serializable(book)
        except Exception as e:
            raise ValueError(f"Failed to save workbook '{book_identifier}': {str(e)}")
    
    @staticmethod
    def get_book_sheets(
        book_identifier: str, 
        pid: Optional[int] = None
    ) -> List[Dict[str, Any]]:
        """
        ワークブック内のシートを取得します。

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
                app = xw.App(pid=pid)
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            return [to_serializable(sheet) for sheet in book.sheets]
        except Exception as e:
            raise ValueError(f"Failed to get sheets for workbook '{book_identifier}': {str(e)}")
