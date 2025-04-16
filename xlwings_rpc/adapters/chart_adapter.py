"""
チャートアダプターモジュール

xlwingsのChartオブジェクトとAPI間のインターフェースを提供します。
"""
from typing import Dict, List, Optional, Any, Union
import os
import logging
import sys  # プラットフォーム判定のために追加
import time  # 遅延処理のために追加
import xlwings as xw
from xlwings_rpc.utils.converters import to_serializable
from xlwings_rpc.adapters.book_adapter import BookAdapter
from xlwings_rpc.adapters.sheet_adapter import SheetAdapter

# ロガーの設定
logger = logging.getLogger(__name__)

# プラットフォーム別のチャートタイプマッピング（Windows環境名→MacOS環境名）
CHART_TYPE_MAPPING = {
    # 基本的なチャートタイプ
    'column': 'column_clustered',
    'column_stacked': 'column_stacked',
    'column_stacked_100': 'column_stacked_100_percent',
    'bar': 'bar_clustered',
    'bar_stacked': 'bar_stacked',
    'bar_stacked_100': 'bar_stacked_100_percent',
    'line': 'line',
    'line_markers': 'line_markers',
    'line_stacked': 'line_stacked',
    'line_stacked_100': 'line_stacked_100_percent',
    'pie': 'pie',
    'doughnut': 'doughnut',
    'scatter': 'scatter_markers',
    'area': 'area',
    'area_stacked': 'area_stacked',
    'area_stacked_100': 'area_stacked_100_percent',
    'radar': 'radar',
    'bubble': 'bubble',
    # 3D系
    '3d_column': '3d_column',
    '3d_bar': '3d_bar',
    '3d_line': '3d_line',
    '3d_pie': '3d_pie',
    '3d_area': '3d_area',
}

# MacOS環境で利用可能なチャートタイプのリスト
MACOS_CHART_TYPES = [
    'area',
    'area_stacked',
    'area_stacked_100_percent',
    'bar_clustered',
    'bar_stacked',
    'bar_stacked_100_percent',
    'bubble',
    'column_clustered',
    'column_stacked',
    'column_stacked_100_percent',
    'doughnut',
    'line',
    'line_markers',
    'line_stacked',
    'line_stacked_100_percent',
    'pie',
    'radar',
    'scatter_markers',
    '3d_area',
    '3d_bar',
    '3d_column',
    '3d_line',
    '3d_pie',
]


class ChartAdapter:
    """
    xlwingsのChartオブジェクトに対するアダプタークラス
    """
    
    @staticmethod
    def get_platform_chart_type(chart_type: str) -> str:
        """
        プラットフォームに適したチャートタイプ名に変換します。
        
        Args:
            chart_type: 元のチャートタイプ名
            
        Returns:
            変換されたチャートタイプ名
        """
        # MacOS環境の場合のみマッピングを適用
        if sys.platform == 'darwin':
            # まず直接マッピングを試す
            if chart_type in CHART_TYPE_MAPPING:
                mapped_type = CHART_TYPE_MAPPING[chart_type]
                logger.debug(f"チャートタイプをマッピング: '{chart_type}' → '{mapped_type}'")
                return mapped_type
            
            # すでにMacOS形式かもしれないのでそのまま返す
            if chart_type in MACOS_CHART_TYPES:
                return chart_type
            
            # それ以外の場合は警告してそのまま返す
            logger.warning(f"未知のチャートタイプ '{chart_type}' が指定されました")
        
        # Windows環境または未知のタイプの場合はそのまま返す
        return chart_type
    
    @staticmethod
    def get_chart_types() -> List[str]:
        """
        サポートされているチャートタイプのリストを返します。
        
        Returns:
            チャートタイプのリスト
        """
        if sys.platform == 'darwin':
            return MACOS_CHART_TYPES
        else:
            # Windows環境の場合はAPIで指定可能なタイプを返す
            return list(CHART_TYPE_MAPPING.keys())
    
    @staticmethod
    def get_charts(book_identifier: str, sheet_identifier: Union[str, int], pid: Optional[int] = None) -> List[Dict[str, Any]]:
        """
        シート上のすべてのチャートを取得します。

        Args:
            book_identifier: ワークブック名またはフルパス
            sheet_identifier: シート名またはインデックス
            pid: ExcelアプリケーションのプロセスID (デフォルト: None)

        Returns:
            チャート情報のリスト
        
        Raises:
            ValueError: ワークブックまたはシートが見つからない場合
        """
        try:
            # シートを取得
            sheet_info = SheetAdapter.get_sheet(book_identifier, sheet_identifier, pid)
            
            # xlwingsオブジェクトを直接取得
            if pid is not None:
                try:
                    app = xw.apps[pid]
                except KeyError:
                    raise ValueError(f"No Excel application found with PID {pid}")
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            sheet = book.sheets[sheet_identifier]
            
            # シート上のすべてのチャートを取得
            charts = []
            for chart in sheet.charts:
                charts.append(to_serializable(chart))
            
            return charts
        except Exception as e:
            logger.exception(f"Error getting charts: {str(e)}")
            raise ValueError(f"Failed to get charts: {str(e)}")
    
    @staticmethod
    def get_chart(book_identifier: str, sheet_identifier: Union[str, int], chart_identifier: Union[str, int], pid: Optional[int] = None) -> Dict[str, Any]:
        """
        特定のチャートを取得します。

        Args:
            book_identifier: ワークブック名またはフルパス
            sheet_identifier: シート名またはインデックス
            chart_identifier: チャート名またはインデックス
            pid: ExcelアプリケーションのプロセスID (デフォルト: None)

        Returns:
            チャート情報
        
        Raises:
            ValueError: ワークブック、シート、またはチャートが見つからない場合
        """
        try:
            # シートを取得
            sheet_info = SheetAdapter.get_sheet(book_identifier, sheet_identifier, pid)
            
            # xlwingsオブジェクトを直接取得
            if pid is not None:
                try:
                    app = xw.apps[pid]
                except KeyError:
                    raise ValueError(f"No Excel application found with PID {pid}")
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            sheet = book.sheets[sheet_identifier]
            
            # チャートを取得
            try:
                if isinstance(chart_identifier, int):
                    chart = sheet.charts[chart_identifier]
                else:
                    chart = sheet.charts[chart_identifier]
            except (IndexError, KeyError):
                raise ValueError(f"Chart '{chart_identifier}' not found")
            
            return to_serializable(chart)
        except Exception as e:
            logger.exception(f"Error getting chart: {str(e)}")
            raise ValueError(f"Failed to get chart: {str(e)}")
    
    @staticmethod
    def add_chart(book_identifier: str, sheet_identifier: Union[str, int], 
                 left: Optional[float] = None, top: Optional[float] = None, 
                 width: Optional[float] = None, height: Optional[float] = None, 
                 pid: Optional[int] = None) -> Dict[str, Any]:
        """
        新しいチャートを追加します。

        Args:
            book_identifier: ワークブック名またはフルパス
            sheet_identifier: シート名またはインデックス
            left: チャートの左位置 (デフォルト: None)
            top: チャートの上位置 (デフォルト: None)
            width: チャートの幅 (デフォルト: None)
            height: チャートの高さ (デフォルト: None)
            pid: ExcelアプリケーションのプロセスID (デフォルト: None)

        Returns:
            新しいチャート情報
        
        Raises:
            ValueError: ワークブックまたはシートが見つからない場合
        """
        try:
            # シートを取得
            sheet_info = SheetAdapter.get_sheet(book_identifier, sheet_identifier, pid)
            
            # xlwingsオブジェクトを直接取得
            if pid is not None:
                try:
                    app = xw.apps[pid]
                except KeyError:
                    raise ValueError(f"No Excel application found with PID {pid}")
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            sheet = book.sheets[sheet_identifier]
            
            # 位置パラメータの準備
            kwargs = {}
            if left is not None:
                kwargs['left'] = left
            if top is not None:
                kwargs['top'] = top
            if width is not None:
                kwargs['width'] = width
            if height is not None:
                kwargs['height'] = height
            
            # チャートを追加
            chart = sheet.charts.add(**kwargs)
            
            return to_serializable(chart)
        except Exception as e:
            logger.exception(f"Error adding chart: {str(e)}")
            raise ValueError(f"Failed to add chart: {str(e)}")
    
    @staticmethod
    def delete_chart(book_identifier: str, sheet_identifier: Union[str, int], 
                    chart_identifier: Union[str, int], pid: Optional[int] = None) -> bool:
        """
        チャートを削除します。

        Args:
            book_identifier: ワークブック名またはフルパス
            sheet_identifier: シート名またはインデックス
            chart_identifier: チャート名またはインデックス
            pid: ExcelアプリケーションのプロセスID (デフォルト: None)

        Returns:
            成功した場合はTrue
        
        Raises:
            ValueError: ワークブック、シート、またはチャートが見つからない場合
        """
        try:
            # シートを取得
            sheet_info = SheetAdapter.get_sheet(book_identifier, sheet_identifier, pid)
            
            # xlwingsオブジェクトを直接取得
            if pid is not None:
                try:
                    app = xw.apps[pid]
                except KeyError:
                    raise ValueError(f"No Excel application found with PID {pid}")
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            sheet = book.sheets[sheet_identifier]
            
            # チャートを取得
            try:
                if isinstance(chart_identifier, int):
                    chart = sheet.charts[chart_identifier]
                else:
                    chart = sheet.charts[chart_identifier]
            except (IndexError, KeyError):
                raise ValueError(f"Chart '{chart_identifier}' not found")
            
            # チャートを削除
            chart.delete()
            
            return True
        except Exception as e:
            logger.exception(f"Error deleting chart: {str(e)}")
            raise ValueError(f"Failed to delete chart: {str(e)}")
    
    @staticmethod
    def set_source_data(book_identifier: str, sheet_identifier: Union[str, int], 
                       chart_identifier: Union[str, int], range_address: str, 
                       pid: Optional[int] = None) -> Dict[str, Any]:
        """
        チャートのデータソースを設定します。

        Args:
            book_identifier: ワークブック名またはフルパス
            sheet_identifier: シート名またはインデックス
            chart_identifier: チャート名またはインデックス
            range_address: データソースの範囲 (例: 'A1:B10')
            pid: ExcelアプリケーションのプロセスID (デフォルト: None)

        Returns:
            更新されたチャート情報
        
        Raises:
            ValueError: ワークブック、シート、またはチャートが見つからない場合
        """
        try:
            # シートを取得
            sheet_info = SheetAdapter.get_sheet(book_identifier, sheet_identifier, pid)
            
            # xlwingsオブジェクトを直接取得
            if pid is not None:
                try:
                    app = xw.apps[pid]
                except KeyError:
                    raise ValueError(f"No Excel application found with PID {pid}")
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            sheet = book.sheets[sheet_identifier]
            
            # チャートを取得
            try:
                if isinstance(chart_identifier, int):
                    chart = sheet.charts[chart_identifier]
                else:
                    chart = sheet.charts[chart_identifier]
            except (IndexError, KeyError):
                raise ValueError(f"Chart '{chart_identifier}' not found")
            
            # 範囲を解析（シートを指定）
            if ':' in range_address:
                # 範囲が指定されている場合
                data_range = sheet.range(range_address)
            else:
                # 単一セルが指定されている場合、そのセルからデータ範囲を拡張
                data_range = sheet.range(range_address).expand()
            
            # データソースを設定
            chart.set_source_data(data_range)
            
            # 画面更新を強制
            if pid is not None:
                try:
                    app = xw.apps[pid]
                    current_state = app.screen_updating
                    app.screen_updating = False
                    app.screen_updating = True
                except Exception as e:
                    logger.warning(f"Failed to force screen update after setting data source: {str(e)}")
            
            return to_serializable(chart)
        except Exception as e:
            logger.exception(f"Error setting chart source data: {str(e)}")
            raise ValueError(f"Failed to set chart source data: {str(e)}")
    
    @staticmethod
    def set_chart_type(book_identifier: str, sheet_identifier: Union[str, int], 
                      chart_identifier: Union[str, int], chart_type: str, 
                      pid: Optional[int] = None) -> Dict[str, Any]:
        """
        チャートのタイプを設定します。

        Args:
            book_identifier: ワークブック名またはフルパス
            sheet_identifier: シート名またはインデックス
            chart_identifier: チャート名またはインデックス
            chart_type: チャートタイプ
            pid: ExcelアプリケーションのプロセスID (デフォルト: None)

        Returns:
            更新されたチャート情報
        
        Raises:
            ValueError: ワークブック、シート、チャート、または無効なチャートタイプが指定された場合
        """
        app = None
        try:
            # シートを取得
            sheet_info = SheetAdapter.get_sheet(book_identifier, sheet_identifier, pid)
            
            # xlwingsオブジェクトを直接取得
            if pid is not None:
                try:
                    app = xw.apps[pid]
                except KeyError:
                    raise ValueError(f"No Excel application found with PID {pid}")
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
                if hasattr(book, 'app'):
                    app = book.app
            
            sheet = book.sheets[sheet_identifier]
            
            # チャートを取得
            try:
                if isinstance(chart_identifier, int):
                    chart = sheet.charts[chart_identifier]
                else:
                    chart = sheet.charts[chart_identifier]
            except (IndexError, KeyError):
                raise ValueError(f"Chart '{chart_identifier}' not found")
            
            # プラットフォームに適したチャートタイプに変換
            platform_chart_type = ChartAdapter.get_platform_chart_type(chart_type)
            
            logger.debug(f"設定するチャートタイプ: '{platform_chart_type}'")
            
            # 一時的に画面更新を無効化（変更の前に）
            screen_updating_state = None
            if app is not None:
                try:
                    screen_updating_state = app.screen_updating
                    app.screen_updating = False
                    logger.debug("画面更新を一時的に無効化")
                except Exception as e:
                    logger.warning(f"Failed to disable screen updating: {str(e)}")
            
            # チャートタイプを設定
            try:
                chart.chart_type = platform_chart_type
                logger.debug(f"チャートタイプを設定: {platform_chart_type}")
                
                # 短い遅延を追加
                time.sleep(0.5)
                logger.debug("遅延を追加: 0.5秒")
                
                # チャートタイプをもう一度設定（一部の環境で必要な場合がある）
                chart.chart_type = platform_chart_type
                logger.debug("チャートタイプを再設定")
                
                # もう少し長い遅延
                time.sleep(0.5)
                logger.debug("追加の遅延: 0.5秒")
                
            except Exception as e:
                # より詳細なエラーメッセージを提供
                if sys.platform == 'darwin':
                    error_msg = (f"Failed to set chart type: '{chart_type}' (mapped to '{platform_chart_type}'). "
                               f"Supported chart types on MacOS: {', '.join(MACOS_CHART_TYPES)}")
                else:
                    error_msg = f"Failed to set chart type: {str(e)}"
                raise ValueError(error_msg)
            finally:
                # 画面更新を元の状態に戻す
                if app is not None and screen_updating_state is not None:
                    try:
                        app.screen_updating = screen_updating_state
                        # 強制的に画面を更新
                        app.screen_updating = False
                        app.screen_updating = True
                        logger.debug("画面更新を強制")
                    except Exception as e:
                        logger.warning(f"Failed to restore screen updating: {str(e)}")
            
            return to_serializable(chart)
        except Exception as e:
            if not isinstance(e, ValueError):
                logger.exception(f"Error setting chart type: {str(e)}")
                raise ValueError(f"Failed to set chart type: {str(e)}")
            raise
    
    @staticmethod
    def export_chart_as_pdf(book_identifier: str, sheet_identifier: Union[str, int], 
                           chart_identifier: Union[str, int], path: Optional[str] = None, 
                           quality: str = "standard", pid: Optional[int] = None) -> Dict[str, Any]:
        """
        チャートをPDFとしてエクスポートします。

        Args:
            book_identifier: ワークブック名またはフルパス
            sheet_identifier: シート名またはインデックス
            chart_identifier: チャート名またはインデックス
            path: 保存先パス (デフォルト: None)
            quality: PDFの品質 ('standard' または 'minimum', デフォルト: 'standard')
            pid: ExcelアプリケーションのプロセスID (デフォルト: None)

        Returns:
            保存されたPDFのパス情報
        
        Raises:
            ValueError: ワークブック、シート、チャートが見つからない場合、または無効な品質が指定された場合
        """
        if quality not in ["standard", "minimum"]:
            raise ValueError(f"Invalid quality value: '{quality}'. Valid values are: 'standard', 'minimum'")
        
        try:
            # シートを取得
            sheet_info = SheetAdapter.get_sheet(book_identifier, sheet_identifier, pid)
            
            # xlwingsオブジェクトを直接取得
            if pid is not None:
                try:
                    app = xw.apps[pid]
                except KeyError:
                    raise ValueError(f"No Excel application found with PID {pid}")
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            sheet = book.sheets[sheet_identifier]
            
            # チャートを取得
            try:
                if isinstance(chart_identifier, int):
                    chart = sheet.charts[chart_identifier]
                else:
                    chart = sheet.charts[chart_identifier]
            except (IndexError, KeyError):
                raise ValueError(f"Chart '{chart_identifier}' not found")
            
            # PDFとしてエクスポート
            result_path = chart.to_pdf(path, quality=quality)
            
            return {"path": result_path}
        except Exception as e:
            logger.exception(f"Error exporting chart as PDF: {str(e)}")
            raise ValueError(f"Failed to export chart as PDF: {str(e)}")
    
    @staticmethod
    def export_chart_as_picture(book_identifier: str, sheet_identifier: Union[str, int], 
                               chart_identifier: Union[str, int], path: Optional[str] = None, 
                               pid: Optional[int] = None) -> Dict[str, Any]:
        """
        チャートを画像としてエクスポートします。

        Args:
            book_identifier: ワークブック名またはフルパス
            sheet_identifier: シート名またはインデックス
            chart_identifier: チャート名またはインデックス
            path: 保存先パス (デフォルト: None)
            pid: ExcelアプリケーションのプロセスID (デフォルト: None)

        Returns:
            保存された画像のパス情報
        
        Raises:
            ValueError: ワークブック、シート、またはチャートが見つからない場合
        """
        try:
            # シートを取得
            sheet_info = SheetAdapter.get_sheet(book_identifier, sheet_identifier, pid)
            
            # xlwingsオブジェクトを直接取得
            if pid is not None:
                try:
                    app = xw.apps[pid]
                except KeyError:
                    raise ValueError(f"No Excel application found with PID {pid}")
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            sheet = book.sheets[sheet_identifier]
            
            # チャートを取得
            try:
                if isinstance(chart_identifier, int):
                    chart = sheet.charts[chart_identifier]
                else:
                    chart = sheet.charts[chart_identifier]
            except (IndexError, KeyError):
                raise ValueError(f"Chart '{chart_identifier}' not found")
            
            # 画像としてエクスポート
            result_path = chart.to_png(path)
            
            return {"path": result_path}
        except Exception as e:
            logger.exception(f"Error exporting chart as picture: {str(e)}")
            raise ValueError(f"Failed to export chart as picture: {str(e)}")
    
    @staticmethod
    def customize_chart(book_identifier: str, sheet_identifier: Union[str, int], 
                       chart_identifier: Union[str, int], settings: Dict[str, Any], 
                       pid: Optional[int] = None) -> Dict[str, Any]:
        """
        チャートの詳細設定をカスタマイズします。
        Windows環境でより多くの設定が利用可能です。

        Args:
            book_identifier: ワークブック名またはフルパス
            sheet_identifier: シート名またはインデックス
            chart_identifier: チャート名またはインデックス
            settings: カスタマイズ設定の辞書
                - title: チャートタイトル
                - has_legend: 凡例を表示するかどうか
                - legend_position: 凡例の位置
                - axis_x: X軸の設定
                - axis_y: Y軸の設定
            pid: ExcelアプリケーションのプロセスID (デフォルト: None)

        Returns:
            更新されたチャート情報
        
        Raises:
            ValueError: ワークブック、シート、またはチャートが見つからない場合
        """
        try:
            # シートを取得
            sheet_info = SheetAdapter.get_sheet(book_identifier, sheet_identifier, pid)
            
            # xlwingsオブジェクトを直接取得
            if pid is not None:
                try:
                    app = xw.apps[pid]
                except KeyError:
                    raise ValueError(f"No Excel application found with PID {pid}")
                book = app.books[book_identifier]
            else:
                book = xw.Book(book_identifier)
            
            sheet = book.sheets[sheet_identifier]
            
            # チャートを取得
            try:
                if isinstance(chart_identifier, int):
                    chart = sheet.charts[chart_identifier]
                else:
                    chart = sheet.charts[chart_identifier]
            except (IndexError, KeyError):
                raise ValueError(f"Chart '{chart_identifier}' not found")
            
            # 各種設定を適用
            # 基本設定（xlwingsの標準API）
            if 'name' in settings:
                chart.name = settings['name']
            
            if 'chart_type' in settings:
                chart_type = settings['chart_type']
                # プラットフォームに適したチャートタイプに変換
                platform_chart_type = ChartAdapter.get_platform_chart_type(chart_type)
                chart.chart_type = platform_chart_type
                
                # 短い遅延を追加
                time.sleep(0.5)
            
            # 位置と大きさ
            if 'left' in settings:
                chart.left = settings['left']
            if 'top' in settings:
                chart.top = settings['top']
            if 'width' in settings:
                chart.width = settings['width']
            if 'height' in settings:
                chart.height = settings['height']
            
            # 拡張設定（COMオブジェクト経由）
            try:
                # タイトル設定
                if 'title' in settings:
                    if sys.platform == 'darwin':
                        # MacOS環境
                        try:
                            chart.api.chart_title.set(settings['title'])
                        except Exception as e:
                            logger.warning(f"Failed to set chart title on MacOS: {str(e)}")
                    else:
                        # Windows環境
                        chart.api[1].SetElement(2)  # 2 = タイトル
                        chart.api[1].ChartTitle.Text = settings['title']
                
                # 凡例設定
                if 'has_legend' in settings:
                    if sys.platform == 'darwin':
                        # MacOS環境
                        try:
                            chart.api.has_legend.set(settings['has_legend'])
                        except Exception as e:
                            logger.warning(f"Failed to set legend on MacOS: {str(e)}")
                    else:
                        # Windows環境
                        chart.api[1].HasLegend = settings['has_legend']
                
                if 'legend_position' in settings and 'has_legend' in settings and settings['has_legend']:
                    position_map = {
                        'bottom': 'bottom',
                        'corner': 'corner',
                        'left': 'left',
                        'right': 'right',
                        'top': 'top'
                    }
                    
                    if settings['legend_position'] in position_map:
                        position = position_map[settings['legend_position']]
                        if sys.platform == 'darwin':
                            # MacOS環境
                            try:
                                chart.api.legend_position.set(position)
                            except Exception as e:
                                logger.warning(f"Failed to set legend position on MacOS: {str(e)}")
                        else:
                            # Windows環境
                            position_codes = {
                                'bottom': -4107,
                                'corner': 2,
                                'left': -4131,
                                'right': -4152,
                                'top': -4160
                            }
                            chart.api[1].Legend.Position = position_codes[settings['legend_position']]
                
                # X軸設定
                if 'axis_x' in settings:
                    axis_x = settings['axis_x']
                    
                    if 'title' in axis_x:
                        if sys.platform == 'darwin':
                            # MacOS環境
                            try:
                                chart.api.axes("category").title.set(axis_x['title'])
                            except Exception as e:
                                logger.warning(f"Failed to set X axis title on MacOS: {str(e)}")
                        else:
                            # Windows環境
                            chart.api[1].Axes(1).HasTitle = True
                            chart.api[1].Axes(1).AxisTitle.Text = axis_x['title']
                    
                    if 'min' in axis_x:
                        if sys.platform == 'darwin':
                            # MacOS環境
                            try:
                                chart.api.axes("category").minimum_scale.set(axis_x['min'])
                            except Exception as e:
                                logger.warning(f"Failed to set X axis minimum on MacOS: {str(e)}")
                        else:
                            # Windows環境
                            chart.api[1].Axes(1).MinimumScale = axis_x['min']
                    
                    if 'max' in axis_x:
                        if sys.platform == 'darwin':
                            # MacOS環境
                            try:
                                chart.api.axes("category").maximum_scale.set(axis_x['max'])
                            except Exception as e:
                                logger.warning(f"Failed to set X axis maximum on MacOS: {str(e)}")
                        else:
                            # Windows環境
                            chart.api[1].Axes(1).MaximumScale = axis_x['max']
                
                # Y軸設定
                if 'axis_y' in settings:
                    axis_y = settings['axis_y']
                    
                    if 'title' in axis_y:
                        if sys.platform == 'darwin':
                            # MacOS環境
                            try:
                                chart.api.axes("value").title.set(axis_y['title'])
                            except Exception as e:
                                logger.warning(f"Failed to set Y axis title on MacOS: {str(e)}")
                        else:
                            # Windows環境
                            chart.api[1].Axes(2).HasTitle = True
                            chart.api[1].Axes(2).AxisTitle.Text = axis_y['title']
                    
                    if 'min' in axis_y:
                        if sys.platform == 'darwin':
                            # MacOS環境
                            try:
                                chart.api.axes("value").minimum_scale.set(axis_y['min'])
                            except Exception as e:
                                logger.warning(f"Failed to set Y axis minimum on MacOS: {str(e)}")
                        else:
                            # Windows環境
                            chart.api[1].Axes(2).MinimumScale = axis_y['min']
                    
                    if 'max' in axis_y:
                        if sys.platform == 'darwin':
                            # MacOS環境
                            try:
                                chart.api.axes("value").maximum_scale.set(axis_y['max'])
                            except Exception as e:
                                logger.warning(f"Failed to set Y axis maximum on MacOS: {str(e)}")
                        else:
                            # Windows環境
                            chart.api[1].Axes(2).MaximumScale = axis_y['max']
            except Exception as e:
                # APIアクセスエラーは警告としてログに記録するが処理は続行
                logger.warning(f"Could not apply extended chart settings: {str(e)}")
            
            # 画面更新を強制
            if pid is not None:
                try:
                    app = xw.apps[pid]
                    current_state = app.screen_updating
                    app.screen_updating = False
                    app.screen_updating = True
                except Exception as e:
                    logger.warning(f"Failed to force screen update after customizing chart: {str(e)}")
            
            return to_serializable(chart)
        except Exception as e:
            logger.exception(f"Error customizing chart: {str(e)}")
            raise ValueError(f"Failed to customize chart: {str(e)}")
