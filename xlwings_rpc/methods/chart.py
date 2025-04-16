"""
チャート関連のRPCメソッド

Excel Chartに関連するJSON-RPCメソッドを実装します。
"""
from typing import Dict, List, Optional, Any, Union
from xlwings_rpc.adapters.chart_adapter import ChartAdapter


class ChartMethods:
    """
    chart.* 名前空間のRPCメソッド実装
    """
    
    @staticmethod
    async def list(params: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        chart.list: シート上のすべてのチャートを取得します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            チャート情報のリスト
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        pid = params.get("pid")
        return ChartAdapter.get_charts(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            pid=pid
        )
    
    @staticmethod
    async def get(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        chart.get: 特定のチャートを取得します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - chart (Union[str, int]): チャート名かインデックス
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            チャート情報
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        chart_identifier = params["chart"]
        pid = params.get("pid")
        return ChartAdapter.get_chart(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            chart_identifier=chart_identifier,
            pid=pid
        )
    
    @staticmethod
    async def add(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        chart.add: 新しいチャートを追加します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - left (Optional[float]): チャートの左位置
                - top (Optional[float]): チャートの上位置
                - width (Optional[float]): チャートの幅
                - height (Optional[float]): チャートの高さ
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            新しいチャート情報
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        left = params.get("left")
        top = params.get("top")
        width = params.get("width")
        height = params.get("height")
        pid = params.get("pid")
        return ChartAdapter.add_chart(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            left=left,
            top=top,
            width=width,
            height=height,
            pid=pid
        )
    
    @staticmethod
    async def delete(params: Dict[str, Any]) -> bool:
        """
        chart.delete: チャートを削除します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - chart (Union[str, int]): チャート名かインデックス
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            成功した場合はTrue
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        chart_identifier = params["chart"]
        pid = params.get("pid")
        return ChartAdapter.delete_chart(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            chart_identifier=chart_identifier,
            pid=pid
        )
    
    @staticmethod
    async def set_source_data(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        chart.set_source_data: チャートのデータソースを設定します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - chart (Union[str, int]): チャート名かインデックス
                - range (str): データソースの範囲 (例: 'A1:B10')
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            更新されたチャート情報
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        chart_identifier = params["chart"]
        range_address = params["range"]
        pid = params.get("pid")
        return ChartAdapter.set_source_data(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            chart_identifier=chart_identifier,
            range_address=range_address,
            pid=pid
        )
    
    @staticmethod
    async def set_chart_type(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        chart.set_chart_type: チャートのタイプを設定します。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - chart (Union[str, int]): チャート名かインデックス
                - type (str): チャートタイプ
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            更新されたチャート情報
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        chart_identifier = params["chart"]
        chart_type = params["type"]
        pid = params.get("pid")
        return ChartAdapter.set_chart_type(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            chart_identifier=chart_identifier,
            chart_type=chart_type,
            pid=pid
        )
    
    @staticmethod
    async def export_as_pdf(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        chart.export_as_pdf: チャートをPDFとしてエクスポートします。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - chart (Union[str, int]): チャート名かインデックス
                - path (Optional[str]): 保存先パス
                - quality (Optional[str]): PDFの品質 ('standard' または 'minimum', デフォルト: 'standard')
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            保存されたPDFのパス情報
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        chart_identifier = params["chart"]
        path = params.get("path")
        quality = params.get("quality", "standard")
        pid = params.get("pid")
        return ChartAdapter.export_chart_as_pdf(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            chart_identifier=chart_identifier,
            path=path,
            quality=quality,
            pid=pid
        )
    
    @staticmethod
    async def export_as_picture(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        chart.export_as_picture: チャートを画像としてエクスポートします。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - chart (Union[str, int]): チャート名かインデックス
                - path (Optional[str]): 保存先パス
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            保存された画像のパス情報
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        chart_identifier = params["chart"]
        path = params.get("path")
        pid = params.get("pid")
        return ChartAdapter.export_chart_as_picture(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            chart_identifier=chart_identifier,
            path=path,
            pid=pid
        )
    
    @staticmethod
    async def customize(params: Dict[str, Any]) -> Dict[str, Any]:
        """
        chart.customize: チャートの詳細設定をカスタマイズします。
        Windows環境でより多くの設定が利用可能です。

        Args:
            params: パラメータオブジェクト
                - book (str): ワークブック名かフルパス
                - sheet (Union[str, int]): シート名かインデックス
                - chart (Union[str, int]): チャート名かインデックス
                - settings (Dict[str, Any]): カスタマイズ設定
                    - title: チャートタイトル
                    - has_legend: 凡例を表示するかどうか
                    - legend_position: 凡例の位置
                    - axis_x: X軸の設定
                    - axis_y: Y軸の設定
                - pid (Optional[int]): ExcelアプリケーションのプロセスID

        Returns:
            更新されたチャート情報
        """
        book_identifier = params["book"]
        sheet_identifier = params["sheet"]
        chart_identifier = params["chart"]
        settings = params["settings"]
        pid = params.get("pid")
        return ChartAdapter.customize_chart(
            book_identifier=book_identifier,
            sheet_identifier=sheet_identifier,
            chart_identifier=chart_identifier,
            settings=settings,
            pid=pid
        )
