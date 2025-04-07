"""
xlwings-rpc メインエントリポイント

コマンドラインからサーバーを起動するためのエントリポイントを提供します。
"""
import argparse
import logging
from xlwings_rpc.server import start_server


def main():
    """
    コマンドライン引数を解析し、サーバーを起動します。
    """
    # ロガーの設定
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    
    logger = logging.getLogger("xlwings-rpc")
    
    # コマンドライン引数の解析
    parser = argparse.ArgumentParser(description="xlwings-rpc server")
    parser.add_argument("--host", default="127.0.0.1", help="Host address to bind")
    parser.add_argument("--port", type=int, default=8000, help="Port to bind")
    parser.add_argument("--log-level", choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
                        default="INFO", help="Logging level")
    
    args = parser.parse_args()
    
    # ログレベルの設定
    logging.getLogger().setLevel(getattr(logging, args.log_level))
    
    logger.info(f"Starting xlwings-rpc server on {args.host}:{args.port}")
    
    # サーバーの起動
    try:
        start_server(args.host, args.port)
    except Exception as e:
        logger.error(f"Error starting server: {str(e)}")
        return 1
    
    return 0


if __name__ == "__main__":
    exit(main())
