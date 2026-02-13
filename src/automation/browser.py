"""ブラウザ設定モジュール"""
from selenium import webdriver


def setup_chrome_options(headless=True):
    """Chromeブラウザのオプションを設定する関数"""
    options = webdriver.ChromeOptions()

    if headless:
        # ヘッドレスモードを有効化
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')  # WindowsでのGPU問題回避
        options.add_argument('--no-sandbox')  # 一部環境でのサンドボックス問題回避
        options.add_argument('--disable-dev-shm-usage')  # 共有メモリ不足問題回避

        # ウィンドウサイズを明示的に設定（ヘッドレスモードではデフォルトが小さい場合あり）
        options.add_argument('--window-size=1920,1080')

        # ユーザーエージェントを設定（ヘッドレスの検出を避けるため）
        options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36')

    # 一般的な設定
    options.add_argument('--disable-extensions')  # 拡張機能を無効化
    options.add_argument('--disable-popup-blocking')  # ポップアップブロックを無効化

    return options
