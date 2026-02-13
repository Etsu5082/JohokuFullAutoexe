"""
城北中央公園テニスコート予約システム
メインエントリーポイント
"""
import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QIcon

from src.gui.main_window import JohokuApp


def main():
    """メインアプリケーションを起動"""
    app = QApplication(sys.argv)

    # アプリケーションスタイルを設定（オプション - システムに応じたスタイルを適用）
    app.setStyle("Fusion")

    # アプリケーションアイコン設定（アイコンファイルがある場合）
    try:
        app.setWindowIcon(QIcon("app_icon.ico"))
    except:
        pass

    # メインウィンドウを作成して表示
    window = JohokuApp()
    window.show()

    # イベントループを開始
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
