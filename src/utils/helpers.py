"""ヘルパー関数モジュール"""
import os
from PyQt5.QtCore import QStandardPaths


def get_writable_dir():
    """書き込み可能なディレクトリを取得する"""
    try:
        # ユーザーのドキュメントディレクトリを試す
        docs_dir = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)
        app_dir = os.path.join(docs_dir, "JohokuTennisApp")

        # ディレクトリが存在しない場合は作成
        if not os.path.exists(app_dir):
            os.makedirs(app_dir)

        # 書き込み可能かテスト
        test_file = os.path.join(app_dir, "test_write.tmp")
        try:
            with open(test_file, 'w') as f:
                f.write("test")
            os.remove(test_file)
            return app_dir
        except:
            pass
    except:
        pass

    # ドキュメントディレクトリが使用できない場合は一時ディレクトリを使用
    try:
        import tempfile
        temp_dir = tempfile.gettempdir()
        app_dir = os.path.join(temp_dir, "JohokuTennisApp")

        if not os.path.exists(app_dir):
            os.makedirs(app_dir)

        return app_dir
    except:
        # 最後の手段としてカレントディレクトリを返す
        return os.getcwd()
