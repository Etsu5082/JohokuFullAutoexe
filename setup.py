"""
城北中央公園テニスコート予約システム用セットアップスクリプト
py2appを使用してMac用アプリケーションを作成します
"""
from setuptools import setup
import os
import PyQt5

pyqt_path = os.path.dirname(PyQt5.__file__)
plugins_path = os.path.join(pyqt_path, 'Qt5', 'plugins')

APP = ['johoku_app.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': False,
    'packages': ['PyQt5', 'pandas', 'selenium', 'webdriver_manager'],
    'includes': ['sip', 'PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtWidgets'],
    'excludes': ['tkinter', 'matplotlib', 'scipy'],
    'qt_plugins': plugins_path,
    'plist': {
        'CFBundleName': '城北中央公園テニスコート予約',
        'CFBundleDisplayName': '城北中央公園テニスコート予約',
        'CFBundleGetInfoString': '城北中央公園テニスコート予約システム',
        'CFBundleIdentifier': 'com.yourcompany.johokuapp',
        'CFBundleVersion': '1.0.0',
        'CFBundleShortVersionString': '1.0.0',
        'NSHumanReadableCopyright': '© 2025 Your Name',
        'NSHighResolutionCapable': True,
    },
    #'iconfile': 'app_icon.icns',  # アイコンファイルがある場合はパスを指定
}

setup(
    name='城北中央公園テニスコート予約',
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)