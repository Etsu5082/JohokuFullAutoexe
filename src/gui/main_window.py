"""メインウィンドウモジュール"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel,
                             QComboBox, QTabWidget, QLineEdit, QTextEdit, QFileDialog,
                             QMessageBox, QGridLayout, QGroupBox, QHBoxLayout, QProgressBar,
                             QCheckBox)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

from ..automation.worker import WorkerThread
from ..utils.helpers import get_writable_dir


class JohokuApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("城北中央公園テニスコート予約システム")
        self.setGeometry(100, 100, 1000, 700)

        # タブウィジェットを作成
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # タブの作成
        self.create_generate_csv_tab()
        self.create_lottery_application_tab()
        self.create_check_lottery_status_tab()
        self.create_lottery_confirm_tab()
        self.create_reservation_check_tab()
        self.create_account_expiry_tab()

        # ワーカースレッド
        self.worker = None

        # フォントの設定
        self.set_font()

    def set_font(self):
        font = QFont()
        font.setPointSize(10)
        self.setFont(font)

    # タブ1: CSVファイル生成
    def create_generate_csv_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        # 説明ラベル
        title_label = QLabel("予約日を配分してCSVファイルを生成")
        title_label.setAlignment(Qt.AlignCenter)
        font = title_label.font()
        font.setPointSize(14)
        font.setBold(True)
        title_label.setFont(font)
        layout.addWidget(title_label)

        layout.addSpacing(10)

        # 入力ファイル選択
        file_layout = QHBoxLayout()
        file_layout.addWidget(QLabel("入力CSVファイル:"))
        self.csv_input_file = QLineEdit("Johoku1.csv")
        file_layout.addWidget(self.csv_input_file)
        self.browse_input_button = QPushButton("参照...")
        self.browse_input_button.clicked.connect(self.browse_input_file)
        file_layout.addWidget(self.browse_input_button)
        layout.addLayout(file_layout)

        # 予約日入力
        dates_label = QLabel("予約日 (YYYY-MM-DD形式、複数の場合はカンマ区切り):")
        layout.addWidget(dates_label)
        self.booking_dates_input = QTextEdit()
        self.booking_dates_input.setMaximumHeight(100)
        layout.addWidget(self.booking_dates_input)

        # 出力ファイル名
        out_layout = QGridLayout()
        out_layout.addWidget(QLabel("出力ファイル1:"), 0, 0)
        self.output_file1 = QLineEdit("Johoku10.csv")
        out_layout.addWidget(self.output_file1, 0, 1)
        out_layout.addWidget(QLabel("出力ファイル2:"), 1, 0)
        self.output_file2 = QLineEdit("Johoku20.csv")
        out_layout.addWidget(self.output_file2, 1, 1)
        layout.addLayout(out_layout)

        # 実行ボタン
        self.generate_button = QPushButton("CSVファイルを生成")
        self.generate_button.setMinimumHeight(40)
        self.generate_button.clicked.connect(self.start_generate_csv)
        layout.addWidget(self.generate_button)

        # プログレスバー
        self.csv_progress = QProgressBar()
        layout.addWidget(self.csv_progress)

        # ログ表示エリア
        log_group = QGroupBox("実行ログ")
        log_layout = QVBoxLayout()
        self.csv_log = QTextEdit()
        self.csv_log.setReadOnly(True)
        log_layout.addWidget(self.csv_log)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        tab.setLayout(layout)
        self.tabs.addTab(tab, "CSVファイル生成")

    # タブ2: 抽選申込
    def create_lottery_application_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        # 説明ラベル
        title_label = QLabel("抽選申込")
        title_label.setAlignment(Qt.AlignCenter)
        font = title_label.font()
        font.setPointSize(14)
        font.setBold(True)
        title_label.setFont(font)
        layout.addWidget(title_label)

        layout.addSpacing(10)

        # CSVファイル選択
        file_layout = QHBoxLayout()
        file_layout.addWidget(QLabel("CSVファイル:"))
        self.lottery_csv_file = QLineEdit("Johoku1.csv")
        file_layout.addWidget(self.lottery_csv_file)
        self.browse_lottery_button = QPushButton("参照...")
        self.browse_lottery_button.clicked.connect(lambda: self.browse_file(self.lottery_csv_file))
        file_layout.addWidget(self.browse_lottery_button)
        layout.addLayout(file_layout)

        # 申込み種類選択
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("申込み種類:"))
        self.apply_type = QComboBox()
        self.apply_type.addItem("申込み1件目")
        self.apply_type.addItem("申込み2件目")
        type_layout.addWidget(self.apply_type)
        layout.addLayout(type_layout)

        # ヘッドレスモード選択（追加）
        self.lottery_headless_checkbox = QCheckBox("ヘッドレスモード（ブラウザ非表示）")
        self.lottery_headless_checkbox.setChecked(True)  # デフォルトはオン
        layout.addWidget(self.lottery_headless_checkbox)

        # 実行ボタン
        self.lottery_button = QPushButton("抽選申込を実行")
        self.lottery_button.setMinimumHeight(40)
        self.lottery_button.clicked.connect(self.start_lottery_application)
        layout.addWidget(self.lottery_button)

        # 停止ボタン
        self.stop_lottery_button = QPushButton("処理を停止")
        self.stop_lottery_button.clicked.connect(self.stop_worker)
        layout.addWidget(self.stop_lottery_button)

        # プログレスバー
        self.lottery_progress = QProgressBar()
        layout.addWidget(self.lottery_progress)

        # スクロール可能なログ表示エリア
        log_group = QGroupBox("実行ログ")
        log_layout = QVBoxLayout()
        self.lottery_log = QTextEdit()
        self.lottery_log.setReadOnly(True)
        log_layout.addWidget(self.lottery_log)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        tab.setLayout(layout)
        self.tabs.addTab(tab, "抽選申込")

    # タブ3: 抽選申込状況の確認
    def create_check_lottery_status_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        # 説明ラベル
        title_label = QLabel("抽選申込状況の確認")
        title_label.setAlignment(Qt.AlignCenter)
        font = title_label.font()
        font.setPointSize(14)
        font.setBold(True)
        title_label.setFont(font)
        layout.addWidget(title_label)

        layout.addSpacing(10)

        # CSVファイル選択
        file_layout = QHBoxLayout()
        file_layout.addWidget(QLabel("CSVファイル:"))
        self.check_status_csv_file = QLineEdit("Johoku1.csv")
        file_layout.addWidget(self.check_status_csv_file)
        self.browse_check_status_button = QPushButton("参照...")
        self.browse_check_status_button.clicked.connect(lambda: self.browse_file(self.check_status_csv_file))
        file_layout.addWidget(self.browse_check_status_button)
        layout.addLayout(file_layout)

        # ヘッドレスモード選択（追加）
        self.check_status_headless_checkbox = QCheckBox("ヘッドレスモード（ブラウザ非表示）")
        self.check_status_headless_checkbox.setChecked(True)  # デフォルトはオン
        layout.addWidget(self.check_status_headless_checkbox)

        # 実行ボタン
        self.check_status_button = QPushButton("申込状況を確認")
        self.check_status_button.setMinimumHeight(40)
        self.check_status_button.clicked.connect(self.start_check_lottery_status)
        layout.addWidget(self.check_status_button)

        # 停止ボタン
        self.stop_check_status_button = QPushButton("処理を停止")
        self.stop_check_status_button.clicked.connect(self.stop_worker)
        layout.addWidget(self.stop_check_status_button)

        # プログレスバー
        self.check_status_progress = QProgressBar()
        layout.addWidget(self.check_status_progress)

        # 結果を表示するボタン
        self.show_check_results_button = QPushButton("確認結果を表示")
        self.show_check_results_button.clicked.connect(lambda: self.show_results_file("reservation_info.txt"))
        layout.addWidget(self.show_check_results_button)

        # スクロール可能なログ表示エリア
        log_group = QGroupBox("実行ログ")
        log_layout = QVBoxLayout()
        self.check_status_log = QTextEdit()
        self.check_status_log.setReadOnly(True)
        log_layout.addWidget(self.check_status_log)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        tab.setLayout(layout)
        self.tabs.addTab(tab, "申込状況確認")

    # タブ4: 抽選確定
    def create_lottery_confirm_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        # 説明ラベル
        title_label = QLabel("抽選確定")
        title_label.setAlignment(Qt.AlignCenter)
        font = title_label.font()
        font.setPointSize(14)
        font.setBold(True)
        title_label.setFont(font)
        layout.addWidget(title_label)

        layout.addSpacing(10)

        # CSVファイル選択
        file_layout = QHBoxLayout()
        file_layout.addWidget(QLabel("CSVファイル:"))
        self.confirm_csv_file = QLineEdit("Johoku1.csv")
        file_layout.addWidget(self.confirm_csv_file)
        self.browse_confirm_button = QPushButton("参照...")
        self.browse_confirm_button.clicked.connect(lambda: self.browse_file(self.confirm_csv_file))
        file_layout.addWidget(self.browse_confirm_button)
        layout.addLayout(file_layout)

        # 利用人数選択
        people_layout = QHBoxLayout()
        people_layout.addWidget(QLabel("利用人数:"))
        self.user_count = QLineEdit("6")
        people_layout.addWidget(self.user_count)
        layout.addLayout(people_layout)

        # ヘッドレスモード選択（追加）
        self.confirm_headless_checkbox = QCheckBox("ヘッドレスモード（ブラウザ非表示）")
        self.confirm_headless_checkbox.setChecked(True)  # デフォルトはオン
        layout.addWidget(self.confirm_headless_checkbox)

        # 実行ボタン
        self.confirm_button = QPushButton("抽選確定処理を実行")
        self.confirm_button.setMinimumHeight(40)
        self.confirm_button.clicked.connect(self.start_confirm_lottery)
        layout.addWidget(self.confirm_button)

        # 停止ボタン
        self.stop_confirm_button = QPushButton("処理を停止")
        self.stop_confirm_button.clicked.connect(self.stop_worker)
        layout.addWidget(self.stop_confirm_button)

        # プログレスバー
        self.confirm_progress = QProgressBar()
        layout.addWidget(self.confirm_progress)

        # 結果を表示するボタン
        self.show_confirm_results_button = QPushButton("確定結果を表示")
        self.show_confirm_results_button.clicked.connect(lambda: self.show_results_file("lottery_results.txt"))
        layout.addWidget(self.show_confirm_results_button)

        # スクロール可能なログ表示エリア
        log_group = QGroupBox("実行ログ")
        log_layout = QVBoxLayout()
        self.confirm_log = QTextEdit()
        self.confirm_log.setReadOnly(True)
        log_layout.addWidget(self.confirm_log)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        tab.setLayout(layout)
        self.tabs.addTab(tab, "抽選確定")

    # タブ5: 予約状況の確認
    def create_reservation_check_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        # 説明ラベル
        title_label = QLabel("予約状況の確認")
        title_label.setAlignment(Qt.AlignCenter)
        font = title_label.font()
        font.setPointSize(14)
        font.setBold(True)
        title_label.setFont(font)
        layout.addWidget(title_label)

        layout.addSpacing(10)

        # CSVファイル選択
        file_layout = QHBoxLayout()
        file_layout.addWidget(QLabel("CSVファイル:"))
        self.reservation_csv_file = QLineEdit("Johoku1.csv")
        file_layout.addWidget(self.reservation_csv_file)
        self.browse_reservation_button = QPushButton("参照...")
        self.browse_reservation_button.clicked.connect(lambda: self.browse_file(self.reservation_csv_file))
        file_layout.addWidget(self.browse_reservation_button)
        layout.addLayout(file_layout)

        # ヘッドレスモード選択（追加）
        self.reservation_headless_checkbox = QCheckBox("ヘッドレスモード（ブラウザ非表示）")
        self.reservation_headless_checkbox.setChecked(True)  # デフォルトはオン
        layout.addWidget(self.reservation_headless_checkbox)

        # 実行ボタン
        self.reservation_button = QPushButton("予約状況を確認")
        self.reservation_button.setMinimumHeight(40)
        self.reservation_button.clicked.connect(self.start_check_reservation)
        layout.addWidget(self.reservation_button)

        # 停止ボタン
        self.stop_reservation_button = QPushButton("処理を停止")
        self.stop_reservation_button.clicked.connect(self.stop_worker)
        layout.addWidget(self.stop_reservation_button)

        # プログレスバー
        self.reservation_progress = QProgressBar()
        layout.addWidget(self.reservation_progress)

        # 結果を表示するボタン
        self.show_reservation_results_button = QPushButton("予約状況結果を表示")
        self.show_reservation_results_button.clicked.connect(lambda: self.show_results_file("r_info.txt"))
        layout.addWidget(self.show_reservation_results_button)

        # スクロール可能なログ表示エリア
        log_group = QGroupBox("実行ログ")
        log_layout = QVBoxLayout()
        self.reservation_log = QTextEdit()
        self.reservation_log.setReadOnly(True)
        log_layout.addWidget(self.reservation_log)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        tab.setLayout(layout)
        self.tabs.addTab(tab, "予約状況確認")

    # タブ6: 有効期限の確認
    def create_account_expiry_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        # 説明ラベル
        title_label = QLabel("有効期限の確認")
        title_label.setAlignment(Qt.AlignCenter)
        font = title_label.font()
        font.setPointSize(14)
        font.setBold(True)
        title_label.setFont(font)
        layout.addWidget(title_label)

        layout.addSpacing(10)

        # CSVファイル選択
        file_layout = QHBoxLayout()
        file_layout.addWidget(QLabel("CSVファイル:"))
        self.expiry_csv_file = QLineEdit("Johoku1.csv")
        file_layout.addWidget(self.expiry_csv_file)
        self.browse_expiry_button = QPushButton("参照...")
        self.browse_expiry_button.clicked.connect(lambda: self.browse_file(self.expiry_csv_file))
        file_layout.addWidget(self.browse_expiry_button)
        layout.addLayout(file_layout)

        # ヘッドレスモード選択（追加）
        self.expiry_headless_checkbox = QCheckBox("ヘッドレスモード（ブラウザ非表示）")
        self.expiry_headless_checkbox.setChecked(True)  # デフォルトはオン
        layout.addWidget(self.expiry_headless_checkbox)

        # 実行ボタン
        self.expiry_button = QPushButton("有効期限を確認")
        self.expiry_button.setMinimumHeight(40)
        self.expiry_button.clicked.connect(self.start_check_expiry)
        layout.addWidget(self.expiry_button)

        # 停止ボタン
        self.stop_expiry_button = QPushButton("処理を停止")
        self.stop_expiry_button.clicked.connect(self.stop_worker)
        layout.addWidget(self.stop_expiry_button)

        # プログレスバー
        self.expiry_progress = QProgressBar()
        layout.addWidget(self.expiry_progress)

        # 結果を表示するボタン
        self.show_expiry_results_button = QPushButton("有効期限結果を表示")
        self.show_expiry_results_button.clicked.connect(lambda: self.show_results_file("expiry.txt"))
        layout.addWidget(self.show_expiry_results_button)

        # スクロール可能なログ表示エリア
        log_group = QGroupBox("実行ログ")
        log_layout = QVBoxLayout()
        self.expiry_log = QTextEdit()
        self.expiry_log.setReadOnly(True)
        log_layout.addWidget(self.expiry_log)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        tab.setLayout(layout)
        self.tabs.addTab(tab, "有効期限確認")

    # ファイル選択ダイアログを表示する関数
    def browse_input_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "入力CSVファイルを選択", "", "CSV Files (*.csv)")
        if file_name:
            self.csv_input_file.setText(file_name)

    def browse_file(self, line_edit):
        file_name, _ = QFileDialog.getOpenFileName(self, "CSVファイルを選択", "", "CSV Files (*.csv)")
        if file_name:
            line_edit.setText(file_name)

    # 結果ファイルを表示する関数
    def show_results_file(self, file_name):
        try:
            # 書き込み可能なディレクトリを取得
            writable_dir = get_writable_dir()
            full_path = os.path.join(writable_dir, file_name)

            if os.path.exists(full_path):
                with open(full_path, 'r', encoding='utf-8') as file:
                    content = file.read()

                dialog = QMainWindow(self)
                dialog.setWindowTitle(f"結果: {file_name}")
                dialog.setGeometry(200, 200, 800, 600)

                text_edit = QTextEdit(dialog)
                text_edit.setReadOnly(True)
                text_edit.setText(content)
                dialog.setCentralWidget(text_edit)

                dialog.show()
            else:
                QMessageBox.warning(self, "ファイルが見つかりません", f"{file_name} が見つかりません。先に処理を実行してください。")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"ファイルを開けませんでした: {str(e)}")

    # ワーカースレッドを停止する関数
    def stop_worker(self):
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            QMessageBox.information(self, "停止", "処理を停止しました。実行中の処理が完了するまでお待ちください。")

    # CSVファイル生成処理を開始する関数
    def start_generate_csv(self):
        input_file = self.csv_input_file.text()
        booking_dates_text = self.booking_dates_input.toPlainText().strip()
        out1 = self.output_file1.text()
        out2 = self.output_file2.text()

        # 入力チェック
        if not input_file:
            QMessageBox.warning(self, "入力エラー", "入力CSVファイルを指定してください。")
            return

        if not booking_dates_text:
            QMessageBox.warning(self, "入力エラー", "予約日を入力してください。")
            return

        # 予約日をリストに変換
        booking_dates = [d.strip() for d in booking_dates_text.split(",")]

        # 日付形式の検証
        for date in booking_dates:
            try:
                datetime.strptime(date, "%Y-%m-%d")
            except ValueError:
                QMessageBox.warning(self, "日付形式エラー", f"無効な日付形式です: {date}\n正しい形式は YYYY-MM-DD (例: 2025-07-05) です。")
                return

        # ログをクリア
        self.csv_log.clear()

        # パラメータを設定
        params = {
            "input_file": input_file,
            "booking_dates": booking_dates,
            "out1": out1,
            "out2": out2
        }

        # ワーカースレッドを作成・起動
        self.worker = WorkerThread("generate_csv", params)
        self.worker.update_signal.connect(lambda msg: self.csv_log.append(msg))
        self.worker.progress_signal.connect(self.csv_progress.setValue)
        self.worker.finished_signal.connect(self.on_worker_finished)

        # ボタンの状態を変更
        self.generate_button.setEnabled(False)

        # スレッドを開始
        self.worker.start()

    # 抽選申込処理を開始する関数
    def start_lottery_application(self):
        csv_file = self.lottery_csv_file.text()
        apply_number_text = self.apply_type.currentText()
        headless = self.lottery_headless_checkbox.isChecked()  # ヘッドレスモード設定を取得

        # 入力チェック
        if not csv_file:
            QMessageBox.warning(self, "入力エラー", "CSVファイルを指定してください。")
            return

        # ファイルの存在確認
        if not os.path.exists(csv_file):
            QMessageBox.warning(self, "ファイルエラー", f"ファイル {csv_file} が見つかりません。")
            return

        # ログをクリア
        self.lottery_log.clear()

        # 確認ダイアログを表示
        message = (f"CSVファイル: {csv_file}\n"
                  f"申込み種類: {apply_number_text}\n"
                  f"ヘッドレスモード: {'有効' if headless else '無効'}\n\n"
                  f"処理を開始しますか？")

        reply = QMessageBox.question(self, "確認", message,
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            # パラメータを設定
            params = {
                "csv_file": csv_file,
                "apply_number_text": apply_number_text,
                "headless": headless  # ヘッドレスモード設定を追加
            }

            # ワーカースレッドを作成・起動
            self.worker = WorkerThread("lottery_application", params)
            self.worker.update_signal.connect(lambda msg: self.lottery_log.append(msg))
            self.worker.progress_signal.connect(self.lottery_progress.setValue)
            self.worker.finished_signal.connect(self.on_worker_finished)

            # ボタンの状態を変更
            self.lottery_button.setEnabled(False)

            # スレッドを開始
            self.worker.start()

    # 抽選申込状況確認処理を開始する関数
    def start_check_lottery_status(self):
        csv_file = self.check_status_csv_file.text()
        headless = self.check_status_headless_checkbox.isChecked()  # ヘッドレスモード設定を取得

        # 入力チェック
        if not csv_file:
            QMessageBox.warning(self, "入力エラー", "CSVファイルを指定してください。")
            return

        # ファイルの存在確認
        if not os.path.exists(csv_file):
            QMessageBox.warning(self, "ファイルエラー", f"ファイル {csv_file} が見つかりません。")
            return

        # ログをクリア
        self.check_status_log.clear()

        # パラメータを設定
        params = {
            "csv_file": csv_file,
            "headless": headless  # ヘッドレスモード設定を追加
        }

        # ワーカースレッドを作成・起動
        self.worker = WorkerThread("check_lottery_status", params)
        self.worker.update_signal.connect(lambda msg: self.check_status_log.append(msg))
        self.worker.progress_signal.connect(self.check_status_progress.setValue)
        self.worker.finished_signal.connect(self.on_worker_finished)

        # ボタンの状態を変更
        self.check_status_button.setEnabled(False)

        # スレッドを開始
        self.worker.start()

    # 抽選確定処理を開始する関数
    def start_confirm_lottery(self):
        csv_file = self.confirm_csv_file.text()
        user_count = self.user_count.text()
        headless = self.confirm_headless_checkbox.isChecked()  # ヘッドレスモード設定を取得

        # 入力チェック
        if not csv_file:
            QMessageBox.warning(self, "入力エラー", "CSVファイルを指定してください。")
            return

        # ファイルの存在確認
        if not os.path.exists(csv_file):
            QMessageBox.warning(self, "ファイルエラー", f"ファイル {csv_file} が見つかりません。")
            return

        # 利用人数のチェック（数値かどうか）
        try:
            int(user_count)
        except ValueError:
            QMessageBox.warning(self, "入力エラー", "利用人数は数値で入力してください。")
            return

        # ログをクリア
        self.confirm_log.clear()

        # 確認ダイアログを表示
        message = (f"CSVファイル: {csv_file}\n"
                  f"利用人数: {user_count}\n"
                  f"ヘッドレスモード: {'有効' if headless else '無効'}\n\n"
                  f"抽選確定処理を開始しますか？")

        reply = QMessageBox.question(self, "確認", message,
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            # パラメータを設定
            params = {
                "csv_file": csv_file,
                "user_count": user_count,
                "headless": headless  # ヘッドレスモード設定を追加
            }

            # ワーカースレッドを作成・起動
            self.worker = WorkerThread("confirm_lottery", params)
            self.worker.update_signal.connect(lambda msg: self.confirm_log.append(msg))
            self.worker.progress_signal.connect(self.confirm_progress.setValue)
            self.worker.finished_signal.connect(self.on_worker_finished)

            # ボタンの状態を変更
            self.confirm_button.setEnabled(False)

            # スレッドを開始
            self.worker.start()

    # 予約状況確認処理を開始する関数
    def start_check_reservation(self):
        csv_file = self.reservation_csv_file.text()
        headless = self.reservation_headless_checkbox.isChecked()  # ヘッドレスモード設定を取得

        # 入力チェック
        if not csv_file:
            QMessageBox.warning(self, "入力エラー", "CSVファイルを指定してください。")
            return

        # ファイルの存在確認
        if not os.path.exists(csv_file):
            QMessageBox.warning(self, "ファイルエラー", f"ファイル {csv_file} が見つかりません。")
            return

        # ログをクリア
        self.reservation_log.clear()

        # パラメータを設定
        params = {
            "csv_file": csv_file,
            "headless": headless  # ヘッドレスモード設定を追加
        }

        # ワーカースレッドを作成・起動
        self.worker = WorkerThread("check_reservation", params)
        self.worker.update_signal.connect(lambda msg: self.reservation_log.append(msg))
        self.worker.progress_signal.connect(self.reservation_progress.setValue)
        self.worker.finished_signal.connect(self.on_worker_finished)

        # ボタンの状態を変更
        self.reservation_button.setEnabled(False)

        # スレッドを開始
        self.worker.start()

    # 有効期限確認処理を開始する関数
    def start_check_expiry(self):
        csv_file = self.expiry_csv_file.text()
        headless = self.expiry_headless_checkbox.isChecked()  # ヘッドレスモード設定を取得

        # 入力チェック
        if not csv_file:
            QMessageBox.warning(self, "入力エラー", "CSVファイルを指定してください。")
            return

        # ファイルの存在確認
        if not os.path.exists(csv_file):
            QMessageBox.warning(self, "ファイルエラー", f"ファイル {csv_file} が見つかりません。")
            return

        # ログをクリア
        self.expiry_log.clear()

        # パラメータを設定
        params = {
            "csv_file": csv_file,
            "headless": headless  # ヘッドレスモード設定を追加
        }

        # ワーカースレッドを作成・起動
        self.worker = WorkerThread("check_expiry", params)
        self.worker.update_signal.connect(lambda msg: self.expiry_log.append(msg))
        self.worker.progress_signal.connect(self.expiry_progress.setValue)
        self.worker.finished_signal.connect(self.on_worker_finished)

        # ボタンの状態を変更
        self.expiry_button.setEnabled(False)

        # スレッドを開始
        self.worker.start()

    # ワーカースレッド終了時の処理
    def on_worker_finished(self, success, message):
        # ボタンの状態を元に戻す
        self.generate_button.setEnabled(True)
        self.lottery_button.setEnabled(True)
        self.check_status_button.setEnabled(True)
        self.confirm_button.setEnabled(True)
        self.reservation_button.setEnabled(True)
        self.expiry_button.setEnabled(True)

        if success:
            QMessageBox.information(self, "完了", message)
        else:
            QMessageBox.warning(self, "エラー", message)
