import sys
import os
import pandas as pd
import time
import random
import calendar
import re
from datetime import datetime
from collections import defaultdict
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QLabel, QComboBox, QTabWidget, 
                             QLineEdit, QTextEdit, QFileDialog, QMessageBox,
                             QGridLayout, QGroupBox, QHBoxLayout, QDateEdit,
                             QTableWidget, QTableWidgetItem, QHeaderView, QProgressBar,
                             QScrollArea, QSizePolicy, QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QDate, QStandardPaths
from PyQt5.QtGui import QFont, QIcon

# Seleniumのインポート
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# ウェブサイトのURL（グローバル変数として定義）
URL = "https://kouen.sports.metro.tokyo.lg.jp/web/"

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

# 書き込み可能なディレクトリを取得する関数
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

# バックグラウンド処理用のスレッドクラス
# バックグラウンド処理用のスレッドクラス
class WorkerThread(QThread):
    update_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(bool, str)
    
    def __init__(self, task_type, params=None):
        super().__init__()
        self.task_type = task_type
        self.params = params if params else {}
        self.is_running = True
    
    def run(self):
        try:
            if self.task_type == "generate_csv":
                self.generate_csv_files()
            elif self.task_type == "lottery_application":
                self.run_lottery_application()
            elif self.task_type == "check_lottery_status":
                self.check_lottery_status()
            elif self.task_type == "confirm_lottery":
                self.confirm_lottery_selection()
            elif self.task_type == "check_reservation":
                self.check_reservation_status()
            elif self.task_type == "check_expiry":
                self.check_account_expiry()
            
            self.finished_signal.emit(True, "処理が正常に完了しました。")
        except Exception as e:
            self.update_signal.emit(f"エラーが発生しました: {str(e)}")
            self.finished_signal.emit(False, f"エラーが発生しました: {str(e)}")
    
    def stop(self):
        self.is_running = False
        self.wait()
    
    # CSVファイル生成機能
    def generate_csv_files(self):
        try:
            input_file = self.params.get("input_file", "Johoku1.csv")
            booking_dates = self.params.get("booking_dates", [])
            out1 = self.params.get("out1", "Johoku10.csv")
            out2 = self.params.get("out2", "Johoku20.csv")
            
            self.update_signal.emit(f"入力ファイル {input_file} を読み込んでいます...")
            
            if not os.path.exists(input_file):
                self.update_signal.emit(f"{input_file} が見つかりません。")
                self.finished_signal.emit(False, f"{input_file} が見つかりません。")
                return
            
            df = pd.read_csv(input_file)
            if len(df) == 0:
                self.update_signal.emit("ユーザーCSVが空です。")
                self.finished_signal.emit(False, "ユーザーCSVが空です。")
                return
            
            self.update_signal.emit(f"{len(df)}人のユーザー情報を読み込みました。")
            self.update_signal.emit(f"予約日を分配します: {booking_dates}")
            
            # 予約日を分配する関数
            df_all = self.distribute_dates(df, booking_dates)
            
            # 分割したCSVを保存する
            user_count = len(df)
            df1 = df_all.iloc[:user_count]
            df2 = df_all.iloc[user_count:]
            
            df1.to_csv(out1, index=False)
            df2.to_csv(out2, index=False)
            
            self.update_signal.emit(f"出力完了:\n{out1}\n{out2}")
        except Exception as e:
            self.update_signal.emit(f"CSV生成中にエラーが発生しました: {str(e)}")
            raise
    
    # 予約日を分配する関数
    def distribute_dates(self, base_df, booking_dates):
        df_all = pd.concat([base_df.copy(), base_df.copy()], ignore_index=True)
        total = len(df_all)
        base = total // len(booking_dates)
        remainder = total % len(booking_dates)
        distribution = [base + (1 if i < remainder else 0) for i in range(len(booking_dates))]
        
        self.update_signal.emit(f"日付分配: 合計{total}人を{len(booking_dates)}日に分配します")
        self.update_signal.emit(f"各日付の予約数: {distribution}")

        new_dates = []
        for date, count in zip(booking_dates, distribution):
            new_dates.extend([date] * count)
        df_all["booking_date"] = new_dates
        return df_all
    
    # 抽選申込の実行
    def run_lottery_application(self):
        csv_file = self.params.get("csv_file", "Johoku1.csv")
        apply_number_text = self.params.get("apply_number_text", "申込み1件目")
        headless = self.params.get("headless", True)  # ヘッドレスモード設定
        
        self.update_signal.emit(f"CSVファイル {csv_file} から予約情報を読み込んでいます...")
        self.update_signal.emit(f"ヘッドレスモード: {'有効' if headless else '無効'}")
        
        # CSVからデータを読み込み
        users_data = pd.read_csv(csv_file, dtype={
            'user_number': str,
            'password': str,
            'booking_date': str,
            'time_code': str
        })
        
        total_users = len(users_data)
        self.update_signal.emit(f"{total_users}人のユーザー情報を読み込みました。")
        
        # Chromeブラウザの起動
        self.update_signal.emit("Chromeブラウザを起動しています...")
        options = setup_chrome_options(headless)  # ヘッドレスモード設定を渡す
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        driver.get("about:blank")
        
        try:
            for index, row in users_data.iterrows():
                if not self.is_running:
                    self.update_signal.emit("処理が中断されました。")
                    break
                
                user_number = row['user_number']
                password = row['password']
                booking_date = row['booking_date']
                time_code = row['time_code']
                
                # booking_date を正しく分解（例: 2025-05-02 -> 年=2025, 月=5, 日=2）
                date_parts = booking_date.split('-')
                year = int(date_parts[0])
                month = int(date_parts[1])
                booking_day = int(date_parts[2])
                
                # 月の最終日を取得
                month_end = calendar.monthrange(year, month)[1]

                progress = int((index / total_users) * 100)
                self.progress_signal.emit(progress)
                
                self.update_signal.emit(f"\nユーザー {user_number} の予約処理を開始します... ({index+1}/{total_users})")
                self.update_signal.emit(f"予約日: {year}年{month}月{booking_day}日, 月末: {month_end}日")
                self.update_signal.emit(f"申込み種類: {apply_number_text}")

                # 新しいタブを開く
                driver.execute_script("window.open('');")
                new_tab = driver.window_handles[-1]
                driver.switch_to.window(new_tab)

                # 選択された申込み種類を使用
                success = self.handle_booking_process(driver, user_number, password, booking_day, time_code, apply_number_text, month_end)
                
                if success:
                    self.update_signal.emit(f"ユーザー {user_number} の全処理が完了しました。")
                else:
                    self.update_signal.emit(f"ユーザー {user_number} の処理は失敗しました。次のユーザーに進みます。")

                # エラーが発生していた場合に備えて、タブの状態を確認・修復
                try:
                    current_handle = driver.current_window_handle
                except:
                    # ブラウザを再起動
                    try:
                        driver.quit()
                    except:
                        pass
                    # Chromeブラウザの起動(再)
                    options = setup_chrome_options(headless)  # ヘッドレスモード設定を渡す
                    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
                    driver.get("about:blank")

                # ユーザー間の待機時間
                time.sleep(random.uniform(1.0, 3.0))
            
            # 最終的な進捗状況を100%に設定
            self.progress_signal.emit(100)
            self.update_signal.emit("全ての予約処理が完了しました。")
            
        except Exception as e:
            self.update_signal.emit(f"実行中にエラーが発生しました: {str(e)}")
            raise
        finally:
            try:
                driver.quit()
            except:
                pass

    # 既存の機能を呼び出す実装部分（元のスクリプトから必要な関数を実装）
    def human_like_mouse_move(self, driver, element):
        """より人間らしいマウスの動きをシミュレート"""
        actions = ActionChains(driver)
        
        # 現在のマウス位置から要素まで、途中で数回停止しながら移動
        for _ in range(3):
            # 要素までの途中の位置にランダムに移動
            actions.move_by_offset(
                random.randint(-100, 100),
                random.randint(-100, 100)
            )
            actions.pause(random.uniform(0.1, 0.3))
        
        # 最終的に要素まで移動
        actions.move_to_element(element)
        actions.pause(random.uniform(0.1, 0.2))
        actions.perform()

    def human_like_click(self, driver, element):
        """より人間らしいクリック操作をシミュレート"""
        try:
            # まず要素まで自然に移動
            self.human_like_mouse_move(driver, element)
            
            # クリック前に少し待機（人間らしい遅延）
            time.sleep(random.uniform(0.1, 0.3))
            
            # クリック
            element.click()
            
            # クリック後に少し待機
            time.sleep(random.uniform(0.1, 0.2))
        except Exception as e:
            self.update_signal.emit(f"人間らしいクリックに失敗しました: {str(e)}")
            # 通常のクリックにフォールバック
            element.click()

    def navigate_to_date(self, driver, booking_day, month_end):
        """
        カレンダー上で指定された日を選択するためのセル位置(day_in_week)を計算します。
        """
        def click_next_week_with_retry(max_retries=3):
            """次の週ボタンを安全にクリックし、例外が発生した場合は再試行する"""
            for attempt in range(max_retries):
                try:
                    # 要素が表示され、クリック可能になるまで待機
                    wait = WebDriverWait(driver, 10)
                    next_week_button = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//button[@id='next-week']"))
                    )
                    # JavaScriptを使用して直接クリック
                    driver.execute_script("arguments[0].click();", next_week_button)
                    
                    # クリック後にページが更新されるのを待機
                    time.sleep(1.5)
                    return True
                except StaleElementReferenceException:
                    if attempt < max_retries - 1:
                        self.update_signal.emit(f"StaleElementReferenceException が発生しました。再試行 {attempt + 1}/{max_retries}")
                        time.sleep(2)
                        continue
                    else:
                        self.update_signal.emit("最大再試行回数を超えました")
                        return False
                except Exception as e:
                    if attempt < max_retries - 1:
                        self.update_signal.emit(f"次の週ボタンのクリックに失敗しました: {e} - 再試行 {attempt + 1}/{max_retries}")
                        time.sleep(2)
                        continue
                    else:
                        self.update_signal.emit(f"次の週ボタンのクリックに失敗しました: {e} - 最大再試行回数を超えました")
                        return False
        
        try:
            if booking_day >= 29:
                # 29日以降の処理
                success_count = 0
                for _ in range(4):
                    if click_next_week_with_retry():
                        success_count += 1
                    else:
                        self.update_signal.emit(f"ナビゲーション失敗。{success_count}/4 回成功")
                        
                if success_count < 4:
                    self.update_signal.emit("警告: すべてのナビゲーションが成功しませんでした")
                    
                # 月末の日数に応じた例外処理
                if month_end == 31:
                    day_mapping = {29: 5, 30: 6, 31: 7}
                    day_in_week = day_mapping.get(booking_day)
                    if day_in_week is None:
                        raise ValueError(f"無効な予約日: {booking_day}")
                elif month_end == 30:
                    day_mapping = {29: 6, 30: 7}
                    day_in_week = day_mapping.get(booking_day)
                    if day_in_week is None:
                        raise ValueError(f"無効な予約日: {booking_day}")
                elif month_end == 29:
                    if booking_day == 29:
                        day_in_week = 7
                    else:
                        raise ValueError(f"無効な予約日: {booking_day}")
                else:
                    raise ValueError(f"無効な月末日: {month_end}")
            else:
                # 1日～28日の場合
                weeks_to_advance = (booking_day - 1) // 7
                day_in_week = (booking_day - 1) % 7 + 1
                
                success_count = 0
                for _ in range(weeks_to_advance):
                    if click_next_week_with_retry():
                        success_count += 1
                    else:
                        self.update_signal.emit(f"ナビゲーション失敗。{success_count}/{weeks_to_advance} 回成功")
                        
                if success_count < weeks_to_advance:
                    self.update_signal.emit(f"警告: すべてのナビゲーション({weeks_to_advance}回)が成功しませんでした")
                        
            return day_in_week
        except Exception as e:
            self.update_signal.emit(f"カレンダーナビゲーションエラー: {e}")
            # エラー発生時の画面キャプチャ
            try:
                driver.save_screenshot(f"calendar_nav_error.png")
            except:
                pass
            raise

    def check_for_captcha(self, driver):
        """reCAPTCHAの有無をチェックする"""
        try:
            # まずアラートの存在をチェック
            try:
                alert = Alert(driver)
                alert_text = alert.text
                # Captcha特有のメッセージかチェック
                if "確認のため、チェックを入れてから" in alert_text:
                    alert.accept()
                    return True
            except:
                pass

            # reCAPTCHAの要素を探す
            captcha_iframe = driver.find_elements(By.CSS_SELECTOR, "iframe[src*='recaptcha']")
            if captcha_iframe:
                return True
            return False
        except Exception as e:
            self.update_signal.emit(f"Captchaチェック中にエラーが発生: {str(e)}")
            return False

    def handle_booking_process(self, driver, user_number, password, booking_day, time_code, apply_number_text, month_end, max_retries=3):
        """予約処理を実行する関数"""
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                # サイトにアクセス
                driver.get(URL)
                time.sleep(1.0)

                # ログイン
                wait = WebDriverWait(driver, 60)
                login_button = wait.until(EC.element_to_be_clickable((By.ID, "btn-login")))
                login_button.click()

                user_number_field = wait.until(EC.presence_of_element_located((By.NAME, "userId")))
                password_field = driver.find_element(By.NAME, "password")
                
                user_number_field.send_keys(user_number)
                password_field.send_keys(password)
                password_field.send_keys(Keys.RETURN)

                WebDriverWait(driver, 60).until_not(EC.presence_of_element_located((By.ID, "btn-login")))
                time.sleep(0.5)

                # 「抽選」タブをクリック
                lottery_tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@data-target='#modal-menus']")))
                driver.execute_script("arguments[0].click();", lottery_tab)
                time.sleep(0.5)

                # 「抽選申込み」ボタンをクリック
                lottery_application_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '抽選申込み')]")))
                driver.execute_script("arguments[0].click();", lottery_application_button)
                time.sleep(0.5)

                # 「テニス（人工芝）」の申込みボタンをクリック
                artificial_grass_tennis_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//tr[td[contains(text(), 'テニス（人工芝')]]//button[contains(text(), '申込み')]")))
                driver.execute_script("arguments[0].click();", artificial_grass_tennis_button)
                time.sleep(random.uniform(1.0, 2.0))

                # 公園選択（「城北中央公園」）
                park_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "bname")))
                Select(park_dropdown).select_by_visible_text("城北中央公園")
                time.sleep(2)

                # 施設選択（「テニス（人工芝）」）
                facility_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "iname")))
                Select(facility_dropdown).select_by_visible_text("テニス（人工芝・照明有）")
                time.sleep(2)

                # 日付が見つかるまで翌週ボタンを押す
                day_in_week = self.navigate_to_date(driver, booking_day, month_end)

                # 日付と時間を選択する部分
                time_index = int(time_code)
                # 待機を追加して前の操作が完了するのを待つ
                time.sleep(1.0)

                # 日付のセルを見つける
                xpath = f'//*[@id="usedate-bheader-{time_index}"]/td[{day_in_week}]'
                cell = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))

                # セルの現在の状態をチェック（すでに選択されているかどうか）
                cell_class = cell.get_attribute("class")
                self.update_signal.emit(f"クリック前のセルのクラス: {cell_class}")

                # まだ選択されていない場合のみクリック
                if "selected" not in cell_class.lower() and "active" not in cell_class.lower():
                    self.update_signal.emit(f"日付時間の選択: 時間帯={time_index}, 曜日={day_in_week}")
                    driver.execute_script("arguments[0].click();", cell)
                    # クリック後に待機時間を確保
                    time.sleep(1.5)
                else:
                    self.update_signal.emit(f"セルはすでに選択されています。クリックをスキップします。")

                # アラートをチェック
                try:
                    alert = driver.switch_to.alert
                    alert_text = alert.text
                    self.update_signal.emit(f"予期せぬアラートが表示されています: {alert_text}")
                    alert.accept()
                    time.sleep(0.5)
                    
                    # アラートが「利用時間帯を選択して下さい」の場合、もう一度クリックするが、注意して行う
                    if "利用時間帯を選択して下さい" in alert_text:
                        self.update_signal.emit("時間帯選択をやり直します。")
                        
                        # セルを再取得して状態を確認
                        cell = driver.find_element(By.XPATH, xpath)
                        cell_class = cell.get_attribute("class")
                        
                        # 選択されていない場合のみクリック
                        if "selected" not in cell_class.lower() and "active" not in cell_class.lower():
                            driver.execute_script("arguments[0].click();", cell)
                            time.sleep(1.0)
                except:
                    # アラートがなければ続行
                    pass

                # 申込みボタンをクリック
                try:
                    apply_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '申込み')]")))
                    driver.execute_script("arguments[0].click();", apply_button)
                    time.sleep(0.5)
                except Exception as e:
                    self.update_signal.emit(f"申込みボタンのクリックに失敗: {str(e)}")
                    # 画面をキャプチャして状況を確認
                    try:
                        driver.save_screenshot(f"apply_button_error_{user_number}.png")
                    except:
                        pass
                    raise e

                # ここからキャプチャ監視対象の処理
                try:
                    # 申込み番号を選択
                    apply_number_select = wait.until(EC.element_to_be_clickable((By.ID, "apply")))
                    driver.execute_script("arguments[0].scrollIntoView(true);", apply_number_select)
                    time.sleep(0.3)
                    driver.execute_script("arguments[0].click();", apply_number_select)
                    Select(apply_number_select).select_by_visible_text(apply_number_text)
                    time.sleep(random.uniform(1.0, 2.0))
                except NoSuchElementException as e:
                    # 修正: apply_number_textがエラーメッセージに含まれるかチェック
                    if apply_number_text in str(e):
                        self.update_signal.emit(f"ユーザー {user_number} は既に {apply_number_text} で申し込み済みのようです。次のユーザーに進みます。")
                        return True
                    raise e

                # 確認画面で申込みボタンをクリック
                confirm_apply_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '申込み')]")))
                driver.execute_script("arguments[0].click();", confirm_apply_button)
                time.sleep(random.uniform(1.0, 2.0))

                # アラートのOKをクリック
                try:
                    WebDriverWait(driver, 10).until(EC.alert_is_present())
                    Alert(driver).accept()
                    time.sleep(random.uniform(2.0, 3.0))
                except:
                    pass

                # 確認画面で申込みボタンをクリック（JavaScriptを使用）
                confirm_apply_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '申込み')]")))
                driver.execute_script("arguments[0].click();", confirm_apply_button)
                time.sleep(random.uniform(1.0, 2.0))

                # アラートのOKをクリック
                try:
                    WebDriverWait(driver, 10).until(EC.alert_is_present())
                    Alert(driver).accept()
                    time.sleep(random.uniform(2.0, 3.0))
                except:
                    pass

                # Captchaチェック
                if self.check_for_captcha(driver):
                    self.update_signal.emit(f"Captchaが検出されました。ユーザー {user_number} の処理を再試行します。(試行回数: {retry_count + 1}/{max_retries})")
                    retry_count += 1
                    
                    try:
                        driver.save_screenshot(f"captcha_detected_{user_number}_retry_{retry_count}.png")
                    except:
                        pass

                    driver.close()
                    
                    remaining_tabs = driver.window_handles
                    if remaining_tabs:
                        driver.switch_to.window(remaining_tabs[0])
                    
                    driver.execute_script("window.open('');")
                    new_tab = driver.window_handles[-1]
                    driver.switch_to.window(new_tab)
                    
                    time.sleep(random.uniform(20.0, 30.0))
                    continue
                
                # 予約完了確認
                try:
                    completion_message = driver.find_element(By.XPATH, "//div[contains(text(), '申込みが完了しました')]")
                    if completion_message:
                        self.update_signal.emit(f"ユーザー {user_number} の予約処理が正常に完了しました。")
                        return True
                except:
                    pass

                return True

            except Exception as e:
                self.update_signal.emit(f"予約プロセス中にエラーが発生: {user_number}, エラー: {type(e).__name__}, {str(e)}")
                
                try:
                    driver.save_screenshot(f"error_process_{user_number}_retry_{retry_count}.png")
                except:
                    pass
                    
                retry_count += 1
                
                if retry_count < max_retries:
                    self.update_signal.emit(f"リトライを実行します。({retry_count}/{max_retries})")
                    try:
                        driver.close()
                        
                        remaining_tabs = driver.window_handles
                        if remaining_tabs:
                            driver.switch_to.window(remaining_tabs[0])
                        
                        driver.execute_script("window.open('');")
                        new_tab = driver.window_handles[-1]
                        driver.switch_to.window(new_tab)
                        
                        time.sleep(random.uniform(20.0, 30.0))
                    except Exception as tab_error:
                        self.update_signal.emit(f"タブの切り替え中にエラーが発生: {str(tab_error)}")
                        try:
                            driver.quit()
                        except:
                            pass

                        # Chromeブラウザの起動(再)
                        options = setup_chrome_options(self.params.get("headless", True))  # ヘッドレスモード設定を渡す
                        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
                        driver.get("about:blank")
                else:
                    self.update_signal.emit(f"最大リトライ回数に達しました。ユーザー {user_number} の処理をスキップします。")
                    return False

        return False

    # 抽選申込状況の確認処理
    def check_lottery_status(self):
        csv_file = self.params.get("csv_file", "Johoku1.csv")
        headless = self.params.get("headless", True)  # ヘッドレスモード設定
        
        # CSVファイルからデータを読み込み
        self.update_signal.emit(f"ファイル {csv_file} からユーザー情報を読み込んでいます...")
        self.update_signal.emit(f"ヘッドレスモード: {'有効' if headless else '無効'}")
        
        users_data = pd.read_csv(csv_file, dtype={
            'user_number': str,
            'password': str
        })
        
        total_users = len(users_data)
        self.update_signal.emit(f"{total_users}人のユーザー情報を読み込みました。")
        
        # 日付と時刻の組み合わせを保存するリスト
        reservation_list = []
        # ログインに失敗したアカウントを保存するリスト
        failed_logins = []
        # 申込がされていないアカウントを保存するリスト
        no_bookings = []
        # 申込が1つのみのアカウントを保存するリスト
        one_booking = []
        # 各ユーザーの予約数を追跡する辞書
        user_booking_count = defaultdict(int)
        
        # 書き込み可能なディレクトリを取得
        writable_dir = get_writable_dir()
        # 結果ファイルを初期化
        output_file = os.path.join(writable_dir, "reservation_info.txt")
        self.update_signal.emit(f"出力ファイル: {output_file}")
        
        with open(output_file, "w", encoding="utf-8") as file:
            file.write("=== 抽選申込状況の確認 ===\n")
            file.write(f"実行日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        # Chromeブラウザの起動
        self.update_signal.emit("Chromeブラウザを起動しています...")
        options = setup_chrome_options(headless)  # ヘッドレスモード設定を渡す
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        try:
            for index, row in users_data.iterrows():
                if not self.is_running:
                    self.update_signal.emit("処理が中断されました。")
                    break
                    
                user_number = row['user_number']
                password = row['password']
                user_name = row.get('Name', '不明')  # Name列がない場合は'不明'を使用
                
                progress = int((index / total_users) * 100)
                self.progress_signal.emit(progress)
                
                self.update_signal.emit(f"\nユーザー {user_number} の処理を開始します... ({index+1}/{total_users})")
                
                # 新しいタブを開く
                driver.execute_script("window.open('');")
                # 新しいタブのハンドルを取得
                new_tab = driver.window_handles[-1]
                # 新しいタブに切り替え
                driver.switch_to.window(new_tab)
                
                login_successful = False
                modal_successful = False
                
                try:
                    # サイトにアクセス
                    driver.get(URL)
                    
                    # 「ログイン」ボタンの表示まで待機
                    wait = WebDriverWait(driver, 10)
                    login_button = wait.until(EC.element_to_be_clickable((By.ID, "btn-login")))
                    login_button.click()
                    
                    # ログインフォームの表示を待機
                    user_number_field = wait.until(EC.presence_of_element_located((By.NAME, "userId")))
                    password_field = driver.find_element(By.NAME, "password")
                    
                    # 利用者番号とパスワードを入力
                    user_number_field.send_keys(user_number)
                    password_field.send_keys(password)
                    password_field.send_keys(Keys.RETURN)  # エンターキーで送信
                    
                    # ログイン後にユーザーメニューが表示されるまで待機
                    try:
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//a[@id='userName']"))
                        )
                        self.update_signal.emit(f"ログイン成功: {user_number}")
                        login_successful = True
                    except Exception as e:
                        self.update_signal.emit(f"ユーザーメニューの表示に失敗: {user_number} - エラー詳細: {e}")
                        failed_logins.append((user_number, password, user_name))
                        continue
                    
                    # モーダルを表示して「抽選申込みの確認」リンクをクリック
                    try:
                        # 「抽選」メニューをクリックしてモーダルを表示
                        lottery_menu = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//a[@data-target='#modal-menus']"))
                        )
                        lottery_menu.click()
                        
                        # モーダル内の「抽選申込みの確認」リンクをクリック
                        confirm_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//a[text()='抽選申込みの確認']"))
                        )
                        confirm_button.click()
                        self.update_signal.emit(f"抽選申込みの確認ボタンをクリック: {user_number}")
                        modal_successful = True
                        
                        # 利用日と時刻の情報を取得
                        try:
                            # テーブルが存在するか確認（表が表示されない場合もある）
                            table_exists = len(driver.find_elements(By.XPATH, "//table[@class='table sp-block-table']//tbody")) > 0
                            
                            if table_exists:
                                rows = driver.find_elements(By.XPATH, "//table[@class='table sp-block-table']//tbody//tr")
                                row_count = len(rows)
                            else:
                                rows = []
                                row_count = 0
                            
                            # ファイルに書き込み
                            with open(output_file, "a", encoding="utf-8") as file:
                                file.write(f"利用者番号: {user_number}\n")
                                file.write(f"パスワード: {password}\n")
                                file.write(f"利用者氏名: {user_name}\n")
                                
                                if row_count == 0:
                                    file.write("申込情報なし\n")
                                    no_bookings.append((user_number, password, user_name))
                                    user_booking_count[(user_number, password, user_name)] = 0
                                else:
                                    booking_count = 0
                                    for row in rows:
                                        status = row.find_element(By.XPATH, "./td[2]").text.strip()
                                        category = row.find_element(By.XPATH, "./td[3]").text.strip()
                                        facility = row.find_element(By.XPATH, "./td[4]").text.strip()
                                        date = row.find_element(By.XPATH, "./td[5]").text.strip()
                                        time = row.find_element(By.XPATH, "./td[6]").text.strip()
                                        file.write(f"状況: {status}\n")
                                        file.write(f"分類: {category}\n")
                                        file.write(f"公園・施設: {facility}\n")
                                        file.write(f"利用日: {date}\n")
                                        file.write(f"時刻: {time}\n")
                                        
                                        # 日付と時刻をリストに追加
                                        reservation_list.append((date, time))
                                        booking_count += 1
                                    
                                    # ユーザーの予約数を記録
                                    user_booking_count[(user_number, password, user_name)] = booking_count
                                    
                                    # 申込みが1つだけの場合
                                    if booking_count == 1:
                                        one_booking.append((user_number, password, user_name))
                                
                                file.write("---------------\n")
                        except Exception as e:
                            self.update_signal.emit(f"予約情報の取得に失敗しました: {user_number} - エラー詳細: {e}")
                            # モーダル表示には成功しているので、予約情報なしと判断
                            with open(output_file, "a", encoding="utf-8") as file:
                                file.write(f"利用者番号: {user_number}\n")
                                file.write(f"パスワード: {password}\n")
                                file.write(f"利用者氏名: {user_name}\n")
                                file.write("申込情報なし（表示エラー）\n")
                                file.write("---------------\n")
                            no_bookings.append((user_number, password, user_name))
                            user_booking_count[(user_number, password, user_name)] = 0
                            
                    except Exception as e:
                        self.update_signal.emit(f"抽選申込みの確認ボタンのクリックに失敗しました: {user_number} - エラー詳細: {e}")
                        failed_logins.append((user_number, password, user_name))
                    
                    # 次のログイン試行前に1秒間待機
                    time.sleep(1)
                        
                except Exception as e:
                    self.update_signal.emit(f"処理中にエラーが発生しました: {user_number} - エラー詳細: {e}")
                    if not login_successful:
                        failed_logins.append((user_number, password, user_name))
                    elif not modal_successful:
                        failed_logins.append((user_number, password, user_name))
            
            # 最終的な進捗状況を100%に設定
            self.progress_signal.emit(100)
            
            # 予約情報を集計してカウント
            reservation_count = pd.Series(reservation_list).value_counts()
            
            # 日本語の日付形式（例: 2024年4月10日）を解析してdatetimeオブジェクトに変換する関数
            def parse_japanese_date(date_str):
                pattern = r'(\d+)年(\d+)月(\d+)日'
                match = re.match(pattern, date_str)
                if match:
                    year, month, day = map(int, match.groups())
                    return datetime(year, month, day)
                return datetime(9999, 12, 31)  # パースできない場合のフォールバック
                
            # reservation_countから辞書リストを作成
            reservation_data = []
            for (date, time), count in reservation_count.items():
                reservation_data.append({
                    'date_str': date,
                    'time': time,
                    'count': count
                })
                
            # datetimeオブジェクトでソート
            if reservation_data:
                reservation_data.sort(key=lambda x: parse_japanese_date(x['date_str']))
            
            # 集計結果をテキストファイルに書き込み
            with open(output_file, "a", encoding="utf-8") as file:
                file.write("=== 予約回数集計結果（日付順） ===\n")
                for item in reservation_data:
                    file.write(f"利用日: {item['date_str']}, 時刻: {item['time']}, 回数: {item['count']}\n")
                
                file.write("\n=== ログインに失敗したアカウント ===\n")
                for user_number, password, user_name in failed_logins:
                    file.write(f"利用者番号: {user_number}, パスワード: {password}, 氏名: {user_name}\n")
                    
                file.write("\n=== 申込みがされていないアカウント ===\n")
                for user_number, password, user_name in no_bookings:
                    file.write(f"利用者番号: {user_number}, パスワード: {password}, 氏名: {user_name}\n")
                    
                file.write("\n=== 申込みが1つだけのアカウント ===\n")
                for user_number, password, user_name in one_booking:
                    file.write(f"利用者番号: {user_number}, パスワード: {password}, 氏名: {user_name}\n")
                    
                # 各ユーザーの予約数を記録
                file.write("\n=== 各ユーザーの申込み数 ===\n")
                for (user_number, password, user_name), count in sorted(user_booking_count.items(), key=lambda x: x[1]):
                    file.write(f"利用者番号: {user_number}, 氏名: {user_name}, 申込み数: {count}\n")
            
            # 集計結果を表示
            summary = f"\n=== 集計結果 ===\n"
            summary += f"合計確認ユーザー数: {len(users_data)}\n"
            summary += f"ログイン失敗数: {len(failed_logins)}\n"
            summary += f"申込みなしユーザー数: {len(no_bookings)}\n"
            summary += f"申込み1つのみユーザー数: {len(one_booking)}\n"
            summary += f"確認された予約総数: {sum(item['count'] for item in reservation_data)}\n"
            summary += f"\n詳細な情報は {output_file} に保存されました。"
            
            self.update_signal.emit(summary)
            
        except Exception as e:
            self.update_signal.emit(f"予約確認処理中にエラーが発生しました: {str(e)}")
            raise
        finally:
            try:
                driver.quit()
            except:
                pass

    # 抽選確定処理
    def confirm_lottery_selection(self):
        csv_file = self.params.get("csv_file", "Johoku1.csv")
        user_count = self.params.get("user_count", "6")
        headless = self.params.get("headless", True)  # ヘッドレスモード設定
        
        # ヘッドレスモード情報をログに出力
        self.update_signal.emit(f"ヘッドレスモード: {'有効' if headless else '無効'}")
        
        # 書き込み可能なディレクトリを取得
        writable_dir = get_writable_dir()
        # 結果を書き込むファイル名
        output_file = os.path.join(writable_dir, "lottery_results.txt")
        self.update_signal.emit(f"出力ファイル: {output_file}")
        
        # CSVファイルからデータを読み込み
        self.update_signal.emit(f"ファイル {csv_file} からユーザー情報を読み込んでいます...")
        users_data = pd.read_csv(csv_file, dtype={
            'user_number': str,
            'password': str
        })
        
        total_users = len(users_data)
        self.update_signal.emit(f"{total_users}人のユーザー情報を読み込みました。")
        
        with open(output_file, "w", encoding="utf-8") as file:
            file.write("===== 抽選確定処理結果 =====\n")
            file.write(f"実行日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        # Chromeブラウザの起動
        self.update_signal.emit("Chromeブラウザを起動しています...")
        options = setup_chrome_options(headless)  # ヘッドレスモード設定を渡す
        options.add_argument("--disable-popup-blocking")  # ポップアップを無効化
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        
        try:
            for index, row in users_data.iterrows():
                if not self.is_running:
                    self.update_signal.emit("処理が中断されました。")
                    break
                    
                user_number = row['user_number']
                password = row['password']
                # 氏名情報の取得（'Kana'または'Name'があれば使用、なければuser_numberを使用）
                user_name = row.get('Kana', row.get('Name', user_number))
                
                progress = int((index / total_users) * 100)
                self.progress_signal.emit(progress)
                
                self.update_signal.emit(f"\nユーザー {user_number} ({user_name}) の処理を開始します... ({index+1}/{total_users})")
                
                # 新しいタブを開く
                driver.execute_script("window.open('');")
                # 新しいタブのハンドルを取得
                new_tab = driver.window_handles[-1]
                # 新しいタブに切り替え
                driver.switch_to.window(new_tab)
                
                try:
                    # サイトにアクセス
                    driver.get(URL)
                    
                    # 「ログイン」ボタンの表示まで待機
                    wait = WebDriverWait(driver, 10)
                    login_button = wait.until(EC.element_to_be_clickable((By.ID, "btn-login")))
                    login_button.click()
                    
                    # ログインフォームの表示を待機
                    user_number_field = wait.until(EC.presence_of_element_located((By.NAME, "userId")))
                    password_field = driver.find_element(By.NAME, "password")
                    
                    # 利用者番号とパスワードを入力
                    user_number_field.send_keys(user_number)
                    password_field.send_keys(password)
                    password_field.send_keys(Keys.RETURN)  # エンターキーで送信
                    
                    # ログイン後に「ログイン」ボタンが存在しないことを確認
                    WebDriverWait(driver, 10).until_not(
                        EC.presence_of_element_located((By.ID, "btn-login"))
                    )
                    
                    self.update_signal.emit(f"ログイン成功: {user_number}")
                    
                    # モーダルを表示して「抽選結果」リンクをクリック
                    try:
                        # 「抽選」メニューをクリックしてモーダルを表示
                        lottery_menu = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//a[@data-target='#modal-menus']"))
                        )
                        driver.execute_script("arguments[0].click();", lottery_menu)
                        
                        # モーダル内の「抽選結果」リンクをクリック
                        result_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//a[text()='抽選結果']"))
                        )
                        driver.execute_script("arguments[0].click();", result_button)
                        self.update_signal.emit(f"抽選結果ボタンをクリック: {user_number}")
                        
                        # 当選結果のテーブルが表示されるまで待機
                        try:
                            WebDriverWait(driver, 3).until(
                                EC.presence_of_element_located((By.XPATH, "//table[@class='table sp-block-table']/tbody/tr"))
                            )
                            
                            # 当選結果の情報を取得し、選択ボタンをクリック
                            rows = driver.find_elements(By.XPATH, "//table[@class='table sp-block-table']/tbody/tr")
                            
                            if rows:
                                with open(output_file, "a", encoding="utf-8") as file:
                                    file.write(f"ユーザー: {user_name} (ID: {user_number})\n")
                                
                                for row in rows:
                                    try:
                                        booking_date = row.find_element(By.XPATH, ".//td[2]/label/span[2]").text
                                        booking_time = row.find_element(By.XPATH, ".//td[3]/label").text
                                        
                                        # ファイルに書き込む
                                        with open(output_file, "a", encoding="utf-8") as file:
                                            file.write(f"  日付: {booking_date}, 時間: {booking_time}\n")
                                        self.update_signal.emit(f"当選情報: {user_name},{booking_date},{booking_time}")
                                        
                                        # 選択ボタンをクリック (JavaScriptでクリック)
                                        select_button = row.find_element(By.XPATH, ".//input[@name='checkElect']")
                                        driver.execute_script("arguments[0].click();", select_button)
                                    except Exception as e:
                                        self.update_signal.emit(f"行の処理に失敗: {str(e)}")
                                
                                # 確認ボタンをクリック (JavaScriptでクリック)
                                try:
                                    confirm_button = driver.find_element(By.ID, "btn-go")
                                    driver.execute_script("arguments[0].click();", confirm_button)
                                    self.update_signal.emit(f"確認ボタンをクリック: {user_number}")
                                    
                                    # 利用人数の入力ページが表示されるまで待機
                                    WebDriverWait(driver, 10).until(
                                        EC.presence_of_element_located((By.XPATH, "//input[@name='applyNum']"))
                                    )
                                    
                                    # 利用人数を入力
                                    user_count_inputs = driver.find_elements(By.XPATH, "//input[@name='applyNum']")
                                    for input_field in user_count_inputs:
                                        input_field.clear()  # 既存の入力をクリア
                                        input_field.send_keys(user_count)  # 指定された利用人数を設定
                                    
                                    # 確認ボタンをクリック (JavaScriptでクリック)
                                    final_confirm_button = driver.find_element(By.XPATH, "//button[contains(text(), '確認')]")
                                    driver.execute_script("arguments[0].click();", final_confirm_button)
                                    self.update_signal.emit(f"最終確認ボタンをクリック: {user_number}")
                                    
                                    # ポップアップの確認とOKボタンをクリック
                                    try:
                                        alert = WebDriverWait(driver, 5).until(EC.alert_is_present())
                                        alert.accept()
                                        self.update_signal.emit(f"ポップアップのOKボタンをクリック: {user_number}")
                                        with open(output_file, "a", encoding="utf-8") as file:
                                            file.write("  処理結果: 確定成功\n\n")
                                    except:
                                        self.update_signal.emit(f"ポップアップは表示されませんでした: {user_number}")
                                        with open(output_file, "a", encoding="utf-8") as file:
                                            file.write("  処理結果: 確定処理完了（ポップアップなし）\n\n")
                                except Exception as e:
                                    self.update_signal.emit(f"確定処理中にエラー: {str(e)}")
                                    with open(output_file, "a", encoding="utf-8") as file:
                                        file.write(f"  処理結果: 確定処理エラー - {str(e)}\n\n")
                            else:
                                self.update_signal.emit(f"ユーザー {user_number} に当選情報がありません")
                                with open(output_file, "a", encoding="utf-8") as file:
                                    file.write(f"ユーザー: {user_name} (ID: {user_number})\n")
                                    file.write("  当選情報なし\n\n")
                        except Exception as e:
                            self.update_signal.emit(f"当選テーブルが見つかりません: {user_number} - {str(e)}")
                            with open(output_file, "a", encoding="utf-8") as file:
                                file.write(f"ユーザー: {user_name} (ID: {user_number})\n")
                                file.write("  当選テーブルなし\n\n")
                        
                    except Exception as e:
                        self.update_signal.emit(f"抽選結果の処理に失敗しました: {user_number} - エラー詳細: {e}")
                        with open(output_file, "a", encoding="utf-8") as file:
                            file.write(f"ユーザー: {user_name} (ID: {user_number})\n")
                            file.write(f"  エラー: 抽選結果の処理に失敗 - {str(e)}\n\n")
                    
                    # 次のログイン試行前に待機
                    time.sleep(1)
                    
                except Exception as e:
                    self.update_signal.emit(f"エラーが発生しました: {str(e)}")
                    with open(output_file, "a", encoding="utf-8") as file:
                        file.write(f"ユーザー: {user_name} (ID: {user_number})\n")
                        file.write(f"  エラー: {str(e)}\n\n")
            
            # 最終的な進捗状況を100%に設定
            self.progress_signal.emit(100)
            
            # 処理完了メッセージ
            self.update_signal.emit("\n抽選確定処理が完了しました")
            self.update_signal.emit(f"結果は {output_file} に保存されました")
            
        except Exception as e:
            self.update_signal.emit(f"抽選確定処理中にエラーが発生しました: {str(e)}")
            raise
        finally:
            try:
                driver.quit()
            except:
                pass

    # 予約状況の確認処理
    def check_reservation_status(self):
        csv_file = self.params.get("csv_file", "Johoku1.csv")
        headless = self.params.get("headless", True)  # ヘッドレスモード設定
        
        # ヘッドレスモード情報をログに出力
        self.update_signal.emit(f"ヘッドレスモード: {'有効' if headless else '無効'}")
        
        # CSVファイルからデータを読み込み
        self.update_signal.emit(f"ファイル {csv_file} からユーザー情報を読み込んでいます...")
        users_data = pd.read_csv(csv_file, dtype={
            'user_number': str,
            'password': str
        })
        
        total_users = len(users_data)
        self.update_signal.emit(f"{total_users}人のユーザー情報を読み込みました。")
        
        # 日付と時刻の組み合わせを保存するリスト
        reservation_list = []
        # ログインに失敗したアカウントを保存するリスト
        failed_logins = []
        
        # 書き込み可能なディレクトリを取得
        writable_dir = get_writable_dir()
        # 結果を書き込むファイル名
        result_file = os.path.join(writable_dir, "r_info.txt")
        self.update_signal.emit(f"出力ファイル: {result_file}")
        
        # ファイルが存在する場合は削除
        if os.path.exists(result_file):
            os.remove(result_file)
        
        # 結果ファイルの初期化
        with open(result_file, "w", encoding="utf-8") as file:
            file.write(f"=== 予約状況確認 ===\n")
            file.write(f"実行日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        # Chromeブラウザの起動
        self.update_signal.emit("Chromeブラウザを起動しています...")
        options = setup_chrome_options(headless)  # ヘッドレスモード設定を渡す
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        
        try:
            for index, row in users_data.iterrows():
                if not self.is_running:
                    self.update_signal.emit("処理が中断されました。")
                    break
                    
                user_number = row['user_number']
                password = row['password']
                user_name = row.get('Name', '不明')  # Name列がない場合は'不明'を使用
                
                progress = int((index / total_users) * 100)
                self.progress_signal.emit(progress)
                
                self.update_signal.emit(f"\nユーザー {user_number} の処理を開始します... ({index+1}/{total_users})")
                
                # 新しいタブを開く
                driver.execute_script("window.open('');")
                new_tab = driver.window_handles[-1]
                driver.switch_to.window(new_tab)
                
                try:
                    # サイトにアクセス
                    driver.get(URL)
                    self.update_signal.emit(f"サイトにアクセス: {URL}")
                    
                    # 「ログイン」ボタンの表示まで待機
                    wait = WebDriverWait(driver, 10)
                    login_button = wait.until(EC.element_to_be_clickable((By.ID, "btn-login")))
                    login_button.click()
                    self.update_signal.emit("ログインボタンをクリック")
                    
                    # ログインフォームの表示を待機
                    user_number_field = wait.until(EC.presence_of_element_located((By.NAME, "userId")))
                    password_field = driver.find_element(By.NAME, "password")
                    
                    # 利用者番号とパスワードを入力
                    user_number_field.send_keys(user_number)
                    password_field.send_keys(password)
                    password_field.send_keys(Keys.RETURN)  # エンターキーで送信
                    self.update_signal.emit(f"ログイン情報入力: {user_number}")
                    
                    # ログイン後にユーザーメニューが表示されるまで待機
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//a[@id='userName']"))
                    )
                    self.update_signal.emit(f"ログイン成功: {user_number}")
                    
                    # 「予約の確認」メニューを開く
                    lottery_menu = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//a[@data-target='#modal-reservation-menus']"))
                    )
                    lottery_menu.click()
                    self.update_signal.emit("予約メニューをクリック")
                    
                    confirm_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//a[text()='予約の確認']"))
                    )
                    confirm_button.click()
                    self.update_signal.emit(f"予約の確認ボタンをクリック: {user_number}")
                    
                    # 一旦待機して画面を読み込む
                    time.sleep(2)
                    
                    # ファイルに書き込む準備
                    with open(result_file, "a", encoding="utf-8") as file:
                        file.write(f"利用者番号: {user_number}\n")
                        file.write(f"利用者氏名: {user_name}\n")
                    
                    # テーブルの存在を確認（存在しない場合もエラーにしない）
                    # find_elementsはリストを返すので、長さをチェックする
                    tables = driver.find_elements(By.ID, "rsvacceptlist")
                    has_table = len(tables) > 0
                    
                    if has_table:
                        # テーブルが存在する場合
                        table = tables[0]
                        self.update_signal.emit("予約テーブルを確認")
                        
                        # テーブルの各行を取得
                        rows = table.find_elements(By.XPATH, ".//tr")
                        row_count = len(rows)
                        self.update_signal.emit(f"テーブル行数: {row_count}")
                        
                        if row_count <= 1:  # ヘッダー行のみの場合
                            self.update_signal.emit("テーブルはありますが、予約情報が存在しません。")
                            with open(result_file, "a", encoding="utf-8") as file:
                                file.write("予約情報が存在しません。\n")
                        else:
                            # ヘッダー行以外のデータがある場合
                            with open(result_file, "a", encoding="utf-8") as file:
                                for row in rows[1:]:  # 最初の行はヘッダーなのでスキップ
                                    # 各行のtdタグを取得
                                    cols = row.find_elements(By.XPATH, ".//td[@class='keep-wide']")
                                    if len(cols) >= 2:  # 必要な情報があるかチェック
                                        # 日付と時間を取得
                                        date_text = cols[0].text.strip()  # 利用日を取得
                                        time_str = cols[1].text.strip()  # 時間を取得（変数名を変更）
                                        
                                        # 書き込み
                                        file.write(f"利用日: {date_text}\n")
                                        file.write(f"時刻: {time_str}\n")
                                        file.write("\n")
                                        
                                        reservation_list.append((date_text, time_str, user_name, user_number))
                            self.update_signal.emit(f"予約情報をファイルに書き込み完了: {user_number}")
                    else:
                        # テーブルが存在しない場合
                        self.update_signal.emit("予約テーブルが存在しません（予約なし）")
                        with open(result_file, "a", encoding="utf-8") as file:
                            file.write("予約情報が存在しません。\n")
                    
                    # 必ず区切り線を書き込む
                    with open(result_file, "a", encoding="utf-8") as file:
                        file.write("---------------\n")
                    
                except Exception as e:
                    self.update_signal.emit(f"ユーザー {user_number} の処理中にエラーが発生しました - エラー詳細: {e}")
                    failed_logins.append((user_number, password, user_name))
                    with open(result_file, "a", encoding="utf-8") as file:
                        file.write(f"エラー: {str(e)}\n")
                        file.write("---------------\n")
                
                # 次のログイン試行前に待機
                time.sleep(0.1)
            
            # 最終的な進捗状況を100%に設定
            self.progress_signal.emit(100)
            
            # 予約情報がある場合は集計処理
            try:
                if reservation_list:
                    # 予約情報をDataFrameに変換し、ソートする
                    df = pd.DataFrame(reservation_list, columns=['利用日', '時刻', '氏名', '利用者番号'])
                    
                    # 日付と時刻のフォーマットを修正
                    df['利用日'] = df['利用日'].apply(lambda x: x.replace('\n', ' ').strip())
                    df['時刻'] = df['時刻'].apply(lambda x: x.split('～')[0].strip() if '～' in x else x)
                    
                    # 日付をdatetimeオブジェクトに変換する関数
                    def parse_date(date_str):
                        # 月、日、年を個別に抽出
                        month_match = re.search(r'(\d+)月', date_str)
                        day_match = re.search(r'(\d+)日', date_str)
                        year_match = re.search(r'(\d{4})年', date_str)
                        
                        if month_match and day_match and year_match:
                            month = int(month_match.group(1))
                            day = int(day_match.group(1))
                            year = int(year_match.group(1))
                            return datetime(year, month, day)
                        else:
                            return pd.NaT  # 解析できない場合は NaT (Not a Time) を返す
                    
                    # 日付を datetime オブジェクトに変換
                    df['利用日'] = df['利用日'].apply(parse_date)
                    
                    # 無効な日付を削除
                    df = df.dropna(subset=['利用日'])
                    
                    # ソート
                    df.sort_values(by=['利用日', '時刻'], inplace=True)
                    
                    # 集計結果をテキストファイルに書き込み
                    with open(result_file, "a", encoding="utf-8") as file:
                        file.write("\n=== 予約回数集計結果 ===\n")
                        if df.empty:
                            file.write("有効な予約情報がありません。\n")
                        else:
                            grouped = df.groupby(['利用日', '時刻'])
                            for (date, time_val), group in grouped:
                                file.write(f"利用日: {date.strftime('%Y年%m月%d日')}, 時刻: {time_val}, 面数: {len(group)}\n")
                                for _, row in group.iterrows():
                                    file.write(f"\t利用者氏名: {row['氏名']}, 利用者番号: {row['利用者番号']}\n")
                else:
                    self.update_signal.emit("予約情報が存在しません。")
                    with open(result_file, "a", encoding="utf-8") as file:
                        file.write("\n=== 予約回数集計結果 ===\n")
                        file.write("予約情報が存在しません。\n")
            except Exception as e:
                self.update_signal.emit(f"集計処理中にエラーが発生しました: {e}")
                with open(result_file, "a", encoding="utf-8") as file:
                    file.write("\n=== 予約回数集計結果 ===\n")
                    file.write(f"集計処理中にエラーが発生しました: {e}\n")
            
            # ログイン失敗したアカウントの情報を出力
            if failed_logins:
                self.update_signal.emit("\nログインに失敗したアカウント:")
                with open(result_file, "a", encoding="utf-8") as file:
                    file.write("\n=== ログインに失敗したアカウント ===\n")
                    for user_number, password, user_name in failed_logins:
                        file.write(f"利用者番号: {user_number}, 氏名: {user_name}\n")
                        self.update_signal.emit(f"利用者番号: {user_number}, 氏名: {user_name}")
            
            self.update_signal.emit("\n予約状況の確認が完了しました")
            self.update_signal.emit(f"結果は {result_file} に保存されました")
        
        except Exception as e:
            self.update_signal.emit(f"予約状況確認処理中にエラーが発生しました: {str(e)}")
            raise
        finally:
            try:
                driver.quit()
            except:
                pass

    # 有効期限の確認処理
    def check_account_expiry(self):
        csv_file = self.params.get("csv_file", "Johoku1.csv")
        headless = self.params.get("headless", True)  # ヘッドレスモード設定
        
        # CSVファイルからデータを読み込み
        self.update_signal.emit(f"ファイル {csv_file} からユーザー情報を読み込んでいます...")
        self.update_signal.emit(f"ヘッドレスモード: {'有効' if headless else '無効'}")
        
        users_data = pd.read_csv(csv_file, dtype={
            'user_number': str,
            'password': str
        })
        
        total_users = len(users_data)
        self.update_signal.emit(f"{total_users}人のユーザー情報を読み込みました。")
        
        # 書き込み可能なディレクトリを取得
        writable_dir = get_writable_dir()
        # 結果を書き込むファイル名（フルパス）
        output_file = os.path.join(writable_dir, "expiry.txt")
        self.update_signal.emit(f"出力ファイル: {output_file}")
        
        # 結果を一時的にリストに保存
        results = []
        
        # Chromeブラウザの起動
        self.update_signal.emit("Chromeブラウザを起動しています...")
        options = setup_chrome_options(headless)  # ヘッドレスモード設定を渡す
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-popup-blocking')
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        wait = WebDriverWait(driver, 10)
        
        try:
            self.update_signal.emit(f"=== アカウント有効期限の確認 ===")
            self.update_signal.emit(f"実行日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
            for index, row in users_data.iterrows():
                if not self.is_running:
                    self.update_signal.emit("処理が中断されました。")
                    break
                    
                user_number = row['user_number']
                password = row['password']
                # 'Kana'または'Name'があれば使用、なければuser_numberを使用
                user_name = row.get('Kana', row.get('Name', user_number))
                
                progress = int((index / total_users) * 100)
                self.progress_signal.emit(progress)
                
                self.update_signal.emit(f"\nユーザー {user_number} の処理を開始します... ({index+1}/{total_users})")
                
                # 新しいタブを開く
                driver.execute_script("window.open('');")
                driver.switch_to.window(driver.window_handles[-1])
                
                try:
                    # サイトにアクセス
                    driver.get(URL)
                    
                    # ログインボタンクリック
                    login_button = wait.until(EC.element_to_be_clickable((By.ID, "btn-login")))
                    login_button.click()
                    
                    # ログインフォーム入力
                    user_number_field = wait.until(EC.presence_of_element_located((By.NAME, "userId")))
                    password_field = driver.find_element(By.NAME, "password")
                    
                    user_number_field.send_keys(user_number)
                    password_field.send_keys(password)
                    password_field.send_keys(Keys.RETURN)
                    
                    # ログイン成功確認
                    wait.until_not(EC.presence_of_element_located((By.ID, "btn-login")))
                    
                    # マイメニューのドロップダウンを表示
                    dropdown_menu = wait.until(EC.element_to_be_clickable((By.ID, "userName")))
                    dropdown_menu.click()
                    
                    # 利用者情報の変更・削除・更新リンクをクリック
                    user_info_link = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '利用者情報の変更・削除・更新')]"))
                    )
                    user_info_link.click()
                    
                    # ページ遷移の完了を待機
                    time.sleep(2)
                    
                    # 有効期限の情報を取得
                    try:
                        # 有効期限を特定のXPathで探す
                        expiry_element = wait.until(
                            EC.presence_of_element_located((By.XPATH, 
                                "//th[.//label[@for='validEndYMD']]/following-sibling::td"
                            ))
                        )
                        expiry_info = expiry_element.text.strip()
                        
                        self.update_signal.emit(f"有効期限を取得: {user_number} - {expiry_info}")
                        
                        # 日付をdatetimeオブジェクトに変換
                        try:
                            # "2025年2月28日" -> datetime
                            year = int(expiry_info[:4])
                            month = int(expiry_info[5:expiry_info.index("月")])
                            day = int(expiry_info[expiry_info.index("月")+1:expiry_info.index("日")])
                            expiry_date = datetime(year, month, day)
                            
                            # 結果をリストに追加
                            results.append({
                                'user_number': user_number,
                                'user_name': user_name,
                                'expiry_info': expiry_info,
                                'expiry_date': expiry_date
                            })
                        except Exception as e:
                            self.update_signal.emit(f"日付解析エラー: {expiry_info} - {str(e)}")
                            # 解析に失敗しても情報は保存
                            results.append({
                                'user_number': user_number,
                                'user_name': user_name,
                                'expiry_info': expiry_info,
                                'expiry_date': datetime(9999, 12, 31)  # 遠い未来の日付
                            })
                        
                    except Exception as e:
                        self.update_signal.emit(f"有効期限の取得に失敗: {str(e)}")
                        # 失敗した場合も結果に追加
                        results.append({
                            'user_number': user_number,
                            'user_name': user_name,
                            'expiry_info': "取得失敗",
                            'expiry_date': datetime(9999, 12, 31)  # 遠い未来の日付
                        })
                
                except Exception as e:
                    self.update_signal.emit(f"ユーザー {user_number} の処理中にエラーが発生: {str(e)}")
                    results.append({
                        'user_number': user_number,
                        'user_name': user_name,
                        'expiry_info': "エラー発生",
                        'expiry_date': datetime(9999, 12, 31)  # 遠い未来の日付
                    })
                
                # 次のユーザーの処理前に待機
                time.sleep(0.5)
            
            # 最終的な進捗状況を100%に設定
            self.progress_signal.emit(100)
            
            # 日付でソート
            results.sort(key=lambda x: x['expiry_date'])
            
            # ソート済みデータを書き込む
            with open(output_file, "w", encoding="utf-8") as file:
                file.write("利用者番号,氏名,有効期限\n")
                for result in results:
                    file.write(f"{result['user_number']},{result['user_name']},{result['expiry_info']}\n")
            
            self.update_signal.emit("\nすべてのデータの書き込みが完了しました")
            self.update_signal.emit(f"結果は {output_file} に保存されました")
            
            # 今日から2週間以内に有効期限が切れるユーザーを表示
            today = datetime.now()
            from datetime import timedelta
            two_weeks_later = today + timedelta(days=14)  # 今日から2週間後
            
            self.update_signal.emit("\n=== 有効期限が2週間以内に切れるユーザー ===")
            expiring_soon = [r for r in results if r['expiry_date'] <= two_weeks_later and r['expiry_date'] != datetime(9999, 12, 31)]
            
            if expiring_soon:
                for result in expiring_soon:
                    self.update_signal.emit(f"利用者番号: {result['user_number']}, 氏名: {result['user_name']}, 有効期限: {result['expiry_info']}")
            else:
                self.update_signal.emit("2週間以内に有効期限が切れるユーザーはいません。")
        
        except Exception as e:
            self.update_signal.emit(f"有効期限確認処理中にエラーが発生しました: {str(e)}")
            raise
        finally:
            try:
                driver.quit()
            except:
                pass

            
# メインアプリケーションクラス
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

# メイン関数
def main():
    app = QApplication(sys.argv)
    
    # アプリケーションスタイルを設定（オプション - システムに応じたスタイルを適用）
    app.setStyle("Fusion")
    
    # アプリケーションアイコン設定（アイコンファイルがある場合）
    # app.setWindowIcon(QIcon("icon.png"))
    
    # メインウィンドウを作成して表示
    window = JohokuApp()
    window.show()
    
    # イベントループを開始
    sys.exit(app.exec_())

# このスクリプトが直接実行された場合にのみ main() を実行
if __name__ == "__main__":
    main()