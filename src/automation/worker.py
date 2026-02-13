"""バックグラウンド処理用のワーカースレッドモジュール"""
import os
import pandas as pd
import time
import random
import calendar
import re
import logging
import subprocess
from datetime import datetime, timedelta
from collections import defaultdict
from PyQt5.QtCore import QThread, pyqtSignal
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

from ..config import URL
from ..utils.helpers import get_writable_dir
from .browser import setup_chrome_options

# webdriver-managerのログを無効化（警告ダイアログを非表示に）
os.environ['WDM_LOG_LEVEL'] = '0'
os.environ['WDM_LOG'] = 'false'
os.environ['WDM_PRINT_FIRST_LINE'] = 'False'
logging.getLogger('WDM').setLevel(logging.ERROR)


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
        service = Service(ChromeDriverManager().install(), log_output=subprocess.DEVNULL)
        driver = webdriver.Chrome(service=service, options=options)
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
                    service = Service(ChromeDriverManager().install(), log_output=subprocess.DEVNULL)
                    driver = webdriver.Chrome(service=service, options=options)
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
                        service = Service(ChromeDriverManager().install(), log_output=subprocess.DEVNULL)
                        driver = webdriver.Chrome(service=service, options=options)
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
        service = Service(ChromeDriverManager().install(), log_output=subprocess.DEVNULL)
        driver = webdriver.Chrome(service=service, options=options)
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
        service = Service(ChromeDriverManager().install(), log_output=subprocess.DEVNULL)
        driver = webdriver.Chrome(service=service, options=options)

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
        service = Service(ChromeDriverManager().install(), log_output=subprocess.DEVNULL)
        driver = webdriver.Chrome(service=service, options=options)

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

        # 結果を一時的にリストに保存（ソート用）
        results = []
        # ログインに失敗したアカウントを保存するリスト
        failed_logins = []

        # ファイルの初期化（ヘッダー行を書き込み）
        with open(output_file, "w", encoding="utf-8") as file:
            file.write("利用者番号,氏名,有効期限\n")

        # Chromeブラウザの起動
        self.update_signal.emit("Chromeブラウザを起動しています...")
        options = setup_chrome_options(headless)  # ヘッドレスモード設定を渡す
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-popup-blocking')
        service = Service(ChromeDriverManager().install(), log_output=subprocess.DEVNULL)
        driver = webdriver.Chrome(service=service, options=options)
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

                login_successful = False

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

                    # アラートが表示された場合は受け入れて次へ
                    time.sleep(1)
                    try:
                        alert = Alert(driver)
                        alert_text = alert.text
                        self.update_signal.emit(f"アラート検出: {user_number} - {alert_text}")
                        alert.accept()
                        # アラートが出たということはログイン失敗
                        failed_logins.append((user_number, password, user_name))
                        result = {
                            'user_number': user_number,
                            'user_name': user_name,
                            'expiry_info': f"ログイン失敗({alert_text})",
                            'expiry_date': datetime(9999, 12, 31)
                        }
                        results.append(result)
                        with open(output_file, "a", encoding="utf-8") as file:
                            file.write(f"{result['user_number']},{result['user_name']},{result['expiry_info']}\n")
                        continue
                    except:
                        # アラートがない場合は通常処理
                        pass

                    # ログイン成功確認
                    try:
                        wait.until_not(EC.presence_of_element_located((By.ID, "btn-login")))
                        self.update_signal.emit(f"ログイン成功: {user_number}")
                        login_successful = True
                    except Exception as e:
                        self.update_signal.emit(f"ログイン失敗: {user_number} - 次のユーザーに移行します")
                        failed_logins.append((user_number, password, user_name))
                        result = {
                            'user_number': user_number,
                            'user_name': user_name,
                            'expiry_info': "ログイン失敗",
                            'expiry_date': datetime(9999, 12, 31)
                        }
                        results.append(result)
                        # リアルタイムでファイルに書き込み
                        with open(output_file, "a", encoding="utf-8") as file:
                            file.write(f"{result['user_number']},{result['user_name']},{result['expiry_info']}\n")
                        continue

                    # マイメニューのドロップダウンを表示
                    try:
                        dropdown_menu = wait.until(EC.element_to_be_clickable((By.ID, "userName")))
                        dropdown_menu.click()
                    except Exception as e:
                        self.update_signal.emit(f"メニュー表示失敗: {user_number} - 次のユーザーに移行します")
                        result = {
                            'user_number': user_number,
                            'user_name': user_name,
                            'expiry_info': "メニュー表示失敗",
                            'expiry_date': datetime(9999, 12, 31)
                        }
                        results.append(result)
                        with open(output_file, "a", encoding="utf-8") as file:
                            file.write(f"{result['user_number']},{result['user_name']},{result['expiry_info']}\n")
                        continue

                    # 利用者情報の変更・削除・更新リンクをクリック
                    try:
                        user_info_link = wait.until(
                            EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '利用者情報の変更・削除・更新')]"))
                        )
                        user_info_link.click()
                    except Exception as e:
                        self.update_signal.emit(f"利用者情報リンククリック失敗: {user_number} - 次のユーザーに移行します")
                        result = {
                            'user_number': user_number,
                            'user_name': user_name,
                            'expiry_info': "リンククリック失敗",
                            'expiry_date': datetime(9999, 12, 31)
                        }
                        results.append(result)
                        with open(output_file, "a", encoding="utf-8") as file:
                            file.write(f"{result['user_number']},{result['user_name']},{result['expiry_info']}\n")
                        continue

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
                            result = {
                                'user_number': user_number,
                                'user_name': user_name,
                                'expiry_info': expiry_info,
                                'expiry_date': expiry_date
                            }
                            results.append(result)
                            # リアルタイムでファイルに書き込み
                            with open(output_file, "a", encoding="utf-8") as file:
                                file.write(f"{result['user_number']},{result['user_name']},{result['expiry_info']}\n")
                        except Exception as e:
                            self.update_signal.emit(f"日付解析エラー: {expiry_info} - {str(e)}")
                            # 解析に失敗しても情報は保存
                            result = {
                                'user_number': user_number,
                                'user_name': user_name,
                                'expiry_info': expiry_info,
                                'expiry_date': datetime(9999, 12, 31)  # 遠い未来の日付
                            }
                            results.append(result)
                            # リアルタイムでファイルに書き込み
                            with open(output_file, "a", encoding="utf-8") as file:
                                file.write(f"{result['user_number']},{result['user_name']},{result['expiry_info']}\n")

                    except Exception as e:
                        self.update_signal.emit(f"有効期限の取得に失敗: {user_number} - 次のユーザーに移行します")
                        # 失敗した場合も結果に追加
                        result = {
                            'user_number': user_number,
                            'user_name': user_name,
                            'expiry_info': "取得失敗",
                            'expiry_date': datetime(9999, 12, 31)  # 遠い未来の日付
                        }
                        results.append(result)
                        # リアルタイムでファイルに書き込み
                        with open(output_file, "a", encoding="utf-8") as file:
                            file.write(f"{result['user_number']},{result['user_name']},{result['expiry_info']}\n")
                        continue

                except Exception as e:
                    self.update_signal.emit(f"ユーザー {user_number} の処理中にエラーが発生: {str(e)}")
                    # ログインに成功していない場合は失敗リストに追加
                    if not login_successful:
                        failed_logins.append((user_number, password, user_name))
                    result = {
                        'user_number': user_number,
                        'user_name': user_name,
                        'expiry_info': "エラー発生" if login_successful else "ログイン失敗",
                        'expiry_date': datetime(9999, 12, 31)  # 遠い未来の日付
                    }
                    results.append(result)
                    # リアルタイムでファイルに書き込み
                    with open(output_file, "a", encoding="utf-8") as file:
                        file.write(f"{result['user_number']},{result['user_name']},{result['expiry_info']}\n")

                # 次のユーザーの処理前に待機
                time.sleep(0.5)

            # 最終的な進捗状況を100%に設定
            self.progress_signal.emit(100)

            # 日付でソートしてファイルを再作成（ソート済み）
            results.sort(key=lambda x: x['expiry_date'])

            # ソート済みデータでファイルを再作成
            with open(output_file, "w", encoding="utf-8") as file:
                file.write("利用者番号,氏名,有効期限\n")
                for result in results:
                    file.write(f"{result['user_number']},{result['user_name']},{result['expiry_info']}\n")

            self.update_signal.emit("\nすべてのデータを日付順にソートしました")
            self.update_signal.emit(f"結果は {output_file} に保存されました")

            # ログイン失敗したアカウントの情報を出力
            if failed_logins:
                self.update_signal.emit("\n=== ログインに失敗したアカウント ===")
                with open(output_file, "a", encoding="utf-8") as file:
                    file.write("\n=== ログインに失敗したアカウント ===\n")
                    for user_number, password, user_name in failed_logins:
                        file.write(f"利用者番号: {user_number}, 氏名: {user_name}\n")
                        self.update_signal.emit(f"利用者番号: {user_number}, 氏名: {user_name}")

            # 今日から2週間以内に有効期限が切れるユーザーを表示
            today = datetime.now()
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
