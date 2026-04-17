#!/usr/bin/env python
# -*- coding: utf-8 -*-

import totp
import os
import sys
import glob
from datetime import datetime
import time
import shutil
import argparse
import selenium.webdriver as webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from config_loader import load_config
from month_utils import resolve_target_month
from path_utils import expand_path

def get_base_path() -> str:
    # PyInstallerで実行されているかどうかをチェック
    if getattr(sys, "frozen", False):
        # EXEの実行ファイルのパスを取得
        return os.path.dirname(sys.executable)
    else:
        # スクリプトの実行ファイルのパスを取得
        return os.path.dirname(os.path.abspath(__file__))
    
def wait_for_download(tmp_download_dir: str, timeout_second: int = 60) -> str:
    """ダウンロード完了まで待機し、ファイルパスを返す"""
    for i in range(timeout_second + 1):
        download_fileName = glob.glob(f'{tmp_download_dir}\\*.*')
        if download_fileName:
            extension = os.path.splitext(download_fileName[0])
            if ".crdownload" not in extension[1]:
                time.sleep(2)
                return download_fileName[0]
        if i >= timeout_second:
            raise Exception("ダウンロードに失敗しました。")
        time.sleep(1)
    raise Exception("ダウンロードに失敗しました。")

def download_attendance(url, username, password, secret_key):
    # Initialize TOTP
    totp_generator = totp.Totp(secret_key)

    tmp_download_dir = os.path.join(get_base_path(), 'Download', datetime.now().strftime('%Y%m%d_%H%M%S'))
    download_dir = os.path.join(get_base_path(), 'Download')
    os.makedirs(tmp_download_dir, exist_ok=True)

    # ------ ChromeDriver のオプション ------
    service = Service()
    service.creation_flags = 0x08000000  # ヘッドレスモードで DevTools listening on ws:~~ を表示させない
    options = Options()
    options.add_argument("--headless=new")  # ヘッドレスモードで起動する
    prefs = {
        "credentials_enable_service": False,  # パスワード保存のポップアップを無効
        "savefile.default_directory": tmp_download_dir,  # ダイアログ(名前を付けて保存)の初期ディレクトリを指定
        "download.default_directory": tmp_download_dir,  # ダウンロード先を指定
        "download_bubble.partial_view_enabled": False,  # ダウンロードが完了したときの通知(吹き出し/下部表示)を無効にする
        "plugins.always_open_pdf_externally": True,  # Chromeの内部PDFビューアを使わない(URLにアクセスすると直接ダウンロードされる)
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    }
    options.add_experimental_option("prefs", prefs)

    # ------ ChromeDriver の起動 ------
    driver = webdriver.Chrome(service=service, options=options)


    try:
        # Navigate to the login page
        driver.get(url)

        # Find the username and password fields and enter the credentials
        driver.find_element(By.XPATH, "//input[@id='username']").send_keys(username)
        driver.find_element(By.XPATH, "//input[@id='password']").send_keys(password)

        driver.find_element(By.XPATH, "//input[@value='ログイン']").click()

        # Wait for the page to load
        driver.implicitly_wait(2)  # Wait for elements to load

        # Find the TOTP field and enter the generated TOTP code
        totp_code = totp_generator.generate_totp()
        driver.find_element(By.XPATH, "//input[@id='tc']").send_keys(totp_code)

        # Submit the form
        driver.find_element(By.XPATH, "//input[@value=' 検証 ']").click()

        # Wait for the page to load and download attendance
        driver.implicitly_wait(2)  # Wait for elements to load

        Select(driver.find_element(By.XPATH, "//select[@id='enc']")).select_by_value("MS932")
        Select(driver.find_element(By.XPATH, "//select[@id='xf']")).select_by_value("localecsv")

        driver.implicitly_wait(1)
        driver.find_element(By.XPATH, "//input[@value='エクスポート']").click()

        wait_for_download(tmp_download_dir)

        download_file = wait_for_download(tmp_download_dir)
        filePath = shutil.move(download_file, os.path.join(download_dir, os.path.basename(download_file)))
        return filePath

    finally:
        # Close the browser
        if driver:
            driver.quit()
        if os.path.exists(tmp_download_dir):
            shutil.rmtree(tmp_download_dir)

def main():
    parser = argparse.ArgumentParser(description="AppsFSからWeb勤怠をダウンロード")
    parser.add_argument("--config", default="config/config.yaml", help="config.yaml のパス")
    args = parser.parse_args()
    config = load_config(args.config)

    username = os.environ.get("APPS_ID")
    if not username:
        raise RuntimeError("AppsFS ID not found")
    password = os.environ.get("APPS_PASSWORD")
    if not password:
        raise RuntimeError("AppsFS Password not found")
    secret_key = os.environ.get("APPS_TOTP_SECRET")
    if not secret_key:
        raise RuntimeError("TOTP Secret not found")
    
    apps_url = f'{config["apps"]["base_url"]}/{config["apps"]["filter_id"]}{config["apps"]["csv_suffix"]}'
    vars = resolve_target_month(config)
    csvPath = expand_path(config["paths"]["web_attendance_file"], vars)

    shutil.move(download_attendance(apps_url, username, password, secret_key), csvPath)

if __name__ == "__main__":
    main()
    