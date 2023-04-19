import requests
from bs4 import BeautifulSoup
import os
import zipfile

# 現在のChromeバージョンを取得する関数
def get_chrome_version():
    url = 'https://www.whatismybrowser.com/guides/the-latest-version/chrome'
    res = requests.get(url)
    soup = BeautifulSoup(res.text, 'html.parser')
    return soup.select_one('.browser--version').text.strip()

# 最新のchromedriverをダウンロードする関数
def download_chromedriver(version):
    download_url = f"https://chromedriver.storage.googleapis.com/{version}/chromedriver_win32.zip"
    res = requests.get(download_url)
    with open("chromedriver.zip", "wb") as f:
        f.write(res.content)
    with zipfile.ZipFile("chromedriver.zip") as zip_f:
        zip_f.extractall()

# chromedriverのバージョンチェックと自動更新
def check_driver_update():
    current_version = os.popen('chromedriver --version').read().split(' ')[1].strip()
    chrome_version = get_chrome_version().split(' ')[0]
    if current_version != chrome_version:
        print(f"[INFO] Current chromedriver version: {current_version}")
        print(f"[INFO] Latest chromedriver version: {chrome_version}")
        download_chromedriver(chrome_version)

check_driver_update()
