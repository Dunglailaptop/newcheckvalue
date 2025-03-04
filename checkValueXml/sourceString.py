from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from rx.subject import BehaviorSubject
import rx
import time
import json
import os
import threading

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "source", "prescriptionDetail.json")


CONFIG_PATH_ICON_APP = os.path.join(BASE_DIR, "source", "crawlLogo.ico")
CONFIG_PATH_IMAGE_BACKGROUND_APP = os.path.join(BASE_DIR, "source", "crawldata.png")
CONFIG_PATH_IMAGE_BACKGROUND_APP_3 = os.path.join(BASE_DIR, "source", "BG.png")
CONFIG_PATH_IMAGE_BACKGROUND_APP_2 = os.path.join(BASE_DIR, "source", "BG2.jpg")
CONFIG_PATH_INVOICEOUTPATIENTDETAIL = os.path.join(BASE_DIR, "source", "invoiceoutpatientdetail.json")
CONFIG_PATH = os.path.join(BASE_DIR, "source", "stt.json")
terminal_destroy = any



Appkey = "55abbe84f48701bc8b3873c72c804bac7a70b3ed2_2942_webapp"
Appkey2 = "55abbe84f48701bc8b3873c72c804bac7a70b3ed2_2942_webapp"
ConnectStr = {
        "host": "192.168.0.127",
        "port": "5432",
        "database": "nhidong2_crawl",
        "user": "sa",
        "password": "It@7477"
    }
DataGet = {
        "host": "localhost",
        "port": "5222",
        "database": "crawl_nhidong2",
        "user": "postgres",
        "password": "1234"
    }
secretKey = ""
secretKey2 = ""
driverGlobal = any
driverGlobal2 = any
contentType = "application/json; charset=utf-8"
urlGetPrescriptiondetail = "http://192.168.0.77/api/doctor_pkg/load_pkg_last_detail_presc"
urlGetinvoiceoutdetail = "http://192.168.0.80/api/patients/getOutPatientInvoiceDetail"
urlApiGetListXML = "http://192.168.0.77/api/tool_send_BHYT/findPaging_nd2"
urlGetDetailXML = "http://192.168.0.77/api/tool_send_BHYT/readXML"
Any = object()

def _login_(usernameStr, passwordStr):
    global driverGlobal, secretKey
    try:
        form = driverGlobal.find_element(By.NAME, "LoginForm")
        username = driverGlobal.find_element(By.NAME, 'username')
        password = driverGlobal.find_element(By.NAME, 'password')
        username.send_keys(usernameStr)
        password.send_keys(passwordStr)
        # ấn nút login
        form.submit()
        time.sleep(0.3)
        for request in reversed(driverGlobal.requests):
            if request.response:
                if 'user/login' in request.url:
                    # format cho dữ liệu về dạng json
                    try:
                        secur = json.loads(request.response.body)
                        # đọc data
                        keyUser = secur.get('data', {}).get('security', {})
                        secretKey = keyUser.get('secret')
                        break
                    except:
                        continue
        return secretKey
    except:
        return ""

def _initSelenium_():
    global driverGlobal
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option("prefs", {
        "profile.default_content_setting_values.autom   atic_downloads": 1,
        "download.prompt_for_download": False,
        "profile.default_content_settings.popups": 0,
        "safebrowsing.enabled": "false",
        "safebrowsing.disable_download_protection": True
    })
    prefs = {"credentials_enable_service": False,
                            "profile.password_manager_enabled": False}
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--headless")  # Chạy trình duyệt trong chế độ headless
    chrome_options.add_argument("--disable-gpu")  # Tăng tốc độ trên các hệ điều hành không có GPU
    chrome_options.add_argument("--window-size=1920x1080")  # Thiết lập kích thước cửa sổ mặc định
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_argument("--blink-settings=imagesEnabled=false")
    driver = webdriver.Chrome(options=chrome_options)
    url = "http://192.168.0.77/dist/#!/login"
    try:
        driver.get(url)
        driverGlobal = driver
        return True
    except Exception as e:
        print(e)
        return False
    
def _destroySelenium_():
    global driverGlobal
    driverGlobal.close()
    
    

