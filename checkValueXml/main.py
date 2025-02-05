import cv2
import numpy as np
import easyocr
import ddddocr
from PIL import Image
import pytesseract
from PIL import Image
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import time
import io
import os
import requests
import base64
import qrcode
from io import BytesIO

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


def read_3():
    image_path = 'temp.png'
    image = cv2.imread(image_path)
    if image is None:
        print("Không thể đọc được ảnh.")
        exit()

    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    # Tạo mask nhị phân bằng OTSU
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]

    # Trường hợp chữ màu đậm, nền trắng => có thể cần đảo ngược
    # thresh = cv2.bitwise_not(thresh)

    # Tạo kernel để loại bỏ đường kẻ (thử kernel nhỏ để "mở" các nét mảnh)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 5))
    clean = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel)

    # Hoặc nếu đường kẻ có nhiều hướng, bạn phải thử kernel dạng (3x3) hoặc (2x2) để không làm mất chữ
    # kernel = np.ones((2,2), np.uint8)
    # clean = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel)

    # Phóng to ảnh
    height, width = clean.shape
    clean = cv2.resize(clean, (width*2, height*2), interpolation=cv2.INTER_CUBIC)

    # Áp dụng Tesseract
    custom_config = r'--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'
    text = pytesseract.image_to_string(clean, config=custom_config)

    print("Kết quả OCR (loại bỏ đường kẻ mảnh):")
    print(text)



def read_2():
    # Đường dẫn đến ảnh cần xử lý
    image_path = "temp.png"

    # Đọc ảnh bằng OpenCV
    image = cv2.imread(image_path)

    # Chuyển ảnh sang định dạng grayscale (xám) để nhận diện tốt hơn
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    # Tiền xử lý ảnh (tùy chỉnh nếu cần)
    gray = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]

    # Lưu ảnh đã xử lý tạm thời
    temp_image = "temp.png"
    cv2.imwrite(temp_image, gray)

    # Sử dụng Tesseract OCR để trích xuất văn bản
    text = pytesseract.image_to_string(Image.open(temp_image), lang="eng")

    # Hiển thị kết quả
    print("Văn bản nhận diện được:")
    print(text)

def read_captcha(image_path):
    # Xử lý ảnh
    processed = preprocess_image(image_path)
    
    # Lưu ảnh đã xử lý tạm thời
    temp_image_path = "temp.png"
    cv2.imwrite(temp_image_path, processed)
    
    # Đọc text bằng pytesseract
    image = Image.open(temp_image_path).convert("L")  # Chuyển sang grayscale
    result = pytesseract.image_to_string(image, config="--psm 6").strip()
    
    return result

def preprocess_image(image):
    # Đọc ảnh
    img = cv2.imread(image)
    
    # Tách kênh màu đỏ rõ hơn
    _, r, _ = cv2.split(img)
    
    # Tăng độ tương phản
    r = cv2.convertScaleAbs(r, alpha=1.5, beta=0)
    
    # Áp dụng kỹ thuật thresholding cải tiến
    _, binary = cv2.threshold(r, 127, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    
    # Tăng kích thước để dễ nhận dạng
    scaled = cv2.resize(binary, None, fx=2, fy=2)
    
    # Làm dày nét chữ
    kernel = np.ones((2,2), np.uint8)
    dilated = cv2.dilate(scaled, kernel, iterations=1)
    
    return dilated

def login():
    url = "https://gdbhyt.baohiemxahoi.gov.vn/DashboardXml1"
    urlchecklist = "https://gdbhyt.baohiemxahoi.gov.vn/DanhSachHSKCB/Index"
    ma_bv = "79408"
    username = "079201003497"
    password = "079201003497@Dung"

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)

    try:
        driver.get(url)
        wait = WebDriverWait(driver, 10)

        # Nhập Mã BV
        ma_bv_input = wait.until(EC.visibility_of_element_located((By.ID, "macskcb")))
        ma_bv_input.send_keys(ma_bv)

        # Nhập Tên Đăng Nhập
        username_input = driver.find_element(By.ID, "username")
        username_input.send_keys(username)

        # Nhập Mật Khẩu
        password_input = driver.find_element(By.ID, "password")
        password_input.send_keys(password)

        # Chụp ảnh CAPTCHA
        captcha_img = driver.find_element(By.ID, "Captcha_IMG1")
        captcha_path = "captcha.png"
        captcha_img.screenshot(captcha_path)

        # Tạo mã QR từ hình ảnh CAPTCHA
        qr = qrcode.QRCode()
        qr.add_data(captcha_path)
        qr.make(fit=True)

        # Hiển thị mã QR trong terminal
        qr.print_ascii()

      
        check = True
        while check:
            # Nhập mã CAPTCHA thủ công
            captcha_code = input("Nhập mã CAPTCHA từ ảnh: ")

            # Nhập mã CAPTCHA vào ô
            captcha_input = driver.find_element(By.ID, "Captcha_TB_I")
            captcha_input.send_keys(captcha_code)
            time.sleep(1)   
            # Nhấn nút đăng nhập
            login_button = driver.find_element(By.CLASS_NAME, "btn_dangNhap")
            login_button.click()
            print("Đang đăng nhập...")
            time.sleep(5)

            # Kiểm tra xem có lỗi gì không
            try:
                error_message = driver.find_element(By.ID, "alert").text
                if error_message:
                    check = True
                    print(f"Lỗi đăng nhập: {error_message}")
                else:
                    check = False
                    print("Đăng nhập thành công!")
            except:
                check = False
                print("Đăng nhập thành công!")
        driver.get(urlchecklist)
        time.sleep(1)
        
        buttonchose = driver.find_element(By.ID, "cb_TrangThaiTT_B-1")
        buttonchose.click()
        
        time.sleep(1)

        cho_danh_sach = driver.find_element(By.ID, "cb_TrangThaiTT_DDD_L_LBI0T0")
        cho_danh_sach.click()

        time.sleep(1)

        timkiem = driver.find_element(By.ID,"bt_TimKiem")
        timkiem.click()
        
    except Exception as e:
        print(f"Lỗi: {e}")
    finally:
        input("Nhấn Enter để đóng trình duyệt...")
        driver.quit()

# Gọi hàm login

login()
