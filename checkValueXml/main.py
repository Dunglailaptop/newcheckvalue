from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import time
import qrcode
import json
import threading
import sourceString as sour
import base64
import customtkinter
import customtkinter as ctk
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
import xml.etree.ElementTree as ET
import os
import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
from datetime import datetime
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "source", "stt.json")

#center window
def center_window(window,width=600,height=400):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x=(screen_width / 2) - (width / 2)
    y=(screen_height / 2 ) - (height / 2)
    window.geometry(f'{width}x{height}+{int(x)}+{int(y)}')

#GET LIST FILE XML 
def getlistfilexml(tree):
    # Xóa dữ liệu cũ trong Treeview trước khi tải mới
    for item in tree.get_children():
        tree.delete(item)

    # Chọn nhiều file XML
    file_paths = filedialog.askopenfilenames(title="Chọn file XML", filetypes=[("XML Files", "*.xml")])

    # Danh sách lưu kết quả JSON
    json_data = []
    stt_counter = 1  # Bắt đầu STT từ 1

    if file_paths:
        for file_path in file_paths:
            try:
                # Parse XML file gốc
                xml_tree = ET.parse(file_path)
                root = xml_tree.getroot()

                # Tìm các FILEHOSO có LOAIHOSO = XML1
                for filehoso in root.findall(".//FILEHOSO"):
                    loai_hoso = filehoso.find("LOAIHOSO").text.strip()
                    if loai_hoso == "XML1":
                        noidung_base64 = filehoso.find("NOIDUNGFILE").text.strip()

                        if noidung_base64:  # Kiểm tra nếu có nội dung Base64
                            try:
                                # Giải mã Base64
                                xml_decoded = base64.b64decode(noidung_base64).decode("utf-8")

                                # Parse XML đã giải mã
                                decoded_root = ET.fromstring(xml_decoded)

                                # Lấy dữ liệu cần thiết
                                ma_bn = decoded_root.find("MA_BN").text.strip() if decoded_root.find("MA_BN") is not None else ""
                                ma_the_bhyt = decoded_root.find("MA_THE_BHYT").text.strip() if decoded_root.find("MA_THE_BHYT") is not None else ""
                                ngay_xuatvien = decoded_root.find("NGAY_RA").text.strip() if decoded_root.find("NGAY_RA") is not None else ""
                                # Thêm vào danh sách JSON
                                json_data.append({
                                    "STT": stt_counter,
                                    "MA_BN": ma_bn,
                                    "MA_THE_BHYT": ma_the_bhyt,
                                    "NGAY_RA":ngay_xuatvien,
                                    "TRANGTHAI": 0
                                })

                                stt_counter += 1  # Tăng STT
                            except Exception as decode_error:
                                print(f"Lỗi giải mã Base64 trong file {file_path}: {decode_error}")

            except Exception as e:
                print(f"Lỗi khi xử lý file {file_path}: {e}")

        # Ghi dữ liệu ra file JSON
        output_file = "output.json"
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(json_data, f, indent=4, ensure_ascii=False)

        print(f"Đã lưu kết quả vào {output_file}")

        # Load dữ liệu lên bảng
        loaddatatable(json_data, tree)  # Truyền danh sách JSON thay vì file
    else:
        print("Không có file nào được chọn!")


#setup treeview table
def setupTreeviewTable(root,tree):
    # Tạo cửa sổ Tkinter
    # Tạo Treeview
 

    # Đặt tiêu đề cho cột
    tree.heading("STT", text="STT")
    tree.heading("MA_BN", text="Mã Bệnh Nhân")
    tree.heading("MA_THE_BHYT", text="Mã Thẻ BHYT")
    tree.heading("TRANGTHAI", text="Trạng Thái")

    # Điều chỉnh độ rộng của cột
    tree.column("STT", width=50, anchor="center")
    tree.column("MA_BN", width=150, anchor="center")
    tree.column("MA_THE_BHYT", width=150, anchor="center")
    tree.column("TRANGTHAI", width=100, anchor="center")

   

    # Thêm thanh cuộn (Scroll Bar)
    scrollbar = ttk.Scrollbar(root, orient="vertical", command=tree.yview)
    tree.configure(yscroll=scrollbar.set)
    scrollbar.pack(side="right", fill="y")

    # Hiển thị Treeview
    tree.pack(expand=True, fill="both")

#load table
def loaddatatable(json_file,tree):
    # Xóa dữ liệu cũ trước khi nạp mới
    for item in tree.get_children():
        tree.delete(item)
     # Đọc dữ liệu từ file JSON và hiển thị lên Treeview
    try:
        for row in json_file:
            tree.insert("", "end", values=(row["STT"], row["MA_BN"], row["MA_THE_BHYT"], row["TRANGTHAI"]))

    except Exception as e:
        print(f"Lỗi khi đọc file JSON: {e}")
#setup app
def setupapp():
    # Tạo cửa sổ chính
    app = ctk.CTk()
    app.title("Ứng dụng đọc kết quả xml")
    center_window(app)

    # Tạo bảng (Treeview)
    columns = ("STT", "MA_BN", "MA_THE_BHYT", "TRANGTHAI")
    tree = ttk.Treeview(app, columns=columns, show="headings")
    setupTreeviewTable(app,tree)

  
    label = ctk.CTkLabel(app,100,20,bg_color="red",text="5",text_color="white")
    label.pack(pady=5)


    # Hàm xử lý khi nhấn nút
    def on_button_click():
        messagebox.showinfo("Thông báo", "Bạn đã nhấn nút!")

    def on_button_2_click():
        messagebox.showinfo("Thông báo", "Bạn đã nhấn nút thứ hai!")

    # Tạo nút bấm
    button1 = ctk.CTkButton(app, text="Chọn file", command=lambda: open_export(app))
    button1.pack(side="left", padx=5, pady=5)

    button2 = ctk.CTkButton(app, text="Đọc kết quả", command=lambda: login())
    button2.pack(side="left", padx=5, pady=5)

    
    # Chạy ứng dụng
    app.mainloop()

#load file dem
def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH,'r') as f:
            return json.load(f)
    return {"STT":1}

def find_record_by_stt(file_path, stt):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
            
            if not isinstance(data, list):
                raise ValueError("JSON không chứa danh sách hồ sơ hợp lệ.")
            
            for record in data:
                if record.get("STT") == stt:
                    return record
            
            return None  # Không tìm thấy
    except (FileNotFoundError, json.JSONDecodeError, ValueError) as e:
        print(f"Lỗi: {e}")
        return None

def update_record_status(file_path, stt, new_status):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
            
            if not isinstance(data, list):
                raise ValueError("JSON không chứa danh sách hồ sơ hợp lệ.")
            
            updated = False
            for record in data:
                if record.get("STT") == stt:
                    record["TRANGTHAI"] = new_status
                    updated = True
                    break
            
            if updated:
                with open(file_path, 'w', encoding='utf-8') as file:
                    json.dump(data, file, ensure_ascii=False, indent=4)
                print(f"Cập nhật trạng thái cho STT {stt} thành {new_status} thành công.")
            else:
                print(f"Không tìm thấy hồ sơ với STT {stt}.")
    except (FileNotFoundError, json.JSONDecodeError, ValueError) as e:
        print(f"Lỗi: {e}")


def save_config(config):
    with open(CONFIG_PATH, 'w') as f:
        json.dump(config, f)

def update_json(stt_value):
    config = load_config()
    config["STT"] = stt_value
    save_config(config)

def scrape_table_data(driver, timeget, checkcount):
    # Chờ bảng load xong, tránh bị lỗi dữ liệu chập chờn
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "tr.dxgvDataRow_EIS"))
        )
    except:
        print("Không tìm thấy dữ liệu trong bảng")
        return False, True

    time.sleep(2)
    
    rows = driver.find_elements(By.CSS_SELECTOR, "tr.dxgvDataRow_EIS")
    
    # Nếu không tìm thấy hàng nào, trả về False
    if not rows:
        print("Không có dữ liệu trong bảng.")
        return False, True

    check = False
    records = None

    # if len(rows) >= 1 & len(rows) <= 2:
    checkcount = True
    row = rows[0]  # Chỉ có 1 dòng, lấy dòng đầu tiên
    
    for rowchil in rows:
        columns = rowchil.find_elements(By.TAG_NAME, "td")
        if len(columns) > 10:  # Đảm bảo có ít nhất 11 cột
            records = columns[11].text.strip()
            if timeget == convert_date_format(records):
                check = True
        else:
            print("Lỗi: Không đủ số cột trong bảng.")
    # else:
    #     checkcount = False  # Nếu có nhiều hơn 1 dòng, đánh dấu là False

    return check, checkcount

def convert_date_format(date_str):
    """Chuyển đổi định dạng '06/02/2025 08:25' thành '20250206'"""
    try:
        date_obj = datetime.strptime(date_str, "%d/%m/%Y %H:%M")
        return date_obj.strftime("%Y%m%d")
    except ValueError:
        print("Định dạng ngày không hợp lệ.")
        return None

def truncate_timestamp(timestamp):
    """Loại bỏ 4 số cuối cùng của chuỗi timestamp"""
    return timestamp[:-4] if len(timestamp) > 4 else timestamp

def convert_timestamp_to_date(timestamp_str):
    """Chuyển đổi '202502062150' thành '06/02/2025'"""
    try:
        date_obj = datetime.strptime(timestamp_str[:8], "%Y%m%d")  # Chỉ lấy 8 ký tự đầu
        return date_obj.strftime("%d/%m/%Y")
    except ValueError:
        print("Định dạng timestamp không hợp lệ.")
        return None


def CHECKVLAUE(driver, timkiembtn):
    config = load_config()
    STT = config["STT"]
    checkstatusfail = False
    checkcount = True

    while True:
        try:
            file_path = "output.json"
            result = find_record_by_stt(file_path, STT)

            if len(result) > 0:
                datetimeneedcheck = truncate_timestamp(result["NGAY_RA"])

                def get_element(by, value, timeout=10, retries=3):
                    """ Lấy phần tử với retry nếu gặp lỗi 'stale element reference' """
                    for _ in range(retries):
                        try:
                            return WebDriverWait(driver, timeout).until(
                                EC.presence_of_element_located((by, value))
                            )
                        except Exception as e:
                            print(f"Lỗi khi lấy phần tử {value}: {str(e)} - Thử lại...")
                            time.sleep(1)  # Đợi trước khi thử lại
                    raise Exception(f"Lỗi: Không tìm thấy phần tử {value} sau {retries} lần thử")

                # Chuyển frame nếu cần thiết
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame("frame_name")  # Thay "frame_name" bằng tên frame đúng
                except:
                    pass  # Nếu không có frame thì bỏ qua

                # Tìm và nhập dữ liệu vào ô MA_THE_BHYT
                txtValue = get_element(By.ID, "gvDanhSachHoSo_DXFREditorcol4_I")
                driver.execute_script("arguments[0].value = '';", txtValue)
                txtValue.send_keys(result["MA_THE_BHYT"])
                time.sleep(1)

                # Tìm và nhập dữ liệu vào ô NGAY_RA
                txtValue2 = get_element(By.ID, "gvDanhSachHoSo_DXFREditorcol10_I")
                driver.execute_script("arguments[0].value = '';", txtValue2)
                txtValue2.send_keys(convert_timestamp_to_date(result["NGAY_RA"]))
                txtValue2.send_keys(Keys.ENTER) 
                time.sleep(1)

                #nut tim kiem
                # timkiembtn.click()   
                # Chờ xử lý và kiểm tra kết quả
                check1, check2 = scrape_table_data(driver, datetimeneedcheck, checkcount)

                if check1:
                    print(f"Dữ liệu: {STT} - {result['MA_THE_BHYT']}")
                    update_record_status(file_path, STT, 1)

                    # Tìm lại phần tử trước khi xóa dữ liệu (tránh lỗi stale)
                    txtValue = get_element(By.ID, "gvDanhSachHoSo_DXFREditorcol4_I")
                    txtValue2 = get_element(By.ID, "gvDanhSachHoSo_DXFREditorcol10_I")

                    driver.execute_script("arguments[0].value = '';", txtValue)
                    driver.execute_script("arguments[0].value = '';", txtValue2)
                    time.sleep(1)

                else:
                    update_record_status(file_path, STT, 0)
                    print("Không có dữ liệu")

                # Kiểm tra điều kiện để tiếp tục vòng lặp
                if not checkstatusfail and check2:
                    STT = int(STT) + 1
                    update_json(stt_value=STT)
                else:
                    checkcount = False
            else:
                break

        except Exception as e:
            print("Xảy ra lỗi: " + str(e))
            time.sleep(2)  # Chờ trước khi thử lại
            driver.refresh()  # Làm mới trang nếu gặp lỗi

# Biến toàn cục để giữ cửa sổ nhập CAPTCHA
captcha_window = None
captcha_code = None

def load_image(file_path,captcha_window):
 

    # Tạo cửa sổ mới
    captcha_window = tk.Toplevel()
    captcha_window.title("Nhập mã CAPTCHA")

    # Load và hiển thị ảnh
    img = Image.open(file_path)
    img = img.resize((200, 80), Image.Resampling.LANCZOS)  # Resize phù hợp
    img_tk = ImageTk.PhotoImage(img)

    # Hiển thị ảnh
    image_label = tk.Label(captcha_window, image=img_tk)
    image_label.image = img_tk  # Giữ tham chiếu để tránh bị xóa bộ nhớ
    image_label.pack(pady=10)

    # Ô nhập CAPTCHA
    entry = tk.Entry(captcha_window, font=("Arial", 14))
    entry.pack(pady=10)
    captcha_var = tk.StringVar()
    # Nút xác nhận
    def submit_captcha():
        captcha_var.set(entry.get())  # Cập nhật giá trị vào biến StringVar
        captcha_window.destroy()  # Đóng cửa sổ sau khi nhập
        
       # Đóng cửa sổ sau khi nhập

    btn_submit = tk.Button(captcha_window, text="Xác nhận", command=submit_captcha)
    btn_submit.pack(pady=5)
    captcha_window.wait_variable(captcha_var)
    return captcha_var.get() 


def open_export(app):
    import mainExportXML as mainex
    mainex.settupAppBeginStart(app)

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

        # Nhập CAPTCHA
        check = True
        while check:
            root = tk.Tk()
            root.withdraw()  # Hide tkinter root window
            code = load_image(captcha_path, root)
            root.destroy()

            # Nhập CAPTCHA vào ô
            captcha_input = driver.find_element(By.ID, "Captcha_TB_I")
            captcha_input.clear()
            captcha_input.send_keys(code)
            time.sleep(1)

            # Nhấn nút đăng nhập
            login_button = driver.find_element(By.CLASS_NAME, "btn_dangNhap")
            login_button.click()
            print("Đang đăng nhập...")
            time.sleep(5)

            # Kiểm tra lỗi đăng nhập
            try:
                error_message = driver.find_element(By.ID, "alert").text
                if error_message:
                    print(f"Lỗi đăng nhập: {error_message}")
                    check = True  # Thử nhập lại CAPTCHA nếu có lỗi
                else:
                    check = False
                    print("Đăng nhập thành công!")
            except:
                check = False
                print("Đăng nhập thành công!")

        time.sleep(3)
        driver.get(urlchecklist)
        time.sleep(1)
        
        hamxuly1(driver)


    except Exception as e:
        print(f"Lỗi: {e}")
    finally:
        input("Nhấn Enter để đóng trình duyệt...")
        driver.quit()


def hamxuly1(driver):
        # Mở dropdown trạng thái
    buttonchose = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "cb_TrangThaiTT_B-1"))
    )

    # Cuộn đến phần tử để đảm bảo phần tử không bị che khuất
    driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", buttonchose)

    # Dùng ActionChains để nhấn vào phần tử
    actions = ActionChains(driver)
    actions.move_to_element(buttonchose).click().perform()
    time.sleep(1)

    # Chọn "Chờ danh sách"
    cho_danh_sach = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "cb_TrangThaiTT_DDD_L_LBI0T0"))
    )
    cho_danh_sach.click()
    time.sleep(1)

    # Nhấn nút tìm kiếm
    timkiem = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "bt_TimKiem"))
    )
    driver.execute_script("arguments[0].scrollIntoView();", timkiem)
    time.sleep(1)
    timkiem.click()

    # Gọi hàm xử lý tiếp
    CHECKVLAUE(driver, timkiem)

# Gọi hàm login

setupapp()
