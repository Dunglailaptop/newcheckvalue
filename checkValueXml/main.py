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
from typing import List, Dict, Any
from PIL import Image, ImageTk
from datetime import datetime, timezone
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from tkcalendar import DateEntry
from rx.subject import Subject
import pandas as pd
from unidecode import unidecode

treemain = Any
cols = Any
folderurl = ""
datacount = 0

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "source", "stt.json")

#center window
def center_window(window,width=900,height=400):
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
def load_json_to_treeview(tree, important_columns=None):

    """Đọc file JSON, chỉ hiển thị các cột cần thiết và tự động điều chỉnh kích thước cột."""
    global folderurl,datacount
    file_path = os.path.join(folderurl, "dataJson.json")

    if not os.path.exists(file_path):
        print(f"⚠️ File {file_path} không tồn tại! Hãy chọn thư mục trước.")
        return

    try:
        with open(file_path, "r", encoding="utf-8") as file:
            data = json.load(file)
        datacount = len(data)
        if not data:
            print("⚠️ JSON rỗng, không thể cập nhật dữ liệu!")
            return

        # Nếu không có danh sách cột cần thiết, lấy toàn bộ cột trong JSON
        all_columns = list(data[0].keys())[:31]  # Giới hạn 31 cột đầu tiên
        columns = important_columns if important_columns else all_columns

        # Lọc bỏ những cột không có trong dữ liệu
        columns = [col for col in columns if col in all_columns]

        # Nếu không còn cột nào hợp lệ
        if not columns:
            print("⚠️ Không có cột nào hợp lệ để hiển thị!")
            return

        # Xóa tất cả cột hiện có trong Treeview
        tree["columns"] = columns
        for col in columns:
            tree.heading(col, text=col)

        # Xóa dữ liệu cũ trong Treeview
        tree.delete(*tree.get_children())

        # Lưu trữ dữ liệu & đo kích thước cột
        row_data = []
        max_widths = {col: len(col) * 10 for col in columns}  # Bắt đầu với tên cột

        for item in data:
            values = [item.get(col, "N/A") for col in columns]  # Auto-fill "N/A" nếu thiếu
            row_data.append(values)

            # Cập nhật chiều rộng cột dựa trên dữ liệu
            for col, value in zip(columns, values):
                max_widths[col] = max(max_widths[col], len(str(value)) * 8)

        # Thêm dữ liệu vào Treeview
        for values in row_data:
            tree.insert("", "end", values=values)

        # Cập nhật chiều rộng cột tự động
        for col in columns:
            tree.column(col, width=max(100, min(max_widths[col], 300)), anchor="center")  # Giới hạn tối đa 300px

        print(f"✅ Đã cập nhật dữ liệu từ {file_path} vào Treeview với các cột cần thiết!")

    except Exception as e:
        print(f"❌ Lỗi khi đọc file JSON: {e}")




#hàm tạo treeview
def setupTreeviewTable(root):
    """Tạo Treeview rỗng, chưa có cột"""
    # Tạo Frame chứa Treeview
    frame = ttk.Frame(root)
    frame.pack(expand=True, fill="both", padx=10, pady=10)

    # Tạo Treeview rỗng (chưa có cột)
    tree = ttk.Treeview(frame, show="headings")

    # Thêm thanh cuộn
    scrollbar_y = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    scrollbar_x = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)

    # Định vị layout
    tree.grid(row=0, column=0, sticky="nsew")
    scrollbar_y.grid(row=0, column=1, sticky="ns")
    scrollbar_x.grid(row=1, column=0, sticky="ew")

    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)

    return tree


def setupTreeview(tree,label):
    global folderurl,datacount
    folderurl = get_urlFolder() + "/"
    # Khi có dữ liệu, cập nhật cột và dữ liệu
    load_json_to_treeview(tree,important_columns=["stt","insurance_code", "patient_code", "ngay_ra","trangthai"])
    label.configure(text=datacount)


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

#hàm lấy ngày chọn
def get_date(cal):
    selected_date = cal.get_date()  # Lấy ngày từ combobox (YYYY-MM-DD)
    now_utc = datetime.now(timezone.utc)  # Lấy thời gian hiện tại UTC

    # Chuyển đổi ngày từ combobox sang datetime nhưng giữ nguyên thời gian của now_utc
    selected_datetime = datetime.strptime(str(selected_date), "%Y-%m-%d").replace(
        hour=now_utc.hour, minute=now_utc.minute, second=now_utc.second,
        microsecond=now_utc.microsecond, tzinfo=timezone.utc
    )

    formatted_date = selected_datetime.isoformat(timespec="milliseconds").replace("+00:00", "Z")
    print(f"Ngày đã chọn (UTC): {formatted_date}")
    return formatted_date



def chon_file_excel(app):
    """
    Hiển thị hộp thoại để người dùng chọn một file Excel và trả về đường dẫn file.
    """
   
   
    file_path = filedialog.askopenfilename(
        title="Chọn file Excel",
        filetypes=[("Excel Files", "*.xlsx;*.xls")]
    )

    return file_path

def chuan_hoa_ten_cot(ten_cot):
    """
    Chuẩn hóa tên cột: loại bỏ dấu, xóa khoảng trắng, viết thường.
    """
    ten_cot = unidecode(ten_cot)  # Bỏ dấu tiếng Việt
    ten_cot = ten_cot.replace(" ", "").lower()  # Xóa khoảng trắng, viết thường
    return ten_cot

def json_to_excel(json_file, excel_file, columns=None):
    """
    Chuyển đổi file JSON thành file Excel, chỉ xuất các cột được chỉ định.

    :param json_file: Đường dẫn file JSON đầu vào
    :param excel_file: Đường dẫn file Excel đầu ra
    :param columns: Danh sách các cột cần xuất (list), nếu None thì xuất tất cả
    """
    try:
        # Đọc dữ liệu từ file JSON
        df = pd.read_json(json_file)

        # Nếu có danh sách cột, chỉ giữ lại các cột đó
        if columns:
            df = df[columns]

        # Ghi dữ liệu ra file Excel
        df.to_excel(excel_file, index=False, engine="openpyxl")

        print(f"Đã tạo file Excel: {excel_file}")

    except Exception as e:
        print(f"Lỗi: {e}")

def excel_to_json(file_path, output_json):
    """
    Chuyển đổi file Excel thành JSON, chuẩn hóa tên cột.

    :param file_path: Đường dẫn file Excel
    :param output_json: Đường dẫn file JSON đầu ra
    """
    try:
        # Đọc file Excel
        df = pd.read_excel(file_path, engine="openpyxl")

        # Chuẩn hóa tên cột
        df.columns = [chuan_hoa_ten_cot(col) for col in df.columns]

        # Chuyển DataFrame thành danh sách JSON
        data_json = df.to_dict(orient="records")

        # Ghi vào file JSON
        with open(output_json, "w", encoding="utf-8") as f:
            json.dump(data_json, f, ensure_ascii=False, indent=4)

        print(f"Đã tạo file JSON: {output_json}")

    except Exception as e:
        print(f"Lỗi: {e}")

def converttojsonexportkqxml(app):
    global folderurl
    file_path = os.path.join(folderurl, "dataJson.json")

    # Kiểm tra đường dẫn hợp lệ
    if not folderurl or not file_path:
        messagebox.showinfo(title="LỖI!", message="VUI LÒNG CHỌN FILE JSON ĐÃ XUẤT")
        return
    
    print(f"✅ Bắt đầu xử lý với folder: {folderurl}")

    try:
        # Load dữ liệu JSON
        data = loadfind_json()
        if not data:
            print("❌ Lỗi: Không có dữ liệu trong dataJson.json")
            return
        
        print(f"🔍 Số lượng bản ghi: {len(data)}")

        # Chọn file Excel
        path = chon_file_excel(app)
        if not path:
            print("❌ Lỗi: Không chọn được file Excel")
            return
        
        print(f"📂 File Excel được chọn: {path}")

        # Chuyển đổi Excel -> JSON
        excel_to_json(path, "dataXmlResult.json")

        for item in data:
            record = kiemtraketqua("dataXmlResult.json", item["patient_code"], item["insurance_code"], item["ngay_ra"])
            
            if record:
                print(f"✅ Có kết quả trả về cho STT {item['stt']}")
                update_json_data_kq(file_path, item["stt"], 1)
            else:
                print(f"❌ Không có kết quả cho STT {item['stt']}")
                update_json_data_kq(file_path, item["stt"], 0)

        # Xuất file Excel
        columns_can_xuat = ["stt", "insurance_code", "patient_code", "ins_transaction_code", "ngay_ra", "trangthai"]
        output_excel_path = os.path.join(folderurl, "dataKQ.xlsx")
        json_to_excel(file_path, output_excel_path, columns_can_xuat)

        print(f"✅ File Excel đã xuất: {output_excel_path}")

    except Exception as e:
        print(f"🚨 Lỗi trong quá trình xử lý: {e}")
    


def update_json_data_kq(file_path,stt,trangthaiup):
    """
    Cập nhật dữ liệu XML trong tệp JSON theo số thứ tự (stt).
    
    Args:
        file_path (str): Đường dẫn đến tệp JSON.
        stt (int): Số thứ tự của phần tử cần cập nhật.
        xml1, xml2, xml3, xml4, xml5, xml7 (str): Dữ liệu cần mã hóa và cập nhật.
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    for item in data:
        if item.get("stt") == stt:
            item["trangthai"] = trangthaiup
            break
    
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)


def loadfind_json():
    global folderurl
    data = []
    file_path = os.path.join(folderurl, "dataJson.json")
    with open(file_path, 'r', encoding='utf-8') as file:
        data = json.load(file)
        return data
    return data    

def chuyen_doi_ngay(ngay_goc):
    """
    Chuyển đổi ngày từ dạng 'DD/MM/YYYY HH:MM' thành 'DDMMYYYY'.
    
    :param ngay_goc: Chuỗi ngày gốc hoặc kiểu datetime
    :return: Chuỗi ngày dạng 'DDMMYYYY'
    """
    try:
        ngay_dt = pd.to_datetime(ngay_goc, errors='coerce')  # Chuyển thành datetime
        if pd.isna(ngay_dt):
            return None  # Trả về None nếu lỗi
        return ngay_dt.strftime("%d%m%Y")  # Trả về chuỗi định dạng DDMMYYYY
    except Exception:
        return None
    
def kiem_tra_record(record, mabhyt, mabn, ngayravien):
    ngay_record = chuyen_doi_ngay(record.get("ngayra"))
    ngay_input = chuyen_doi_ngay(ngayravien)

    # print(f"🔍 Kiểm tra: {record.get('mathe')} == {mabhyt}")
    # print(f"🔍 Kiểm tra: {record.get('mabn')} ({type(record.get('mabn'))}) == {mabn} ({type(mabn)})")
    # print(f"🔍 Kiểm tra ngày: {ngay_record} == {ngay_input}")

    return (record.get("mathe").strip() == mabhyt.strip() and
            str(record.get("mabn")).strip() == str(mabn).strip() and
            ngay_record == ngay_input)

def kiemtraketqua(file_path, mabn,mabhyt,ngayravien):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
            
            if not isinstance(data, list):
                raise ValueError("JSON không chứa danh sách hồ sơ hợp lệ.")
            
            for record in data:
                # if record.get("stt") == 2655:
                #     print("ở dây:"+str(chuyen_doi_ngay(record.get("ngayra")))+"||"+str(chuyen_doi_ngay(ngayravien)))
                if kiem_tra_record(record,mabhyt,mabn,ngayravien):
                    return record
            
            return None  # Không tìm thấy
    except (FileNotFoundError, json.JSONDecodeError, ValueError) as e:
        print(f"Lỗi: {e}")
        return None
#setup app


def setupapp():
    global datacount
    # Tạo cửa sổ chính
    app = ctk.CTk()
    app.title("Ứng dụng đọc kết quả xml")
    center_window(app)
    try:
        tree = setupTreeviewTable(app)
     
        
        label = ctk.CTkLabel(app,100,20,bg_color="red",text=str(datacount),text_color="white")
        label.pack(pady=5)
    except Exception as e:
        print("lỗi"+str(e))


    # Hàm xử lý khi nhấn nút
    def on_button_click():
        messagebox.showinfo("Thông báo", "Bạn đã nhấn nút!")

    def on_button_2_click():
        messagebox.showinfo("Thông báo", "Bạn đã nhấn nút thứ hai!")

    selected_date = tk.StringVar()
    cal = DateEntry(app, width=20, background='darkblue', foreground='white', borderwidth=2)
    cal.pack(side="left", padx=5, pady=5)    

    # Tạo nút bấm
    button1 = ctk.CTkButton(app, text="Lấy danh sách", command=lambda: open_export(app,1,cal))
    button1.pack(side="left", padx=5, pady=5)

    button2 = ctk.CTkButton(app, text="Xuất xml", command=lambda: open_export(app,2,cal))
    button2.pack(side="left", padx=5, pady=5)

    
    button3 = ctk.CTkButton(app, text="Chọn Folder", command=lambda: setupTreeview(tree,label))
    button3.pack(side="left", padx=5, pady=5)

    button4 = ctk.CTkButton(app, text="Kiểm tra kết quả và xuất file kết quả", command=lambda: converttojsonexportkqxml(app))
    button4.pack(side="left", padx=5, pady=5)

        # Lấy ngày hiện tại
  

    # btn = tk.Button(app, text="Lấy ngày", command=lambda: get_date(selected_date,cal))
    # btn.pack(pady=5)



    
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

def get_urlFolder():
    folder_path = filedialog.askdirectory(title="Chọn thư mục")
    print("Đường dẫn thư mục đã chọn:", folder_path)
    return folder_path


def open_export(app,type,cal):
    global folderurl
    datechoose = get_date(cal)
    print(str(datechoose))
    if folderurl != '':
        if type == 1:
            import mainExportXML as mainex
            mainex.settupAppBeginStart(app,type,datechoose,folderurl)
        else: 
            import mainExportXML as mainex
            mainex.settupAppBeginStart(app,type,datechoose,folderurl)
    else:
        print("Vui lòng chọn folder")

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


