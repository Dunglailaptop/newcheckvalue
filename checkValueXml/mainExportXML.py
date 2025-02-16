import psycopg2
import sourceString as sour
import tkinter as tk
import time
import threading
import os
import json
import logging
import requests
import customtkinter
import base64
from tkinter import messagebox
from psycopg2 import sql
from PIL import ImageTk, Image
from tkinter import Menu
from typing import List, Dict, Any
import csv
from datetime import datetime
import xml.etree.ElementTree as ET
import json
import base64
from handle import button_click_subject

folderurl = ""
CSV_FILE = ""
PROGRESS_FILE = ""
data_jsonurl = ""

tree = Any
datetimechoose =  ""
app = sour.Any
type = 0
terminal_window = Any
terminal_text = Any
script_thread = Any
def log_terminal(message):
        global terminal_text
        terminal_text.insert(tk.END,message + "\n")
        terminal_text.see(tk.END)  # Scroll to the end
        terminal_text.update_idletasks()
#hàm chạy chính các sự kiện main
def run_script(terminaltext):
    global terminal_window, terminal_text
    terminal_text = terminaltext
    log_terminal("...........chay con server 70...............")
    log_terminal("...........khởi động chương trình...........")
    log_terminal("...........Khởi động chomre.................")
    bTry = False
    errorChrome = 1
    while bTry == False:
          bTry = sour._initSelenium_()
          if bTry == False:
             if errorChrome >= 10:
                 log_terminal(".......khởi động chrome thất bại quá nhiều lần! tắt chương trình.......")
                 terminal_window.destroy()
                #  app.destroy()
             else:
                 errorChrome += 1
                 log_terminal("....khởi động chomre thất bại thử lại lần nữa.......")
    time.sleep(3)
    log_terminal("........Duyệt website thành công.........")
    sour._login_("quyen.ngoq", "74777477")
     # ấn nút login
    time.sleep(0.5)
    log_terminal(".........................Sao chép userkey thành công.....................................")
    headers = {
                    "Appkey": sour.Appkey,
                    "Userkey": sour.secretKey,
                    "Authorization": sour.secretKey,
                    "Content-Type": sour.contentType
                }
    log_terminal(".........................Đang tiến hành thu thập! Vui lòng chờ...........................")
    loading = True
    def loading_animation():
        chars = "/—\\|"
        i = 0
        while loading:
            log_terminal('\r' + 'Vui lòng đợi quá trình tải dữ liệu đang được diễn ra... ' + chars[i % len(chars)])
            time.sleep(0.1)
            i += 1
      # Thread cho animation
    loading_thread = threading.Thread(target=loading_animation)
    loading_thread.daemon = True  # Đảm bảo thread sẽ kết thúc khi chương trình chính kết thúc
    # Bắt đầu animation
    loading_thread.start()
    # Biến global để kiểm soát animation
    try:
        if type == 1:
            run_api_and_save(headers) 
            messagebox.showinfo(title="Thành công!",message="Hoàn thành cào dữ liệu bệnh nhân! Kết nối PostgreSQL đã đóng!.....")
        else:
            hamchaylistxml(headers,data_jsonurl)
            messagebox.showinfo(title="Thành công!",message="Hoàn thành cào dữ liệu bệnh nhân! Kết nối PostgreSQL đã đóng!.....")
    finally:
        loading = False
        loading_thread.join()
    sour._destroySelenium_()
    on_closing()
    # app.deiconify()

def load_config():
    global PROGRESS_FILE
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE,'r') as f:
            return json.load(f)
    return {"page_value": 1, "record_value": 0, "total_page": 1}

def save_config(config):
    global PROGRESS_FILE
    with open(PROGRESS_FILE, 'w') as f:
        json.dump(config, f)

def update_file_json(l4_value, l6_value, total_page):
    config = load_config()
    config['page_value'] = l4_value
    config['record_value'] = l6_value
    config["total_page"] = total_page
    save_config(config)
   
def replace_nulls_with_string(data):
    if isinstance(data, dict):
        # Xử lý dictionary
        for key, value in data.items():
            if value is None:
                data[key] = "None"
    elif isinstance(data, list):
        # Xử lý list chứa dictionary hoặc tuple
        for i in range(len(data)):
            if isinstance(data[i], dict):
                # Nếu phần tử là dictionary, xử lý như dictionary
                data[i] = replace_nulls_with_string(data[i])
            elif isinstance(data[i], tuple):
                # Nếu phần tử là tuple, xử lý từng giá trị trong tuple
                data[i] = tuple("None" if v is None else v for v in data[i])
    elif isinstance(data, tuple):
        # Xử lý trực tiếp tuple
        return tuple("None" if v is None else v for v in data)
    return data

    
def csv_to_json(csv_file_path, json_file_path=None):
    """
    Chuyển đổi tệp CSV thành JSON, thêm số thứ tự cho từng phần tử.
    
    Args:
        csv_file_path (str): Đường dẫn đến tệp CSV.
        json_file_path (str, optional): Đường dẫn lưu tệp JSON. Nếu None, trả về dữ liệu JSON.
    
    Returns:
        dict or None: Dữ liệu JSON nếu không lưu vào tệp.
    """
    data = []
    
    with open(csv_file_path, mode='r', encoding='utf-8') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        for index, row in enumerate(csv_reader, start=1):
            row["stt"] = index
            row["trangthai"]= 0
            row["xml1"] = ""
            row["xml2"] = ""
            row["xml3"] = ""
            row["xml4"] = ""
            row["xml5"] = ""
            row["xml7"] = ""
            data.append(row)
    
    if json_file_path:
        with open(json_file_path, mode='w', encoding='utf-8') as json_file:
            json.dump(data, json_file, indent=4, ensure_ascii=False)
    else:
        return data

#hàm check kiem tra sự tồn tại
def check_Count_value(prescriptionid,numberdataget):
    try:
        prescription_id = int(prescriptionid)
        con_params = sour.ConnectStr
        conn = psycopg2.connect(**con_params)
        cur = conn.cursor()
        queyStr = f"SELECT CASE WHEN COUNT(*) >= {numberdataget} THEN 1 ELSE 0 END as is_enough FROM prescription_details WHERE prescription_id = {prescription_id};"
        # queyStr = f"select count(*) from prescription_details where prescription_id = {prescription_id};"
        cur.execute(queyStr)
        countPrescription = cur.fetchone()[0]
        return int(countPrescription)
    except Exception as e:
        print("lỗi sảy ra trong quá trình lấy số lượng của prescriptiondetail" + str(e))
    finally:
        if conn:
            cur.close()
            conn.close()    


# Đường dẫn file CSV và file lưu tiến trình
# Hàm lưu tiến trình

def save_progress(page, record, total_page, total_record):
    with open(PROGRESS_FILE, "w") as f:
        json.dump({
            "page_value": page,
            "record_value": record,
            "total_page": total_page,
            "total_record": total_record
        }, f)

# Hàm tải tiến trình trước đó
def load_progress():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, "r") as f:
            return json.load(f)
    return {"page_value": 1, "record_value": 0, "total_page": 0, "total_record": 0}

# Hàm lấy tổng số trang và phần tử (CHỈ GỌI 1 LẦN)
def get_total_pages(header, payload):
    response = requests.get(sour.urlApiGetListXML, headers=header, json=payload)
    if response.status_code == 200:
        data = response.json()
        total_page = int(data["data"]["paging"]["total_page"])
        total_record = int(data["data"]["paging"]["total_record"])
        return total_page, total_record
    return 0, 0

# Hàm ghi dữ liệu vào CSV
def write_to_csv(data):
    file_exists = os.path.isfile(CSV_FILE)
    with open(CSV_FILE, mode="a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=data[0].keys())
        if not file_exists:
            writer.writeheader()  # Ghi tiêu đề nếu file chưa tồn tại
        writer.writerows(data)

# Hàm gọi API và lưu dữ liệu
def run_api_and_save(header):
    global data_jsonurl
    progress = load_progress()
    page_value = progress["page_value"]
    record_value = progress["record_value"]
    total_page = progress["total_page"]
    total_record = progress["total_record"]

    payload = {
        "date_from": datetimechoose,
        "date_to": datetimechoose,
        "item_status": "2",
        "item_type": "0",
        "page": 1,  # Chỉ gọi 1 lần để lấy tổng số trang
        "patient_code": "",
        "row_per_page": 20,
        "tt_hs_ntru": 0
    }

    # Chỉ gọi API 1 lần để lấy tổng số trang & phần tử nếu chưa có
    if total_record == 0:
        total_page, total_record = get_total_pages(header, payload)
        print(f"Tổng số phần tử: {total_record}, Tổng số trang: {total_page}")
        save_progress(page_value, record_value, total_page, total_record)  # Lưu lại

    # Chạy API cho từng trang
    while page_value <= total_page:
        payload["page"] = page_value  # Cập nhật page trong payload

        try:
            response = requests.get(sour.urlApiGetListXML, headers=header, json=payload)
            dataTT = response.json()

            if response.status_code == 200 and dataTT["error"]["code"] == 200:
                data = dataTT["data"]["res"]
                if data:
                    write_to_csv(data)
                    record_value += len(data)
                    save_progress(page_value, record_value, total_page, total_record)  # Cập nhật tiến trình

                print(f"Page {page_value} - Đã ghi {record_value}/{total_record} phần tử")
                page_value += 1
            else:
                print(f"Lỗi API: {dataTT['error']['message']}")
                break

        except Exception as e:
            print(f"Lỗi trong quá trình lấy dữ liệu: {e}")
            break  # Tránh vòng lặp vô hạn nếu có lỗi lớn

    print("Hoàn thành việc lấy dữ liệu!")
    csv_to_json(CSV_FILE,data_jsonurl)
    print("đang chuyển đổi từ csv sang json")
    button_click_subject.on_next(tree)
   


def convert_date_format(date_str):
    """
    Chuyển đổi ngày từ định dạng 'dd/MM/yyyy HH:mm' sang 'yyyy/MM/dd'.
    
    Args:
        date_str (str): Ngày cần chuyển đổi.
    
    Returns:
        str: Ngày đã chuyển đổi.
    """
    date_obj = datetime.strptime(date_str, "%d/%m/%Y %H:%M")
    return date_obj.strftime("%Y/%m/%d")

def convert_date_compact(date_str):
    """
    Chuyển đổi ngày từ định dạng 'dd/MM/yyyy HH:mm' sang 'yyyyMMdd'.
    
    Args:
        date_str (str): Ngày cần chuyển đổi.
    
    Returns:
        str: Ngày đã chuyển đổi.
    """
    date_obj = datetime.strptime(date_str, "%d/%m/%Y %H:%M")
    return date_obj.strftime("%Y%m%d")

def load_config_STT():
    global folderurl
    URL = folderurl + "stt.json"
    if os.path.exists(URL):
        with open(URL,'r') as f:
            return json.load(f)
    return {"STT":1}

def save_config_STT(config):
    global folderurl
    URL = folderurl + "stt.json"
    with open(URL, 'w') as f:
        json.dump(config, f)

def update_json(stt_value):
    config = load_config_STT()
    config["STT"] = stt_value
    save_config_STT(config)

def encode_base64(input_str):
    """
    Mã hóa một chuỗi thành base64.
    
    Args:
        input_str (str): Chuỗi cần mã hóa.
    
    Returns:
        str: Chuỗi đã mã hóa base64.
    """
    if not input_str:  # Kiểm tra nếu input_str là None hoặc ""
        return ""
    encoded_bytes = base64.b64encode(input_str.encode("utf-8"))
    return encoded_bytes.decode("utf-8")

def update_json_dataxml(file_path, stt, xml1, xml2, xml3, xml4, xml5, xml7):
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
            item["xml1"] = encode_base64(xml1)
            item["xml2"] = encode_base64(xml2)
            item["xml3"] = encode_base64(xml3)
            item["xml4"] = encode_base64(xml4)
            item["xml5"] = encode_base64(xml5)
            item["xml7"] = encode_base64(xml7)
            break
    
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)


def getAll_JSON(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
            
            if not isinstance(data, list):
                raise ValueError("JSON không chứa danh sách hồ sơ hợp lệ.")
            
            if len(data) > 0:
              return data
            
            return None  # Không tìm thấy
    except (FileNotFoundError, json.JSONDecodeError, ValueError) as e:
        print(f"Lỗi: {e}")
        return None

def find_record_by_stt(file_path, stt):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
            
            if not isinstance(data, list):
                raise ValueError("JSON không chứa danh sách hồ sơ hợp lệ.")
            
            for record in data:
                if record.get("stt") == stt:
                    return record
            
            return None  # Không tìm thấy
    except (FileNotFoundError, json.JSONDecodeError, ValueError) as e:
        print(f"Lỗi: {e}")
        return None



def hamchaylistxml(header,json_file_path):
    config = load_config_STT()
    stt_value = config["STT"]
    demkiemtra = 0
    checkstatus = False
    checkupdate = False
    while (True):
        try:
                data = find_record_by_stt(json_file_path,stt_value)
                if len(data) > 0:
                    dateFirst = convert_date_format(data["ngay_ra"]) +"/"
                    dataEnd = convert_date_compact(data["ngay_ra"])

                    pathxml = str(dateFirst)+data["patient_code"]+"-"+data["pay_receipt_id"]+"-"+str(dataEnd)+"/"
                    payload = {
                            "enum_payment_type": data["enum_payment_type"],
                            "item_status": "2",
                            "outPatient": True,
                            "pathXML": pathxml,
                            "pay_receipt_id": data["pay_receipt_id"]
                    }
                    print(payload)
                    responsexml = requests.get(sour.urlGetDetailXML,headers= header,json=payload)
                    dataTT = responsexml.json()
                    Errors = dataTT["error"]
                    checkApi = Errors["code"]
                   
                    if responsexml.status_code == 200 and int(checkApi) == 200:
                             try:
                                 datRes = dataTT['data']
                                 if len(datRes) > 0:
                                     log_terminal("......BẮT ĐẦU GHI DATA VÔ NHA......")
                                     xml_keys = ["Xml2", "Xml3", "Xml4", "Xml5", "Xml7"]
                                     valid_extra_xml_exists = any(datRes[key] not in [None, ""] for key in xml_keys)
                                     if datRes["Xml1"] not in [None, ""] and valid_extra_xml_exists:
                                        update_json_dataxml(json_file_path,stt_value,datRes["Xml1"],datRes["Xml2"],datRes["Xml3"],datRes["Xml4"],datRes["Xml5"],datRes["Xml7"])
                                        demkiemtra = 0
                                     else:
                                       checkstatus = True
                                       demkiemtra += 1
                                     #hamm them du lieu vao database postgresql
                             except Exception as e:
                                 print(f"loi khi them du lieu vao database......")
                                 checkstatus = True
                    else:
                             print(f"loi khi lay du lieu chi tiet toa thuoc" + str(e))
                             checkstatus = True
                    if not checkstatus:
                        stt_value += 1
                        update_json(stt_value)
                    if demkiemtra == 5: 
                        stt_value += 1
                        update_json(stt_value)
                else:
                    checkupdate = False
                    break
        except Exception as e:
            print("lỗi trong quá trình ghi dữ liệu:"+str(e))
            break
    if not checkupdate:
        data = getAll_JSON(json_file_path)
        create_fileXML(data)
   

def create_xml_file(data, filename):
    giamdinhhs = ET.Element("GIAMDINHHS")
    
    thongtindonvi = ET.SubElement(giamdinhhs, "THONGTINDONVI")
    macskcb = ET.SubElement(thongtindonvi, "MACSKCB")
    macskcb.text = "79408"
    
    thongtinhoso = ET.SubElement(giamdinhhs, "THONGTINHOSO")
    ngaylap = ET.SubElement(thongtinhoso, "NGAYLAP")
    soluonghoso = ET.SubElement(thongtinhoso, "SOLUONGHOSO")
    soluonghoso.text = "1"
    
    danhsachhoso = ET.SubElement(thongtinhoso, "DANHSACHHOSO")
    hoso = ET.SubElement(danhsachhoso, "HOSO")
    
    # Lọc các key xml1 đến xml7 có giá trị
    for i in range(1, 8):
        key = f"xml{i}"
        if key in data and data[key]:
            filehoso = ET.SubElement(hoso, "FILEHOSO")
            loaihoso = ET.SubElement(filehoso, "LOAIHOSO")
            loaihoso.text = key.upper()
            
            noidungfile = ET.SubElement(filehoso, "NOIDUNGFILE")
            noidungfile.text = data[key]
    
    tree = ET.ElementTree(giamdinhhs)
    tree.write(filename, encoding="utf-8", xml_declaration=True)

def create_fileXML(data_list):
   global folderurl
   for data in data_list:
        filename = os.path.join(folderurl,f"79408_{data['ma_lk']}.xml")
        create_xml_file(data, filename)
        print(f"File {filename} đã được tạo.")

     
       

#hàm lấy dự liệu lên từ database của bảng prescription
def get_list_data_prescription(header):
    config = load_config()
    page_value =config["page_value"]
    record_value = config["record_value"]
    checkstatusfail = False
    while(True):
        try:
            conn_params = sour.ConnectStr
            conn = psycopg2.connect(**conn_params)
            cur = conn.cursor()
            queryStr = f"SELECT * FROM prescription order by stt asc OFFSET {record_value} LIMIT 20;"
            cur.execute(queryStr)
            listdata = cur.fetchall()
            if len(listdata) > 0 :
                for item in listdata:
                    prescription_id = item[54]
                    payload = {
                         "prescription_id": prescription_id
                    }
                    responseChiTietToaThuoc = requests.get(sour.urlGetPrescriptiondetail, headers=header,json=payload)
                    dataTT = responseChiTietToaThuoc.json()
                    Errors = dataTT["error"]
                    checkApi = Errors["code"]
                    if responseChiTietToaThuoc.status_code == 200 and int(checkApi) == 200:
                        try:
                            data = dataTT['data']
                            numbergetdata = len(data)
                            if len(data) > 0:
                                log_terminal("......BẮT ĐẦU GHI DATA VÔ NHA......")
                                add_detail_prescription(data)
                                #hamm them du lieu vao database postgresql
                        except Exception as e:
                            print(f"loi khi them du lieu vao database......")
                    else:
                        print(f"loi khi lay du lieu chi tiet toa thuoc" + str(e))
                        checkstatusfail = True
                if not checkstatusfail:
                    p = int(page_value) + 1
                    pSub = p - 1
                    rc = pSub * 20 - 1
                    update_file_json(l4_value=p, l6_value=rc)
                    page_value = p
                    record_value = rc
                    log_terminal(f"tổng page vlaue/record value:{page_value}/{record_value}")
                    checkstatusfail = False
            else:
                break
        except Exception as e:
           print("Lỗi xảy ra trong quá trình truy cập CSDL... : "+ str(e))                  
        finally:
            if conn:
                 cur.close()
                 conn.close()



def on_closing():
    global terminal_window,terminal_text
    terminal_window.destroy()
   
    # app.deiconify()
    
# Function to open a terminal window embedded in the tab
def open_terminal_window():
    global terminal_window
    terminal_window = tk.Toplevel()  # Tạo cửa sổ mới
    terminal_window.title("Terminal Output")
    terminal_window.geometry("600x400")

    terminal_text = tk.Text(terminal_window, bg="black", fg="green", insertbackground="green")
    terminal_text.pack(expand=True, fill="both")
     
    return terminal_window, terminal_text

# Function to start the script in a separate thread
def start_script_thread():
    global script_thread
    terminal_window, terminal_text = open_terminal_window()
    script_thread = threading.Thread(target=run_script, args=(terminal_text,))
    script_thread.start()

# Function to set up the main application in the tab
def settupAppBeginStart(new_tab,types,date,urlFolderSave):
   global type, datetimechoose,folderurl,CSV_FILE,PROGRESS_FILE,data_jsonurl
   type = types
   folderurl = urlFolderSave
   datetimechoose = str(date)
   CSV_FILE = urlFolderSave + "data.csv"
   PROGRESS_FILE = urlFolderSave + "progress.json"
   data_jsonurl = urlFolderSave + "dataJson.json"
   start_script_thread()


def open_new_terminal_new():
    """Mở một cửa sổ terminal mới và chạy script"""
    terminal_window = tk.Toplevel()  # Tạo cửa sổ mới
    terminal_window.title("Terminal Output")
    terminal_window.geometry("600x400")

    terminal_text = tk.Text(terminal_window, bg="black", fg="green", insertbackground="green")
    terminal_text.pack(expand=True, fill="both")

    # Chạy script trong thread để không bị treo GUI
    script_thread = threading.Thread(target=run_script, args=(terminal_text, terminal_window))
    script_thread.start()