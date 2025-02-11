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
from tkinter import messagebox
from psycopg2 import sql
from PIL import ImageTk, Image
from tkinter import Menu
from typing import List, Dict, Any
import csv

app = sour.Any
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
        run_api_and_save(headers) 
        messagebox.showinfo(title="Thành công!",message="Hoàn thành cào dữ liệu bệnh nhân! Kết nối PostgreSQL đã đóng!.....")
    finally:
        loading = False
        loading_thread.join()
    sour._destroySelenium_()
    
    # app.deiconify()

def load_config():
    if os.path.exists(sour.CONFIG_PATH):
        with open(sour.CONFIG_PATH,'r') as f:
            return json.load(f)
    return {"page_value": 1, "record_value": 0, "total_page": 1}

def save_config(config):
    with open(sour.CONFIG_PATH, 'w') as f:
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

    
# Thiết lập logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def add_detail_prescription(datalist: List[Dict[str, Any]]):
    conn_params = sour.ConnectStr  # Giả sử đây là một đối tượng chứa thông tin kết nối
    conn = None
    cur = None

    try:
        conn = psycopg2.connect(**conn_params)
        cur = conn.cursor()

        if len(datalist) > 0:
            for index, d in enumerate(datalist, start=1):
                try:
                    # Chuẩn bị dữ liệu cho drug_material_of_prescription
                    drug_material_data = (
                        d["drug_material_id"], d["ma_hoat_chat_ax"], d["hoat_chat_ax"],
                        d["ma_duong_dung_ax"], d["duong_dung_ax"], d["bhyt_so_dk_gpnk"],
                        d["bhyt_ham_luong"], d["code"], d["code_insurance"],
                        d["enum_insurance_type"], d["proprietary_name"],
                        d["insurance_name"], d["ten_thuongmai"], d["drug_original_name_id"],
                        d["original_names"], d["default_usage_id"], d["enum_usage"],
                        d["enum_unit_import_sell"], d["unit_usage_id"], d["allow_auto_cal"],
                        d["num_of_day"], d["max_usage"], d["num_of_time"], d["unit_volume_id"],
                        d["volume_value"], d["disable"], d["enum_item_type"], d["made_in"],
                        d["country_name"], d["include_children"], d["insurance_support"],
                        d["cancer_drug"], d["price"], d["ham_luong"], d["dong_goi"],
                        d["tac_dung"], d["chi_dinh"], d["chong_chi_dinh"], d["tac_dung_phu"],
                        d["lieu_luong"], d["poison_type_id"], d["pharmacology_id"],
                        d["manufacturer_id"], d["ten_hang_sx"], d["renumeration_price"],
                        d["so_dk_gpnk"], d["price_bv"], d["price_qd"],
                        d["latest_import_price_vat"], d["latest_import_price"],
                        d["is_bhyt_in_surgery"], d["stt_tt"], d["bv_ap_thau"], d["stt_dmt"],
                        d["bhyt_effect_date"], d["bhyt_exp_effect_date"],
                        d["ngay_hieu_luc_hop_dong"], d["goi_thau_bhyt"], d["phan_nhom_bhyt"],
                        d["insurance_drug_material_id"], d["bhyt_loai_thuoc"],
                        d["bhyt_loai_thau"], d["bhyt_nha_thau"], d["bhyt_nha_thau_bak"],
                        d["bhyt_quyet_dinh"], d["bhyt_so_luong"], d["bhxh_id"],
                        d["creator_id"], d["created_at"], d["modifier_id"], d["updated_at"],
                        d["bhxh_pay_percent"], d["service_group_cost_code"],
                        d["ma_thuoc_dqg"], d["khu_dieu_tri"], d["note"], d["dang_bao_che"],
                        d["locked"], d["code_atc"], d["co_han_dung"], d["t_trantt"],
                        d["bk_enum_item_type"], d["is_control"], d["nhom_thuoc"],
                        d["nhom_duoc_ly"], d["phan_nhom_thuoc_id"], d["dst_price"],
                        d["im_price"], d["is_special_dept"], d["thang_tuoi_chi_dinh"],
                        d["max_one_times"], d["max_one_times_by_weight"],
                        d["min_one_times_by_weight"], d["max_one_day"],
                        d["max_one_day_by_weight"], d["min_one_day_by_weight"],
                        d["thuoc_ra_le"], d["gia_temp"], d["is_inventory"],
                        d["loai_thuan_hop"], d["khong_thanh_toan_rieng"],
                        d["is_used_event"], d["bhyt_nha_thau_id"], d["bhyt_nha_thau_code"],
                        d["so_luong_cho_nhap"], d["so_luong_da_nhan"],
                        d["is_used_event_idm"], d["dose_quantity"], d["dose_unit"],
                        d["thoi_gian_bao_quan"], d["ten_theo_thau"],
                        d["prescription_item_id"], d["prescription_id"], d["medicine_id"],
                        d["usage_title"], d["usage_num"], d["dosage"], d["time"],
                        d["quantity_num"], d["confirm_sell_num"], d["quantity_title"],
                        d["dosage_title"], d["morning"], d["noon"], d["afternoon"],
                        d["evening"], d["paid"], d["is_bhyt"], d["bhyt_pay_percent"],
                        d["is_bhbl"], d["bhbl_percent"], d["insurance_company_id"],
                        d["bhbl_amount"], d["bhbl_must_buy_full"], d["status"],
                        d["num_per_time"], d["is_deleted"], d["solan_ngay"],
                        d["is_max_one_times"], d["is_max_one_times_by_weight"],
                        d["is_max_one_day"], d["is_max_one_day_by_weight"],
                        d["is_min_one_day_by_weight"], d["is_min_one_times_by_weight"],
                        d["is_duply_original_name"], d["warning_note_doctor"],
                        d["bhyt_store"], d["canh_bao_thang_tuoi_chi_dinh"],
                        d["loai_ke_toa"], d["buoi_uong"], d["da_cap"], d["quantity_use"],
                        d["quantity_remain"], d["ngay_dung_thuoc"], d["order_by"]
                    )

                    # logging.info(f"Đang xử lý bản ghi thứ {index}/{len(datalist)}")
                    # logging.debug(f"Số lượng trường dữ liệu: {len(drug_material_data)}")

                    # Sử dụng sql.SQL để tránh SQL injection
                    insert_query = sql.SQL("""
                        INSERT INTO prescription_details (
                            drug_material_id, ma_hoat_chat_ax, hoat_chat_ax, 
                            ma_duong_dung_ax, duong_dung_ax, bhyt_so_dk_gpnk, bhyt_ham_luong, 
                            code, code_insurance, enum_insurance_type, proprietary_name, 
                            insurance_name, ten_thuongmai, drug_original_name_id, original_names, 
                            default_usage_id, enum_usage, enum_unit_import_sell, unit_usage_id, 
                            allow_auto_cal, num_of_day, max_usage, num_of_time, unit_volume_id, 
                            volume_value, disable, enum_item_type, made_in, country_name, 
                            include_children, insurance_support, cancer_drug, price, ham_luong, 
                            dong_goi, tac_dung, chi_dinh, chong_chi_dinh, tac_dung_phu, lieu_luong, 
                            poison_type_id, pharmacology_id, manufacturer_id, ten_hang_sx, 
                            renumeration_price, so_dk_gpnk, price_bv, price_qd, latest_import_price_vat, 
                            latest_import_price, is_bhyt_in_surgery, stt_tt, bv_ap_thau, stt_dmt, 
                            bhyt_effect_date, bhyt_exp_effect_date, ngay_hieu_luc_hop_dong, 
                            goi_thau_bhyt, phan_nhom_bhyt, insurance_drug_material_id, bhyt_loai_thuoc, 
                            bhyt_loai_thau, bhyt_nha_thau, bhyt_nha_thau_bak, bhyt_quyet_dinh, 
                            bhyt_so_luong, bhxh_id, creator_id, created_at, modifier_id, updated_at, 
                            bhxh_pay_percent, service_group_cost_code, ma_thuoc_dqg, khu_dieu_tri, 
                            note, dang_bao_che, locked, code_atc, co_han_dung, t_trantt, 
                            bk_enum_item_type, is_control, nhom_thuoc, nhom_duoc_ly, phan_nhom_thuoc_id, 
                            dst_price, im_price, is_special_dept, thang_tuoi_chi_dinh, max_one_times, 
                            max_one_times_by_weight, min_one_times_by_weight, max_one_day, 
                            max_one_day_by_weight, min_one_day_by_weight, thuoc_ra_le, gia_temp, 
                            is_inventory, loai_thuan_hop, khong_thanh_toan_rieng, is_used_event, 
                            bhyt_nha_thau_id, bhyt_nha_thau_code, so_luong_cho_nhap, so_luong_da_nhan, 
                            is_used_event_idm, dose_quantity, dose_unit, thoi_gian_bao_quan, ten_theo_thau,
                            prescription_item_id, prescription_id, medicine_id, usage_title, 
                            usage_num, dosage, time, quantity_num, confirm_sell_num, 
                            quantity_title, dosage_title, morning, noon, afternoon, evening, 
                            paid, is_bhyt, bhyt_pay_percent, is_bhbl, bhbl_percent, 
                            insurance_company_id, bhbl_amount, bhbl_must_buy_full, status, 
                            num_per_time, is_deleted, solan_ngay, is_max_one_times, 
                            is_max_one_times_by_weight, is_max_one_day, is_max_one_day_by_weight, 
                            is_min_one_day_by_weight, is_min_one_times_by_weight, 
                            is_duply_original_name, warning_note_doctor, bhyt_store, 
                            canh_bao_thang_tuoi_chi_dinh, loai_ke_toa, buoi_uong, da_cap, 
                            quantity_use, quantity_remain, ngay_dung_thuoc, order_by
                        ) 
                        VALUES ({})
                    """).format(
                        sql.SQL(', ').join([sql.Placeholder()] * len(drug_material_data))
                    )

                    cur.execute(insert_query, drug_material_data)
                    conn.commit()
                    logging.info(f"Bản ghi thứ {index} đã được chèn thành công")
                    log_terminal(f"Bản ghi thứ {index} đã được chèn thành công")

                except Exception as e:
                    conn.rollback()
                    logging.error(f"Lỗi khi chèn bản ghi thứ {index}: {e}")
                    # Có thể thêm xử lý lỗi cụ thể ở đây nếu cần

    except psycopg2.Error as e:
        logging.error(f"Lỗi kết nối database: {e}")
    finally:
        if cur:
            cur.close()
        if conn:
            conn.close()
        logging.info("Kết nối database đã đóng")

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
CSV_FILE = "data.csv"
PROGRESS_FILE = "progress.json"
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
    progress = load_progress()
    page_value = progress["page_value"]
    record_value = progress["record_value"]
    total_page = progress["total_page"]
    total_record = progress["total_record"]

    payload = {
        "date_from": "2025-02-10T13:43:32.800Z",
        "date_to": "2025-02-10T13:43:32.801Z",
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

  

def hamchaylistxml(header):
    config = load_config()
    page_value =config["page_value"]
    record_value = config["record_value"]
    checkstatusfail = False
    total_records = 424
    while(True):
        try:
            payload = {
                "date_from": "2025-02-10T13:43:32.800Z",
                "date_to": "2025-02-10T13:43:32.801Z",
                "item_status": "2",
                "item_type": "0",
                "page": page_value,
                "patient_code": "",
                "row_per_page": 20,
                "tt_hs_ntru": 0
            }
            responsexml = requests.get(sour.urlGetinvoiceoutdetail,headers= header,json=payload)
            dataTT = responsexml.json()
            Errors = dataTT["error"]
            checkApi = Errors["code"]
            if responsexml.status_code == 200 and int(checkApi) == 200:
                    try:
                        data = dataTT['data']
                        numbergetdata = len(data)
                        if len(data) > 0:
                            log_terminal("......BẮT ĐẦU GHI DATA VÔ NHA......")
                            #hàm ghi csv ở đây
                            #hamm them du lieu vao database postgresql
                    except Exception as e:
                        print(f"loi khi them du lieu vao database......")
            else:
                    print(f"loi khi lay du lieu chi tiet toa thuoc" + str(e))
                    checkstatusfail = True
        except Exception as e:
            print("lỗi trong quá trình ghi dữ liệu:"+e)
       

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




    
# Function to open a terminal window embedded in the tab
def open_terminal_window(new_tab):
       # Clear any existing widgets in the tab
    for widget in new_tab.winfo_children():
        widget.destroy()
    terminal_window = tk.Frame(new_tab)
    terminal_window.pack(expand=True, fill="both")

    terminal_text = tk.Text(terminal_window, bg="black", fg="green", insertbackground="green")
    terminal_text.pack(expand=True, fill="both")
  
    return terminal_window, terminal_text

# Function to start the script in a separate thread
def start_script_thread(new_tab):
    global script_thread
    terminal_window, terminal_text = open_terminal_window(new_tab)
    script_thread = threading.Thread(target=run_script, args=(terminal_text,))
    script_thread.start()

# Function to set up the main application in the tab
def settupAppBeginStart(new_tab):
   start_script_thread(new_tab) 


