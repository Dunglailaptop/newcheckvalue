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
import os
import json
from datetime import datetime

folderurl = ""

def chuyen_doi_ngay(ngay_goc):
    """Chuyển đổi ngày từ 'DD/MM/YYYY HH:MM' thành 'DDMMYYYY'."""
    try:
        return datetime.strptime(ngay_goc, "%d/%m/%Y %H:%M").strftime("%d%m%Y")
    except ValueError:
        return None

def kiem_tra_record(record, mabhyt, mabn, ngayravien):
    """So sánh dữ liệu nhanh hơn bằng cách chuyển đổi trước."""
    return (
        record.get("mathe") == mabhyt and
        str(record.get("mabn")) == str(mabn) and
        chuyen_doi_ngay(record.get("ngayra")) == ngayravien
    )

def kiemtraketqua(file_path, mabn, mabhyt, ngayravien):
    """Tìm kiếm nhanh hơn bằng cách dùng dictionary."""
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            data = json.load(file)
            if not isinstance(data, list):
                raise ValueError("JSON không chứa danh sách hồ sơ hợp lệ.")
            
            ngayravien = chuyen_doi_ngay(ngayravien)  # Chuyển đổi trước để tăng tốc

            # Tạo dict để tìm kiếm nhanh hơn
            records_dict = {
                (record.get("mathe"), str(record.get("mabn")), chuyen_doi_ngay(record.get("ngayra"))): record
                for record in data
            }
            
            return records_dict.get((mabhyt, str(mabn), ngayravien))

    except (FileNotFoundError, json.JSONDecodeError, ValueError) as e:
        print(f"Lỗi: {e}")
        return None


def loadfind_json():
    global folderurl
    data = []
    file_path = os.path.join(folderurl, "dataJson.json")
    with open(file_path, 'r', encoding='utf-8') as file:
        data = json.load(file)
        return data
    return data    

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

def converttojsonexportkqxml(app):
    """Hàm chính chạy tối ưu hơn."""
    global folderurl
    file_path = os.path.join(folderurl, "dataJson.json")

    if not folderurl or not file_path:
        print("❌ LỖI: Vui lòng chọn file JSON đã xuất!")
        return

    print(f"✅ Bắt đầu xử lý với folder: {folderurl}")

    try:
        data = loadfind_json()
        if not data:
            print("❌ Không có dữ liệu trong dataJson.json")
            return
        
        print(f"🔍 Số lượng bản ghi: {len(data)}")

        path = chon_file_excel(app)
        if not path:
            print("❌ Lỗi: Không chọn được file Excel")
            return
        
        print(f"📂 File Excel được chọn: {path}")

        excel_to_json(path, "dataXmlResult.json")

        # Đọc dữ liệu một lần duy nhất
        with open("dataXmlResult.json", "r", encoding="utf-8") as f:
            json_data = json.load(f)

        # Tạo dictionary để tra cứu nhanh
        records_dict = {
            (record.get("mathe"), str(record.get("mabn")), chuyen_doi_ngay(record.get("ngayra"))): record
            for record in json_data
        }

        # Duyệt và cập nhật kết quả nhanh hơn
        for item in data:
            ngay_ra = chuyen_doi_ngay(item["ngay_ra"])  # Chuyển đổi trước
            key = (item["insurance_code"], str(item["patient_code"]), ngay_ra)
            record = records_dict.get(key)

            if record:
                print(f"✅ Có kết quả trả về cho STT {item['stt']}")
                item["trangthai"] = 1
            else:
                print(f"❌ Không có kết quả cho STT {item['stt']}")
                item["trangthai"] = 0

        # Ghi lại toàn bộ JSON chỉ 1 lần
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

        # Xuất file Excel
        columns_can_xuat = ["stt", "insurance_code", "patient_code", "ins_transaction_code", "ngay_ra", "trangthai"]
        output_excel_path = os.path.join(folderurl, "dataKQ.xlsx")
        json_to_excel(file_path, output_excel_path, columns_can_xuat)

        print(f"✅ File Excel đã xuất: {output_excel_path}")

    except Exception as e:
        print(f"🚨 Lỗi trong quá trình xử lý: {e}")
