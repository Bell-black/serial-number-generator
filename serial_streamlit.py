import streamlit as st
import hashlib
import pandas as pd
import os
import barcode
from barcode.writer import SVGWriter
from datetime import datetime
import zipfile
import streamlit.components.v1 as components
from google.oauth2 import service_account
import json
import gspread
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# --------------------------
# 기본 설정
# --------------------------
alpha_dict = {
    '1': 'A', '2': 'C', '3': 'D', '4': 'E', '5': 'F',
    '6': 'H', '7': 'J', '8': 'K', '9': 'L', '0': 'M',
    '10': 'M', '11': 'N', '12': 'P'
}

used_codes = set()
model_code_cache = {}
model_map_file = "model_map.csv"

maker_dict = {
    "닝보 타이웨이": "NB", "리앤텍": "HL", "마라타": "MT", "웨이슬라": "VS",
    "킹크린": "KE", "푸산 데코": "DC", "헝쉰전자": "HX", "화유": "HU", "중산 커리신": "ZK"
}

category_dict = {
    "무선 진공 청소기": "MC", "무선 물걸레 청소기": "AC", "가습기": "MH",
    "공기청정기": "AP", "제습기": "DH", "선풍기": "MF", "에어프라이어": "AF",
    "블렌더": "MB", "헤어 드라이기": "MS", "음식물 처리기": "FP"
}

# --------------------------
# Google Sheets 연결 설정
# --------------------------
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = "1O3aZxhweHlcjt5nIFKPu-1WERxPzl6Tjt7PUr3DraDo"
info = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
client = gspread.authorize(creds)
sheet = client.open_by_key(SPREADSHEET_ID).sheet1

# --------------------------
# 유틸 함수
# --------------------------
def num_to_alpha(num): return ''.join(alpha_dict[d] for d in str(num))
def model_to_number(name): return int(hashlib.sha256(name.upper().encode()).hexdigest(), 16)
def number_to_code(num): return chr(ord('A') + num % 676 // 26) + chr(ord('A') + num % 26)

def get_unique_code(name):
    name = name.upper()
    if name in model_code_cache:
        return model_code_cache[name]
    base = model_to_number(name)
    for offset in range(676):
        code = number_to_code(base + offset)
        if code not in used_codes:
            used_codes.add(code)
            model_code_cache[name] = code
            return code
    raise Exception("모든 코드가 소진되었습니다!")

def generate_serial(maker, category, model_code, year, month, order, seq):
    return f"{maker}{category}{model_code}{num_to_alpha(year[-1])}{alpha_dict[month]}{str(order).zfill(2)}{seq}"

def generate_barcode_svg(serial):
    CODE128 = barcode.get_barcode_class('code128')
    writer = SVGWriter()
    writer.set_options({
        "module_width": 0.6,
        "module_height": 80.0,
        "font_size": 20,
        "text_distance": 5.0,
        "quiet_zone": 2.0,
        "write_text": True
    })
    barcode_obj = CODE128(serial, writer=writer)
    filename = barcode_obj.save(f"barcode_{serial}")
    return filename

def save_model_mapping(name, code):
    try:
        if os.path.exists(model_map_file):
            df = pd.read_csv(model_map_file)
            if not ((df['모델코드'] == code) & (df['모델명'] == name)).any():
                df = pd.concat([df, pd.DataFrame([{"모델코드": code, "모델명": name}])], ignore_index=True)
        else:
            df = pd.DataFrame([{"모델코드": code, "모델명": name}])
        df.to_csv(model_map_file, index=False)
    except Exception as e:
        print(f"[모델 매핑 저장 오류] {e}")

def append_serial_to_sheet(serial_data: dict):
    try:
        row = [
            serial_data.get("시리얼넘버"),
            serial_data.get("제조사"),
            serial_data.get("제품 카테고리"),
            serial_data.get("모델명"),
            serial_data.get("제조년도"),
            serial_data.get("제조월"),
            serial_data.get("주문차수"),
            serial_data.get("생산순서")
        ]
        sheet.append_row(row)
    except Exception as e:
        st.error(f"[❌ Google Sheets 저장 실패] {e}")

def search_serial_from_sheet(serial_number: str):
    try:
        records = sheet.get_all_records()
        for row in records:
            if row.get("시리얼넘버") == serial_number:
                return row
        return None
    except Exception as e:
        st.error(f"[❌ Google Sheets 조회 실패] {e}")
        return None

def guess_full_year(d):
    try:
        year = datetime.now().year // 10 * 10 + int(d)
        return str(year - 10 if year > datetime.now().year + 1 else year)
    except: return "Unknown"

def lookup_model_name(code):
    try:
        if not os.path.exists(model_map_file): return "(매핑 없음)"
        df = pd.read_csv(model_map_file)
        return df[df['모델코드'] == code].iloc[0]['모델명'] if not df[df['모델코드'] == code].empty else "(매핑 없음)"
    except Exception as e: return f"(에러: {e})"

def decode_serial(serial):
    try:
        maker_code = serial[0:2]
        category_code = serial[2:4]
        model_code = serial[4:6]
        year_alpha = serial[6]
        month_alpha = serial[7]
        order = serial[8:10]
        sequence = serial[10:]

        rev_maker = {v: k for k, v in maker_dict.items()}
        rev_category = {v: k for k, v in category_dict.items()}
        rev_alpha = {v: k for k, v in alpha_dict.items()}

        year_digit = rev_alpha.get(year_alpha, None)
        full_year = guess_full_year(year_digit) if year_digit else "Unknown"

        month_digit = rev_alpha.get(month_alpha, None)
        if month_digit and month_digit.isdigit():
            month = month_digit.zfill(2)
        else:
            month = "Unknown"

        model_name = lookup_model_name(model_code)

        return {
            "제조사": rev_maker.get(maker_code, "알 수 없음"),
            "카테고리": rev_category.get(category_code, "알 수 없음"),
            "모델 코드": model_code,
            "모델명": model_name,
            "제조년도": full_year,
            "제조월": month,
            "주문차수": order,
            "생산순서": sequence
        }

    except Exception as e:
        return {"오류": str(e)}