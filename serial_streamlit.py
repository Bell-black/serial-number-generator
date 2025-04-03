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
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import json
import streamlit as st
import tempfile

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
    filename = barcode_obj.save(f"barcode_{serial}")  # 확장자 자동 포함됨
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

def save_to_excel(data):
    fname = "serial_numbers_streamlit.xlsx"
    if os.path.exists(fname):
        df = pd.read_excel(fname)
        df = pd.concat([df, pd.DataFrame(data)], ignore_index=True)
    else:
        df = pd.DataFrame(data)
    df.to_excel(fname, index=False)
    return os.path.abspath(fname)

# 기존 save_to_excel 함수 아래 또는 근처에 붙이기
def upload_excel_to_drive(filepath, folder_id=None):
    try:
        # ✅ secrets에서 서비스 계정 정보 불러오기
        info = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
        creds = service_account.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/drive"]
        )
        service = build("drive", "v3", credentials=creds)

        file_metadata = {
            "name": os.path.basename(filepath),
            "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }

        if folder_id:
            file_metadata["parents"] = [folder_id]

        media = MediaFileUpload(filepath, mimetype=file_metadata["mimeType"], resumable=True)

        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id"
        ).execute()

        st.info(f"✅ 파일이 Google Drive에 업로드되었습니다. 파일 ID: {file.get('id')}")
        return file.get("id")

    except Exception as e:
        st.error(f"[❌ Google Drive 업로드 실패] {e}")
        return None

# 시리얼 넘버 생성 버튼 내부에서 파일 저장 후 Google Drive 업로드 연결
def save_to_excel(data):
    filename = "serial_numbers_streamlit.xlsx"
    if os.path.exists(filename):
        df_existing = pd.read_excel(filename)
        df_new = pd.concat([df_existing, pd.DataFrame(data)], ignore_index=True)
    else:
        df_new = pd.DataFrame(data)
    df_new.to_excel(filename, index=False)

    # ✅ 저장 후 Google Drive에 업로드
    upload_excel_to_drive(filename, folder_id="serial_uploads")

    return os.path.abspath(filename)
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

        # 제조사, 카테고리, 알파벳 코드 역변환
        rev_maker = {v: k for k, v in maker_dict.items()}
        rev_category = {v: k for k, v in category_dict.items()}
        rev_alpha = {v: k for k, v in alpha_dict.items()}  # 🔧 전체 키 포함 (0~12월까지)

        # 제조년도 해석
        year_digit = rev_alpha.get(year_alpha, None)
        full_year = guess_full_year(year_digit) if year_digit else "Unknown"

        # 제조월 해석
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
def search_serial_from_excel(serial):
    filename = "serial_numbers_streamlit.xlsx"
    if not os.path.exists(filename):
        return None
    try:
        df = pd.read_excel(filename)
        result = df[df['시리얼넘버'] == serial]
        return result.to_dict(orient="records")[0] if not result.empty else None
    except Exception as e:
        return {"오류": str(e)}


# --------------------------
# Streamlit UI
# --------------------------
st.set_page_config(page_title="시리얼 넘버 생성기", layout="centered")
st.title("📦 시리얼 넘버 자동 생성기")
st.caption("💡 각 입력 필드는 Enter 대신 Tab 키로 이동하세요.")

if 'clicked' not in st.session_state:
    st.session_state.clicked = False

# 입력 필드
maker_name = st.selectbox("제조사", list(maker_dict.keys()), key="maker")
category_name = st.selectbox("제품 카테고리", list(category_dict.keys()), key="category")

model = st.text_input("모델명", key="model")
if st.session_state.clicked and not model:
    st.error("모델명을 입력해주세요.")

year = st.text_input("제조년도 (예: 2025)", key="year")
if st.session_state.clicked and (not year.isdigit() or len(year) != 4):
    st.error("제조년도는 4자리 숫자로 입력해주세요.")

month = st.text_input("제조월 (1~12)", key="month")
if st.session_state.clicked and (not month.isdigit() or not (1 <= int(month) <= 12)):
    st.error("제조월은 1~12 사이의 숫자로 입력해주세요.")

order = st.text_input("주문차수", key="order")
if st.session_state.clicked and not order.isdigit():
    st.error("주문차수는 숫자로 입력해주세요.")

start_num = st.text_input("시작 번호", key="start")
if st.session_state.clicked and not start_num.isdigit():
    st.error("시작 번호는 숫자로 입력해주세요.")

end_num = st.text_input("끝 번호", key="end")
if st.session_state.clicked and not end_num.isdigit():
    st.error("끝 번호는 숫자로 입력해주세요.")

# ... 기존 코드 생략 ...

# ... 기존 코드 생략 ...
import streamlit.components.v1 as components

# 시리얼 넘버 생성
if st.button("✅ 시리얼 넘버 생성"):
    st.session_state.clicked = True

    valid = all([
        model, year.isdigit() and len(year) == 4,
        month.isdigit() and 1 <= int(month) <= 12,
        order.isdigit(), start_num.isdigit(), end_num.isdigit()
    ])

    if not valid:
        st.warning("입력값을 다시 확인해주세요.")
    else:
        start = int(start_num)
        end = int(end_num)
        if start < 1 or end < start:
            st.error("시작 번호와 끝 번호를 다시 확인해주세요.")
        else:
            try:
                model_code = get_unique_code(model)
                save_model_mapping(model, model_code)
                maker_code = maker_dict[maker_name]
                category_code = category_dict[category_name]

                results, records, serial_list = [], [], []
                for i in range(start, end + 1):
                    seq = str(i).zfill(5)
                    serial = generate_serial(maker_code, category_code, model_code, year, month.lstrip("0"), order, seq)
                    svg_path = generate_barcode_svg(serial)
                    results.append((serial, svg_path))
                    serial_list.append(serial)
                    records.append({
                        "시리얼넘버": serial, "제조사": maker_name, "제품 카테고리": category_name,
                        "모델명": model, "제조년도": year, "제조월": month,
                        "주문차수": order, "생산순서": seq
                    })

                save_to_excel(records)
                st.success(f"총 {len(results)}개의 시리얼 넘버를 생성했습니다.")

                # 세션 상태에 시리얼 넘버 리스트 저장
                st.session_state["serial_list"] = serial_list

                if len(results) > 1:
                    zip_name = "barcodes_download.zip"
                    with zipfile.ZipFile(zip_name, 'w') as zipf:
                        for _, path in results:
                            zipf.write(path)
                    with open(zip_name, "rb") as zf:
                        st.download_button("ZIP 파일 다운로드", data=zf, file_name=zip_name, mime="application/zip")
                else:
                    serial, path = results[0]
                    with open(path, "rb") as f:
                        st.download_button(f"{serial} 바코드 다운로드", data=f, file_name=os.path.basename(path), mime="image/svg+xml")
            except Exception as e:
                st.error(f"에러 발생: {e}")

# ✅ 생성된 시리얼 넘버 리스트 출력 (텍스트창 + 클립보드 복사 버튼)
if "serial_list" in st.session_state:
    serial_text = "\n".join(st.session_state["serial_list"])
    st.text_area("📄 생성된 시리얼 넘버 목록", value=serial_text, height=200, disabled=True)
    components.html(f"""
        <button onclick=\"navigator.clipboard.writeText(`{serial_text}`); alert('시리얼 넘버가 클립보드에 복사되었습니다!')\"
                style=\"margin-top: 10px; padding: 8px 16px; font-size: 16px; cursor: pointer; border-radius: 6px;\">
            📋 복사하기
        </button>
    """, height=60)

# 시리얼 넘버 조회
st.subheader("🔍 시리얼 넘버 조회")
decode_input = st.text_input("시리얼 넘버 입력 (최대 15자리)", max_chars=15, key="decode_input")
if st.button("조회"):
    if decode_input:
        serial = decode_input.strip()
        record = search_serial_from_excel(serial)
        if record:
            st.success("📄 등록된 시리얼 넘버입니다.")
            for k, v in record.items():
                st.write(f"{k}: {v}")
        else:
            st.error("❌ 조회하신 시리얼 넘버는 존재하지 않는 시리얼 넘버입니다.")
    else:
        st.warning("시리얼 넘버를 입력해주세요.")

