import customtkinter as ctk
import tkinter.messagebox
import hashlib
import os
import barcode
import pandas as pd
import zipfile
import datetime
import subprocess
from xml.etree import ElementTree as ET

# 시리얼 생성 관련 설정
alpha_dict = {
    '1': 'A', '2': 'C', '3': 'D', '4': 'E', '5': 'F',
    '6': 'H', '7': 'J', '8': 'K', '9': 'L', '0': 'M',
    '10': 'M', '11': 'N', '12': 'P'
}

used_codes = set()
model_code_cache = {}
model_map_file = "model_map.csv"

last_saved_file = ""

def num_to_alpha(num):
    return ''.join(alpha_dict[digit] for digit in str(num))

def model_to_number(model_name):
    model_name = model_name.upper()
    h = hashlib.sha256(model_name.encode()).hexdigest()
    return int(h, 16)

def number_to_code(num):
    num = num % 676
    first = chr(ord('A') + num // 26)
    second = chr(ord('A') + num % 26)
    return first + second

def get_unique_code(model_name):
    model_name = model_name.upper()
    if model_name in model_code_cache:
        return model_code_cache[model_name]
    base_num = model_to_number(model_name)
    for offset in range(676):
        code = number_to_code(base_num + offset)
        if code not in used_codes:
            used_codes.add(code)
            model_code_cache[model_name] = code
            return code
    raise Exception("모든 코드가 소진되었습니다! (676개 제한)")

def generate_serial(maker, category, model_code, year, month, order, seq):
    year_alpha = num_to_alpha(year[-1])
    month_alpha = alpha_dict[month]
    order_number = str(order).zfill(2)
    return f"{maker}{category}{model_code}{year_alpha}{month_alpha}{order_number}{seq}"

def generate_barcode(serial):
    CODE128 = barcode.get_barcode_class('code128')
    writer = barcode.writer.SVGWriter()
    writer.set_options({
        "module_width": 0.6,
        "module_height": 80.0,
        "font_size": 20,
        "text_distance": 5.0,
        "quiet_zone": 2.0,
        "write_text": True
    })
    barcode_img = CODE128(serial, writer=writer)
    filename = barcode_img.save(f'barcode_{serial}')
    return filename

def save_to_excel(data):
    filename = "serial_numbers_gui.xlsx"
    if os.path.exists(filename):
        df_existing = pd.read_excel(filename)
        df_new = pd.concat([df_existing, pd.DataFrame(data)], ignore_index=True)
    else:
        df_new = pd.DataFrame(data)
    df_new.to_excel(filename, index=False)
    return os.path.abspath(filename)

def zip_svg_files(serial_list, model_name, year, month, order):
    short_date = datetime.datetime.now().strftime('%y%m%d')
    zip_filename = f"serial-number_{short_date}_{model_name}_{year}년_{month}월_{order}차.zip"
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for serial in serial_list:
            file = f"barcode_{serial}.svg"
            if os.path.exists(file):
                zipf.write(file)
                os.remove(file)
    return os.path.abspath(zip_filename)

def save_model_mapping(model_name, model_code):
    try:
        if os.path.exists(model_map_file):
            df = pd.read_csv(model_map_file)
            if not ((df['모델코드'] == model_code) & (df['모델명'] == model_name)).any():
                df = pd.concat([df, pd.DataFrame([{"모델코드": model_code, "모델명": model_name}])], ignore_index=True)
                df.to_csv(model_map_file, index=False)
        else:
            df = pd.DataFrame([{"모델코드": model_code, "모델명": model_name}])
            df.to_csv(model_map_file, index=False)
    except Exception as e:
        print(f"[모델 매핑 저장 오류] {e}")

def lookup_model_name(code):
    if not os.path.exists(model_map_file):
        return "(매핑 없음)"
    try:
        df = pd.read_csv(model_map_file)
        row = df[df['모델코드'] == code]
        if not row.empty:
            return row.iloc[0]['모델명']
        else:
            return "(매핑 없음)"
    except Exception as e:
        return f"(에러: {e})"

def guess_full_year(last_digit):
    from datetime import datetime
    try:
        current_year = datetime.now().year
        current_decade = current_year // 10 * 10
        candidate_year = current_decade + int(last_digit)
        if candidate_year > current_year + 1:
            candidate_year -= 10
        return str(candidate_year)
    except:
        pass
    return "Unknown"

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
        rev_alpha = {v: k for k, v in alpha_dict.items() if len(k) == 1}

        year_digit = rev_alpha.get(year_alpha, None)
        full_year = guess_full_year(year_digit) if year_digit else 'Unknown'
        month = rev_alpha.get(month_alpha, 'Unknown')
        if isinstance(month, str) and month.isdigit():
            month = month.zfill(2)

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

# 제조사 및 카테고리 코드
maker_dict = {
    "닝보 타이웨이": "NB",
    "리앤텍": "HL",
    "마라타": "MT",
    "웨이슬라": "VS",
    "킹크린": "KE",
    "푸산 데코": "DC",
    "헝쉰전자": "HX",
    "화유": "HU",
    "중산 커리신": "ZK"
}

category_dict = {
    "무선 진공 청소기": "MC",
    "무선 물걸레 청소기": "AC",
    "가습기": "MH",
    "공기청정기": "AP",
    "제습기": "DH",
    "선풍기": "MF",
    "에어프라이어": "AF",
    "블렌더": "MB",
    "헤어 드라이기": "MS",
    "음식물 처리기": "FP"
}

# 전체 GUI 앱 클래스 및 실행
class SerialApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("에어메이드 시리얼넘버 생성기")
        self.geometry("600x800")
        self.build_ui()

    def build_ui(self):
        container = ctk.CTkFrame(self)
        container.pack(anchor="w", padx=20, pady=10, fill="both", expand=True)

        ctk.CTkLabel(container, text="제조사 선택").pack(anchor="w")
        self.maker_menu = ctk.CTkOptionMenu(container, values=list(maker_dict.keys()))
        self.maker_menu.pack(anchor="w", pady=5)

        ctk.CTkLabel(container, text="제품 카테고리 선택").pack(anchor="w")
        self.category_menu = ctk.CTkOptionMenu(container, values=list(category_dict.keys()))
        self.category_menu.pack(anchor="w", pady=5)

        self.entry_model = self.make_labeled_entry(container, "모델명")
        self.entry_year = self.make_labeled_entry(container, "제조년도 (예: 2025)", "년")
        self.entry_month = self.make_labeled_entry(container, "제조월 (1~12 또는 01~12)", "월")
        self.entry_order = self.make_labeled_entry(container, "주문차수", "차")
        self.entry_start = self.make_labeled_entry(container, "시작 번호", "부터")
        self.entry_end = self.make_labeled_entry(container, "끝 번호", "까지")

        button_frame = ctk.CTkFrame(container)
        button_frame.pack(anchor="w", pady=10)

        self.generate_btn = ctk.CTkButton(button_frame, text="시리얼 넘버 생성", command=self.generate_serials)
        self.generate_btn.pack(side="left", padx=(0, 10))

        self.open_folder_btn = ctk.CTkButton(button_frame, text="저장 위치 열기", command=self.open_saved_folder, state="disabled")
        self.open_folder_btn.pack(side="left")

        self.output_box = ctk.CTkTextbox(container, height=200)
        self.output_box.pack(fill="both", pady=(10,10))

        decode_frame = ctk.CTkFrame(container)
        decode_frame.pack(anchor="w", pady=(5, 10), fill="x")
        ctk.CTkLabel(decode_frame, text="시리얼 넘버 조회:").pack(side="left")
        self.decode_entry = ctk.CTkEntry(decode_frame, width=200)
        self.decode_entry.pack(side="left", padx=(5, 5))
        decode_btn = ctk.CTkButton(decode_frame, text="조회", command=self.decode_serial_ui)
        decode_btn.pack(side="left")

    def make_labeled_entry(self, parent, label, suffix_text=None):
        frame = ctk.CTkFrame(parent)
        frame.pack(anchor="w", pady=(10,0))
        ctk.CTkLabel(frame, text=label).pack(side="left")
        entry = ctk.CTkEntry(frame, width=120)
        entry.pack(side="left", padx=(10, 5))
        if suffix_text:
            ctk.CTkLabel(frame, text=suffix_text).pack(side="left")
        return entry

    def open_saved_folder(self):
        open_folder(os.path.dirname(last_saved_file))

    def decode_serial_ui(self):
        serial = self.decode_entry.get().strip()
        if not serial:
            tkinter.messagebox.showerror("입력 오류", "시리얼 넘버를 입력해주세요")
            return
        info = decode_serial(serial)
        result = "\n".join(f"{k}: {v}" for k, v in info.items())
        self.output_box.delete("1.0", "end")
        self.output_box.insert("end", result + "\n")

    def generate_serials(self):
        global last_saved_file

        maker_name = self.maker_menu.get()
        category_name = self.category_menu.get()
        model = self.entry_model.get().strip()
        year = self.entry_year.get().strip()
        month = self.entry_month.get().lstrip("0")
        order = self.entry_order.get().strip()
        start = self.entry_start.get().strip()
        end = self.entry_end.get().strip()

        missing_fields = []
        if not model: missing_fields.append("모델명")
        if not year: missing_fields.append("제조년도")
        if not month: missing_fields.append("제조월")
        if not order: missing_fields.append("주문차수")
        if not start: missing_fields.append("시작 번호")
        if not end: missing_fields.append("끝 번호")

        if missing_fields:
            tkinter.messagebox.showerror("입력 오류", f"다음 항목을 입력해주세요: {', '.join(missing_fields)}")
            return

        try:
            start_num = int(start)
            end_num = int(end)
            if not (1 <= start_num <= 99999 and 1 <= end_num <= 99999 and start_num <= end_num):
                raise ValueError("시작/끝 번호는 1~99999 사이의 숫자이며 시작이 끝보다 작거나 같아야 합니다.")

            model_code = get_unique_code(model)
            save_model_mapping(model, model_code)
            maker_code = maker_dict[maker_name]
            category_code = category_dict[category_name]

            self.output_box.delete("1.0", "end")
            records = []
            serial_list = []
            for i in range(start_num, end_num + 1):
                seq = str(i).zfill(5)
                serial = generate_serial(maker_code, category_code, model_code, year, month, order, seq)
                generate_barcode(serial)
                self.output_box.insert("end", serial + "\n")
                records.append({
                    "시리얼넘버": serial,
                    "제조사": maker_name,
                    "제품 카테고리": category_name,
                    "모델명": model,
                    "제조년도": year,
                    "제조월": month,
                    "주문차수": order,
                    "생산순서": seq
                })
                serial_list.append(serial)

            excel_path = save_to_excel(records)
            last_saved_file = excel_path
            self.output_box.insert("end", f"\n[엑셀 저장 완료] {excel_path}\n")

            if len(serial_list) >= 3:
                zip_path = zip_svg_files(serial_list, model, year, month, order)
                self.output_box.insert("end", f"[압축 완료] {zip_path}\n")
                last_saved_file = zip_path

            self.open_folder_btn.configure(state="normal", fg_color="#009b77")
            tkinter.messagebox.showinfo("생성 완료", f"총 {len(serial_list)}개의 시리얼 넘버가 생성되었습니다.")

        except ValueError as ve:
            tkinter.messagebox.showerror("입력 오류", str(ve))
        except Exception as e:
            tkinter.messagebox.showerror("에러", str(e))

if __name__ == '__main__':
    app = SerialApp()
    app.mainloop()
