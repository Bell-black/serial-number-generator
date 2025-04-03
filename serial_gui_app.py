import customtkinter as ctk
import tkinter.messagebox
import hashlib
import os
import barcode
import pandas as pd
import zipfile
import datetime
import subprocess

# 시리얼 생성 관련 설정
alpha_dict = {
    '1': 'A', '2': 'C', '3': 'D', '4': 'E', '5': 'F',
    '6': 'H', '7': 'J', '8': 'K', '9': 'L', '0': 'M',
    '10': 'M', '11': 'N', '12': 'P'
}

used_codes = set()
model_code_cache = {}

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
        "module_width": 0.3,
        "module_height": 20.0,
        "font_size": 12,
        "text_distance": 3.0,
        "quiet_zone": 5.0
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
                os.remove(file)  # 압축 후 원본 삭제
    return os.path.abspath(zip_filename)

def open_folder(path):
    try:
        if os.name == 'nt':
            os.startfile(path)
        else:
            subprocess.call(['open', path])
    except Exception as e:
        print(f"폴더 열기 실패: {e}")

# 제조사와 카테고리 매핑
maker_dict = {
    "닝보 타이웨이": "NB",
    "리앤텍": "LA",
    "마라타": "MT",
    "웨이슬라": "SV",
    "킹크린": "KE",
    "푸산 데코": "DC",
    "헝쉰전자": "HX",
    "화유": "HU",
    "중산 커리신": "KR"
}

category_dict = {
    "무선 진공 청소기": "MC",
    "무선 물걸레 청소기": "AC",
    "가습기": "MH",
    "공기청정기": "AP",
    "제습기": "DH",
    "선풍기": "MF",
    "에어프라이어": "AF",
    "블렌더": "MB"
}

# GUI 창
class SerialApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("시리얼넘버 생성기")
        self.geometry("600x620")
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
        self.entry_quantity = self.make_labeled_entry(container, "생성 개수", "개")

        button_frame = ctk.CTkFrame(container)
        button_frame.pack(anchor="w", pady=10)

        self.generate_btn = ctk.CTkButton(button_frame, text="시리얼 넘버 생성", command=self.generate_serials)
        self.generate_btn.pack(side="left", padx=(0, 10))

        self.open_folder_btn = ctk.CTkButton(button_frame, text="저장 위치 열기", command=self.open_saved_folder, state="disabled")
        self.open_folder_btn.pack(side="left")

        self.output_box = ctk.CTkTextbox(container, height=200)
        self.output_box.pack(fill="both", pady=(10,10))

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

    def generate_serials(self):
        global last_saved_file

        maker_name = self.maker_menu.get()
        category_name = self.category_menu.get()
        model = self.entry_model.get().strip()
        year = self.entry_year.get().strip()
        month = self.entry_month.get().lstrip("0")
        order = self.entry_order.get().strip()
        quantity = self.entry_quantity.get().strip()

        missing_fields = []
        if not model: missing_fields.append("모델명")
        if not year: missing_fields.append("제조년도")
        if not month: missing_fields.append("제조월")
        if not order: missing_fields.append("주문차수")
        if not quantity: missing_fields.append("생성 개수")

        if missing_fields:
            tkinter.messagebox.showerror("입력 오류", f"다음 항목을 입력해주세요: {', '.join(missing_fields)}")
            return

        try:
            qty = int(quantity)
            model_code = get_unique_code(model)
            maker_code = maker_dict[maker_name]
            category_code = category_dict[category_name]

            self.output_box.delete("1.0", "end")
            records = []
            serial_list = []
            for i in range(qty):
                seq = str(i + 1).zfill(5)
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

            if qty >= 3:
                zip_path = zip_svg_files(serial_list, model, year, month, order)
                self.output_box.insert("end", f"[압축 완료] {zip_path}\n")
                last_saved_file = zip_path

            self.open_folder_btn.configure(state="normal", fg_color="#009b77")
            tkinter.messagebox.showinfo("생성 완료", f"총 {qty}개의 시리얼 넘버가 생성되었습니다.")

        except ValueError:
            tkinter.messagebox.showerror("입력 오류", "생성 개수는 숫자로 입력해주세요.")
        except Exception as e:
            tkinter.messagebox.showerror("에러", str(e))

if __name__ == "__main__":
    app = SerialApp()
    app.mainloop()