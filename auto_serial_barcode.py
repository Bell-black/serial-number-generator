import barcode
import os
import pandas as pd
import hashlib
import zipfile
import datetime

# 회사의 최종 알파벳 치환 기준 적용
alpha_dict = {
    '1': 'A', '2': 'C', '3': 'D', '4': 'E', '5': 'F',
    '6': 'H', '7': 'J', '8': 'K', '9': 'L', '0': 'M',
    '10': 'M', '11': 'N', '12': 'P'
}

# 제조사명 → 코드 매핑 (실제 사용 제조사만)
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

# 제품 카테고리명 → 코드 매핑
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

# 중복 방지를 위한 모델 코드 저장소
used_codes = set()
model_code_cache = {}

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

def choose_from_list(title, options):
    print(f"\n[{title}]")
    for idx, key in enumerate(options.keys(), start=1):
        print(f"{idx}. {key}")
    while True:
        try:
            choice = int(input(f"{title} 번호 입력: "))
            if 1 <= choice <= len(options):
                selected_key = list(options.keys())[choice - 1]
                return selected_key, options[selected_key]
        except ValueError:
            pass
        print("유효한 번호를 입력해주세요.")

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
    print(f"[생성완료] 바코드 저장: {filename}.svg")

def zip_barcode_files(serial_list):
    date_str = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    zip_filename = f"barcodes_{date_str}.zip"
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for serial in serial_list:
            file = f"barcode_{serial}.svg"
            if os.path.exists(file):
                zipf.write(file)
    print(f"[ZIP 생성 완료] {zip_filename}")

def get_next_seq():
    if os.path.exists("latest_serial.txt"):
        with open("latest_serial.txt", "r") as f:
            latest = int(f.read().strip())
    else:
        latest = 0
    return latest + 1

def update_latest_seq(seq):
    with open("latest_serial.txt", "w") as f:
        f.write(str(seq))

def save_to_excel(data):
    filename = "serial_numbers.xlsx"
    try:
        if os.path.exists(filename):
            df_existing = pd.read_excel(filename)
            df_new = pd.concat([df_existing, pd.DataFrame(data)], ignore_index=True)
        else:
            df_new = pd.DataFrame(data)
        df_new.to_excel(filename, index=False)
        print(f"[엑셀 저장 완료] 파일명: {filename}")
    except PermissionError:
        print(f"[오류] 엑셀 파일이 열려 있어서 저장할 수 없습니다. '{filename}' 파일을 닫고 다시 실행해주세요.")

def main():
    print("=== 시리얼넘버 자동생성기 (제조사/제품 카테고리 코드 선택형) ===")
    maker_input, maker = choose_from_list("제조사", maker_dict)
    category_input, category = choose_from_list("제품 카테고리", category_dict)
    model_name = input("모델명 입력 (예: AMH-9000): ")
    model_code = get_unique_code(model_name)
    year = input("제조년도 입력 (4자리 숫자, 예: 2025): ")
    month = input("제조월 입력 (숫자 1~12): ").lstrip("0")
    order = input("주문차수 입력 (숫자): ")
    quantity = int(input("생성할 시리얼 개수 입력: "))

    next_seq = get_next_seq()
    records = []
    serial_list = []

    for i in range(quantity):
        seq_num = str(next_seq + i).zfill(5)
        serial = generate_serial(maker, category, model_code, year, month, order, seq_num)
        print(f"[시리얼 생성] {serial}")
        generate_barcode(serial)
        serial_list.append(serial)

        records.append({
            "제조사": maker_input,
            "제조사 코드": maker,
            "제품 카테고리": category_input,
            "카테고리 코드": category,
            "모델명": model_name,
            "모델 코드": model_code,
            "제조년도": year,
            "제조월": month,
            "주문차수": order,
            "생산순서": seq_num,
            "시리얼넘버": serial
        })

    save_to_excel(records)
    update_latest_seq(next_seq + quantity - 1)

    if quantity >= 30:
        zip_barcode_files(serial_list)

if __name__ == "__main__":
    main()
