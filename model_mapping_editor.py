import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os

CSV_FILE = "model_map.csv"

class ModelMapEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("모델명 매핑 편집기")
        self.root.geometry("400x400")

        self.tree = ttk.Treeview(root, columns=("code", "name"), show="headings")
        self.tree.heading("code", text="모델코드")
        self.tree.heading("name", text="모델명")
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        self.load_data()

        entry_frame = tk.Frame(root)
        entry_frame.pack(pady=5)

        tk.Label(entry_frame, text="모델코드").grid(row=0, column=0)
        self.code_entry = tk.Entry(entry_frame, width=10)
        self.code_entry.grid(row=0, column=1, padx=5)

        tk.Label(entry_frame, text="모델명").grid(row=0, column=2)
        self.name_entry = tk.Entry(entry_frame, width=20)
        self.name_entry.grid(row=0, column=3, padx=5)

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=5)

        tk.Button(btn_frame, text="추가", command=self.add_entry).pack(side="left", padx=5)
        tk.Button(btn_frame, text="삭제", command=self.delete_entry).pack(side="left", padx=5)
        tk.Button(btn_frame, text="저장", command=self.save_data).pack(side="left", padx=5)

    def load_data(self):
        self.tree.delete(*self.tree.get_children())
        if os.path.exists(CSV_FILE):
            df = pd.read_csv(CSV_FILE)
            for _, row in df.iterrows():
                self.tree.insert("", "end", values=(row["모델코드"], row["모델명"]))

    def add_entry(self):
        code = self.code_entry.get().strip().upper()
        name = self.name_entry.get().strip()
        if not code or not name:
            messagebox.showwarning("입력 오류", "모델코드와 모델명을 모두 입력해주세요.")
            return
        for item in self.tree.get_children():
            if self.tree.item(item, "values")[0] == code:
                messagebox.showwarning("중복 코드", "이미 존재하는 모델코드입니다.")
                return
        self.tree.insert("", "end", values=(code, name))
        self.code_entry.delete(0, "end")
        self.name_entry.delete(0, "end")

    def delete_entry(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("선택 없음", "삭제할 항목을 선택해주세요.")
            return
        for item in selected:
            self.tree.delete(item)

    def save_data(self):
        data = []
        for item in self.tree.get_children():
            code, name = self.tree.item(item, "values")
            data.append({"모델코드": code, "모델명": name})
        df = pd.DataFrame(data)
        df.to_csv(CSV_FILE, index=False)
        messagebox.showinfo("저장 완료", "CSV 파일이 저장되었습니다.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ModelMapEditor(root)
    root.mainloop()
