import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pyperclip
import keyboard
from tkinterdnd2 import DND_FILES, TkinterDnD

class ExcelSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 查詢工具")
        self.root.geometry("400x250")  # Made window slightly taller

        self.df = None
        self.current_file = None  # Store current file path

        # 拖曳提示
        self.label = tk.Label(root, text="請拖曳 Excel 檔案到此視窗\n或按下按鈕選擇檔案", pady=20)
        self.label.pack()

        # 目前檔案顯示
        self.file_label = tk.Label(root, text="目前檔案: 未載入", wraplength=350)
        self.file_label.pack(pady=5)

        # Frame for buttons
        self.button_frame = tk.Frame(root)
        self.button_frame.pack(pady=5)

        # 選擇檔案按鈕
        self.button = tk.Button(self.button_frame, text="選擇檔案", command=self.load_file)
        self.button.pack(side=tk.LEFT, padx=5)

        # 移除檔案按鈕
        self.remove_button = tk.Button(self.button_frame, text="移除檔案", command=self.remove_file, state=tk.DISABLED)
        self.remove_button.pack(side=tk.LEFT, padx=5)

        # 啟用快捷鍵說明
        self.tip = tk.Label(root, text="熱鍵: Ctrl+Alt+S → 查詢選取文字", fg="blue")
        self.tip.pack(pady=10)

        # 註冊快捷鍵
        keyboard.add_hotkey("ctrl+alt+s", self.search_selected_text)

        # 支援拖曳
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.drop_file)

    def load_file(self, file_path=None):
        if not file_path:
            file_path = filedialog.askopenfilename(filetypes=[("Excel/CSV Files", "*.xlsx *.xls *.csv")])
        if file_path:
            try:
                # Check file extension
                if file_path.lower().endswith('.csv'):
                    self.df = pd.read_csv(file_path, dtype=str)
                else:
                    self.df = pd.read_excel(file_path, dtype=str)
                self.current_file = file_path
                self.file_label.config(text=f"目前檔案: {file_path}")
                self.remove_button.config(state=tk.NORMAL)
                messagebox.showinfo("成功", f"已載入檔案:\n{file_path}")
            except Exception as e:
                messagebox.showerror("錯誤", f"無法讀取檔案: {e}")

    def remove_file(self):
        self.df = None
        self.current_file = None
        self.file_label.config(text="目前檔案: 未載入")
        self.remove_button.config(state=tk.DISABLED)
        messagebox.showinfo("移除", "已移除檔案")

    def drop_file(self, event):
        path = event.data.strip("{}")  # 去掉拖曳的花括號
        self.load_file(path)

    def search_selected_text(self):
        if self.df is None:
            messagebox.showwarning("警告", "請先載入 Excel 檔案")
            return

        # 取得目前選取文字
        keyboard.send("ctrl+c")
        self.root.after(200, self.do_search)

    def do_search(self):
        text = pyperclip.paste().strip()
        if not text:
            messagebox.showwarning("警告", "未取得選取文字")
            return

        result = self.df[self.df.iloc[:,0] == text]  # A 欄 = 第一欄
        if not result.empty:
            row = result.iloc[0]
            info = f"A: {row[0]}\nB: {row[1]}\nC: {row[2]}"
            messagebox.showinfo("查詢結果", info)
        else:
            messagebox.showinfo("查無資料", f"Excel A欄未找到: {text}")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = ExcelSearchApp(root)
    root.mainloop()
