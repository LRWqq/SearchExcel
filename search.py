import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk  # Add this import for Combobox
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
        self.button = tk.Button(self.button_frame, text="選擇(拖曳)檔案", command=self.load_file)
        self.button.pack(side=tk.LEFT, padx=5)

        # 移除檔案按鈕
        self.remove_button = tk.Button(self.button_frame, text="移除檔案", command=self.remove_file, state=tk.DISABLED)
        self.remove_button.pack(side=tk.LEFT, padx=5)

        # Add column selection frame
        self.column_frame = tk.Frame(root)
        self.column_frame.pack(pady=5)
        
        # Search column selection with Combobox
        self.search_col_label = tk.Label(self.column_frame, text="搜尋欄位:")
        self.search_col_label.pack(side=tk.LEFT, padx=5)
        self.search_col_var = tk.StringVar()
        self.search_col_combo = ttk.Combobox(self.column_frame, 
                                           textvariable=self.search_col_var,
                                           width=5,
                                           state='readonly')
        self.search_col_combo.pack(side=tk.LEFT, padx=5)
        self.search_col_combo['values'] = ['A']  # Default value, will update when file loaded

        # Show all results checkbox
        self.show_all_var = tk.BooleanVar(value=False)
        self.show_all_check = tk.Checkbutton(root, text="顯示所有符合結果", 
                                            variable=self.show_all_var)
        self.show_all_check.pack(pady=5)

        # 啟用快捷鍵說明
        self.tip = tk.Label(root, text=" 選取欲查詢文字：Ctrl+C -> Ctrl+Alt+S ", fg="blue")
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
                
                # Update column selection dropdown
                num_cols = len(self.df.columns)
                col_letters = [chr(ord('A') + i) for i in range(num_cols)]
                self.search_col_combo['values'] = col_letters
                self.search_col_combo.set(col_letters[0])  # Set first column as default
                
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
        self.search_col_combo['values'] = ['A']  # Reset column choices
        self.search_col_combo.set('A')
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

        # Get column index from combobox
        search_col = ord(self.search_col_var.get()) - ord('A')
        
        # Search in specified column
        result = self.df[self.df.iloc[:, search_col] == text]

        if result.empty:
            messagebox.showinfo("查無資料", f"在 {self.search_col_var.get()} 欄未找到: {text}")
            return

        # Show results
        if self.show_all_var.get():
            # Get all matches when show_all is selected
            result = self.df[self.df.iloc[:, search_col] == text]
            
            if result.empty:
                messagebox.showinfo("查無資料", f"在 {self.search_col_var.get()} 欄未找到: {text}")
                return

            total_matches = len(result)
            info = f"找到 {total_matches} 筆符合結果:\n\n"
            for idx, row in result.iterrows():
                info += f"第 {idx+1} 筆:\n"
                for col_idx, value in enumerate(row):
                    col_letter = chr(ord('A') + col_idx)
                    info += f"{col_letter}: {value}\n"
                info += "-------------------\n"
        else:
            # Only check for up to 2 matches when not showing all
            matches = []
            for idx, value in enumerate(self.df.iloc[:, search_col]):
                if value == text:
                    matches.append(idx)
                    if len(matches) > 1:  # Stop after finding second match
                        break
            
            if not matches:
                messagebox.showinfo("查無資料", f"在 {self.search_col_var.get()} 欄未找到: {text}")
                return

            # Show first result
            row = self.df.iloc[matches[0]]
            info = "查詢結果:\n\n"
            for col_idx, value in enumerate(row):
                col_letter = chr(ord('A') + col_idx)
                info += f"{col_letter}: {value}\n"
            
            # Add note if there are more matches
            if len(matches) > 1:
                info += "\n注意: 尚有其他符合結果，請勾選「顯示所有符合結果」以查看全部"

        # Show results in scrollable text window
        result_window = tk.Toplevel(self.root)
        result_window.title("查詢結果")
        result_window.geometry("400x300")
        
        # Center the window relative to main window
        x = self.root.winfo_x() + (self.root.winfo_width() - 400) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 300) // 2
        result_window.geometry(f"+{x}+{y}")
        
        # Make window modal but keep main window active
        result_window.grab_set()
        result_window.focus_force()

        text_widget = tk.Text(result_window, wrap=tk.WORD, height=15, width=40)
        scrollbar = tk.Scrollbar(result_window, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)

        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        text_widget.insert(tk.END, info)
        text_widget.configure(state='disabled')

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = ExcelSearchApp(root)
    root.mainloop()
