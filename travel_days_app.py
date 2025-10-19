import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta
import csv
import os
import webbrowser 
# 引入 relativedelta，用於精確處理年數計算 (自動處理閏年)
from dateutil.relativedelta import relativedelta 

# CSV 檔案名稱
DATA_FILE = "bno_travel_data.csv"

# 顏色設定
COLOR_NORMAL = "black"
COLOR_YELLOW = "orange"
COLOR_ORANGE = "#FF6600"
COLOR_RED = "red"


class TravelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("BNO Visa 離境日數計算")
        # 略為增加高度以容納底部的儲存/Read Me資訊
        self.root.geometry("650x480") 
        self.root.resizable(False, False)
        self.root.configure(bg="#f0f0f0")

        self.rows = [] 
        self.entry_approval = None
        self.entry_arrival = None
        
        self.lbl_total = None
        self.lbl_365 = None
        self.lbl_period = None
        self.lbl_save_date = None

        self.create_widgets()
        self.load_data()
        self.calculate_days() 
        
    # === 新增功能：開啟 Read Me 網頁 ===
    def open_readme(self):
        """在預設瀏覽器中開啟指定的 GitHub 網頁"""
        url = 'https://github.com/ICHTBAA/bno_visa_cal'
        try:
            webbrowser.open_new_tab(url)
        except Exception as e:
            messagebox.showerror("瀏覽器錯誤", f"無法開啟網頁：{url}\n錯誤信息: {e}")

    # === 輔助方法 ===
    def parse_date(self, text):
        """將字串解析為 datetime 物件"""
        try:
            return datetime.strptime(text.strip(), "%Y-%m-%d").date() # 使用 .date() 避免時間影響
        except:
            return None

    def validate_date(self, entry):
        """檢查 Entry 中的日期格式是否正確，並設定背景顏色"""
        value = entry.get().strip()
        if value == "":
            entry.config(bg="white")
            return True
        try:
            datetime.strptime(value, "%Y-%m-%d")
            entry.config(bg="white")
            return True
        except ValueError:
            entry.config(bg="#ffcccc")
            return False

    def color_label(self, label, remain, thresholds):
        """根據剩餘日數設置標籤顏色"""
        if remain < 0:
            label.config(fg=COLOR_RED)
        elif remain <= thresholds[2]:
            label.config(fg=COLOR_RED)
        elif remain <= thresholds[1]:
            label.config(fg=COLOR_ORANGE)
        elif remain <= thresholds[0]:
            label.config(fg=COLOR_YELLOW)
        else:
            label.config(fg=COLOR_NORMAL)

    # === 介面創建與管理 (Grid 佈局) ===
    def create_widgets(self):
        """創建所有 Tkinter 元件並使用 pack/grid 佈局"""
        
        self.frame_rows = tk.Frame(self.root, bg="#f0f0f0")
        self.frame_rows.pack(pady=10)

        # 批核日/到達日 標籤 (row=0)
        tk.Label(self.frame_rows, text="批核日", bg="#f0f0f0", font=("Microsoft JhengHei", 10)).grid(row=0, column=0, padx=10, pady=(0, 3))
        tk.Label(self.frame_rows, text="到達日", bg="#f0f0f0", font=("Microsoft JhengHei", 10)).grid(row=0, column=1, padx=10, pady=(0, 3))

        # 批核日/到達日 輸入框 (row=1)
        self.entry_approval = tk.Entry(self.frame_rows, width=15, justify="center")
        self.entry_arrival = tk.Entry(self.frame_rows, width=15, justify="center")
        self.entry_approval.grid(row=1, column=0, padx=10, pady=(0, 10))
        self.entry_arrival.grid(row=1, column=1, padx=10, pady=(0, 10))
        self.entry_approval.bind("<FocusOut>", lambda e: self.validate_date(self.entry_approval))
        self.entry_arrival.bind("<FocusOut>", lambda e: self.validate_date(self.entry_arrival))

        # 出國/回國 標籤 (row=2)
        tk.Label(self.frame_rows, text="出國日", bg="#f0f0f0", font=("Microsoft JhengHei", 10)).grid(row=2, column=0, padx=10, pady=(0, 3))
        tk.Label(self.frame_rows, text="回國日", bg="#f0f0f0", font=("Microsoft JhengHei", 10)).grid(row=2, column=1, padx=10, pady=(0, 3))
        tk.Label(self.frame_rows, text="選取刪除", bg="#f0f0f0", font=("Microsoft JhengHei", 10)).grid(row=2, column=2, padx=10, pady=(0, 3))

        tk.Label(self.root, text="日期格式：yyyy-mm-dd", bg="#f0f0f0", fg="#555", font=("Microsoft JhengHei", 9)).pack(pady=(5, 2))

        # --- 按鈕區 ---
        frame_buttons = tk.Frame(self.root, bg="#f0f0f0")
        frame_buttons.pack(pady=5)
        tk.Button(frame_buttons, text="＋ 新增一行", width=10, command=self.add_row).grid(row=0, column=0, padx=5)
        tk.Button(frame_buttons, text="🗑 刪除選取", width=10, command=self.delete_selected).grid(row=0, column=1, padx=5)
        tk.Button(frame_buttons, text="💾 儲存", width=10, command=self.save_data).grid(row=0, column=2, padx=5)
        tk.Button(frame_buttons, text="📊 計算", width=10, command=self.calculate_days).grid(row=0, column=3, padx=5)

        # --- 結果顯示區 ---
        frame_results = tk.Frame(self.root, bg="#f0f0f0")
        frame_results.pack(pady=10)
        
        self.lbl_total = tk.Label(frame_results, text="總離境日數：0", bg="#f0f0f0", font=("Microsoft JhengHei", 10))
        self.lbl_total.pack()
        self.lbl_365 = tk.Label(frame_results, text="任意365日內最多離境：0（剩餘180日）", bg="#f0f0f0", font=("Microsoft JhengHei", 10))
        self.lbl_365.pack()
        
        # 450 日總離境日數的標籤 (將在 calculate_days 中動態更新描述)
        self.lbl_period = tk.Label(frame_results, text="入籍計算期（+1年至+6年）離境日數：N/A", bg="#f0f0f0", font=("Microsoft JhengHei", 10))
        self.lbl_period.pack()

        # --- Footer 區 (儲存日期 + Read Me 按鈕) ---
        frame_footer = tk.Frame(self.root, bg="#f0f0f0")
        frame_footer.pack(fill='x', padx=10, pady=(0, 10))
        
        self.lbl_save_date = tk.Label(frame_footer, text="上次儲存日期：未儲存", bg="#f0f0f0", font=("Microsoft JhengHei", 9), fg="#555")
        self.lbl_save_date.pack(side=tk.LEFT)
        
        # Read Me 按鈕
        tk.Button(frame_footer, text="💡 Read Me", command=self.open_readme).pack(side=tk.RIGHT)

    def add_row(self, out_date="", in_date=""):
        """新增一行出國/回國日期輸入框和 Checkbutton"""
        row = []
        # 計算新行在 grid 上的索引 (從 row=3 開始)
        new_row_index = 3 + len(self.rows) 
        
        entry_out = tk.Entry(self.frame_rows, width=15, justify="center")
        entry_out.grid(row=new_row_index, column=0, padx=10, pady=(0, 5))
        entry_out.bind("<FocusOut>", lambda e: self.validate_date(entry_out))
        entry_out.insert(0, out_date)
        row.append(entry_out)
        
        entry_in = tk.Entry(self.frame_rows, width=15, justify="center")
        entry_in.grid(row=new_row_index, column=1, padx=10, pady=(0, 5))
        entry_in.bind("<FocusOut>", lambda e: self.validate_date(entry_in))
        entry_in.insert(0, in_date)
        row.append(entry_in)
        
        var = tk.BooleanVar()
        chk = tk.Checkbutton(self.frame_rows, variable=var, bg="#f0f0f0")
        chk.grid(row=new_row_index, column=2, padx=5)
        row.append(var) # [Entry_out, Entry_in, BooleanVar]

        self.rows.append(row)

    def redraw_rows(self):
        """重新佈局剩餘的日期行（刪除後使用）"""
        
        # 1. 清除所有動態行 (row >= 3) 的舊佈局
        # 銷毀 Checkbutton，並忘記 Entry 的 grid 佈局
        for widget in self.frame_rows.winfo_children():
            row_info = widget.grid_info().get("row")
            if row_info is not None and row_info >= 3:
                if isinstance(widget, tk.Checkbutton):
                    widget.destroy()
                elif isinstance(widget, tk.Entry):
                    # 只是忘記 grid 佈局，Entry 物件本身仍存在於 self.rows 中
                    widget.grid_forget()

        # 2. 重新佈局剩餘的 Entry 和重新建立 Checkbutton
        for i, row in enumerate(self.rows):
            current_row_index = 3 + i
            
            # 重新佈局 Entry widgets
            # 只有未被刪除的行 (Entry 仍存在) 會在這裡重新 grid
            row[0].grid(row=current_row_index, column=0, padx=10, pady=(0, 5))
            row[1].grid(row=current_row_index, column=1, padx=10, pady=(0, 5))
            
            # 重新建立並佈局 Checkbutton (因為 Checkbutton 在步驟 1 中被銷毀)
            chk = tk.Checkbutton(self.frame_rows, variable=row[2], bg="#f0f0f0")
            chk.grid(row=current_row_index, column=2, padx=5)


    def delete_selected(self):
        """刪除所有被選取的日期行"""
        new_rows = []
        for row in self.rows:
            if row[2].get():
                # 銷毀被選取行的 Entry 元件
                row[0].destroy()
                row[1].destroy()
                # Checkbutton 會在 redraw_rows 中處理
            else:
                # 保留未被選取的行
                new_rows.append(row)
        
        self.rows = new_rows
        
        # 重新佈局剩下的行
        self.redraw_rows()
        
        if not self.rows:
            self.add_row()
        
        self.calculate_days()

    # === 數據管理 ===
    def load_data(self):
        """從 CSV 檔案載入數據"""
        if not os.path.exists(DATA_FILE):
            self.add_row()
            return
            
        with open(DATA_FILE, newline='', encoding="utf-8") as f:
            reader = csv.reader(f)
            data = list(reader)
            if not data:
                self.add_row()
                return

            # 載入批核日/到達日
            self.entry_approval.delete(0, tk.END)
            self.entry_arrival.delete(0, tk.END)
            self.entry_approval.insert(0, data[0][0])
            self.entry_arrival.insert(0, data[0][1])
            if len(data[0]) > 2:
                self.lbl_save_date.config(text=f"上次儲存日期：{data[0][2]}")

            # 清除舊的行並載入旅行記錄
            for row in self.rows:
                row[0].destroy()
                row[1].destroy()
            self.rows = []
            
            for r in data[1:]:
                self.add_row(r[0], r[1])
            if not self.rows:
                self.add_row()

    def save_data(self):
        """將數據儲存到 CSV 檔案"""
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        rows_data = [[self.entry_approval.get(), self.entry_arrival.get(), now]]
        for row in self.rows:
            out_date = row[0].get().strip()
            in_date = row[1].get().strip()
            if out_date or in_date:
                rows_data.append([out_date, in_date])

        with open(DATA_FILE, "w", newline='', encoding="utf-8") as f:
            csv.writer(f).writerows(rows_data)
            
        self.lbl_save_date.config(text=f"上次儲存日期：{now}")
        messagebox.showinfo("儲存完成", "資料已儲存。")
        self.calculate_days()

    # === 核心計算邏輯 (已修正閏年計算問題) ===
    def calculate_days(self):
        """計算所有 BNO 簽證相關的離境日數指標"""
        
        approval = self.parse_date(self.entry_approval.get())
        arrival = self.parse_date(self.entry_arrival.get())
        
        # 處理無效或空日期
        if not approval or not arrival:
            self.lbl_total.config(text="總離境日數：0")
            self.lbl_365.config(text="任意365日內最多離境：N/A")
            # 在沒有有效批核日期的情況下，顯示預設描述
            self.lbl_period.config(text=f"入籍計算期（+1年至+6年）離境日數：N/A")
            return

        if arrival < approval:
            messagebox.showwarning("日期錯誤", "到達日必須晚於或等於批核日。")
            return

        # 整理有效的旅行記錄
        trips = []
        for row in self.rows:
            start = self.parse_date(row[0].get())
            end = self.parse_date(row[1].get())
            
            if start and end:
                if end <= start:
                    messagebox.showwarning("日期錯誤", "回國日必須晚於出國日。")
                    return
                trips.append((start, end))
            elif start or end:
                messagebox.showwarning("日期錯誤", "出國日和回國日必須同時填寫。")
                return

        # 1. 批核日到首次到達日的離境日數：[批核日, 首次到達日 - 1 day]
        initial_departure_days = set()
        current = approval
        while current < arrival:
            initial_departure_days.add(current)
            current += timedelta(days=1)
        
        # 2. 旅行記錄的離境日數：[出國日 + 1 day, 回國日 - 1 day]
        travel_departure_days = set()
        for start, end in trips:
            # 根據 Home Office 指引，離境和抵達日不算作離境日
            current = start + timedelta(days=1)
            while current < end:
                travel_departure_days.add(current)
                current += timedelta(days=1)
        
        all_departure_days = initial_departure_days.union(travel_departure_days)

        # A. 總離境日數
        total_departure_days_count = len(all_departure_days)
        self.lbl_total.config(text=f"總離境日數：{total_departure_days_count}")

        # B. 任意 365 日內的離境日數 (ILR 180天限制)
        max_365 = 0
        # 只檢查到批核日 + 5 年 (大致的簽證期)
        end_of_check_period = approval + relativedelta(years=5) 
        
        current_date = approval
        while current_date < end_of_check_period:
            start_window = current_date
            end_window = current_date + timedelta(days=365)
            
            count = len([d for d in all_departure_days if start_window <= d < end_window])
            
            if count > max_365:
                max_365 = count
                
            current_date += timedelta(days=1)
        
        remain_365 = 180 - max_365

        self.lbl_365.config(text=f"任意365日內最多離境：{max_365}（剩餘 {remain_365} 日）")
        self.color_label(self.lbl_365, remain_365, [50, 30, 10])


        # C. ILR/入籍計算期總離境日數 (五年內 450 天總限制)
        
        # 修正：使用 relativedelta 確保精確計算年數，自動處理閏年
        # 計算新的動態日期範圍：[批核日 + 1 年, 批核日 + 6 年 - 1 天]
        ilr_start = approval + relativedelta(years=1) 
        ilr_end = (approval + relativedelta(years=6)) - timedelta(days=1) 

        # 重新計算落在動態期限內的離境日數
        period_days_count = len([d for d in all_departure_days 
                                 if ilr_start <= d <= ilr_end])
        
        remain_long = 450 - period_days_count
        
        # 格式化顯示的日期
        ilr_start_str = ilr_start.strftime('%Y/%#m/%#d')
        ilr_end_str = ilr_end.strftime('%Y/%#m/%#d')

        self.lbl_period.config(text=f"入籍計算期（{ilr_start_str}–{ilr_end_str}）離境日數：{period_days_count}（剩餘 {remain_long} 日）")
        self.color_label(self.lbl_period, remain_long, [120, 60, 30])


if __name__ == "__main__":
    # 設置 Tkinter 在 Mac 或 Windows 上的高 DPI 兼容性
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass

    root = tk.Tk()
    app = TravelApp(root)
    root.mainloop()