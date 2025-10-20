import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta
import csv
import os
import webbrowser 
from dateutil.relativedelta import relativedelta 

DATA_FILE = "bno_travel_data.csv"

COLOR_NORMAL = "black"
COLOR_YELLOW = "orange"
COLOR_ORANGE = "#FF6600"
COLOR_RED = "red"


class TravelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("BNO Visa 離境日數計算")
        self.root.geometry("780x890") 
        self.root.resizable(False, True) 
        self.root.configure(bg="#f0f0f0")
        
        self.is_saved = True
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.rows = [] 
        self.entry_approval = None
        self.entry_arrival = None
        
        self.lbl_total = None
        self.lbl_max_365 = None 
        self.lbl_period = None
        self.lbl_final_year = None 
        self.lbl_save_date = None
        self.lbl_past_365 = None 
        self.max_365_periods = [] 
        
        self.canvas = None
        self.scrollable_frame = None 
        
        self.create_widgets()
        self.load_data()
        self.calculate_days() 
    
    def on_closing(self):
        if not self.is_saved:
            if messagebox.askyesnocancel("儲存確認", "有未儲存的資料，您想在退出前儲存嗎？"):
                self.save_data()
                self.root.destroy()
            elif messagebox.askyesno("退出確認", "您確定要退出且不儲存資料嗎？"):
                self.root.destroy()
        else:
            self.root.destroy()
    
    def auto_hyphenate_date(self, event):
        entry = event.widget
        value = entry.get()
        cursor_pos = entry.index(tk.INSERT)
        
        if len(value) == 4 and value.isdigit() and value[-1] != '-' and cursor_pos == 4:
            entry.insert(4, "-")
            
        elif len(value) == 7 and value.count('-') == 1 and cursor_pos == 7:
            if value[0:4].isdigit() and value[5:7].isdigit():
                entry.insert(7, "-")
    
    def mark_unsaved(self, event=None):
        self.is_saved = False
        
    def open_readme(self):
        url = 'https://github.com/ICHTBAA/bno_visa_cal'
        try:
            webbrowser.open_new_tab(url)
        except Exception as e:
            messagebox.showerror("瀏覽器錯誤", f"無法開啟網頁：{url}\n錯誤信息: {e}")

    def parse_date(self, text):
        try:
            return datetime.strptime(text.strip(), "%Y-%m-%d").date()
        except:
            return None

    def validate_date(self, entry):
        value = entry.get().strip()
        if value == "":
            entry.config(bg="white")
            self.mark_unsaved() 
            return True
        try:
            datetime.strptime(value, "%Y-%m-%d")
            entry.config(bg="white")
            self.mark_unsaved()
            return True
        except ValueError:
            entry.config(bg="#ffcccc")
            self.mark_unsaved()
            return False

    def color_label(self, label, remain, thresholds):
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

    def create_widgets(self):
        
        frame_static_top = tk.Frame(self.root, bg="#f0f0f0")
        frame_static_top.pack(pady=10)

        tk.Label(frame_static_top, text="批核日", bg="#f0f0f0", font=("Microsoft JhengHei", 10)).grid(row=0, column=0, padx=10, pady=(0, 3))
        tk.Label(frame_static_top, text="到達日", bg="#f0f0f0", font=("Microsoft JhengHei", 10)).grid(row=0, column=1, padx=10, pady=(0, 3))
        
        tk.Label(frame_static_top, text="日期格式：yyyy-mm-dd", bg="#f0f0f0", fg="#555", font=("Microsoft JhengHei", 9)).grid(row=0, column=2, padx=(10, 5), pady=(0, 3), columnspan=3, sticky="e") 

        self.entry_approval = tk.Entry(frame_static_top, width=15, justify="center")
        self.entry_arrival = tk.Entry(frame_static_top, width=15, justify="center")
        
        self.entry_approval.grid(row=1, column=0, padx=10, pady=(0, 10))
        self.entry_arrival.grid(row=1, column=1, padx=10, pady=(0, 10))
        
        tk.Label(frame_static_top, text="", width=10, bg="#f0f0f0").grid(row=1, column=2, padx=5, pady=(0, 10)) 
        tk.Label(frame_static_top, text="", width=15, bg="#f0f0f0").grid(row=1, column=3, padx=10, pady=(0, 10)) 
        tk.Label(frame_static_top, text="", width=5, bg="#f0f0f0").grid(row=1, column=4, padx=5, pady=(0, 10)) 

        self.entry_approval.bind("<FocusOut>", lambda e: self.validate_date(self.entry_approval))
        self.entry_arrival.bind("<FocusOut>", lambda e: self.validate_date(self.entry_arrival))
        
        self.entry_approval.bind("<KeyRelease>", self.auto_hyphenate_date)
        self.entry_arrival.bind("<KeyRelease>", self.auto_hyphenate_date)
        self.entry_approval.bind("<KeyRelease>", self.mark_unsaved, add='+')
        self.entry_arrival.bind("<KeyRelease>", self.mark_unsaved, add='+')
        
        frame_scroll_container = tk.Frame(self.root)
        frame_scroll_container.pack(fill="both", expand=True, padx=10)

        scrollbar = tk.Scrollbar(frame_scroll_container, orient="vertical")
        
        self.canvas = tk.Canvas(frame_scroll_container, yscrollcommand=scrollbar.set, bg="#f0f0f0")

        scrollbar.config(command=self.canvas.yview)
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.scrollable_frame = tk.Frame(self.canvas, bg="#f0f0f0")

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw", 
                                   tags="self.scrollable_frame")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.config(
                scrollregion=self.canvas.bbox("all")
            )
        )
        
        tk.Label(self.scrollable_frame, text="出國日", bg="#f0f0f0", font=("Microsoft JhengHei", 10)).grid(row=0, column=0, padx=10, pady=(0, 3))
        tk.Label(self.scrollable_frame, text="回國日", bg="#f0f0f0", font=("Microsoft JhengHei", 10)).grid(row=0, column=1, padx=10, pady=(0, 3))
        tk.Label(self.scrollable_frame, text="365日離境", bg="#f0f0f0", font=("Microsoft JhengHei", 10)).grid(row=0, column=2, padx=5, pady=(0, 3))         tk.Label(self.scrollable_frame, text="活動", bg="#f0f0f0", font=("Microsoft JhengHei", 10)).grid(row=0, column=3, padx=10, pady=(0, 3)) # Moved to column 3
        tk.Label(self.scrollable_frame, text="刪除", bg="#f0f0f0", font=("Microsoft JhengHei", 10)).grid(row=0, column=4, padx=10, pady=(0, 3)) # Moved to column 4

        frame_buttons = tk.Frame(self.root, bg="#f0f0f0")
        frame_buttons.pack(pady=5)
        tk.Button(frame_buttons, text="＋ 新增一行", width=10, command=self.add_row).grid(row=0, column=0, padx=5)
        tk.Button(frame_buttons, text="🗑 刪除選取", width=10, command=self.delete_selected).grid(row=0, column=1, padx=5)
        tk.Button(frame_buttons, text="💾 儲存", width=10, command=self.save_data).grid(row=0, column=2, padx=5)
        tk.Button(frame_buttons, text="📊 計算", width=10, command=self.calculate_days).grid(row=0, column=3, padx=5)

        frame_results = tk.Frame(self.root, bg="#f0f0f0")
        frame_results.pack(pady=10)
        
        result_font = ("Microsoft JhengHei", 9) 
        
        self.lbl_total = tk.Label(frame_results, text="總離境日數：0", bg="#f0f0f0", font=result_font)
        self.lbl_total.pack()
        
        self.lbl_past_365 = tk.Label(frame_results, text="過去365日離境日數：N/A", bg="#f0f0f0", font=result_font)
        self.lbl_past_365.pack() 
        
        self.lbl_max_365 = tk.Label(frame_results, text="任意365日內最多離境：0（剩餘180日）", bg="#f0f0f0", font=result_font)
        self.lbl_max_365.pack()
        
        self.lbl_period = tk.Label(frame_results, text="入籍計算期（+1年至+6年）離境日數：N/A", bg="#f0f0f0", font=result_font)
        self.lbl_period.pack()

        self.lbl_final_year = tk.Label(frame_results, text="最後一年離境日數：N/A", bg="#f0f0f0", font=result_font)
        self.lbl_final_year.pack()

        frame_footer = tk.Frame(self.root, bg="#f0f0f0")
        frame_footer.pack(fill='x', padx=10, pady=(0, 10))
        
        self.lbl_save_date = tk.Label(frame_footer, text="上次儲存日期：未儲存", bg="#f0f0f0", font=("Microsoft JhengHei", 9), fg="#555")
        self.lbl_save_date.pack(side=tk.LEFT)
        
        tk.Button(frame_footer, text="💡 Read Me", command=self.open_readme).pack(side=tk.RIGHT)

    def add_row(self, out_date="", in_date="", activity=""): 
        row = []
        new_row_index = 1 + len(self.rows) 
        
        parent_frame = self.scrollable_frame
        
        entry_out = tk.Entry(parent_frame, width=15, justify="center")
        entry_out.grid(row=new_row_index, column=0, padx=10, pady=(0, 5))
        entry_out.bind("<FocusOut>", lambda e: self.validate_date(entry_out))
        entry_out.bind("<KeyRelease>", self.auto_hyphenate_date)
        entry_out.bind("<KeyRelease>", self.mark_unsaved, add='+')
        entry_out.insert(0, out_date)
        row.append(entry_out)
        
        entry_in = tk.Entry(parent_frame, width=15, justify="center")
        entry_in.grid(row=new_row_index, column=1, padx=10, pady=(0, 5))
        entry_in.bind("<FocusOut>", lambda e: self.validate_date(entry_in))
        entry_in.bind("<KeyRelease>", self.auto_hyphenate_date)
        entry_in.bind("<KeyRelease>", self.mark_unsaved, add='+')
        entry_in.insert(0, in_date)
        row.append(entry_in)

        lbl_365_count = tk.Label(parent_frame, text="-", bg="#f0f0f0", width=10, font=("Microsoft JhengHei", 10))
        lbl_365_count.grid(row=new_row_index, column=2, padx=5, pady=(0, 5))
        row.append(lbl_365_count)
        
        entry_activity = tk.Entry(parent_frame, width=15, justify="left")
        entry_activity.grid(row=new_row_index, column=3, padx=10, pady=(0, 5))
        entry_activity.bind("<KeyRelease>", self.mark_unsaved)
        entry_activity.insert(0, activity)
        row.append(entry_activity) 
        
        var = tk.BooleanVar()
        chk = tk.Checkbutton(parent_frame, variable=var, bg="#f0f0f0", command=self.mark_unsaved)
        chk.grid(row=new_row_index, column=4, padx=5)
        row.append(var) 

        self.rows.append(row)
        
        self.scrollable_frame.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def redraw_rows(self):
        
        parent_frame = self.scrollable_frame
        
        for widget in parent_frame.winfo_children():
            row_info = widget.grid_info().get("row")
            if row_info is not None and row_info >= 1:
                if isinstance(widget, tk.Checkbutton):
                    widget.destroy()
                elif isinstance(widget, tk.Entry) or isinstance(widget, tk.Label): 
                    widget.grid_forget()

        for i, row in enumerate(self.rows):
            current_row_index = 1 + i
            
            row[0].grid(row=current_row_index, column=0, padx=10, pady=(0, 5)) 
            row[1].grid(row=current_row_index, column=1, padx=10, pady=(0, 5)) 
            row[2].grid(row=current_row_index, column=2, padx=5, pady=(0, 5)) 
            row[3].grid(row=current_row_index, column=3, padx=10, pady=(0, 5)) 
            
            chk = tk.Checkbutton(parent_frame, variable=row[4], bg="#f0f0f0", command=self.mark_unsaved)
            chk.grid(row=current_row_index, column=4, padx=5)
            
        parent_frame.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))


    def delete_selected(self):
        new_rows = []
        deleted = False
        for row in self.rows:
            if row[4].get(): 
                for widget in row[:4]: 
                    widget.destroy()
                deleted = True
            else:
                new_rows.append(row)
        
        if deleted:
            self.mark_unsaved()
            
        self.rows = new_rows
        
        self.redraw_rows()
        
        if not self.rows:
            self.add_row()
        
        self.calculate_days()

    def load_data(self):
        if not os.path.exists(DATA_FILE):
            self.add_row()
            return
            
        with open(DATA_FILE, newline='', encoding="utf-8") as f:
            reader = csv.reader(f)
            data = list(reader)
            if not data:
                self.add_row()
                return

            self.entry_approval.delete(0, tk.END)
            self.entry_arrival.delete(0, tk.END)
            self.entry_approval.insert(0, data[0][0])
            self.entry_arrival.insert(0, data[0][1])
            if len(data[0]) > 2:
                self.lbl_save_date.config(text=f"上次儲存日期：{data[0][2]}")

            for row in self.rows:
                pass 
            self.rows = []
            
            for r in data[1:]:
                out_date = r[0] if len(r) > 0 else ""
                in_date = r[1] if len(r) > 1 else ""
                activity = r[2] if len(r) > 2 else "" # Load activity (Still CSV index 2)
                self.add_row(out_date, in_date, activity)
            if not self.rows:
                self.add_row()
                
            self.is_saved = True 

    def save_data(self):
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        rows_data = [[self.entry_approval.get(), self.entry_arrival.get(), now]]
        for row in self.rows:
            out_date = row[0].get().strip()
            in_date = row[1].get().strip()
            activity = row[3].get().strip() # Get activity (Now UI index 3)
            if out_date or in_date:
                rows_data.append([out_date, in_date, activity]) # CSV index 2

        with open(DATA_FILE, "w", newline='', encoding="utf-8") as f:
            csv.writer(f).writerows(rows_data)
            
        self.lbl_save_date.config(text=f"上次儲存日期：{now}")
        self.is_saved = True 
        messagebox.showinfo("儲存完成", "資料已儲存。")
        self.calculate_days()

    def calculate_days(self):
        
        approval = self.parse_date(self.entry_approval.get())
        arrival = self.parse_date(self.entry_arrival.get())
        
        if not approval or not arrival:
            self.lbl_total.config(text="總離境日數：0")
            self.lbl_past_365.config(text="過去365日離境日數：N/A")
            self.lbl_max_365.config(text="任意365日內最多離境：N/A")
            self.lbl_period.config(text=f"入籍計算期（+1年至+6年）離境日數：N/A")
            self.lbl_final_year.config(text="最後一年離境日數：N/A") 
            return

        if arrival < approval:
            messagebox.showwarning("日期錯誤", "到達日必須晚於或等於批核日。")
            return

        trips = []
        for row in self.rows:
            start = self.parse_date(row[0].get())
            end = self.parse_date(row[1].get())
            if start and end:
                if end <= start:
                    messagebox.showwarning("日期錯誤", "回國日必須晚於出國日。")
                    return
                activity = row[3].get() if len(row) > 4 and isinstance(row[3], tk.Entry) else "" 
                trips.append((start, end, activity))
            elif start or end:
                messagebox.showwarning("日期錯誤", "出國日和回國日必須同時填寫。")
                return

        initial_departure_days = set()
        current = approval
        while current < arrival:
            initial_departure_days.add(current)
            current += timedelta(days=1)
        
        travel_departure_days = set()
        for start, end, _ in trips:
            current = start + timedelta(days=1)
            while current < end:
                travel_departure_days.add(current)
                current += timedelta(days=1)
        
        all_departure_days = initial_departure_days.union(travel_departure_days)

        total_departure_days_count = len(all_departure_days)
        self.lbl_total.config(text=f"總離境日數：{total_departure_days_count}")
        
        for row in self.rows:
            entry_in = row[1].get()
            lbl_365_count = row[2] 
            end = self.parse_date(entry_in)
            
            if end: 
                end_of_check = end
                start_of_check = end_of_check - timedelta(days=365) + timedelta(days=1)
                
                count_365 = len([d for d in all_departure_days 
                                 if start_of_check <= d <= end_of_check])
                
                color_fg = COLOR_NORMAL
                if count_365 > 180:
                    color_fg = COLOR_RED
                elif count_365 >= 150:
                    color_fg = COLOR_ORANGE
                
                lbl_365_count.config(text=str(count_365), fg=color_fg)
            else:
                lbl_365_count.config(text="-", fg=COLOR_NORMAL)
        
        today = datetime.now().date()
        past_365_start = today - timedelta(days=365)
        
        past_365_days_count = len([d for d in all_departure_days 
                                   if past_365_start <= d < today])
                                   
        past_365_remain = 180 - past_365_days_count 

        past_365_start_str = past_365_start.strftime('%Y/%#m/%#d')
        today_str = today.strftime('%Y/%#m/%#d')
        
        self.lbl_past_365.config(text=f"過去365日離境日數（{past_365_start_str}–{today_str}）：{past_365_days_count} 日")
        self.color_label(self.lbl_past_365, past_365_remain, [50, 30, 10])

        self.max_365_periods = []
        max_365_count = 0
        
        end_of_check_period = approval + relativedelta(years=5) 
        
        current_date = approval
        while current_date < end_of_check_period:
            start_window = current_date
            end_window = current_date + timedelta(days=365) - timedelta(days=1) 
            
            count = len([d for d in all_departure_days if start_window <= d <= end_window])
            
            if count >= 150: 
                period_str = f"{start_window.strftime('%Y/%#m/%#d')}–{end_window.strftime('%Y/%#m/%#d')}"
                self.max_365_periods.append((count, period_str))
                
            if count > max_365_count:
                max_365_count = count
                
            current_date += timedelta(days=1)
            
        self.max_365_periods.sort(key=lambda x: x[0], reverse=True)
        
        display_periods = []
        closest_to_180_count = -1
        closest_period = "N/A"
        
        for count, period in self.max_365_periods:
            if count > 180:
                if len(display_periods) < 5: 
                    display_periods.append(f"{period}：{count}日 (超額 {count - 180} 日)")
            
            if 180 >= count > closest_to_180_count:
                closest_to_180_count = count
                closest_period = period
        
        if display_periods:
            result_text = "任意365日內最多離境：\n" + "\n".join(display_periods)
            if closest_to_180_count != -1:
                 result_text += f"\n最接近180日（未超額）：{closest_period}：{closest_to_180_count}日"
            remain_max_365 = 180 - max_365_count
            result_text += f"\n最高紀錄：{max_365_count}日 (剩餘 {remain_max_365} 日)"
        elif closest_to_180_count != -1:
            remain_max_365 = 180 - max_365_count
            result_text = f"任意365日內最多離境：\n{closest_period}：{closest_to_180_count}日 (剩餘 {remain_max_365} 日)"
        else:
            remain_max_365 = 180 - max_365_count
            result_text = f"任意365日內最多離境：{max_365_count}日 (剩餘 {remain_max_365} 日)"
            
        self.lbl_max_365.config(text=result_text)
        self.color_label(self.lbl_max_365, remain_max_365, [50, 30, 10])

        naturalisation_application_date = approval + relativedelta(years=6)
        
        ilr_end = naturalisation_application_date - timedelta(days=1)
        
        ilr_start = ilr_end - relativedelta(years=5) + timedelta(days=1) 

        period_days_count = len([d for d in all_departure_days 
                                 if ilr_start <= d <= ilr_end])
        
        remain_long = 450 - period_days_count
        
        ilr_start_str = ilr_start.strftime('%Y/%#m/%#d')
        ilr_end_str = ilr_end.strftime('%Y/%#m/%#d')

        self.lbl_period.config(text=f"入籍計算期（{ilr_start_str}–{ilr_end_str}）離境日數：{period_days_count}（剩餘 {remain_long} 日）")
        self.color_label(self.lbl_period, remain_long, [120, 60, 30])
        
        final_year_end = ilr_end
        final_year_start = ilr_end - relativedelta(years=1) + timedelta(days=1)
        
        final_year_days_count = len([d for d in all_departure_days 
                                     if final_year_start <= d <= final_year_end])
        
        final_year_remain = 90 - final_year_days_count

        final_year_start_str = final_year_start.strftime('%Y/%#m/%#d')
        final_year_end_str = final_year_end.strftime('%Y/%#m/%#d')
        
        self.lbl_final_year.config(text=f"最後一年離境日數（{final_year_start_str}–{final_year_end_str}）：{final_year_days_count}（剩餘 {final_year_remain} 日）")
        self.color_label(self.lbl_final_year, final_year_remain, [30, 15, 5])


if __name__ == "__main__":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass

    root = tk.Tk()
    app = TravelApp(root)
    root.mainloop()
