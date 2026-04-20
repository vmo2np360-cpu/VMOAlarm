"""
VMO鬧鐘
功能：
- 開機自動啟動（可選）
- 支援工作日、假日、每天、指定日期鬧鐘
- 內置香港公眾假期（2024-2028），使用者可手動增刪
- 鬧鐘觸發時彈出醒目視窗，必須手動確認
- 關閉主視窗即縮小到工作列，不會退出
"""

import sqlite3
import datetime
import os
import sys
import winsound
from tkinter import *
from tkinter import ttk, messagebox, simpledialog

# ==================== 香港公眾假期數據（2024-2028）====================
HONGKONG_HOLIDAYS_2024_2028 = {
    # 2024
    "2024-01-01": "元旦", "2024-02-10": "農曆年初一", "2024-02-11": "農曆年初二",
    "2024-02-12": "農曆年初三", "2024-03-29": "耶穌受難節", "2024-03-30": "耶穌受難節翌日",
    "2024-04-01": "復活節星期一", "2024-04-04": "清明節", "2024-05-01": "勞動節",
    "2024-05-15": "佛誕", "2024-06-10": "端午節", "2024-07-01": "香港特別行政區成立紀念日",
    "2024-09-18": "中秋節翌日", "2024-10-01": "國慶日", "2024-10-11": "重陽節",
    "2024-12-25": "聖誕節", "2024-12-26": "聖誕節後第一個周日",
    # 2025
    "2025-01-01": "元旦", "2025-01-29": "農曆年初一", "2025-01-30": "農曆年初二",
    "2025-01-31": "農曆年初三", "2025-04-04": "清明節", "2025-04-18": "耶穌受難節",
    "2025-04-19": "耶穌受難節翌日", "2025-04-21": "復活節星期一", "2025-05-01": "勞動節",
    "2025-05-05": "佛誕", "2025-05-31": "端午節", "2025-07-01": "香港特別行政區成立紀念日",
    "2025-10-01": "國慶日", "2025-10-07": "中秋節翌日", "2025-10-29": "重陽節",
    "2025-12-25": "聖誕節", "2025-12-26": "聖誕節後第一個周日",
    # 2026
    "2026-01-01": "元旦", "2026-02-17": "農曆年初一", "2026-02-18": "農曆年初二",
    "2026-02-19": "農曆年初三", "2026-04-03": "耶穌受難節", "2026-04-04": "耶穌受難節翌日",
    "2026-04-06": "復活節星期一", "2026-04-05": "清明節", "2026-05-01": "勞動節",
    "2026-05-24": "佛誕", "2026-06-19": "端午節", "2026-07-01": "香港特別行政區成立紀念日",
    "2026-09-18": "中秋節翌日", "2026-10-01": "國慶日", "2026-10-19": "重陽節",
    "2026-12-25": "聖誕節", "2026-12-26": "聖誕節後第一個周日",
    # 2027
    "2027-01-01": "元旦", "2027-02-06": "農曆年初一", "2027-02-07": "農曆年初二",
    "2027-02-08": "農曆年初三", "2027-03-26": "耶穌受難節", "2027-03-27": "耶穌受難節翌日",
    "2027-03-29": "復活節星期一", "2027-04-05": "清明節", "2027-05-01": "勞動節",
    "2027-05-12": "佛誕", "2027-05-31": "端午節", "2027-07-01": "香港特別行政區成立紀念日",
    "2027-09-18": "中秋節翌日", "2027-10-01": "國慶日", "2027-10-19": "重陽節",
    "2027-12-25": "聖誕節", "2027-12-26": "聖誕節後第一個周日",
    # 2028
    "2028-01-01": "元旦", "2028-01-26": "農曆年初一", "2028-01-27": "農曆年初二",
    "2028-01-28": "農曆年初三", "2028-04-14": "耶穌受難節", "2028-04-15": "耶穌受難節翌日",
    "2028-04-17": "復活節星期一", "2028-04-04": "清明節", "2028-05-01": "勞動節",
    "2028-05-03": "佛誕", "2028-05-30": "端午節", "2028-07-01": "香港特別行政區成立紀念日",
    "2028-09-18": "中秋節翌日", "2028-10-01": "國慶日", "2028-10-17": "重陽節",
    "2028-12-25": "聖誕節", "2028-12-26": "聖誕節後第一個周日",
}

# ==================== 資料庫操作 ====================
class Database:
    def __init__(self, db_path="alarm_data.db"):
        if getattr(sys, 'frozen', False):
            self.db_path = os.path.join(os.path.dirname(sys.executable), db_path)
        else:
            self.db_path = db_path
        self.init_db()
        self.migrate_db()

    def init_db(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS alarms (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                time TEXT NOT NULL,
                content TEXT NOT NULL,
                type TEXT NOT NULL CHECK(type IN ('workday', 'holiday', 'everyday', 'specific')),
                specific_date TEXT
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS holidays (
                date TEXT PRIMARY KEY,
                name TEXT NOT NULL
            )
        ''')
        conn.commit()
        conn.close()

    def migrate_db(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(alarms)")
        columns = [col[1] for col in cursor.fetchall()]
        if 'specific_date' not in columns:
            cursor.execute("ALTER TABLE alarms ADD COLUMN specific_date TEXT")
            conn.commit()
        conn.close()

    def get_alarms(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT id, time, content, type, specific_date FROM alarms ORDER BY time")
        alarms = cursor.fetchall()
        conn.close()
        return alarms

    def add_alarm(self, time_str, content, alarm_type, specific_date=None):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO alarms (time, content, type, specific_date) VALUES (?, ?, ?, ?)",
            (time_str, content, alarm_type, specific_date)
        )
        conn.commit()
        conn.close()

    def update_alarm(self, alarm_id, time_str, content, alarm_type, specific_date=None):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE alarms SET time=?, content=?, type=?, specific_date=? WHERE id=?",
            (time_str, content, alarm_type, specific_date, alarm_id)
        )
        conn.commit()
        conn.close()

    def delete_alarm(self, alarm_id):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM alarms WHERE id=?", (alarm_id,))
        conn.commit()
        conn.close()

    def get_holidays(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT date, name FROM holidays ORDER BY date")
        holidays = cursor.fetchall()
        conn.close()
        return holidays

    def add_holiday(self, date_str, name):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("INSERT INTO holidays (date, name) VALUES (?, ?)", (date_str, name))
            conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False
        finally:
            conn.close()

    def delete_holiday(self, date_str):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM holidays WHERE date=?", (date_str,))
        conn.commit()
        conn.close()

    def restore_default_holidays(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM holidays")
        for date_str, name in HONGKONG_HOLIDAYS_2024_2028.items():
            cursor.execute("INSERT OR IGNORE INTO holidays (date, name) VALUES (?, ?)", (date_str, name))
        conn.commit()
        conn.close()

    def is_holiday(self, date_obj):
        if date_obj.weekday() >= 5:
            return True
        date_str = date_obj.strftime("%Y-%m-%d")
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM holidays WHERE date=?", (date_str,))
        exists = cursor.fetchone() is not None
        conn.close()
        return exists

    def is_workday(self, date_obj):
        return date_obj.weekday() < 5 and not self.is_holiday(date_obj)

# ==================== 鬧鐘彈出視窗（必須手動確認）====================
class AlarmPopup(Toplevel):
    def __init__(self, master, title, message):
        super().__init__(master)
        self.title("鬧鐘提醒")
        self.geometry("400x250")
        self.resizable(False, False)
        self.attributes('-topmost', True)
        self.focus_force()
        self.configure(bg='#FF4444')
        main_frame = Frame(self, bg='#FF4444')
        main_frame.pack(expand=True, fill=BOTH, padx=20, pady=20)
        title_label = Label(main_frame, text=title, font=('微軟正黑體', 16, 'bold'),
                           fg='white', bg='#FF4444')
        title_label.pack(pady=(0, 20))
        content_label = Label(main_frame, text=message, font=('微軟正黑體', 14),
                             fg='white', bg='#FF4444', wraplength=350)
        content_label.pack(pady=(0, 30))
        ok_btn = Button(main_frame, text="確認", font=('微軟正黑體', 12),
                       command=self.destroy, bg='white', fg='#FF4444',
                       activebackground='#DDDDDD', width=10)
        ok_btn.pack()
        winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
        self.bind('<Escape>', lambda e: None)

# ==================== 鬧鐘管理器主視窗 ====================
class AlarmManager:
    def __init__(self, root, db):
        self.root = root
        self.db = db
        self.triggered_today = {}
        self.root.title("VMO鬧鐘")
        self.root.geometry("700x450")
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)
        self.create_widgets()
        self.start_check_loop()

    def hide_window(self):
        self.root.withdraw()

    def show_window(self):
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def create_widgets(self):
        toolbar = Frame(self.root)
        toolbar.pack(fill=X, padx=5, pady=5)
        Button(toolbar, text="新增鬧鐘", command=self.add_alarm).pack(side=LEFT, padx=2)
        Button(toolbar, text="編輯鬧鐘", command=self.edit_alarm).pack(side=LEFT, padx=2)
        Button(toolbar, text="刪除鬧鐘", command=self.delete_alarm).pack(side=LEFT, padx=2)
        Button(toolbar, text="假期管理", command=self.manage_holidays).pack(side=LEFT, padx=2)
        Button(toolbar, text="開機自啟設定", command=self.toggle_autostart).pack(side=LEFT, padx=2)
        Button(toolbar, text="顯示視窗", command=self.show_window).pack(side=LEFT, padx=2)

        columns = ("ID", "時間", "內容", "類型", "指定日期")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings")
        self.tree.heading("ID", text="ID")
        self.tree.heading("時間", text="時間")
        self.tree.heading("內容", text="內容")
        self.tree.heading("類型", text="類型")
        self.tree.heading("指定日期", text="指定日期")
        self.tree.column("ID", width=50)
        self.tree.column("時間", width=80)
        self.tree.column("內容", width=300)
        self.tree.column("類型", width=100)
        self.tree.column("指定日期", width=120)

        scrollbar = Scrollbar(self.root, orient=VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.tree.pack(fill=BOTH, expand=True, padx=5, pady=5)

        self.refresh_alarm_list()

    def refresh_alarm_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        type_display = {'workday': '工作日', 'holiday': '假日', 'everyday': '每天', 'specific': '指定日期'}
        for alarm in self.db.get_alarms():
            alarm_id, time_str, content, alarm_type, specific_date = alarm
            display_type = type_display.get(alarm_type, alarm_type)
            date_display = specific_date if specific_date else ''
            self.tree.insert("", END, values=(alarm_id, time_str, content, display_type, date_display))

    def add_alarm(self):
        dialog = Toplevel(self.root)
        dialog.title("新增鬧鐘")
        dialog.geometry("400x350")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        Label(dialog, text="時間 (HH:MM):").pack(pady=(10,0))
        time_entry = Entry(dialog)
        time_entry.pack(pady=5)
        time_entry.insert(0, "09:00")
        Label(dialog, text="提醒內容:").pack(pady=(10,0))
        content_entry = Entry(dialog, width=40)
        content_entry.pack(pady=5)
        content_entry.insert(0, "提醒事項")
        Label(dialog, text="生效類型:").pack(pady=(10,0))
        type_var = StringVar(value="workday")
        def on_type_change(*args):
            if type_var.get() == 'specific':
                date_frame.pack(pady=5, fill=X)
            else:
                date_frame.pack_forget()
        type_var.trace('w', on_type_change)
        radio_frame = Frame(dialog)
        radio_frame.pack()
        ttk.Radiobutton(radio_frame, text="工作日", variable=type_var, value="workday").pack(anchor=W)
        ttk.Radiobutton(radio_frame, text="假日", variable=type_var, value="holiday").pack(anchor=W)
        ttk.Radiobutton(radio_frame, text="每天", variable=type_var, value="everyday").pack(anchor=W)
        ttk.Radiobutton(radio_frame, text="指定日期", variable=type_var, value="specific").pack(anchor=W)
        date_frame = Frame(dialog)
        Label(date_frame, text="日期 (YYYY-MM-DD):").pack(side=LEFT, padx=5)
        specific_date_entry = Entry(date_frame, width=15)
        specific_date_entry.pack(side=LEFT)
        def save():
            time_str = time_entry.get().strip()
            content = content_entry.get().strip()
            alarm_type = type_var.get()
            specific_date = None
            if not time_str or not content:
                messagebox.showerror("錯誤", "時間和內容不能為空")
                return
            try:
                datetime.datetime.strptime(time_str, "%H:%M")
            except ValueError:
                messagebox.showerror("錯誤", "時間格式錯誤，請使用 HH:MM 格式")
                return
            if alarm_type == 'specific':
                specific_date = specific_date_entry.get().strip()
                if not specific_date:
                    messagebox.showerror("錯誤", "請填寫指定日期")
                    return
                try:
                    datetime.datetime.strptime(specific_date, "%Y-%m-%d")
                except ValueError:
                    messagebox.showerror("錯誤", "日期格式錯誤，請使用 YYYY-MM-DD 格式")
                    return
            self.db.add_alarm(time_str, content, alarm_type, specific_date)
            self.refresh_alarm_list()
            dialog.destroy()
        Button(dialog, text="儲存", command=save).pack(pady=20)

    def edit_alarm(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("提示", "請先選擇要編輯的鬧鐘")
            return
        item = self.tree.item(selected[0])
        alarm_id = item['values'][0]
        current_time = item['values'][1]
        current_content = item['values'][2]
        current_type_display = item['values'][3]
        current_specific_date = item['values'][4] if len(item['values']) > 4 else ''
        type_map = {'工作日': 'workday', '假日': 'holiday', '每天': 'everyday', '指定日期': 'specific'}
        current_type = type_map.get(current_type_display, 'workday')
        dialog = Toplevel(self.root)
        dialog.title("編輯鬧鐘")
        dialog.geometry("400x350")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        Label(dialog, text="時間 (HH:MM):").pack(pady=(10,0))
        time_entry = Entry(dialog)
        time_entry.pack(pady=5)
        time_entry.insert(0, current_time)
        Label(dialog, text="提醒內容:").pack(pady=(10,0))
        content_entry = Entry(dialog, width=40)
        content_entry.pack(pady=5)
        content_entry.insert(0, current_content)
        Label(dialog, text="生效類型:").pack(pady=(10,0))
        type_var = StringVar(value=current_type)
        def on_type_change(*args):
            if type_var.get() == 'specific':
                date_frame.pack(pady=5, fill=X)
                specific_date_entry.delete(0, END)
                specific_date_entry.insert(0, current_specific_date)
            else:
                date_frame.pack_forget()
        type_var.trace('w', on_type_change)
        radio_frame = Frame(dialog)
        radio_frame.pack()
        ttk.Radiobutton(radio_frame, text="工作日", variable=type_var, value="workday").pack(anchor=W)
        ttk.Radiobutton(radio_frame, text="假日", variable=type_var, value="holiday").pack(anchor=W)
        ttk.Radiobutton(radio_frame, text="每天", variable=type_var, value="everyday").pack(anchor=W)
        ttk.Radiobutton(radio_frame, text="指定日期", variable=type_var, value="specific").pack(anchor=W)
        date_frame = Frame(dialog)
        Label(date_frame, text="日期 (YYYY-MM-DD):").pack(side=LEFT, padx=5)
        specific_date_entry = Entry(date_frame, width=15)
        specific_date_entry.pack(side=LEFT)
        if current_type == 'specific':
            date_frame.pack(pady=5, fill=X)
            specific_date_entry.insert(0, current_specific_date)
        def save():
            time_str = time_entry.get().strip()
            content = content_entry.get().strip()
            alarm_type = type_var.get()
            specific_date = None
            if not time_str or not content:
                messagebox.showerror("錯誤", "時間和內容不能為空")
                return
            try:
                datetime.datetime.strptime(time_str, "%H:%M")
            except ValueError:
                messagebox.showerror("錯誤", "時間格式錯誤")
                return
            if alarm_type == 'specific':
                specific_date = specific_date_entry.get().strip()
                if not specific_date:
                    messagebox.showerror("錯誤", "請填寫指定日期")
                    return
                try:
                    datetime.datetime.strptime(specific_date, "%Y-%m-%d")
                except ValueError:
                    messagebox.showerror("錯誤", "日期格式錯誤")
                    return
            self.db.update_alarm(alarm_id, time_str, content, alarm_type, specific_date)
            self.refresh_alarm_list()
            dialog.destroy()
        Button(dialog, text="儲存", command=save).pack(pady=20)

    def delete_alarm(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("提示", "請先選擇要刪除的鬧鐘")
            return
        if messagebox.askyesno("確認", "確定要刪除這個鬧鐘嗎？"):
            item = self.tree.item(selected[0])
            alarm_id = item['values'][0]
            self.db.delete_alarm(alarm_id)
            self.refresh_alarm_list()

    def manage_holidays(self):
        holiday_win = Toplevel(self.root)
        holiday_win.title("假期管理")
        holiday_win.geometry("550x450")
        holiday_win.transient(self.root)
        columns = ("日期", "名稱")
        tree = ttk.Treeview(holiday_win, columns=columns, show="headings")
        tree.heading("日期", text="日期")
        tree.heading("名稱", text="名稱")
        tree.pack(fill=BOTH, expand=True, padx=5, pady=5)
        def refresh_holidays():
            for item in tree.get_children():
                tree.delete(item)
            for date_str, name in self.db.get_holidays():
                tree.insert("", END, values=(date_str, name))
        refresh_holidays()
        btn_frame = Frame(holiday_win)
        btn_frame.pack(fill=X, padx=5, pady=5)
        def add_holiday():
            date_str = simpledialog.askstring("新增假期", "請輸入日期 (YYYY-MM-DD):", parent=holiday_win)
            if date_str:
                try:
                    datetime.datetime.strptime(date_str, "%Y-%m-%d")
                except ValueError:
                    messagebox.showerror("錯誤", "日期格式錯誤")
                    return
                name = simpledialog.askstring("新增假期", "請輸入假期名稱:", parent=holiday_win)
                if name:
                    if self.db.add_holiday(date_str, name):
                        refresh_holidays()
                    else:
                        messagebox.showerror("錯誤", "該日期已存在")
        def delete_holiday():
            selected = tree.selection()
            if not selected:
                return
            item = tree.item(selected[0])
            date_str = item['values'][0]
            if messagebox.askyesno("確認", f"確定要刪除假期 {date_str} 嗎？"):
                self.db.delete_holiday(date_str)
                refresh_holidays()
        def restore_default():
            if messagebox.askyesno("確認", "恢復預設假期將清除所有手動修改的假期，並恢復內置香港法定假期（2024-2028）。\n確定要繼續嗎？"):
                self.db.restore_default_holidays()
                refresh_holidays()
                messagebox.showinfo("成功", "已恢復預設假期")
        Button(btn_frame, text="新增假期", command=add_holiday).pack(side=LEFT, padx=2)
        Button(btn_frame, text="刪除假期", command=delete_holiday).pack(side=LEFT, padx=2)
        Button(btn_frame, text="恢復預設假期", command=restore_default).pack(side=LEFT, padx=2)
        Button(btn_frame, text="關閉", command=holiday_win.destroy).pack(side=RIGHT, padx=2)
        Label(holiday_win, text="提示：週六、週日自動視為假日，無需新增。\n恢復預設假期會覆蓋所有當前假期資料。",
              fg="gray", font=("微軟正黑體", 9)).pack(pady=5)

    def toggle_autostart(self):
        """設定開機自動啟動（透過登錄檔）"""
        try:
            import winreg
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            app_name = "VMOAlarm"
            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
            else:
                exe_path = os.path.abspath(__file__)
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ | winreg.KEY_WRITE)
            try:
                winreg.QueryValueEx(key, app_name)
                winreg.DeleteValue(key, app_name)
                messagebox.showinfo("提示", "已關閉開機自動啟動")
            except FileNotFoundError:
                winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, exe_path)
                messagebox.showinfo("提示", "已開啟開機自動啟動")
            winreg.CloseKey(key)
        except Exception as e:
            messagebox.showerror("錯誤", f"設定失敗：{str(e)}")

    def start_check_loop(self):
        self.check_alarms()
        self.root.after(30000, self.start_check_loop)

    def check_alarms(self):
        now = datetime.datetime.now()
        current_time = now.strftime("%H:%M")
        today = now.date()
        for alarm_id, last_date in list(self.triggered_today.items()):
            if last_date != today:
                del self.triggered_today[alarm_id]
        alarms = self.db.get_alarms()
        for alarm in alarms:
            alarm_id, alarm_time, content, alarm_type, specific_date = alarm
            if alarm_time != current_time:
                continue
            if alarm_id in self.triggered_today:
                continue
            should_trigger = False
            if alarm_type == 'everyday':
                should_trigger = True
            elif alarm_type == 'workday':
                if self.db.is_workday(today):
                    should_trigger = True
            elif alarm_type == 'holiday':
                if self.db.is_holiday(today):
                    should_trigger = True
            elif alarm_type == 'specific':
                if specific_date == today.strftime("%Y-%m-%d"):
                    should_trigger = True
            if should_trigger:
                self.triggered_today[alarm_id] = today
                self.root.after(100, lambda c=content: AlarmPopup(self.root, "鬧鐘提醒", c))

# ==================== 主程式 ====================
def main():
    try:
        db = Database()
        root = Tk()
        alarm_manager = AlarmManager(root, db)
        root.mainloop()
    except Exception as e:
        with open("alarm_error.log", "w") as f:
            f.write(str(e))

if __name__ == "__main__":
    main()