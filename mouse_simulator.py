import tkinter as tk
from tkinter import ttk, messagebox
import pyautogui
import time
import random
import threading

# 自訂對話框，用於設定目標點的點擊次數與等待秒數
class TargetSettingDialog(tk.Toplevel):
    def __init__(self, master, x, y):
        """
        建構子
        
        :param master: 父視窗
        :param x: 使用者點選的 X 座標
        :param y: 使用者點選的 Y 座標
        """
        super().__init__(master)
        self.x = x
        self.y = y
        self.result = None  # 儲存 (click_count, wait_time)
        self.title(f"設定目標點 ({x}, {y}) 詳細資料")
        self.geometry("300x150")
        
        # 點擊次數輸入
        label_click = tk.Label(self, text="點擊次數：")
        label_click.grid(row=0, column=0, padx=5, pady=5)
        self.entry_click = tk.Entry(self)
        self.entry_click.grid(row=0, column=1, padx=5, pady=5)
        
        # 移動後等待秒數輸入
        label_wait = tk.Label(self, text="等待秒數：")
        label_wait.grid(row=1, column=0, padx=5, pady=5)
        self.entry_wait = tk.Entry(self)
        self.entry_wait.grid(row=1, column=1, padx=5, pady=5)
        
        # 確定與取消按鈕
        btn_ok = tk.Button(self, text="確定", command=self.on_ok)
        btn_ok.grid(row=2, column=0, padx=5, pady=5)
        btn_cancel = tk.Button(self, text="取消", command=self.on_cancel)
        btn_cancel.grid(row=2, column=1, padx=5, pady=5)
        
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.grab_set()  # 封鎖其他視窗操作直到此對話框關閉
        self.wait_window(self)

    def on_ok(self):
        """ 使用者按下確定後，取得輸入值 """
        try:
            click_count = int(self.entry_click.get())
            wait_time = float(self.entry_wait.get())
        except Exception:
            messagebox.showerror("錯誤", "請輸入有效的數值")
            return
        self.result = (click_count, wait_time)
        self.destroy()

    def on_cancel(self):
        """ 使用者取消設定 """
        self.destroy()


# 主 GUI 介面類別
class MouseClickSimulatorGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("滑鼠點擊模擬器設定")
        self.targets = []  # 儲存所有目標點資料，每個項目為 (x, y, click_count, wait_time)
        self.is_running = False  # 控制模擬是否正在運行
        self.is_paused = False   # 控制模擬是否暫停
        self.disable_check = False  # 移動期間暫時停用介入檢查
        self.simulation_thread = None

        # 介面排版設定
        # 將 Treeview 設定為佔據前 6 個欄位
        self.tree = ttk.Treeview(self.root, columns=("x", "y", "clicks", "wait"), show="headings")
        self.tree.heading("x", text="X座標")
        self.tree.heading("y", text="Y座標")
        self.tree.heading("clicks", text="點擊次數")
        self.tree.heading("wait", text="等待秒數")
        self.tree.grid(row=0, column=0, columnspan=6, padx=10, pady=10)
        # 新增「?」按鈕，位於第一列最右側
        self.help_button = tk.Button(self.root, text="?", command=self.show_help)
        self.help_button.grid(row=0, column=6, padx=5, pady=5, sticky="ne")

        # 重複次數設定
        label_repeat = tk.Label(self.root, text="重複執行次數：")
        label_repeat.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.entry_repeat = tk.Entry(self.root)
        self.entry_repeat.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.entry_repeat.insert(0, "1")

        # 操作按鈕：新增目標點、移除選取目標、開始模擬、暫停/繼續模擬、停止模擬
        btn_add = tk.Button(self.root, text="新增目標點", command=self.add_target)
        btn_add.grid(row=2, column=0, padx=5, pady=10)
        btn_remove = tk.Button(self.root, text="移除選取目標", command=self.remove_target)
        btn_remove.grid(row=2, column=1, padx=5, pady=10)
        self.start_button = tk.Button(self.root, text="開始模擬", command=self.start_simulation)
        self.start_button.grid(row=2, column=2, padx=5, pady=10)
        self.pause_button = tk.Button(self.root, text="暫停/繼續模擬", command=self.toggle_pause, state=tk.DISABLED)
        self.pause_button.grid(row=2, column=3, padx=5, pady=10)
        self.stop_button = tk.Button(self.root, text="停止模擬", command=self.stop_simulation, state=tk.DISABLED)
        self.stop_button.grid(row=2, column=4, padx=5, pady=10)

        # 顯示目前模擬執行次數的標籤，跨越全部欄位
        self.iteration_label = tk.Label(self.root, text="目前執行次數：0")
        self.iteration_label.grid(row=3, column=0, columnspan=7, pady=10)

        self.root.mainloop()

    def show_help(self):
        """ 顯示操作說明 """
        help_text = (
            "操作說明:\n\n"
            "1. 啟動程式\n"
            "   - 執行檔案：雙擊產生的 .exe 檔案，即可啟動「滑鼠點擊模擬器設定」的 GUI 介面。\n\n"
            "2. 設定目標點\n"
            "   - 新增目標點：\n"
            "       a. 點選介面上「新增目標點」按鈕。\n"
            "       b. 畫面會出現全螢幕的半透明覆蓋視窗，提示「請點選目標點」。\n"
            "       c. 在你希望模擬點擊的位置，使用滑鼠左鍵點擊。點擊後，覆蓋視窗會自動關閉。\n"
            "       d. 隨後會跳出設定對話框，請依序輸入：\n"
            "            - 點擊次數：在該目標點上要執行的點擊次數。\n"
            "            - 等待秒數：完成該目標點操作後的等待秒數。\n"
            "       e. 點選「確定」後，該目標點會加入下方的列表中。\n\n"
            "   - 移除目標點：\n"
            "       - 若要刪除目標點，請選取列表中的目標點後點選「移除選取目標」按鈕。\n\n"
            "3. 設定執行次數\n"
            "   - 在「重複執行次數」輸入框中，輸入模擬流程重複執行的次數。\n\n"
            "4. 開始模擬\n"
            "   - 點選「開始模擬」按鈕，程式會依照設定的次數，依序對所有目標點執行滑鼠移動、點擊與等待操作。\n"
            "   - 模擬進度會顯示在介面下方（例如：目前執行次數：2 / 5）。\n\n"
            "5. 自動暫停與使用者介入\n"
            "   - 在模擬過程中，若你手動移動滑鼠，使其偏離預期目標超過 20 像素，模擬將自動暫停，並在進度標籤上顯示提示。\n\n"
            "6. 暫停 / 繼續模擬\n"
            "   - 當模擬暫停時，按鈕文字會顯示「繼續模擬」。點選後，程式會先自動將滑鼠移動回目標位置，再繼續模擬。\n\n"
            "7. 停止模擬\n"
            "   - 隨時點選「停止模擬」按鈕可中斷模擬流程。停止後，狀態與按鈕將重置，進度標籤會顯示「模擬已停止」。\n\n"
            "小提示:\n"
            "   - 請在設定目標點時確認位置與參數正確，避免模擬動作錯誤。\n"
            "   - 如需介入或調整，可使用暫停功能暫停模擬，再進行必要操作後繼續。"
        )
        help_window = tk.Toplevel(self.root)
        help_window.title("操作說明")
        help_window.geometry("600x400")

        # 建立捲動式文字區域顯示說明內容
        text_area = tk.Text(help_window, wrap=tk.WORD)
        text_area.insert(tk.END, help_text)
        text_area.config(state=tk.DISABLED)
        text_area.pack(expand=True, fill=tk.BOTH)

        scrollbar = tk.Scrollbar(text_area)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_area.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=text_area.yview)

    def add_target(self):
        """ 使用者按下 '新增目標點' 按鈕後，顯示全螢幕覆蓋視窗捕捉點選位置 """
        overlay = tk.Toplevel(self.root)
        overlay.attributes('-fullscreen', True)
        overlay.attributes('-topmost', True)
        overlay.attributes("-alpha", 0.3)  # 設定視窗透明度 30%
        overlay.config(bg='gray')
        label = tk.Label(overlay, text="請點選目標點", font=("Arial", 24), bg="gray")
        label.pack(expand=True)
        overlay.bind("<Button-1>", lambda event: self.on_overlay_click(event, overlay))

    def on_overlay_click(self, event, overlay):
        """
        捕捉使用者點選的座標後，關閉覆蓋視窗並進入目標點詳細設定
        """
        x, y = event.x_root, event.y_root
        overlay.destroy()
        self.prompt_target_details(x, y)

    def prompt_target_details(self, x, y):
        """
        顯示對話框讓使用者輸入該目標點的點擊次數與等待秒數，
        並將資料儲存到目標點列表中。
        """
        dialog = TargetSettingDialog(self.root, x, y)
        if dialog.result:
            click_count, wait_time = dialog.result
            self.targets.append((x, y, click_count, wait_time))
            self.update_treeview()

    def update_treeview(self):
        """ 更新目標點列表視圖 """
        for item in self.tree.get_children():
            self.tree.delete(item)
        for target in self.targets:
            x, y, clicks, wait_time = target
            self.tree.insert("", "end", values=(x, y, clicks, wait_time))

    def remove_target(self):
        """ 移除使用者在列表中選取的目標點 """
        selected = self.tree.selection()
        if not selected:
            messagebox.showerror("錯誤", "請選取要移除的目標點")
            return
        for sel in selected:
            values = self.tree.item(sel, "values")
            try:
                x, y, clicks, wait_time = values
                target = (int(x), int(y), int(clicks), float(wait_time))
                if target in self.targets:
                    self.targets.remove(target)
            except Exception:
                continue
            self.tree.delete(sel)

    def start_simulation(self):
        """ 開始執行滑鼠點擊模擬 """
        try:
            repeat = int(self.entry_repeat.get())
        except Exception:
            messagebox.showerror("錯誤", "請輸入有效的重複次數")
            return
        if not self.targets:
            messagebox.showerror("錯誤", "請新增至少一個目標點")
            return

        self.is_running = True
        self.is_paused = False

        self.start_button.config(state=tk.DISABLED)
        self.pause_button.config(state=tk.NORMAL, text="暫停模擬")
        self.stop_button.config(state=tk.NORMAL)

        # 以新執行緒執行模擬，避免 GUI 凍結
        self.simulation_thread = threading.Thread(target=self.run_simulation, args=(repeat,))
        self.simulation_thread.start()

    def toggle_pause(self):
        """ 切換暫停與繼續模擬 """
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.pause_button.config(text="繼續模擬")
            self.iteration_label.config(text="已暫停")
        else:
            self.pause_button.config(text="暫停模擬")
            # 使用者點選繼續後，等待內部偵測到 is_paused 為 False，
            # 內部會先移動滑鼠至目前任務目標位置再繼續執行

    def stop_simulation(self):
        """ 停止模擬 """
        self.is_running = False
        self.is_paused = False
        self.start_button.config(state=tk.NORMAL)
        self.pause_button.config(state=tk.DISABLED, text="暫停模擬")
        self.stop_button.config(state=tk.DISABLED)
        self.iteration_label.config(text="模擬已停止")

    def wait_until_resumed(self, target_x, target_y):
        """
        當模擬處於暫停狀態時，等待使用者點選繼續，
        並在繼續前先將滑鼠移動到目標位置，
        移動期間暫時禁用介入檢查。
        """
        while self.is_paused:
            time.sleep(0.1)
        # 使用者已點擊繼續，先移動滑鼠到目標位置
        self.disable_check = True
        pyautogui.moveTo(target_x, target_y, duration=0.5)
        self.disable_check = False

    def check_intervention(self, expected_x, expected_y):
        """
        檢查目前滑鼠位置是否與預期目標點偏離超過閾值，
        若是則自動觸發暫停模擬功能
        """
        if self.disable_check:
            return
        current_x, current_y = pyautogui.position()
        threshold = 20  # 閾值 (像素)
        if abs(current_x - expected_x) > threshold or abs(current_y - expected_y) > threshold:
            if not self.is_paused:
                self.is_paused = True
                self.root.after(0, lambda: self.pause_button.config(text="繼續模擬"))
                self.root.after(0, lambda: self.iteration_label.config(text="已暫停：偵測到使用者移動滑鼠"))

    def safe_sleep(self, duration, expected_x, expected_y):
        """
        以小間隔的 sleep 方式等待，並持續檢查使用者是否移動滑鼠
        :param duration: 等待秒數
        :param expected_x: 預期的 X 座標
        :param expected_y: 預期的 Y 座標
        """
        start_time = time.time()
        while time.time() - start_time < duration:
            if not self.is_running:
                break
            if self.is_paused:
                self.wait_until_resumed(expected_x, expected_y)
            self.check_intervention(expected_x, expected_y)
            time.sleep(0.1)

    def run_simulation(self, repeat):
        """
        執行模擬動作：依照設定的重複次數執行所有目標點動作，
        並在介面上更新目前執行到第幾次。
        """
        for iteration in range(1, repeat + 1):
            if not self.is_running:
                break
            self.root.after(0, lambda i=iteration, r=repeat: self.iteration_label.config(text=f"目前執行次數：{i} / {r}"))
            for target in self.targets:
                if not self.is_running:
                    break
                x, y, click_count, wait_time = target
                # 若目前暫停，等待使用者繼續並移動至目標位置
                if self.is_paused:
                    self.wait_until_resumed(x, y)
                # 移動至目標點（移動期間禁用介入檢查）
                self.disable_check = True
                pyautogui.moveTo(x, y, duration=0.5)
                self.disable_check = False
                self.check_intervention(x, y)
                for _ in range(click_count):
                    if not self.is_running:
                        break
                    if self.is_paused:
                        self.wait_until_resumed(x, y)
                    self.check_intervention(x, y)
                    pyautogui.click()
                    interval = random.uniform(0.2, 0.5)
                    self.safe_sleep(interval, x, y)
                self.safe_sleep(wait_time, x, y)
        self.is_running = False
        self.root.after(0, lambda: self.start_button.config(state=tk.NORMAL))
        self.root.after(0, lambda: self.pause_button.config(state=tk.DISABLED, text="暫停模擬"))
        self.root.after(0, lambda: self.stop_button.config(state=tk.DISABLED))
        self.root.after(0, lambda: self.iteration_label.config(text="模擬完成"))

if __name__ == "__main__":
    MouseClickSimulatorGUI()
