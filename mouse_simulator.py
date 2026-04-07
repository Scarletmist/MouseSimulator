import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyautogui
import time
import random
import threading
import json

# ── 選用套件 ─────────────────────────────────────────────────
try:
    from screeninfo import get_monitors
    HAS_SCREENINFO = True
except ImportError:
    HAS_SCREENINFO = False

# ── 點擊類型對照表 ────────────────────────────────────────────
CLICK_TYPES_DISPLAY  = ["左鍵單擊", "左鍵雙擊", "右鍵單擊"]
CLICK_TYPE_TO_KEY    = {"左鍵單擊": "left", "左鍵雙擊": "double", "右鍵單擊": "right"}
CLICK_KEY_TO_DISPLAY = {v: k for k, v in CLICK_TYPE_TO_KEY.items()}


# ══════════════════════════════════════════════════════════════
#  目標點設定 / 編輯對話框
# ══════════════════════════════════════════════════════════════
class TargetSettingDialog(tk.Toplevel):
    def __init__(self, master, x, y, prefill=None):
        """
        :param prefill: (click_count, wait_time, click_type_key)  ← 編輯模式時傳入
        """
        super().__init__(master)
        self.result = None
        self.title(f"{'編輯' if prefill else '設定'}目標點 ({x}, {y})")
        self.geometry("300x200")
        self.resizable(False, False)

        # 點擊次數
        tk.Label(self, text="點擊次數：").grid(row=0, column=0, padx=8, pady=6, sticky="e")
        self.entry_click = tk.Entry(self, width=12)
        self.entry_click.grid(row=0, column=1, padx=8, pady=6)

        # 等待秒數
        tk.Label(self, text="等待秒數：").grid(row=1, column=0, padx=8, pady=6, sticky="e")
        self.entry_wait = tk.Entry(self, width=12)
        self.entry_wait.grid(row=1, column=1, padx=8, pady=6)

        # ── 功能 4：點擊類型 ──────────────────────────────────
        tk.Label(self, text="點擊類型：").grid(row=2, column=0, padx=8, pady=6, sticky="e")
        self.click_type_var = tk.StringVar(value="左鍵單擊")
        self.combo_type = ttk.Combobox(
            self, textvariable=self.click_type_var,
            values=CLICK_TYPES_DISPLAY, state="readonly", width=10
        )
        self.combo_type.grid(row=2, column=1, padx=8, pady=6, sticky="w")

        # 預填（編輯模式）
        if prefill:
            self.entry_click.insert(0, str(prefill[0]))
            self.entry_wait.insert(0, str(prefill[1]))
            self.click_type_var.set(CLICK_KEY_TO_DISPLAY.get(prefill[2], "左鍵單擊"))

        tk.Button(self, text="確定", command=self.on_ok).grid(row=3, column=0, padx=8, pady=8)
        tk.Button(self, text="取消", command=self.on_cancel).grid(row=3, column=1, padx=8, pady=8)

        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.grab_set()
        self.wait_window(self)

    def on_ok(self):
        try:
            click_count = int(self.entry_click.get())
            wait_time   = float(self.entry_wait.get())
            if click_count <= 0 or wait_time < 0:
                raise ValueError
        except Exception:
            messagebox.showerror("錯誤", "點擊次數需為正整數，等待秒數需 ≥ 0")
            return
        self.result = (click_count, wait_time, CLICK_TYPE_TO_KEY[self.click_type_var.get()])
        self.destroy()

    def on_cancel(self):
        self.destroy()


# ══════════════════════════════════════════════════════════════
#  主介面
# ══════════════════════════════════════════════════════════════
class MouseClickSimulatorGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("滑鼠點擊模擬器設定")

        # targets 每筆：(x, y, click_count, wait_time, click_type_key)
        self.targets           = []
        self.is_running        = False
        self.is_paused         = False
        self.disable_check     = False
        self.resume_skip       = False
        self.simulation_thread = None

        # 拖曳排序暫存
        self._drag_item        = None
        self._drag_last_target = None

        self._build_ui()
        self.root.mainloop()

    # ──────────────────────────────────────────────────────────
    #  UI 建構
    # ──────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Treeview（含「點擊類型」欄）────────────────────────
        cols = ("x", "y", "clicks", "wait", "type")
        self.tree = ttk.Treeview(self.root, columns=cols, show="headings", height=10)
        for col, text, width in [
            ("x",      "X 座標",   75),
            ("y",      "Y 座標",   75),
            ("clicks", "點擊次數",  75),
            ("wait",   "等待秒數",  75),
            ("type",   "點擊類型",  90),
        ]:
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor="center")
        self.tree.grid(row=0, column=0, columnspan=7, padx=10, pady=10, sticky="nsew")

        # ── 功能 3：高亮標籤 ────────────────────────────────────
        self.tree.tag_configure("running", background="#b8e0b8", foreground="#1a5c1a")

        # 雙擊編輯
        self.tree.bind("<Double-1>",        self._on_tree_double_click)
        # ── 功能 2：拖曳排序 ────────────────────────────────────
        self.tree.bind("<ButtonPress-1>",   self._drag_start)
        self.tree.bind("<B1-Motion>",       self._drag_motion)
        self.tree.bind("<ButtonRelease-1>", self._drag_release)

        # 說明按鈕
        tk.Button(self.root, text="?", command=self.show_help).grid(
            row=0, column=7, padx=5, pady=5, sticky="ne"
        )

        # ── 第 1 列：重複次數 ＋ 移動時間範圍 ──────────────────
        tk.Label(self.root, text="重複執行次數：").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.entry_repeat = tk.Entry(self.root, width=6)
        self.entry_repeat.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.entry_repeat.insert(0, "1")

        # ── 功能 5：移動時間隨機範圍 ────────────────────────────
        tk.Label(self.root, text="移動時間（秒）：").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.entry_move_min = tk.Entry(self.root, width=5)
        self.entry_move_min.grid(row=1, column=3, padx=2, pady=5, sticky="w")
        self.entry_move_min.insert(0, "0.3")
        tk.Label(self.root, text="～").grid(row=1, column=4, padx=0)
        self.entry_move_max = tk.Entry(self.root, width=5)
        self.entry_move_max.grid(row=1, column=5, padx=2, pady=5, sticky="w")
        self.entry_move_max.insert(0, "0.7")

        # ── 第 2 列：操作按鈕 ──────────────────────────────────
        for text, cmd, col in [
            ("新增目標點",   self.add_target,        0),
            ("移除選取目標", self.remove_target,      1),
            ("儲存設定",     self.save_config,        2),
            ("載入設定",     self.load_config,        3),
        ]:
            tk.Button(self.root, text=text, command=cmd).grid(row=2, column=col, padx=4, pady=8)

        self.start_button = tk.Button(self.root, text="開始模擬", command=self.start_simulation)
        self.start_button.grid(row=2, column=4, padx=4, pady=8)
        self.pause_button = tk.Button(self.root, text="暫停模擬", command=self.toggle_pause, state=tk.DISABLED)
        self.pause_button.grid(row=2, column=5, padx=4, pady=8)
        self.stop_button  = tk.Button(self.root, text="停止模擬", command=self.stop_simulation, state=tk.DISABLED)
        self.stop_button.grid(row=2, column=6, padx=4, pady=8)

        # 狀態列
        self.iteration_label = tk.Label(self.root, text="目前執行次數：0")
        self.iteration_label.grid(row=3, column=0, columnspan=8, pady=8)

    # ──────────────────────────────────────────────────────────
    #  功能 2：拖曳排序
    # ──────────────────────────────────────────────────────────
    def _drag_start(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self._drag_item        = item
            self._drag_last_target = item

    def _drag_motion(self, event):
        if not self._drag_item:
            return
        target = self.tree.identify_row(event.y)
        if not target or target == self._drag_item:
            return
        src_idx = self.tree.index(self._drag_item)
        tgt_idx = self.tree.index(target)
        # 交換 self.targets
        self.targets[src_idx], self.targets[tgt_idx] = \
            self.targets[tgt_idx], self.targets[src_idx]
        # 移動 Treeview 列
        self.tree.move(self._drag_item, "", tgt_idx)
        self._drag_last_target = target

    def _drag_release(self, event):
        self.tree.selection_remove(self.tree.selection())
        self._drag_item        = None
        self._drag_last_target = None

    # ──────────────────────────────────────────────────────────
    #  功能 3：高亮目前執行目標點
    # ──────────────────────────────────────────────────────────
    def _highlight_row(self, idx):
        items = self.tree.get_children()
        for i, item in enumerate(items):
            self.tree.item(item, tags=("running",) if i == idx else ())

    def _clear_highlight(self):
        for item in self.tree.get_children():
            self.tree.item(item, tags=())

    # ──────────────────────────────────────────────────────────
    #  多螢幕覆蓋視窗（新增目標點）
    # ──────────────────────────────────────────────────────────
    def _get_monitor_bounds(self):
        if HAS_SCREENINFO:
            return [(m.x, m.y, m.width, m.height) for m in get_monitors()]
        w, h = pyautogui.size()
        return [(0, 0, w, h)]

    def add_target(self):
        overlays = []
        for idx, (mx, my, mw, mh) in enumerate(self._get_monitor_bounds()):
            ov = tk.Toplevel(self.root)
            ov.overrideredirect(True)
            ov.attributes('-topmost', True)
            ov.attributes('-alpha', 0.3)
            ov.config(bg='gray')
            ov.geometry(f"{mw}x{mh}+{mx}+{my}")
            tk.Label(
                ov, text="請點選目標點（按 Esc 取消）",
                font=("Arial", 24), bg="gray", fg="white"
            ).place(relx=0.5, rely=0.5, anchor="center")
            ov.bind("<Button-1>", lambda e, ovs=overlays: self._on_overlay_click(e, ovs))
            ov.bind("<Escape>",   lambda e, ovs=overlays: self._close_overlays(ovs))
            overlays.append(ov)
        if overlays:
            overlays[0].focus_set()

    def _on_overlay_click(self, event, overlays):
        x, y = event.x_root, event.y_root
        self._close_overlays(overlays)
        self._prompt_target_details(x, y)

    def _close_overlays(self, overlays):
        for ov in overlays:
            try:
                ov.destroy()
            except Exception:
                pass

    # ──────────────────────────────────────────────────────────
    #  Treeview 工具
    # ──────────────────────────────────────────────────────────
    def _prompt_target_details(self, x, y, idx=None, prefill=None):
        """新增（idx=None）或編輯（idx=列索引）目標點"""
        dialog = TargetSettingDialog(self.root, x, y, prefill=prefill)
        if dialog.result:
            click_count, wait_time, click_type = dialog.result
            entry = (x, y, click_count, wait_time, click_type)
            if idx is None:
                self.targets.append(entry)
            else:
                self.targets[idx] = entry
            self.update_treeview()

    def _on_tree_double_click(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        sel    = selected[0]
        values = self.tree.item(sel, "values")
        if not values:
            return
        try:
            x, y       = int(values[0]), int(values[1])
            clicks     = int(values[2])
            wait_time  = float(values[3])
            click_type = CLICK_TYPE_TO_KEY.get(values[4], "left")
        except Exception:
            return
        idx = self.tree.index(sel)
        self._prompt_target_details(x, y, idx=idx, prefill=(clicks, wait_time, click_type))

    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        for t in self.targets:
            x, y, clicks, wait_time, click_type = t
            self.tree.insert("", "end", values=(
                x, y, clicks, wait_time,
                CLICK_KEY_TO_DISPLAY.get(click_type, click_type)
            ))

    def remove_target(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showerror("錯誤", "請選取要移除的目標點")
            return
        indices = sorted([self.tree.index(s) for s in selected], reverse=True)
        for idx in indices:
            self.targets.pop(idx)
        self.update_treeview()

    # ──────────────────────────────────────────────────────────
    #  功能 1：儲存 / 載入 JSON 設定檔
    # ──────────────────────────────────────────────────────────
    def save_config(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON 設定檔", "*.json"), ("所有檔案", "*.*")],
            title="儲存設定"
        )
        if not path:
            return
        data = {
            "repeat":   self.entry_repeat.get(),
            "move_min": self.entry_move_min.get(),
            "move_max": self.entry_move_max.get(),
            "targets": [
                {"x": t[0], "y": t[1], "click_count": t[2],
                 "wait_time": t[3], "click_type": t[4]}
                for t in self.targets
            ],
        }
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("儲存成功", f"設定已儲存至：\n{path}")
        except Exception as e:
            messagebox.showerror("儲存失敗", str(e))

    def load_config(self):
        path = filedialog.askopenfilename(
            filetypes=[("JSON 設定檔", "*.json"), ("所有檔案", "*.*")],
            title="載入設定"
        )
        if not path:
            return
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
            self.targets = [
                (int(t["x"]), int(t["y"]), int(t["click_count"]),
                 float(t["wait_time"]), t.get("click_type", "left"))
                for t in data.get("targets", [])
            ]
            for entry, val in [
                (self.entry_repeat,   data.get("repeat",   "1")),
                (self.entry_move_min, data.get("move_min", "0.3")),
                (self.entry_move_max, data.get("move_max", "0.7")),
            ]:
                entry.delete(0, tk.END)
                entry.insert(0, str(val))
            self.update_treeview()
            messagebox.showinfo("載入成功", f"已從以下路徑載入設定：\n{path}")
        except Exception as e:
            messagebox.showerror("載入失敗", f"設定檔格式錯誤：{e}")

    # ──────────────────────────────────────────────────────────
    #  模擬控制
    # ──────────────────────────────────────────────────────────
    def start_simulation(self):
        try:
            repeat = int(self.entry_repeat.get())
            if repeat <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror("錯誤", "請輸入有效的重複次數（正整數）")
            return
        # ── 功能 5：驗證移動時間範圍 ──────────────────────────
        try:
            move_min = float(self.entry_move_min.get())
            move_max = float(self.entry_move_max.get())
            if move_min < 0 or move_max < move_min:
                raise ValueError
        except Exception:
            messagebox.showerror("錯誤", "移動時間需為數字，且最小值 ≥ 0、最大值 ≥ 最小值")
            return
        if not self.targets:
            messagebox.showerror("錯誤", "請新增至少一個目標點")
            return

        self.is_running  = True
        self.is_paused   = False
        self.resume_skip = False

        self.start_button.config(state=tk.DISABLED)
        self.pause_button.config(state=tk.NORMAL, text="暫停模擬")
        self.stop_button.config(state=tk.NORMAL)

        self.simulation_thread = threading.Thread(
            target=self.run_simulation,
            args=(repeat, move_min, move_max),
            daemon=True
        )
        self.simulation_thread.start()

    def toggle_pause(self):
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.pause_button.config(text="繼續模擬")
            self.root.after(0, lambda: self.iteration_label.config(text="已暫停"))
        else:
            self.resume_skip = True
            self.pause_button.config(text="暫停模擬")

    def stop_simulation(self):
        self.is_running  = False
        self.is_paused   = False
        self.resume_skip = False
        self.start_button.config(state=tk.NORMAL)
        self.pause_button.config(state=tk.DISABLED, text="暫停模擬")
        self.stop_button.config(state=tk.DISABLED)
        self.iteration_label.config(text="模擬已停止")
        self.root.after(0, self._clear_highlight)

    # ──────────────────────────────────────────────────────────
    #  模擬輔助方法
    # ──────────────────────────────────────────────────────────
    def wait_until_resumed(self):
        while self.is_paused and self.is_running:
            time.sleep(0.05)

    def check_intervention(self, expected_x, expected_y):
        if self.disable_check or self.is_paused:
            return
        cx, cy = pyautogui.position()
        if abs(cx - expected_x) > 20 or abs(cy - expected_y) > 20:
            self.is_paused = True
            self.root.after(0, lambda: self.pause_button.config(text="繼續模擬"))
            self.root.after(0, lambda: self.iteration_label.config(
                text="已暫停：偵測到使用者移動滑鼠"))

    def safe_sleep(self, duration, expected_x, expected_y):
        start = time.time()
        while time.time() - start < duration:
            if not self.is_running or self.resume_skip:
                return
            if self.is_paused:
                self.wait_until_resumed()
                return
            self.check_intervention(expected_x, expected_y)
            time.sleep(0.05)

    # ──────────────────────────────────────────────────────────
    #  主模擬迴圈
    # ──────────────────────────────────────────────────────────
    def run_simulation(self, repeat, move_min, move_max):
        for iteration in range(1, repeat + 1):
            if not self.is_running:
                break

            self.root.after(0, lambda i=iteration, r=repeat:
                self.iteration_label.config(text=f"目前執行次數：{i} / {r}"))

            for t_idx, target in enumerate(self.targets):
                if not self.is_running:
                    break

                x, y, click_count, wait_time, click_type = target

                # 若進入此目標點前已暫停，先等待繼續
                if self.is_paused:
                    self.wait_until_resumed()

                # 到達新目標點：重置 resume_skip（前一個目標被跳過後，此處正常執行）
                if self.resume_skip:
                    self.resume_skip = False

                if not self.is_running:
                    break

                # ── 功能 3：高亮目前目標點 ──────────────────────
                self.root.after(0, lambda i=t_idx: self._highlight_row(i))

                # ── 功能 5：隨機移動時間 ＋ easeInOutQuad 曲線 ──
                move_dur = random.uniform(move_min, move_max)
                self.disable_check = True
                pyautogui.moveTo(x, y, duration=move_dur, tween=pyautogui.easeInOutQuad)
                self.disable_check = False

                self.check_intervention(x, y)

                # ── 點擊迴圈 ────────────────────────────────────
                for _ in range(click_count):
                    if not self.is_running or self.resume_skip:
                        break
                    if self.is_paused:
                        self.wait_until_resumed()
                        break

                    self.check_intervention(x, y)
                    if self.resume_skip:
                        break

                    # ── 功能 4：依點擊類型執行 ───────────────────
                    if click_type == "double":
                        pyautogui.doubleClick()
                    elif click_type == "right":
                        pyautogui.rightClick()
                    else:
                        pyautogui.click()

                    self.safe_sleep(random.uniform(0.2, 0.5), x, y)

                # 等待秒數（resume_skip 時略過）
                if not self.resume_skip and self.is_running:
                    self.safe_sleep(wait_time, x, y)

        # ── 模擬結束 ────────────────────────────────────────────
        self.is_running  = False
        self.resume_skip = False
        self.root.after(0, self._clear_highlight)
        self.root.after(0, lambda: self.start_button.config(state=tk.NORMAL))
        self.root.after(0, lambda: self.pause_button.config(state=tk.DISABLED, text="暫停模擬"))
        self.root.after(0, lambda: self.stop_button.config(state=tk.DISABLED))
        self.root.after(0, lambda: self.iteration_label.config(text="模擬完成"))

    # ──────────────────────────────────────────────────────────
    #  說明視窗
    # ──────────────────────────────────────────────────────────
    def show_help(self):
        help_text = (
            "操作說明\n"
            "═══════════════════════════════════════\n\n"
            "【設定目標點】\n"
            "  新增：點選「新增目標點」→ 半透明覆蓋視窗覆蓋所有螢幕 →\n"
            "        在任意螢幕點選目標位置 → 輸入點擊次數、等待秒數、\n"
            "        點擊類型（左鍵單擊 / 左鍵雙擊 / 右鍵單擊）→ 確定。\n"
            "  編輯：在列表中雙擊目標點列，即可修改所有參數。\n"
            "  排序：按住目標點列後上下拖曳，即可調整執行順序。\n"
            "  移除：選取目標點後點選「移除選取目標」。\n"
            "  取消新增：按 Esc 鍵關閉覆蓋視窗。\n\n"
            "【儲存 / 載入設定】\n"
            "  「儲存設定」：將目前所有目標點與參數存為 JSON 檔。\n"
            "  「載入設定」：從 JSON 檔還原目標點與參數。\n\n"
            "【移動時間（秒）】\n"
            "  設定滑鼠移動到每個目標點的時間範圍（最小值 ～ 最大值）。\n"
            "  每次移動會在此範圍內隨機取值，並套用 easeInOutQuad\n"
            "  緩動曲線，讓動作更接近真人操作。\n\n"
            "【重複執行次數】\n"
            "  輸入正整數，模擬將依序重複執行指定次數。\n\n"
            "【開始模擬】\n"
            "  依序對所有目標點執行移動、點擊（依類型）、等待。\n"
            "  目前執行中的目標點會在列表中以綠色高亮顯示。\n\n"
            "【自動暫停】\n"
            "  模擬中手動移動滑鼠超過 20 像素，會自動暫停。\n\n"
            "【繼續模擬】\n"
            "  點選後從被中斷目標點的「下一個」目標點開始執行。\n\n"
            "【停止模擬】\n"
            "  隨時中斷模擬，狀態與按鈕重置。\n\n"
            "【多螢幕】\n"
            "  安裝 screeninfo 套件以支援多螢幕選點：\n"
            "  pip install screeninfo\n"
        )
        win = tk.Toplevel(self.root)
        win.title("操作說明")
        win.geometry("560x500")

        scrollbar = tk.Scrollbar(win)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        text_area = tk.Text(win, wrap=tk.WORD, padx=12, pady=12,
                            yscrollcommand=scrollbar.set)
        text_area.insert(tk.END, help_text)
        text_area.config(state=tk.DISABLED)
        text_area.pack(expand=True, fill=tk.BOTH)
        scrollbar.config(command=text_area.yview)


# ══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    MouseClickSimulatorGUI()
