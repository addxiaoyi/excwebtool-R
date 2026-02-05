import tkinter as tk
from tkinter import messagebox
from typing import List, Dict
from datetime import datetime
from pathlib import Path

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

AUTHOR = "ADD2048"


# =========================
# 数据模型
# =========================
def default_site() -> Dict:
    return {
        "name": "默认站点",
        "url": "",
        "interval": 5,
        "page_mode": "number",
        "max_pages": 1,
        "fields": ["字段 A", "字段 B", "字段 C"],
        "output": "output.xlsx"
    }


# =========================
# 主程序
# =========================
class App(tb.Window):
    def __init__(self):
        super().__init__(themename="darkly")
        self.geometry("1400x820")
        self.title("网页表格自动抓取系统")
        self.overrideredirect(True)

        self.sites: List[Dict] = [default_site()]
        self.current_site = self.sites[0]

        self._drag_x = 0
        self._drag_y = 0

        self._build_ui()

    # ================= UI =================
    def _build_ui(self):
        self._build_titlebar()
        self._build_main()
        self._build_footer()

    def _build_titlebar(self):
        bar = tb.Frame(self, bootstyle="secondary")
        bar.pack(fill=X)

        bar.bind("<ButtonPress-1>", self._start_move)
        bar.bind("<B1-Motion>", self._move)

        tb.Label(
            bar,
            text="网页表格自动抓取系统",
            font=("Segoe UI", 13, "bold")
        ).pack(side=LEFT, padx=12)

        tb.Button(
            bar,
            text="▶ 执行",
            bootstyle=SUCCESS,
            command=self.run_task
        ).pack(side=RIGHT, padx=6)

        tb.Button(
            bar,
            text="✕",
            bootstyle=DANGER,
            width=3,
            command=self.destroy
        ).pack(side=RIGHT)

    def _build_main(self):
        root = tb.Frame(self)
        root.pack(fill=BOTH, expand=True)

        # 左侧站点栏
        left = tb.Frame(root, width=240, bootstyle="dark")
        left.pack(side=LEFT, fill=Y)

        tb.Label(left, text="站点列表", font=("Segoe UI", 11, "bold")).pack(anchor=W, padx=12, pady=10)

        self.site_list = tk.Listbox(
            left,
            bg="#020617",
            fg="#e5e7eb",
            relief=FLAT,
            highlightthickness=0
        )
        self.site_list.pack(fill=BOTH, expand=True, padx=10)
        self.site_list.insert(END, self.current_site["name"])
        self.site_list.selection_set(0)

        # 中央配置区
        center = tb.Frame(root, padding=16)
        center.pack(side=LEFT, fill=BOTH, expand=True)

        self.tabs = tb.Notebook(center)
        self.tabs.pack(fill=BOTH, expand=True)

        self._build_tabs()

        # 右侧仪表盘
        right = tb.Frame(root, width=300, padding=12, bootstyle="secondary")
        right.pack(side=RIGHT, fill=Y)

        tb.Label(right, text="任务仪表盘", font=("Segoe UI", 11, "bold")).pack(anchor=W)

        self.progress = tb.Progressbar(right, mode="determinate")
        self.progress.pack(fill=X, pady=10)

        tb.Label(right, text="执行日志").pack(anchor=W)
        self.log_box = tk.Text(
            right,
            height=18,
            bg="#020617",
            fg="#d1d5db",
            relief=FLAT
        )
        self.log_box.pack(fill=BOTH, expand=True)

    def _build_tabs(self):
        self.tab_basic = tb.Frame(self.tabs, padding=12)
        self.tab_page = tb.Frame(self.tabs, padding=12)
        self.tab_fields = tb.Frame(self.tabs, padding=12)
        self.tab_output = tb.Frame(self.tabs, padding=12)

        self.tabs.add(self.tab_basic, text="基础")
        self.tabs.add(self.tab_page, text="分页")
        self.tabs.add(self.tab_fields, text="字段")
        self.tabs.add(self.tab_output, text="输出")

        self._build_basic_tab()
        self._build_page_tab()
        self._build_fields_tab()
        self._build_output_tab()

    # ================= Tabs =================
    def _build_basic_tab(self):
        tb.Label(self.tab_basic, text="网页 URL").pack(anchor=W)
        self.url_entry = tb.Entry(self.tab_basic)
        self.url_entry.pack(fill=X)

        tb.Label(self.tab_basic, text="抓取间隔（分钟）").pack(anchor=W, pady=(10, 0))
        self.interval_entry = tb.Entry(self.tab_basic, width=8)
        self.interval_entry.insert(0, "5")
        self.interval_entry.pack(anchor=W)

    def _build_page_tab(self):
        tb.Label(self.tab_page, text="分页策略（可视化）").pack(anchor=W)

        self.page_mode = tk.StringVar(value="number")

        for text, value in [
            ("页码模式", "number"),
            ("按钮模式", "button"),
            ("滚动模式", "scroll")
        ]:
            tb.Radiobutton(
                self.tab_page,
                text=text,
                value=value,
                variable=self.page_mode
            ).pack(anchor=W, pady=2)

        tb.Label(self.tab_page, text="最大页数").pack(anchor=W, pady=(10, 0))
        self.max_pages = tb.Entry(self.tab_page, width=6)
        self.max_pages.insert(0, "1")
        self.max_pages.pack(anchor=W)

    def _build_fields_tab(self):
        tb.Label(self.tab_fields, text="字段区域（Notion 风）").pack(anchor=W)

        self.field_list = tk.Listbox(
            self.tab_fields,
            bg="#020617",
            fg="#e5e7eb",
            relief=FLAT,
            selectbackground="#2563eb"
        )
        self.field_list.pack(fill=BOTH, expand=True, pady=6)

        for f in self.current_site["fields"]:
            self.field_list.insert(END, f)

        tb.Label(
            self.tab_fields,
            text="提示：顺序即 Excel 列顺序",
            foreground="#9ca3af"
        ).pack(anchor=W, pady=(6, 0))

    def _build_output_tab(self):
        tb.Label(self.tab_output, text="输出 Excel 文件名").pack(anchor=W)
        self.filename_entry = tb.Entry(self.tab_output)
        self.filename_entry.insert(0, "output.xlsx")
        self.filename_entry.pack(anchor=W)

    # ================= Logic =================
    def log(self, text: str):
        time = datetime.now().strftime("%H:%M:%S")
        self.log_box.insert(END, f"[{time}] {text}\n")
        self.log_box.see(END)
        self.update()

    def run_task(self):
        self.progress["value"] = 0
        self.log("任务开始")

        filename = self.filename_entry.get().strip() or "output.xlsx"
        fields = list(self.field_list.get(0, END))

        self.progress["value"] = 30
        self.log("生成 Excel 文件")

        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"

        # 表头样式 = 网站样式（基础）
        for col, name in enumerate(fields, start=1):
            cell = ws.cell(row=1, column=col, value=name)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")

        self.progress["value"] = 80
        wb.save(filename)

        self.progress["value"] = 100
        self.log(f"完成，生成文件：{Path(filename).resolve()}")

        messagebox.showinfo("完成", "任务执行完成")

    # ================= Window Move =================
    def _start_move(self, e):
        self._drag_x = e.x
        self._drag_y = e.y

    def _move(self, e):
        self.geometry(f"+{e.x_root - self._drag_x}+{e.y_root - self._drag_y}")

    def _build_footer(self):
        tb.Label(
            self,
            text=f"Author : {AUTHOR}",
            anchor=E,
            foreground="#9ca3af"
        ).pack(fill=X, side=BOTTOM, padx=10, pady=4)


# =========================
if __name__ == "__main__":
    App().mainloop()
