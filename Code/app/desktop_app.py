from __future__ import annotations

import json
import os
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from app.core.config_loader import APP_DIR, CONFIG_DIR, expand_path, load_customers, load_settings
from app.core.draft_generator import DraftGenerator
from app.core.logger import setup_logger
from app.storage.repository import DraftRepository


APP_NAME = "发货单邮件草稿自动生成工具"
PROJECT_DIR = APP_DIR
CUSTOMERS_PATH = CONFIG_DIR / "customers.json"
SETTINGS_PATH = CONFIG_DIR / "settings.json"


def _startup_dir() -> Path:
    return Path(os.environ["APPDATA"]) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"


def _startup_file() -> Path:
    return _startup_dir() / "AutoEmailDraftTool.vbs"


class AutoEmailApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1120x720")
        self.minsize(980, 640)
        self.protocol("WM_DELETE_WINDOW", self.hide_window)

        self.is_scanning = False
        self.close_notice_shown = False
        self.selected_customer_index: int | None = None
        self.settings: dict = {}
        self.customers: list[dict] = []
        self.logger = None
        self.repository: DraftRepository | None = None
        self.generator: DraftGenerator | None = None

        self.status_var = tk.StringVar(value="就绪")
        self.customer_summary_var = tk.StringVar()
        self.foxmail_path_var = tk.StringVar()
        self.account_storage_var = tk.StringVar()
        self.autostart_var = tk.BooleanVar(value=self._is_autostart_enabled())
        self.auto_scan_var = tk.BooleanVar(value=True)
        self.scan_interval_var = tk.StringVar(value="60")
        self.last_scan_var = tk.StringVar(value="尚未扫描")
        self.current_scan_is_manual = False

        self.customer_id_var = tk.StringVar()
        self.customer_name_var = tk.StringVar()
        self.watch_dir_var = tk.StringVar()
        self.include_patterns_var = tk.StringVar(value="*送货单*.xlsx, *送货单*.xls")
        self.recursive_var = tk.BooleanVar(value=True)
        self.start_mtime_var = tk.StringVar()
        self.to_var = tk.StringVar()
        self.cc_var = tk.StringVar()

        self._load_runtime()
        self._build_ui()
        self._sync_settings_vars()
        self.refresh_customer_list()
        self.refresh_records()
        self._schedule_auto_scan(initial=True)

    def _load_runtime(self) -> None:
        self.settings = load_settings()
        self.customers = load_customers()
        self.logger = setup_logger(self.settings)
        self.repository = DraftRepository(self.settings["database"]["path"])
        self.repository.init_schema()
        self.generator = DraftGenerator(
            self.settings,
            self.customers,
            self.repository,
            self.logger,
        )

    def _reload_runtime(self) -> None:
        self._load_runtime()
        self._sync_settings_vars()
        self.refresh_customer_list()
        self.refresh_records()

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=14)
        root.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(root, text=APP_NAME, font=("Microsoft YaHei UI", 16, "bold"))
        title.pack(anchor=tk.W)
        ttk.Label(
            root,
            text="业务人员保存 Excel 发货单后，在这里扫描生成 Foxmail 草稿。发送前仍需人工检查。",
        ).pack(anchor=tk.W, pady=(4, 12))

        notebook = ttk.Notebook(root)
        notebook.pack(fill=tk.BOTH, expand=True)

        run_tab = ttk.Frame(notebook, padding=12)
        customers_tab = ttk.Frame(notebook, padding=12)
        settings_tab = ttk.Frame(notebook, padding=12)
        notebook.add(run_tab, text="运行")
        notebook.add(customers_tab, text="客户配置")
        notebook.add(settings_tab, text="系统设置")

        self._build_run_tab(run_tab)
        self._build_customers_tab(customers_tab)
        self._build_settings_tab(settings_tab)

    def _build_run_tab(self, parent: ttk.Frame) -> None:
        status_frame = ttk.LabelFrame(parent, text="当前状态", padding=12)
        status_frame.pack(fill=tk.X)
        ttk.Label(status_frame, textvariable=self.status_var, font=("Microsoft YaHei UI", 11)).pack(anchor=tk.W)
        ttk.Label(status_frame, textvariable=self.customer_summary_var).pack(anchor=tk.W, pady=(6, 0))
        ttk.Label(status_frame, textvariable=self.foxmail_path_var).pack(anchor=tk.W, pady=(2, 0))
        ttk.Label(status_frame, textvariable=self.last_scan_var).pack(anchor=tk.W, pady=(2, 0))

        actions = ttk.Frame(parent)
        actions.pack(fill=tk.X, pady=14)
        self.scan_button = ttk.Button(actions, text="立即扫描并生成草稿", command=self.scan_once)
        self.scan_button.pack(side=tk.LEFT)
        ttk.Button(actions, text="刷新记录", command=self.refresh_records).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(actions, text="打开 Foxmail", command=self.open_foxmail).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(actions, text="打开日志", command=self.open_log).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(actions, text="打开输出目录", command=self.open_output_dir).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(actions, text="退出程序", command=self.exit_app).pack(side=tk.RIGHT)

        table_frame = ttk.LabelFrame(parent, text="最近处理记录", padding=8)
        table_frame.pack(fill=tk.BOTH, expand=True)
        columns = ("id", "customer", "file", "status", "import_status", "msg_id", "updated")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=16)
        headings = {
            "id": "编号",
            "customer": "客户",
            "file": "文件名",
            "status": "状态",
            "import_status": "导入",
            "msg_id": "草稿ID",
            "updated": "更新时间",
        }
        widths = {
            "id": 60,
            "customer": 120,
            "file": 340,
            "status": 90,
            "import_status": 90,
            "msg_id": 80,
            "updated": 160,
        }
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], anchor=tk.W)
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Label(
            parent,
            text="状态 imported/success 表示已进入 Foxmail 草稿箱；failed 表示失败，请打开日志查看原因。",
        ).pack(anchor=tk.W, pady=(10, 0))

    def _build_customers_tab(self, parent: ttk.Frame) -> None:
        left = ttk.Frame(parent)
        left.pack(side=tk.LEFT, fill=tk.Y)
        right = ttk.Frame(parent)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(14, 0))

        ttk.Label(left, text="客户列表").pack(anchor=tk.W)
        self.customer_list = tk.Listbox(left, width=28, height=22, exportselection=False)
        self.customer_list.pack(fill=tk.Y, expand=False, pady=(6, 8))
        self.customer_list.bind("<<ListboxSelect>>", self.on_customer_selected)
        ttk.Button(left, text="新增客户", command=self.new_customer).pack(fill=tk.X)
        ttk.Button(left, text="删除客户", command=self.delete_customer).pack(fill=tk.X, pady=(8, 0))

        form = ttk.LabelFrame(right, text="客户信息", padding=12)
        form.pack(fill=tk.BOTH, expand=True)
        self._entry_row(form, "客户编号", self.customer_id_var, 0)
        self._entry_row(form, "客户名称", self.customer_name_var, 1)
        self._path_row(form, "Excel 文件夹", self.watch_dir_var, 2, self.choose_watch_dir)
        self._entry_row(form, "文件规则", self.include_patterns_var, 3)
        ttk.Checkbutton(
            form,
            text="扫描子文件夹",
            variable=self.recursive_var,
        ).grid(row=4, column=1, sticky=tk.W, pady=6)
        ttk.Label(form, text="起始时间").grid(row=5, column=0, sticky=tk.W, pady=6)
        ttk.Entry(form, textvariable=self.start_mtime_var).grid(row=5, column=1, sticky=tk.EW, pady=6)
        ttk.Button(form, text="设为当前时间", command=self.set_start_mtime_now).grid(row=5, column=2, sticky=tk.E, padx=(8, 0), pady=6)
        self._entry_row(form, "收件人", self.to_var, 6)
        self._entry_row(form, "抄送人", self.cc_var, 7)

        ttk.Label(form, text="邮件正文").grid(row=8, column=0, sticky=tk.NW, pady=6)
        self.body_text = tk.Text(form, height=8, wrap=tk.WORD)
        self.body_text.grid(row=8, column=1, columnspan=2, sticky=tk.NSEW, pady=6)
        form.columnconfigure(1, weight=1)
        form.rowconfigure(8, weight=1)

        ttk.Label(
            form,
            text="提示：起始时间格式为 2026-06-03 09:30:00；只处理修改时间晚于起始时间的 Excel。",
        ).grid(row=9, column=1, sticky=tk.W, pady=(4, 10))
        ttk.Button(form, text="保存客户配置", command=self.save_customer).grid(row=10, column=1, sticky=tk.W)

    def _build_settings_tab(self, parent: ttk.Frame) -> None:
        form = ttk.LabelFrame(parent, text="系统设置", padding=12)
        form.pack(fill=tk.X)
        self.foxmail_exe_var = tk.StringVar()
        self.account_storage_edit_var = tk.StringVar()
        self._path_row(form, "Foxmail 程序", self.foxmail_exe_var, 0, self.choose_foxmail_exe)
        self._path_row(form, "邮箱本地目录", self.account_storage_edit_var, 1, self.choose_account_storage)

        ttk.Checkbutton(
            form,
            text="电脑重启后自动启动",
            variable=self.autostart_var,
            command=self.toggle_autostart,
        ).grid(row=2, column=1, sticky=tk.W, pady=8)
        ttk.Checkbutton(
            form,
            text="自动扫描客户文件夹",
            variable=self.auto_scan_var,
        ).grid(row=3, column=1, sticky=tk.W, pady=8)
        ttk.Label(form, text="扫描间隔秒数").grid(row=4, column=0, sticky=tk.W, pady=6)
        ttk.Entry(form, textvariable=self.scan_interval_var, width=12).grid(row=4, column=1, sticky=tk.W, pady=6)

        ttk.Button(form, text="保存系统设置", command=self.save_settings).grid(row=5, column=1, sticky=tk.W, pady=(8, 0))
        form.columnconfigure(1, weight=1)

        notes = ttk.LabelFrame(parent, text="说明", padding=12)
        notes.pack(fill=tk.X, pady=(14, 0))
        ttk.Label(notes, text="Foxmail 程序通常是 Foxmail.exe。").pack(anchor=tk.W)
        ttk.Label(notes, text="邮箱本地目录通常位于 Foxmail\\Storage\\邮箱账号。").pack(anchor=tk.W, pady=(4, 0))
        ttk.Label(notes, text="开机自启使用无黑窗启动方式，不会弹出 CMD 窗口。").pack(anchor=tk.W, pady=(4, 0))
        ttk.Label(notes, text="建议开启自动扫描，默认每 60 秒扫描一次；扫描未完成时不会重复启动下一轮。").pack(anchor=tk.W, pady=(4, 0))

    def _entry_row(self, parent: ttk.Frame, label: str, variable: tk.StringVar, row: int) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, pady=6)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, columnspan=2, sticky=tk.EW, pady=6)

    def _path_row(
        self,
        parent: ttk.Frame,
        label: str,
        variable: tk.StringVar,
        row: int,
        command,
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, pady=6)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky=tk.EW, pady=6)
        ttk.Button(parent, text="选择", command=command).grid(row=row, column=2, sticky=tk.E, padx=(8, 0), pady=6)

    def _sync_settings_vars(self) -> None:
        names = "、".join(customer["customer_name"] for customer in self.customers) if self.customers else "未配置"
        self.customer_summary_var.set(f"客户配置：{names}")
        self.foxmail_path_var.set(f"Foxmail：{self.settings['foxmail']['foxmail_exe']}")
        self.account_storage_var.set(self.settings["foxmail"]["account_storage"])
        if hasattr(self, "foxmail_exe_var"):
            self.foxmail_exe_var.set(self.settings["foxmail"]["foxmail_exe"])
            self.account_storage_edit_var.set(self.settings["foxmail"]["account_storage"])
            scanner = self.settings.get("scanner", {})
            self.auto_scan_var.set(bool(scanner.get("auto_scan_enabled", True)))
            self.scan_interval_var.set(str(int(scanner.get("scan_interval_seconds", 60))))
        self.autostart_var.set(self._is_autostart_enabled())

    def refresh_customer_list(self) -> None:
        if not hasattr(self, "customer_list"):
            return
        self.customer_list.delete(0, tk.END)
        for customer in self.customers:
            self.customer_list.insert(tk.END, customer.get("customer_name", "未命名客户"))
        if self.customers:
            self.customer_list.selection_set(0)
            self.load_customer_into_form(0)

    def on_customer_selected(self, _event=None) -> None:
        selection = self.customer_list.curselection()
        if selection:
            self.load_customer_into_form(selection[0])

    def load_customer_into_form(self, index: int) -> None:
        self.selected_customer_index = index
        customer = self.customers[index]
        self.customer_id_var.set(customer.get("customer_id", ""))
        self.customer_name_var.set(customer.get("customer_name", ""))
        self.watch_dir_var.set(customer.get("watch_dir", ""))
        self.include_patterns_var.set(", ".join(customer.get("include_patterns", ["*送货单*.xlsx", "*送货单*.xls"])))
        self.recursive_var.set(bool(customer.get("recursive", True)))
        self.start_mtime_var.set(customer.get("start_mtime", ""))
        self.to_var.set(", ".join(customer.get("to", [])))
        self.cc_var.set(", ".join(customer.get("cc", [])))
        self.body_text.delete("1.0", tk.END)
        self.body_text.insert("1.0", customer.get("body", "您好，附件为{subject}，请查收。"))

    def new_customer(self) -> None:
        self.selected_customer_index = None
        self.customer_id_var.set("")
        self.customer_name_var.set("")
        self.watch_dir_var.set("")
        self.include_patterns_var.set("*送货单*.xlsx, *送货单*.xls")
        self.recursive_var.set(True)
        self.start_mtime_var.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.to_var.set("")
        self.cc_var.set("")
        self.body_text.delete("1.0", tk.END)
        self.body_text.insert("1.0", "您好，附件为{subject}，请查收。")

    def delete_customer(self) -> None:
        if self.selected_customer_index is None:
            messagebox.showinfo("提示", "请先选择客户。")
            return
        customer_name = self.customers[self.selected_customer_index].get("customer_name", "")
        if not messagebox.askyesno("确认删除", f"确定删除客户配置：{customer_name}？"):
            return
        del self.customers[self.selected_customer_index]
        self._save_customers_file()
        self._reload_runtime()
        self.status_var.set("客户配置已删除。")

    def save_customer(self) -> None:
        customer_id = self.customer_id_var.get().strip()
        customer_name = self.customer_name_var.get().strip()
        watch_dir = self.watch_dir_var.get().strip()
        if not customer_id or not customer_name or not watch_dir:
            messagebox.showwarning("信息不完整", "客户编号、客户名称、Excel 文件夹不能为空。")
            return
        customer = {
            "customer_id": customer_id,
            "customer_name": customer_name,
            "watch_dir": watch_dir,
            "include_patterns": self._split_values(self.include_patterns_var.get()) or ["*送货单*.xlsx", "*送货单*.xls"],
            "recursive": bool(self.recursive_var.get()),
            "start_mtime": self.start_mtime_var.get().strip(),
            "to": self._split_values(self.to_var.get()),
            "cc": self._split_values(self.cc_var.get()),
            "body": self.body_text.get("1.0", tk.END).strip() or "您好，附件为{subject}，请查收。",
        }
        if self.selected_customer_index is None:
            self.customers.append(customer)
        else:
            self.customers[self.selected_customer_index] = customer
        self._save_customers_file()
        self._reload_runtime()
        self.status_var.set("客户配置已保存。")
        messagebox.showinfo("保存成功", "客户配置已保存。")

    def save_settings(self) -> None:
        try:
            interval = int(self.scan_interval_var.get().strip())
        except ValueError:
            messagebox.showwarning("配置错误", "扫描间隔秒数必须是数字。")
            return
        if interval < 10:
            messagebox.showwarning("配置错误", "扫描间隔建议不少于 10 秒。")
            return
        self.settings["foxmail"]["foxmail_exe"] = self.foxmail_exe_var.get().strip()
        self.settings["foxmail"]["account_storage"] = self.account_storage_edit_var.get().strip()
        self.settings.setdefault("scanner", {})
        self.settings["scanner"]["auto_scan_enabled"] = bool(self.auto_scan_var.get())
        self.settings["scanner"]["scan_interval_seconds"] = interval
        self._save_settings_file()
        self._reload_runtime()
        self.status_var.set("系统设置已保存。")
        messagebox.showinfo("保存成功", "系统设置已保存。")

    def set_start_mtime_now(self) -> None:
        self.start_mtime_var.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    def choose_watch_dir(self) -> None:
        path = filedialog.askdirectory(title="选择 Excel 发货单文件夹")
        if path:
            self.watch_dir_var.set(path)

    def choose_foxmail_exe(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Foxmail.exe",
            filetypes=[("Foxmail", "Foxmail.exe"), ("EXE 文件", "*.exe"), ("所有文件", "*.*")],
        )
        if path:
            self.foxmail_exe_var.set(path)

    def choose_account_storage(self) -> None:
        path = filedialog.askdirectory(title="选择 Foxmail 邮箱本地目录")
        if path:
            self.account_storage_edit_var.set(path)

    def scan_once(self, manual: bool = True) -> None:
        if self.is_scanning:
            return
        self.is_scanning = True
        self.current_scan_is_manual = manual
        self.scan_button.configure(state=tk.DISABLED)
        if manual:
            self.status_var.set("正在扫描客户文件夹，请稍候...")
        else:
            self.status_var.set("自动扫描中...")
        threading.Thread(target=self._scan_worker, daemon=True).start()

    def _scan_worker(self) -> None:
        try:
            assert self.generator is not None
            self.generator.scan_once()
            self.after(0, self._scan_finished, True, "扫描完成，请到 Foxmail 草稿箱检查。")
        except Exception as exc:
            assert self.logger is not None
            self.logger.exception("桌面端扫描失败")
            self.after(0, self._scan_finished, False, str(exc))

    def _scan_finished(self, success: bool, message: str) -> None:
        self.is_scanning = False
        self.scan_button.configure(state=tk.NORMAL)
        self.status_var.set(message)
        self.last_scan_var.set(f"最近扫描：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.refresh_records()
        if not success and self.current_scan_is_manual:
            messagebox.showerror("扫描失败", message)

    def _schedule_auto_scan(self, initial: bool = False) -> None:
        scanner = self.settings.get("scanner", {})
        interval = max(10, int(scanner.get("scan_interval_seconds", 60)))
        delay = 3000 if initial else interval * 1000
        self.after(delay, self._auto_scan_tick)

    def _auto_scan_tick(self) -> None:
        scanner = self.settings.get("scanner", {})
        if bool(scanner.get("auto_scan_enabled", True)) and not self.is_scanning:
            self.scan_once(manual=False)
        self._schedule_auto_scan()

    def refresh_records(self) -> None:
        if not hasattr(self, "tree") or self.repository is None:
            return
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in self.repository.list_recent(50):
            self.tree.insert(
                "",
                tk.END,
                values=(
                    row["id"],
                    row["customer_name"],
                    row["file_name"],
                    row["status"],
                    row["import_status"],
                    row["foxmail_msg_id"] or "",
                    row["updated_at"],
                ),
            )

    def open_foxmail(self) -> None:
        self._open_path(self.settings["foxmail"]["foxmail_exe"])

    def open_log(self) -> None:
        self._open_path(self.settings["logging"]["path"])

    def open_output_dir(self) -> None:
        self._open_path(self.settings["output"]["eml_dir"])

    def toggle_autostart(self) -> None:
        try:
            if self.autostart_var.get():
                self._enable_autostart()
                self.status_var.set("已启用开机自动启动。")
            else:
                self._disable_autostart()
                self.status_var.set("已关闭开机自动启动。")
        except Exception as exc:
            self.autostart_var.set(self._is_autostart_enabled())
            messagebox.showerror("开机启动设置失败", str(exc))

    def _enable_autostart(self) -> None:
        startup = _startup_file()
        startup.parent.mkdir(parents=True, exist_ok=True)
        startup.write_text(self._vbs_content(), encoding="utf-8")

    def _disable_autostart(self) -> None:
        startup = _startup_file()
        if startup.exists():
            startup.unlink()

    def _is_autostart_enabled(self) -> bool:
        return _startup_file().exists()

    def _vbs_content(self) -> str:
        if getattr(sys, "frozen", False):
            command = f'"{sys.executable}"'
            return (
                'Set shell = CreateObject("WScript.Shell")\n'
                f'shell.CurrentDirectory = "{PROJECT_DIR}"\n'
                f'shell.Run "{command}", 0, False\n'
            )
        pythonw = Path(sys.executable).with_name("pythonw.exe")
        exe = pythonw if pythonw.exists() else Path(sys.executable)
        command = f'"{exe}" -m app.main --gui'
        return (
            'Set shell = CreateObject("WScript.Shell")\n'
            f'shell.CurrentDirectory = "{PROJECT_DIR}"\n'
            f'shell.Run "{command}", 0, False\n'
        )

    def hide_window(self) -> None:
        self.iconify()
        if not self.close_notice_shown:
            self.close_notice_shown = True
            messagebox.showinfo(
                "程序仍在运行",
                "窗口已最小化，程序没有退出。如需关闭程序，请点击“退出程序”。",
            )

    def exit_app(self) -> None:
        if self.is_scanning:
            messagebox.showwarning("正在扫描", "当前正在扫描，请等待完成后再退出。")
            return
        if messagebox.askyesno("退出程序", "确定退出发货单邮件草稿自动生成工具？"):
            self.destroy()

    def _open_path(self, path_value: str | Path) -> None:
        path = Path(expand_path(str(path_value)))
        if not path.exists():
            messagebox.showwarning("路径不存在", str(path))
            return
        os.startfile(str(path))

    def _save_customers_file(self) -> None:
        CUSTOMERS_PATH.write_text(
            json.dumps(self.customers, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def _save_settings_file(self) -> None:
        SETTINGS_PATH.write_text(
            json.dumps(self.settings, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def _split_values(self, value: str) -> list[str]:
        return [item.strip() for item in value.split(",") if item.strip()]


def run_desktop_app() -> None:
    app = AutoEmailApp()
    app.mainloop()
