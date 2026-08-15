import os
import sys
import json
import datetime
import traceback
from tkinter import filedialog, messagebox, simpledialog
import tkinter as tk

import pandas as pd
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

# 确保打包时包含 openpyxl 和 xlrd
try:
    import openpyxl  # noqa: F401
except ImportError:
    pass
try:
    import xlrd  # noqa: F401
except ImportError:
    pass


def get_excel_column_letter(n):
    """数字索引（0-based）转 Excel 列字母"""
    n += 1
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def normalize_columns(df):
    """
    规范化列名：处理空列名、重复列名，返回 (df_with_normalized_columns, columns_info)
    columns_info: [(index, original_name, display_name, excel_letter)]
    """
    raw_names = list(df.columns)
    columns_info = []
    seen = {}
    new_columns = []
    for i, raw in enumerate(raw_names):
        if pd.isna(raw) or str(raw).strip() == "":
            original = ""
            base = f"列{i+1}"
        else:
            original = str(raw).strip()
            base = original
        if base in seen:
            seen[base] += 1
            display = f"{base}_{seen[base]}"
        else:
            seen[base] = 0
            display = base
        excel_letter = get_excel_column_letter(i)
        columns_info.append((i, original, display, excel_letter))
        new_columns.append(display)
    df.columns = new_columns
    return df, columns_info


def get_config_dir():
    """获取配置文件目录，优先使用 exe 同目录/config，失败则请求用户选择"""
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    config_dir = os.path.join(base_dir, "config")
    try:
        os.makedirs(config_dir, exist_ok=True)
        test_file = os.path.join(config_dir, ".write_test")
        with open(test_file, "w") as f:
            f.write("test")
        os.remove(test_file)
        return config_dir
    except Exception:
        messagebox.showwarning("配置目录不可用",
                               f"无法在程序目录创建配置文件夹：\n{config_dir}\n\n"
                               "请选择一个可写的目录用于保存配置。")
        chosen = filedialog.askdirectory(title="选择配置保存目录")
        if chosen:
            return chosen
        else:
            import tempfile
            return tempfile.gettempdir()


class ToolTip:
    """简易鼠标悬停提示"""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, event=None):
        if self.tip_window or not self.text:
            return
        x, y = self.widget.winfo_pointerxy()
        x += 15
        y += 20
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = ttk.Label(tw, text=self.text, bootstyle="info", padding=5)
        label.pack()

    def hide(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None


class App:
    def __init__(self):
        self.root = ttk.Window(themename="flatly")
        self.root.title("REL测试数据筛选工具")
        self.root.geometry("1000x800")
        self.root.minsize(900, 750)

        # 字体
        self.default_font = ("Microsoft YaHei UI", 10)
        self.root.option_add("*Font", self.default_font)
        style = ttk.Style()
        style.configure(".", font=self.default_font)
        style.configure("Treeview", font=self.default_font)
        style.configure("TButton", font=self.default_font)

        # 配置目录
        self.config_dir = get_config_dir()
        self.config_file = os.path.join(self.config_dir, "configs.json")
        self.configs = []
        self.current_config_name = None

        # 数据相关
        self.file_path = None
        self.df = None
        self.sheet_names = []
        self.current_sheet = None
        self.header_row = 1
        self.columns_info = []
        self.selected_indices = []

        # 多文件合并相关
        self.multi_files = []
        self.multi_area_visible = False
        self.multi_area_built = False

        # 进度条
        self.progress_bar = None

        self.build_ui()
        self.load_configs()
        self.populate_config_combobox()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    # ---------- UI 构建 ----------
    def build_ui(self):
        # ========== 顶部多文件合并区域 ==========
        self.multi_container = ttk.Frame(self.root, padding=5)
        self.multi_container.pack(fill=tk.X)

        self.btn_toggle_multi = ttk.Button(self.multi_container, text="展开多文件合并", command=self.toggle_multi_area)
        self.btn_toggle_multi.pack(anchor='w', padx=5, pady=2)

        self.multi_frame = ttk.Frame(self.multi_container, padding=5, relief=tk.GROOVE)

        # ========== 主文件操作区域 ==========
        self.top_frame = ttk.Frame(self.root, padding=10)
        self.top_frame.pack(fill=tk.X)
        self.top_frame.columnconfigure(1, weight=1)

        # 第0行：导入文件
        ttk.Label(self.top_frame, text="导入文件：", width=12, anchor='e').grid(row=0, column=0, sticky='e', padx=5, pady=3)
        self.entry_file = ttk.Entry(self.top_frame)
        self.entry_file.grid(row=0, column=1, sticky='ew', padx=5, pady=3)
        ttk.Button(self.top_frame, text="浏览...", command=self.browse_file).grid(row=0, column=2, padx=5, pady=3)

        # 第1行：导出路径
        ttk.Label(self.top_frame, text="导出路径：", width=12, anchor='e').grid(row=1, column=0, sticky='e', padx=5, pady=3)
        self.entry_output = ttk.Entry(self.top_frame)
        self.entry_output.grid(row=1, column=1, sticky='ew', padx=5, pady=3)
        ttk.Button(self.top_frame, text="浏览...", command=self.browse_output).grid(row=1, column=2, padx=5, pady=3)

        # 第2行：Sheet 和表头行
        ttk.Label(self.top_frame, text="Sheet：", width=12, anchor='e').grid(row=2, column=0, sticky='e', padx=5, pady=3)
        self.combo_sheet = ttk.Combobox(self.top_frame, state="readonly", width=20)
        self.combo_sheet.grid(row=2, column=1, sticky='w', padx=5, pady=3)
        self.combo_sheet.bind("<<ComboboxSelected>>", self.on_sheet_change)
        ttk.Label(self.top_frame, text="表头行：").grid(row=2, column=1, sticky='e', padx=(0, 180), pady=3)
        self.spin_header = ttk.Spinbox(self.top_frame, from_=1, to=1000, width=5)
        self.spin_header.set(1)
        self.spin_header.grid(row=2, column=1, sticky='e', padx=(0, 60), pady=3)
        self.spin_header.bind("<Return>", self.on_header_change)
        self.spin_header.bind("<FocusOut>", self.on_header_change)

        # 第3行：搜索列
        ttk.Label(self.top_frame, text="搜索列：", width=12, anchor='e').grid(row=3, column=0, sticky='e', padx=5, pady=3)
        self.entry_search = ttk.Entry(self.top_frame)
        self.entry_search.grid(row=3, column=1, columnspan=2, sticky='ew', padx=5, pady=3)
        self.entry_search.bind("<KeyRelease>", self.filter_left_listbox)

        # ========== 底部状态栏 ==========
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # ========== 底部按钮区域（动态计算固定高度） ==========
        self.bottom_frame = ttk.Frame(self.root, padding=10)
        self.bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)

        self.btn_export = ttk.Button(self.bottom_frame, text="开始筛选并输出 Excel", command=self.export_excel,
                                     bootstyle="info")
        self.btn_export.pack(side=tk.LEFT, padx=5)
        ttk.Button(self.bottom_frame, text="打开输出目录", command=self.open_output_dir).pack(side=tk.LEFT, padx=5)

        # 动态计算并固定高度
        self.root.update_idletasks()
        btn_req_height = self.btn_export.winfo_reqheight()
        bottom_frame_height = btn_req_height + 20 + 4  # padding 10+10 + 余量4
        self.bottom_frame.configure(height=bottom_frame_height)
        self.bottom_frame.pack_propagate(False)

        # ========== 配置管理区域（底部，自适应） ==========
        config_frame = ttk.Frame(self.root, padding=10)
        config_frame.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Label(config_frame, text="选择配置：", width=12, anchor='e').pack(side=tk.LEFT)
        self.combo_config = ttk.Combobox(config_frame, state="readonly", width=25)
        self.combo_config.pack(side=tk.LEFT, padx=5)
        self.combo_config.bind("<<ComboboxSelected>>", self.on_config_select)
        ttk.Button(config_frame, text="保存配置", command=self.save_config).pack(side=tk.LEFT, padx=2)
        ttk.Button(config_frame, text="另存为", command=self.save_config_as).pack(side=tk.LEFT, padx=2)
        ttk.Button(config_frame, text="删除配置", command=self.delete_config).pack(side=tk.LEFT, padx=2)

        # ========== 进度条 ==========
        self.progress_bar = ttk.Progressbar(self.root, mode='indeterminate', bootstyle='info', length=200)

        # ========== 中间列选择区域 ==========
        self.mid_frame = ttk.Frame(self.root, padding=10)
        self.mid_frame.pack(fill=tk.BOTH, expand=True)

        left_frame = ttk.Frame(self.mid_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Label(left_frame, text="所有列（可多选，Ctrl/Shift）").pack(anchor=tk.W)
        self.list_left = tk.Listbox(left_frame, selectmode=tk.EXTENDED, exportselection=False,
                                    height=15, font=self.default_font)
        scroll_left = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.list_left.yview)
        self.list_left.configure(yscrollcommand=scroll_left.set)
        self.list_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_left.pack(side=tk.RIGHT, fill=tk.Y)
        self.list_left.bind("<Double-Button-1>", lambda e: self.add_columns())

        btn_mid_frame = ttk.Frame(self.mid_frame, padding=10)
        btn_mid_frame.pack(side=tk.LEFT, fill=tk.Y)
        btn_width = 8
        ttk.Button(btn_mid_frame, text="全选", width=btn_width, bootstyle="secondary-outline",
                   command=self.select_all_left).pack(pady=2)
        ttk.Button(btn_mid_frame, text="反选", width=btn_width, bootstyle="secondary-outline",
                   command=self.invert_selection_left).pack(pady=2)
        ttk.Button(btn_mid_frame, text="添加 >", width=btn_width, bootstyle="outline-primary",
                   command=self.add_columns).pack(pady=10)
        ttk.Button(btn_mid_frame, text="< 移除", width=btn_width, bootstyle="outline-danger",
                   command=self.remove_columns).pack(pady=2)

        right_frame = ttk.Frame(self.mid_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Label(right_frame, text="已选列（输出顺序）").pack(anchor=tk.W)
        self.list_right = tk.Listbox(right_frame, selectmode=tk.EXTENDED, exportselection=False,
                                     height=15, font=self.default_font)
        scroll_right = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.list_right.yview)
        self.list_right.configure(yscrollcommand=scroll_right.set)
        self.list_right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_right.pack(side=tk.RIGHT, fill=tk.Y)
        self.list_right.bind("<Double-Button-1>", lambda e: self.remove_columns())

        btn_right_frame = ttk.Frame(self.mid_frame, padding=10)
        btn_right_frame.pack(side=tk.LEFT, fill=tk.Y)
        ttk.Button(btn_right_frame, text="上移", width=btn_width, command=self.move_up).pack(pady=2)
        ttk.Button(btn_right_frame, text="下移", width=btn_width, command=self.move_down).pack(pady=2)

    def toggle_multi_area(self):
        """展开/收起多文件合并区域，并动态调整窗口高度"""
        if self.multi_area_visible:
            self.multi_frame.pack_forget()
            self.btn_toggle_multi.config(text="展开多文件合并")
            self.multi_area_visible = False
            self.change_window_height(-220)
        else:
            self.multi_frame.pack(fill=tk.X)
            self.btn_toggle_multi.config(text="收起多文件合并")
            self.multi_area_visible = True
            if not self.multi_area_built:
                self.build_multi_area()
            self.change_window_height(220)

    def change_window_height(self, delta):
        """调整窗口高度，保持宽度不变"""
        try:
            self.root.update_idletasks()
            current_width = self.root.winfo_width()
            current_height = self.root.winfo_height()
            new_height = current_height + delta
            min_height = 750
            if new_height < min_height:
                new_height = min_height
            self.root.geometry(f"{current_width}x{new_height}")
        except Exception as e:
            print(f"调整窗口高度失败：{e}")

    def build_multi_area(self):
        """构建多文件合并区域 UI（首次调用时创建内部控件）"""
        if self.multi_area_built:
            return
        for widget in self.multi_frame.winfo_children():
            widget.destroy()

        container = ttk.Frame(self.multi_frame)
        container.pack(fill=tk.X, expand=True)

        btn_row = ttk.Frame(container)
        btn_row.pack(fill=tk.X, pady=5)
        ttk.Button(btn_row, text="添加文件", command=self.add_multi_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="合并并保存为中转文件", command=self.merge_multi_files, bootstyle="info").pack(side=tk.LEFT, padx=5)

        self.multi_files_canvas = tk.Canvas(container, height=150)
        scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=self.multi_files_canvas.yview)
        self.multi_files_canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.multi_files_canvas.pack(fill=tk.BOTH, expand=True)

        self.multi_files_canvas.bind("<MouseWheel>", self.on_mousewheel)

        self.multi_files_frame = ttk.Frame(self.multi_files_canvas)
        self.multi_files_canvas.create_window((0, 0), window=self.multi_files_frame, anchor='nw')
        self.multi_files_frame.bind("<Configure>", lambda e: self.multi_files_canvas.configure(scrollregion=self.multi_files_canvas.bbox("all")))

        self.multi_area_built = True

    def on_mousewheel(self, event):
        """鼠标滚轮滚动多文件列表"""
        if self.multi_files_canvas:
            self.multi_files_canvas.yview_scroll(int(-event.delta/120), "units")

    def add_multi_files(self):
        """添加多个文件到合并列表"""
        filetypes = [("Excel/CSV 文件", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")]
        paths = filedialog.askopenfilenames(title="选择多个文件", filetypes=filetypes)
        if not paths:
            return
        for path in paths:
            self.add_multi_file_entry(path)

    def add_multi_file_entry(self, path):
        """在合并区域添加一个文件条目，文件名靠左，控件靠右"""
        if not self.multi_area_built:
            self.build_multi_area()
        frame = ttk.Frame(self.multi_files_frame)
        frame.pack(fill=tk.X, padx=5, pady=2)

        full_name = os.path.basename(path)
        label = ttk.Label(frame, text=full_name, width=55, anchor='w')
        label.grid(row=0, column=0, sticky='w', padx=5)
        ToolTip(label, full_name)

        sheet_var = tk.StringVar(value="")
        sheet_combo = ttk.Combobox(frame, textvariable=sheet_var, state="readonly", width=15)
        sheet_combo.grid(row=0, column=1, padx=5)

        header_var = tk.IntVar(value=1)
        header_spin = ttk.Spinbox(frame, from_=1, to=1000, width=5, textvariable=header_var)
        header_spin.grid(row=0, column=2, padx=5)

        remove_btn = ttk.Button(frame, text="移除", bootstyle="outline-danger",
                                command=lambda: self.remove_multi_file(path))
        remove_btn.grid(row=0, column=3, padx=5)

        ext = os.path.splitext(path)[1].lower()
        sheet_names = []
        if ext == ".csv":
            sheet_names = ["单表"]
            sheet_combo.config(state="disabled")
            sheet_var.set("单表")
        else:
            try:
                engine = "openpyxl" if ext == ".xlsx" else "xlrd"
                xls = pd.ExcelFile(path, engine=engine)
                sheet_names = xls.sheet_names
                sheet_combo.config(state="readonly")
                if sheet_names:
                    sheet_var.set(sheet_names[0])
            except Exception as e:
                messagebox.showerror("错误", f"读取文件 {os.path.basename(path)} 的 sheet 失败：{str(e)}")
                frame.destroy()
                return
        sheet_combo['values'] = sheet_names

        for child in frame.winfo_children():
            child.bind("<MouseWheel>", self.on_mousewheel)
        frame.bind("<MouseWheel>", self.on_mousewheel)

        self.multi_files.append({
            "path": path,
            "frame": frame,
            "sheet_var": sheet_var,
            "header_var": header_var,
        })

    def remove_multi_file(self, path):
        """从合并列表移除指定文件"""
        for i, item in enumerate(self.multi_files):
            if item["path"] == path:
                item["frame"].destroy()
                del self.multi_files[i]
                break

    def merge_multi_files(self):
        """合并所有已添加的文件，保存为中转文件并加载"""
        if not self.multi_files:
            messagebox.showwarning("提示", "请先添加至少一个文件")
            return
        self.status_var.set("正在合并文件...")
        self.root.update()
        dataframes = []
        first_original_cols = None
        first_file_name = ""
        for item in self.multi_files:
            path = item["path"]
            sheet = item["sheet_var"].get()
            header_row = item["header_var"].get()
            try:
                ext = os.path.splitext(path)[1].lower()
                if ext == ".csv":
                    raw_data = b""
                    with open(path, "rb") as f:
                        for _ in range(10):
                            line = f.readline()
                            if not line:
                                break
                            raw_data += line
                    enc = None
                    for candidate in ['utf-8-sig', 'gbk', 'gb18030', 'latin1']:
                        try:
                            raw_data.decode(candidate)
                            enc = candidate
                            break
                        except UnicodeDecodeError:
                            continue
                    if enc is None:
                        raise ValueError("无法识别 CSV 文件编码")
                    df = pd.read_csv(path, header=header_row-1, encoding=enc)
                else:
                    engine = "openpyxl" if ext == ".xlsx" else "xlrd"
                    df = pd.read_excel(path, sheet_name=sheet, header=header_row-1, engine=engine)

                original_cols = list(df.columns)
                if first_original_cols is None:
                    first_original_cols = original_cols
                    first_file_name = os.path.basename(path)
                else:
                    if original_cols != first_original_cols:
                        diff_details = []
                        max_len = max(len(first_original_cols), len(original_cols))
                        count = 0
                        for i in range(max_len):
                            col_a = first_original_cols[i] if i < len(first_original_cols) else "<缺失>"
                            col_b = original_cols[i] if i < len(original_cols) else "<缺失>"
                            if col_a != col_b:
                                diff_details.append(f"第{i+1}列：{first_file_name}为'{col_a}'，{os.path.basename(path)}为'{col_b}'")
                                count += 1
                                if count >= 3:
                                    break
                        error_msg = "列不一致，具体差异：\n" + "\n".join(diff_details)
                        if max_len > 3 and count >= 3:
                            error_msg += "\n（其余差异已省略）"
                        raise ValueError(error_msg)
                dataframes.append(df)
            except Exception as e:
                self.status_var.set("合并失败")
                messagebox.showerror("错误", f"合并失败：{str(e)}")
                return

        if not dataframes:
            messagebox.showerror("错误", "没有可合并的数据")
            return

        combined_df = pd.concat(dataframes, axis=0, ignore_index=True)
        first_file_dir = os.path.dirname(self.multi_files[0]["path"])
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        merged_file_path = os.path.join(first_file_dir, f"合并结果_{timestamp}.xlsx")
        try:
            combined_df.to_excel(merged_file_path, index=False)
        except Exception as e:
            self.status_var.set("保存合并文件失败")
            messagebox.showerror("错误", f"保存合并文件失败：{str(e)}")
            return

        self.file_path = merged_file_path
        self.entry_file.delete(0, tk.END)
        self.entry_file.insert(0, self.file_path)
        self.entry_output.delete(0, tk.END)
        self.entry_output.insert(0, first_file_dir)
        self.spin_header.set(1)
        self.toggle_multi_area()
        self.load_file()
        self.status_var.set(f"合并完成，中转文件已保存：{merged_file_path}")
        messagebox.showinfo("成功", f"合并完成，中转文件已保存至：\n{merged_file_path}")

    # ---------- 文件操作 ----------
    def browse_file(self):
        filetypes = [("Excel/CSV 文件", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")]
        path = filedialog.askopenfilename(title="选择导入文件", filetypes=filetypes)
        if path:
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, path)
            self.file_path = path
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, os.path.dirname(path))
            self.load_file()

    def browse_output(self):
        path = filedialog.askdirectory(title="选择导出目录")
        if path:
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, path)

    # ---------- 数据加载 ----------
    def load_file(self):
        """加载单文件（根据 self.file_path）"""
        if not self.file_path:
            return
        ext = os.path.splitext(self.file_path)[1].lower()
        self.status_var.set("正在加载文件...")
        self.root.update()
        try:
            if ext == ".csv":
                self.sheet_names = ["单表"]
                self.combo_sheet['values'] = self.sheet_names
                self.combo_sheet.current(0)
                self.current_sheet = "单表"
                self.combo_sheet.config(state="disabled")
                self.read_headers()
            else:
                engine = "openpyxl" if ext == ".xlsx" else "xlrd"
                xls = pd.ExcelFile(self.file_path, engine=engine)
                self.sheet_names = xls.sheet_names
                self.combo_sheet['values'] = self.sheet_names
                self.combo_sheet.config(state="readonly")
                if self.sheet_names:
                    self.combo_sheet.current(0)
                    self.current_sheet = self.sheet_names[0]
                    self.read_headers()
                else:
                    messagebox.showerror("错误", "文件中没有工作表")
            self.status_var.set("文件加载完成")
        except Exception as e:
            self.status_var.set("加载失败")
            messagebox.showerror("错误", f"文件加载失败：{str(e)}\n{traceback.format_exc()}")

    def read_headers(self):
        """读取当前 sheet 和表头行，刷新列信息"""
        if not self.file_path:
            return
        try:
            header_row = int(self.spin_header.get())
            if header_row < 1:
                raise ValueError("表头行号必须大于0")
        except ValueError:
            messagebox.showerror("错误", "请输入有效的表头行号")
            return

        self.header_row = header_row
        ext = os.path.splitext(self.file_path)[1].lower()
        self.status_var.set("正在读取表头...")
        self.root.update()
        try:
            if ext == ".csv":
                raw_data = b""
                with open(self.file_path, "rb") as f:
                    for _ in range(10):
                        line = f.readline()
                        if not line:
                            break
                        raw_data += line
                enc = None
                for candidate in ['utf-8-sig', 'gbk', 'gb18030', 'latin1']:
                    try:
                        raw_data.decode(candidate)
                        enc = candidate
                        break
                    except UnicodeDecodeError:
                        continue
                if enc is None:
                    raise ValueError("无法识别 CSV 文件编码")
                df_head = pd.read_csv(self.file_path, header=header_row-1, nrows=0, encoding=enc)
                self.df = pd.read_csv(self.file_path, header=header_row-1, encoding=enc)
            else:
                engine = "openpyxl" if ext == ".xlsx" else "xlrd"
                df_head = pd.read_excel(self.file_path, sheet_name=self.current_sheet,
                                        header=header_row-1, nrows=0, engine=engine)
                self.df = pd.read_excel(self.file_path, sheet_name=self.current_sheet,
                                        header=header_row-1, engine=engine)

            if len(self.df) == 0:
                messagebox.showerror("错误", "表头行号超出数据范围或文件无有效数据")
                self.df = None
                return

            self.df, self.columns_info = normalize_columns(self.df)
            self.populate_left_listbox()
            self.selected_indices = []
            self.populate_right_listbox()
            self.status_var.set("表头读取完成")
        except Exception as e:
            self.status_var.set("读取表头失败")
            messagebox.showerror("错误", f"读取表头失败：{str(e)}\n{traceback.format_exc()}")

    def populate_left_listbox(self, filter_text=""):
        self.list_left.delete(0, tk.END)
        filter_lower = filter_text.lower()
        for idx, orig, display, letter in self.columns_info:
            show_text = f"{letter} - {display}"
            if filter_lower in show_text.lower():
                self.list_left.insert(tk.END, show_text)

    def filter_left_listbox(self, event=None):
        self.populate_left_listbox(self.entry_search.get())

    def populate_right_listbox(self):
        self.list_right.delete(0, tk.END)
        for idx in self.selected_indices:
            if 0 <= idx < len(self.columns_info):
                letter = self.columns_info[idx][3]
                display = self.columns_info[idx][2]
                self.list_right.insert(tk.END, f"{letter} - {display}")

    # ---------- 列表操作 ----------
    def select_all_left(self):
        self.list_left.selection_set(0, tk.END)

    def invert_selection_left(self):
        current = set(self.list_left.curselection())
        all_items = set(range(self.list_left.size()))
        new_selection = all_items - current
        self.list_left.selection_clear(0, tk.END)
        for i in new_selection:
            self.list_left.selection_set(i)

    def add_columns(self):
        selected = self.list_left.curselection()
        if not selected:
            messagebox.showinfo("提示", "请先在左侧选择列")
            return
        left_items = self.list_left.get(0, tk.END)
        display_to_index = {}
        for i, (idx, orig, disp, letter) in enumerate(self.columns_info):
            display_to_index[f"{letter} - {disp}"] = idx
        for sel in selected:
            text = left_items[sel]
            idx = display_to_index.get(text)
            if idx is not None and idx not in self.selected_indices:
                self.selected_indices.append(idx)
        self.populate_right_listbox()

    def remove_columns(self):
        selected = self.list_right.curselection()
        if not selected:
            messagebox.showinfo("提示", "请先在右侧选择列")
            return
        for sel in reversed(selected):
            if 0 <= sel < len(self.selected_indices):
                del self.selected_indices[sel]
        self.populate_right_listbox()

    def move_up(self):
        selected = self.list_right.curselection()
        if len(selected) != 1:
            messagebox.showinfo("提示", "请选择一项进行上移")
            return
        pos = selected[0]
        if pos > 0:
            self.selected_indices[pos], self.selected_indices[pos-1] = \
                self.selected_indices[pos-1], self.selected_indices[pos]
            self.populate_right_listbox()
            self.list_right.selection_set(pos-1)

    def move_down(self):
        selected = self.list_right.curselection()
        if len(selected) != 1:
            messagebox.showinfo("提示", "请选择一项进行下移")
            return
        pos = selected[0]
        if pos < len(self.selected_indices)-1:
            self.selected_indices[pos], self.selected_indices[pos+1] = \
                self.selected_indices[pos+1], self.selected_indices[pos]
            self.populate_right_listbox()
            self.list_right.selection_set(pos+1)

    # ---------- 事件处理 ----------
    def on_sheet_change(self, event=None):
        if not self.file_path:
            return
        self.current_sheet = self.combo_sheet.get()
        self.read_headers()

    def on_header_change(self, event=None):
        if not self.file_path:
            return
        self.read_headers()

    # ---------- 配置管理 ----------
    def load_configs(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.configs = data.get("configs", [])
            else:
                self.configs = []
        except Exception as e:
            self.configs = []
            self.status_var.set(f"配置文件读取失败：{str(e)}")

    def save_configs_to_file(self):
        try:
            os.makedirs(self.config_dir, exist_ok=True)
            data = {"configs": self.configs}
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self.status_var.set("配置已保存")
        except Exception as e:
            messagebox.showerror("错误", f"配置保存失败：{str(e)}")

    def populate_config_combobox(self):
        names = [cfg["name"] for cfg in self.configs]
        self.combo_config['values'] = names
        if names:
            self.combo_config.current(0)
            self.current_config_name = names[0]
        else:
            self.current_config_name = None

    def on_config_select(self, event=None):
        name = self.combo_config.get()
        for cfg in self.configs:
            if cfg["name"] == name:
                self.current_config_name = name
                self.apply_config(cfg)
                break

    def apply_config(self, cfg):
        if not self.file_path:
            messagebox.showwarning("提示", "请先导入文件再应用配置")
            return
        self.show_progress()
        try:
            if cfg.get("sheet") in self.sheet_names:
                self.combo_sheet.set(cfg["sheet"])
                self.current_sheet = cfg["sheet"]
            self.spin_header.set(cfg.get("header_row", 1))
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, cfg.get("output_dir", ""))

            saved_cols = cfg.get("selected_columns", [])
            self.read_headers()
            self.selected_indices = []
            for col in saved_cols:
                idx = col.get("index")
                matched = False
                if idx is not None and 0 <= idx < len(self.columns_info):
                    self.selected_indices.append(idx)
                    matched = True
                else:
                    col_name = col.get("name")
                    for i, (orig_idx, orig_name, disp, letter) in enumerate(self.columns_info):
                        if orig_name == col_name or disp == col_name:
                            self.selected_indices.append(i)
                            matched = True
                            break
                if not matched:
                    print(f"警告：配置中的列 '{col.get('name', '')}' 在当前数据中未找到")
            self.populate_right_listbox()
            self.status_var.set(f"已应用配置：{cfg['name']}")
        except Exception as e:
            messagebox.showerror("错误", f"应用配置失败：{str(e)}")
        finally:
            self.hide_progress()

    def show_progress(self):
        if self.progress_bar is None:
            return
        if not self.progress_bar.winfo_manager():
            self.progress_bar.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=2, before=self.bottom_frame)
        self.progress_bar.start(10)

    def hide_progress(self):
        if self.progress_bar is not None:
            self.progress_bar.stop()
            if self.progress_bar.winfo_manager():
                self.progress_bar.pack_forget()

    def save_config(self):
        if not self.file_path:
            messagebox.showwarning("提示", "请先导入文件并选择列")
            return
        if not self.selected_indices:
            messagebox.showwarning("提示", "请选择配置或至少选择一列")
            return
        if self.current_config_name and self.current_config_name in [c["name"] for c in self.configs]:
            for cfg in self.configs:
                if cfg["name"] == self.current_config_name:
                    self.update_config_dict(cfg)
                    break
            self.save_configs_to_file()
            self.populate_config_combobox()
            self.status_var.set(f"配置已更新：{self.current_config_name}")
        else:
            self.save_config_as()

    def save_config_as(self):
        if not self.file_path:
            messagebox.showwarning("提示", "请先导入文件并选择列")
            return
        if not self.selected_indices:
            messagebox.showwarning("提示", "请选择配置或至少选择一列")
            return
        name = simpledialog.askstring("保存配置", "请输入配置名称：", parent=self.root)
        if not name:
            return
        if name in [c["name"] for c in self.configs]:
            messagebox.showerror("错误", "配置名称已存在")
            return
        new_cfg = {}
        self.update_config_dict(new_cfg)
        new_cfg["name"] = name
        self.configs.append(new_cfg)
        self.save_configs_to_file()
        self.populate_config_combobox()
        self.combo_config.set(name)
        self.current_config_name = name
        self.status_var.set(f"配置已保存：{name}")

    def update_config_dict(self, cfg):
        cfg["sheet"] = self.current_sheet or self.combo_sheet.get()
        cfg["header_row"] = int(self.spin_header.get())
        cfg["output_dir"] = self.entry_output.get()
        cfg["selected_columns"] = []
        for idx in self.selected_indices:
            if 0 <= idx < len(self.columns_info):
                orig_name = self.columns_info[idx][1]
                display_name = self.columns_info[idx][2]
                cfg["selected_columns"].append({
                    "index": idx,
                    "name": orig_name if orig_name else display_name,
                })

    def delete_config(self):
        name = self.combo_config.get()
        if not name:
            return
        if messagebox.askyesno("确认删除", f"确定要删除配置“{name}”吗？"):
            self.configs = [c for c in self.configs if c["name"] != name]
            self.save_configs_to_file()
            self.populate_config_combobox()
            self.status_var.set(f"配置已删除：{name}")

    # ---------- 导出 ----------
    def export_excel(self):
        if not self.file_path:
            messagebox.showwarning("提示", "请先导入文件")
            return
        if not self.selected_indices:
            messagebox.showwarning("提示", "请选择配置或至少选择一列")
            return
        if self.df is None:
            messagebox.showerror("错误", "数据未加载，请重新导入文件")
            return
        output_dir = self.entry_output.get().strip()
        if not output_dir:
            output_dir = os.path.dirname(self.file_path)
            if not output_dir:
                messagebox.showerror("错误", "请指定导出路径")
                return
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                messagebox.showerror("错误", f"导出路径不存在且无法创建：{str(e)}")
                return
        base = os.path.splitext(os.path.basename(self.file_path))[0]
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(output_dir, f"{base}_筛选结果_{timestamp}.xlsx")

        self.status_var.set("正在导出...")
        self.root.update()
        try:
            selected_columns = [self.columns_info[i][0] for i in self.selected_indices]
            df_out = self.df.iloc[:, selected_columns].copy()
            original_names = [self.columns_info[i][1] for i in self.selected_indices]
            display_names = [self.columns_info[i][2] for i in self.selected_indices]
            final_names = [orig if orig != "" else disp for orig, disp in zip(original_names, display_names)]
            df_out.columns = final_names
            df_out.to_excel(output_file, index=False)
            self.status_var.set(f"导出成功：{output_file}")
            if messagebox.askyesno("导出成功", f"文件已导出至：\n{output_file}\n\n是否打开文件？"):
                os.startfile(output_file)
        except Exception as e:
            self.status_var.set("导出失败")
            messagebox.showerror("错误", f"导出失败：{str(e)}\n{traceback.format_exc()}")

    def open_output_dir(self):
        output_dir = self.entry_output.get().strip()
        if not output_dir:
            if self.file_path:
                output_dir = os.path.dirname(self.file_path)
            else:
                messagebox.showwarning("提示", "请先导入文件或指定导出路径")
                return
        if os.path.exists(output_dir):
            os.startfile(output_dir)
        else:
            messagebox.showwarning("提示", "输出目录不存在")

    def on_closing(self):
        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = App()
    app.run()
