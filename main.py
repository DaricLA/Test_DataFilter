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

# 尝试导入 openpyxl 和 xlrd，打包时需包含
try:
    import openpyxl  # noqa: F401
except ImportError:
    pass
try:
    import xlrd  # noqa: F401
except ImportError:
    pass


def get_excel_column_letter(n):
    """将数字索引（0-based）转换为 Excel 列字母（A, B, ...）"""
    n += 1
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def get_config_dir():
    """获取配置文件目录，优先使用 exe 同目录/config，失败则请求用户选择目录"""
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    config_dir = os.path.join(base_dir, "config")
    try:
        os.makedirs(config_dir, exist_ok=True)
        # 测试写入权限
        test_file = os.path.join(config_dir, ".write_test")
        with open(test_file, "w") as f:
            f.write("test")
        os.remove(test_file)
        return config_dir
    except Exception:
        # 无法创建或写入，请求用户选择目录
        messagebox.showwarning("配置目录不可用",
                               f"无法在程序目录创建配置文件夹：\n{config_dir}\n\n"
                               "请选择一个可写的目录用于保存配置。")
        chosen = filedialog.askdirectory(title="选择配置保存目录")
        if chosen:
            return chosen
        else:
            # 用户取消，使用临时目录
            import tempfile
            return tempfile.gettempdir()


class App:
    def __init__(self):
        self.root = ttk.Window(themename="flatly")
        self.root.title("Excel 列筛选工具")
        self.root.geometry("1000x700")
        self.root.minsize(900, 600)

        # 字体设置
        self.default_font = ("Microsoft YaHei UI", 10)
        self.root.option_add("*Font", self.default_font)
        style = ttk.Style()
        style.configure(".", font=self.default_font)
        style.configure("Treeview", font=self.default_font)
        style.configure("TButton", font=self.default_font)

        # 配置目录
        self.config_dir = get_config_dir()
        self.config_file = os.path.join(self.config_dir, "configs.json")
        self.configs = []          # 配置列表，元素为 dict
        self.current_config_name = None  # 当前选中的配置名

        # 数据相关
        self.file_path = None
        self.df = None             # 当前完整数据（包含所有列）
        self.sheet_names = []
        self.current_sheet = None
        self.header_row = 1
        self.columns_info = []     # 存储列信息：[(index, original_name, display_name, excel_letter)]
        self.selected_indices = [] # 右侧已选列的索引（0-based）

        # 构建界面
        self.build_ui()

        # 加载配置
        self.load_configs()
        self.populate_config_combobox()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    # ---------- UI 构建 ----------
    def build_ui(self):
        # 顶部文件选择
        top_frame = ttk.Frame(self.root, padding=10)
        top_frame.pack(fill=tk.X)

        # 导入文件行
        row1 = ttk.Frame(top_frame)
        row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text="导入文件：").pack(side=tk.LEFT)
        self.entry_file = ttk.Entry(row1, width=60)
        self.entry_file.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(row1, text="浏览...", command=self.browse_file).pack(side=tk.LEFT, padx=2)

        # 导出路径行
        row2 = ttk.Frame(top_frame)
        row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text="导出路径：").pack(side=tk.LEFT)
        self.entry_output = ttk.Entry(row2, width=60)
        self.entry_output.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(row2, text="浏览...", command=self.browse_output).pack(side=tk.LEFT, padx=2)
        ttk.Button(row2, text="默认导入目录", command=self.set_output_to_source).pack(side=tk.LEFT, padx=2)

        # Sheet 和表头行
        row3 = ttk.Frame(top_frame)
        row3.pack(fill=tk.X, pady=5)
        ttk.Label(row3, text="Sheet：").pack(side=tk.LEFT)
        self.combo_sheet = ttk.Combobox(row3, state="readonly", width=20)
        self.combo_sheet.pack(side=tk.LEFT, padx=5)
        self.combo_sheet.bind("<<ComboboxSelected>>", self.on_sheet_change)
        ttk.Label(row3, text="表头行：").pack(side=tk.LEFT, padx=(15, 0))
        self.spin_header = ttk.Spinbox(row3, from_=1, to=1000, width=5)
        self.spin_header.set(1)
        self.spin_header.pack(side=tk.LEFT, padx=5)
        self.spin_header.bind("<Return>", self.on_header_change)
        self.spin_header.bind("<FocusOut>", self.on_header_change)

        # 搜索列
        row4 = ttk.Frame(top_frame)
        row4.pack(fill=tk.X, pady=2)
        ttk.Label(row4, text="搜索列：").pack(side=tk.LEFT)
        self.entry_search = ttk.Entry(row4, width=30)
        self.entry_search.pack(side=tk.LEFT, padx=5)
        self.entry_search.bind("<KeyRelease>", self.filter_left_listbox)

        # 中间列选择区域
        mid_frame = ttk.Frame(self.root, padding=10)
        mid_frame.pack(fill=tk.BOTH, expand=True)

        # 左侧列表
        left_frame = ttk.Frame(mid_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Label(left_frame, text="所有列（可多选，Ctrl/Shift）").pack(anchor=tk.W)
        self.list_left = tk.Listbox(left_frame, selectmode=tk.EXTENDED, exportselection=False,
                                    height=15, font=self.default_font)
        scroll_left = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.list_left.yview)
        self.list_left.configure(yscrollcommand=scroll_left.set)
        self.list_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_left.pack(side=tk.RIGHT, fill=tk.Y)

        # 中间按钮列
        btn_mid_frame = ttk.Frame(mid_frame, padding=10)
        btn_mid_frame.pack(side=tk.LEFT, fill=tk.Y)
        ttk.Button(btn_mid_frame, text="全选", command=self.select_all_left).pack(pady=2)
        ttk.Button(btn_mid_frame, text="反选", command=self.invert_selection_left).pack(pady=2)
        ttk.Button(btn_mid_frame, text="添加 >", command=self.add_columns).pack(pady=10)
        ttk.Button(btn_mid_frame, text="< 移除", command=self.remove_columns).pack(pady=2)

        # 右侧列表
        right_frame = ttk.Frame(mid_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Label(right_frame, text="已选列（输出顺序）").pack(anchor=tk.W)
        self.list_right = tk.Listbox(right_frame, selectmode=tk.EXTENDED, exportselection=False,
                                     height=15, font=self.default_font)
        scroll_right = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.list_right.yview)
        self.list_right.configure(yscrollcommand=scroll_right.set)
        self.list_right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_right.pack(side=tk.RIGHT, fill=tk.Y)

        # 右侧按钮列（上下移动）
        btn_right_frame = ttk.Frame(mid_frame, padding=10)
        btn_right_frame.pack(side=tk.LEFT, fill=tk.Y)
        ttk.Button(btn_right_frame, text="上移", command=self.move_up).pack(pady=2)
        ttk.Button(btn_right_frame, text="下移", command=self.move_down).pack(pady=2)

        # 配置管理区域
        config_frame = ttk.Frame(self.root, padding=10)
        config_frame.pack(fill=tk.X)
        ttk.Label(config_frame, text="配置：").pack(side=tk.LEFT)
        self.combo_config = ttk.Combobox(config_frame, state="readonly", width=25)
        self.combo_config.pack(side=tk.LEFT, padx=5)
        self.combo_config.bind("<<ComboboxSelected>>", self.on_config_select)
        ttk.Button(config_frame, text="保存配置", command=self.save_config).pack(side=tk.LEFT, padx=2)
        ttk.Button(config_frame, text="另存为", command=self.save_config_as).pack(side=tk.LEFT, padx=2)
        ttk.Button(config_frame, text="删除配置", command=self.delete_config).pack(side=tk.LEFT, padx=2)

        # 底部操作按钮
        bottom_frame = ttk.Frame(self.root, padding=10)
        bottom_frame.pack(fill=tk.X)
        self.btn_export = ttk.Button(bottom_frame, text="开始筛选并输出 Excel", command=self.export_excel,
                                     bootstyle="success")
        self.btn_export.pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_frame, text="打开输出目录", command=self.open_output_dir).pack(side=tk.LEFT, padx=5)

        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    # ---------- 文件操作 ----------
    def browse_file(self):
        filetypes = [("Excel/CSV 文件", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")]
        path = filedialog.askopenfilename(title="选择导入文件", filetypes=filetypes)
        if path:
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, path)
            self.file_path = path
            # 自动设置默认导出路径为导入文件目录
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, os.path.dirname(path))
            self.load_file()

    def browse_output(self):
        path = filedialog.askdirectory(title="选择导出目录")
        if path:
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, path)

    def set_output_to_source(self):
        if self.file_path:
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, os.path.dirname(self.file_path))
        else:
            messagebox.showwarning("提示", "请先导入文件")

    # ---------- 数据加载 ----------
    def load_file(self):
        """加载文件并获取 sheet 列表和初始表头"""
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
                # 读取所有 sheet 名
                if ext == ".xlsx":
                    engine = "openpyxl"
                elif ext == ".xls":
                    engine = "xlrd"
                else:
                    raise ValueError("不支持的文件格式")
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
        """根据当前 sheet 和表头行读取列信息，刷新左侧列表"""
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
                # 改进的编码检测：读取前 10 行判断
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

            # 检查表头行号是否有效（数据行数至少为 1）
            if len(self.df) == 0:
                messagebox.showerror("错误", "表头行号超出数据范围或文件无有效数据")
                self.df = None
                return

            # 获取列信息
            raw_names = list(df_head.columns)
            self.columns_info = []
            seen = {}
            for i, raw in enumerate(raw_names):
                # 处理空列名
                if pd.isna(raw) or str(raw).strip() == "":
                    original = ""
                    base = f"列{i+1}"
                else:
                    original = str(raw).strip()
                    base = original
                # 处理重复
                if base in seen:
                    seen[base] += 1
                    display = f"{base}_{seen[base]}"
                else:
                    seen[base] = 0
                    display = base
                excel_letter = get_excel_column_letter(i)
                self.columns_info.append((i, original, display, excel_letter))

            # 刷新左侧列表
            self.populate_left_listbox()
            # 清空右侧已选列（因为列可能变化）
            self.selected_indices = []
            self.populate_right_listbox()
            self.status_var.set("表头读取完成")
        except Exception as e:
            self.status_var.set("读取表头失败")
            messagebox.showerror("错误", f"读取表头失败：{str(e)}\n{traceback.format_exc()}")

    def populate_left_listbox(self, filter_text=""):
        """刷新左侧列表，支持搜索过滤"""
        self.list_left.delete(0, tk.END)
        filter_lower = filter_text.lower()
        for idx, orig, display, letter in self.columns_info:
            show_text = f"{letter} - {display}"
            if filter_lower in show_text.lower():
                self.list_left.insert(tk.END, show_text)

    def filter_left_listbox(self, event=None):
        text = self.entry_search.get()
        self.populate_left_listbox(text)

    def populate_right_listbox(self):
        """刷新右侧列表，显示当前已选列顺序"""
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
        """将左侧选中的列添加到右侧"""
        selected = self.list_left.curselection()
        if not selected:
            messagebox.showinfo("提示", "请先在左侧选择列")
            return
        # 获取左侧当前显示内容对应的索引（因为过滤后索引可能变化）
        left_items = self.list_left.get(0, tk.END)
        # 建立显示文本到 column_info 索引的映射
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
        """将右侧选中的列移除"""
        selected = self.list_right.curselection()
        if not selected:
            messagebox.showinfo("提示", "请先在右侧选择列")
            return
        for sel in reversed(selected):  # 从后往前删除
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
        self.read_headers()

    # ---------- 配置管理 ----------
    def load_configs(self):
        """从配置文件加载所有配置"""
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
        """将当前配置列表写入文件"""
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
        """应用配置到界面"""
        try:
            # 1. 设置 sheet（如果存在）
            if cfg.get("sheet") in self.sheet_names:
                self.combo_sheet.set(cfg["sheet"])
                self.current_sheet = cfg["sheet"]
            # 2. 设置表头行
            self.spin_header.set(cfg.get("header_row", 1))
            # 3. 设置导出路径
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, cfg.get("output_dir", ""))

            # 4. 暂存配置中的列信息，待读取表头后匹配
            saved_cols = cfg.get("selected_columns", [])

            # 5. 重新读取表头（会清空 selected_indices，但这是必要的）
            self.read_headers()

            # 6. 根据配置匹配列
            self.selected_indices = []
            for col in saved_cols:
                idx = col.get("index")
                matched = False
                # 优先按索引匹配
                if idx is not None and 0 <= idx < len(self.columns_info):
                    self.selected_indices.append(idx)
                    matched = True
                else:
                    # 索引无效，尝试按列名匹配
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

    def save_config(self):
        """保存当前配置（覆盖当前选中或另存为）"""
        if not self.file_path:
            messagebox.showwarning("提示", "请先导入文件并选择列")
            return
        if not self.selected_indices:
            messagebox.showwarning("提示", "请至少选择一列")
            return
        if self.current_config_name and self.current_config_name in [c["name"] for c in self.configs]:
            # 覆盖
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
        """另存为新配置"""
        if not self.file_path:
            messagebox.showwarning("提示", "请先导入文件并选择列")
            return
        if not self.selected_indices:
            messagebox.showwarning("提示", "请至少选择一列")
            return
        # 弹出输入框输入配置名
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
        """将当前界面状态写入配置字典"""
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
            self.current_config_name = None
            self.status_var.set(f"配置已删除：{name}")

    # ---------- 导出 ----------
    def export_excel(self):
        if not self.file_path:
            messagebox.showwarning("提示", "请先导入文件")
            return
        if not self.selected_indices:
            messagebox.showwarning("提示", "请至少选择一列")
            return
        if self.df is None:
            messagebox.showerror("错误", "数据未加载，请重新导入文件")
            return
        output_dir = self.entry_output.get().strip()
        if not output_dir:
            if self.file_path:
                output_dir = os.path.dirname(self.file_path)
            else:
                messagebox.showerror("错误", "请指定导出路径")
                return
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                messagebox.showerror("错误", f"导出路径不存在且无法创建：{str(e)}")
                return
        # 生成输出文件名
        base = os.path.splitext(os.path.basename(self.file_path))[0]
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(output_dir, f"{base}_筛选结果_{timestamp}.xlsx")

        self.status_var.set("正在导出...")
        self.root.update()
        try:
            # 选择列
            selected_columns = [self.columns_info[i][0] for i in self.selected_indices]
            df_out = self.df.iloc[:, selected_columns].copy()
            # 输出
            df_out.to_excel(output_file, index=False)
            self.status_var.set(f"导出成功：{output_file}")
            messagebox.showinfo("成功", f"文件已导出至：\n{output_file}")
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
