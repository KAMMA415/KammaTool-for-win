import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from tkinter.constants import *
import os
import sys
import threading
from pathlib import Path
import json
import queue  # 引入线程安全队列
import re
from datetime import datetime
from PIL import Image, ImageTk

# 导入各个处理器
from excel_processor import ExcelProcessor
from file_processor import FileProcessor
from word_processor import WordProcessor
from pdf_exporter import PDFExporter, ExcelSplitter
from version import APP_VERSION

class ConfigManager:
    """配置管理器，用于保存和加载用户最后使用的路径"""
    def __init__(self, config_file="config.json"):
        self.config_file = config_file
        self.config = self._load()

    def _load(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return {}
        return {}

    def save_config(self):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
        except:
            pass

    def get(self, key, default=None):
        return self.config.get(key, default)

    def set(self, key, value):
        self.config[key] = value

class DongToolGUI:
    DEFAULT_WINDOW_SIZE = (1180, 760)
    MIN_WINDOW_SIZE = (980, 640)
    MAX_STARTUP_SIZE = (1280, 820)

    def _parse_geometry(self, geometry):
        match = re.match(r"^(\d+)x(\d+)([+-]\d+)?([+-]\d+)?$", str(geometry or ""))
        if not match:
            return None
        width, height, x_pos, y_pos = match.groups()
        return int(width), int(height), int(x_pos or 100), int(y_pos or 100)

    def _screen_limited_size(self):
        screen_width = max(self.root.winfo_screenwidth(), self.MIN_WINDOW_SIZE[0])
        screen_height = max(self.root.winfo_screenheight(), self.MIN_WINDOW_SIZE[1])
        max_width = min(self.MAX_STARTUP_SIZE[0], max(self.MIN_WINDOW_SIZE[0], screen_width - 80))
        max_height = min(self.MAX_STARTUP_SIZE[1], max(self.MIN_WINDOW_SIZE[1], screen_height - 100))
        default_width = min(self.DEFAULT_WINDOW_SIZE[0], max_width)
        default_height = min(self.DEFAULT_WINDOW_SIZE[1], max_height)
        return default_width, default_height, max_width, max_height

    def _center_geometry(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x_pos = max((screen_width - width) // 2, 0)
        y_pos = max((screen_height - height) // 2, 0)
        return f"{width}x{height}+{x_pos}+{y_pos}"

    def _clamp_window_position(self, width, height, x_pos, y_pos):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        max_x = max(screen_width - width - 20, 0)
        max_y = max(screen_height - height - 60, 0)
        return min(max(x_pos, 0), max_x), min(max(y_pos, 0), max_y)

    def _safe_startup_geometry(self, geometry=None):
        default_width, default_height, max_width, max_height = self._screen_limited_size()
        parsed = self._parse_geometry(geometry)
        if parsed:
            width, height, x_pos, y_pos = parsed
            size_is_safe = (
                self.MIN_WINDOW_SIZE[0] <= width <= max_width and
                self.MIN_WINDOW_SIZE[1] <= height <= max_height
            )
            if size_is_safe:
                x_pos, y_pos = self._clamp_window_position(width, height, x_pos, y_pos)
                return f"{width}x{height}+{x_pos}+{y_pos}"
        return self._center_geometry(default_width, default_height)

    def __init__(self, root):
        self.root = root
        self.root.title("KammaTool for win")
        config_file = "config.json"
        if getattr(sys, "frozen", False):
            config_file = Path(sys.executable).with_name(config_file)
        self.config = ConfigManager(config_file)
        
        # 加载安全范围内的窗口大小，避免上次关闭时过大的尺寸被反复恢复
        self.root.minsize(*self.MIN_WINDOW_SIZE)
        self.root.geometry(self._safe_startup_geometry(self.config.get('window_geometry')))
        
        # 设置窗口图标
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "Dongdong tool 2.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
                # 为 Logo 准备 PhotoImage
                logo_img = Image.open(icon_path).resize((32, 32), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_img)
            
            # 加载文件工具专用图标
            self.file_icons = {}
            for name in ['rename', 'pdf', 'merge', 'organize']:
                p = os.path.join(os.path.dirname(__file__), f"icon_{name}.png")
                if os.path.exists(p):
                    img = Image.open(p).resize((24, 24), Image.Resampling.LANCZOS)
                    self.file_icons[name] = ImageTk.PhotoImage(img)
        except Exception as e:
            print(f"加载图标失败: {e}")

        # 线程安全的消息队列（必须在处理器初始化之前创建）
        self.msg_queue = queue.Queue()
        
        # 初始化各个处理器
        self.excel_processor = ExcelProcessor()
        self.word_processor = WordProcessor()
        self.file_processor = FileProcessor(self.msg_queue, self.config)
        self.pdf_exporter = PDFExporter()
        self.excel_splitter = ExcelSplitter()
        
        self._create_widgets()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 启动消息队列处理循环
        self.process_msg_queue()
        
    def _create_widgets(self):
        # 顶部工作区选择
        self.workspace_path = tk.StringVar()

        # --- 顶部 Header ---
        header_frame = ttk.Frame(self.root, bootstyle="secondary", padding=10)
        header_frame.pack(fill=X)

        if hasattr(self, 'logo_photo'):
            logo_label = ttk.Label(header_frame, image=self.logo_photo, bootstyle="secondary")
            logo_label.pack(side=LEFT, padx=(5, 10))

        title_label = ttk.Label(header_frame, text="KammaTool for win",
                               font=("微软雅黑", 16, "bold"), bootstyle="inverse-secondary")
        title_label.pack(side=LEFT)

        ver_label = ttk.Label(header_frame, text=f"v{APP_VERSION} 公开版",
                             font=("微软雅黑", 9), bootstyle="inverse-secondary")
        ver_label.pack(side=LEFT, padx=10, pady=(5, 0))

        # 创建主要内容区域：上方Notebook，下方处理日志控制台
        content_frame = ttk.Frame(self.root)
        content_frame.pack(fill=BOTH, expand=YES, padx=15, pady=(15, 0))

        # 上方 Notebook
        self.notebook = ttk.Notebook(content_frame)
        self.notebook.pack(fill=BOTH, expand=YES)

        # 创建各个Tab
        self.tab_file = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_file, text="📁文件工具")
        self._init_file_tab(self.tab_file)

        self.tab_excel = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_excel, text="📊EXCEL工具")
        self._init_excel_tab(self.tab_excel)

        self.tab_word = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_word, text="📝WORD工具")
        self._init_word_tab(self.tab_word)

        # 下方处理日志控制台（固定在窗口底部）
        console_frame = ttk.Labelframe(self.root, text="处理日志控制台", padding=8, bootstyle="info")
        console_frame.pack(fill=X, padx=15, pady=(6, 12))

        # 进度条
        self.progress = ttk.Progressbar(console_frame, mode='determinate', bootstyle="success-striped")
        self.progress.pack(fill=X, pady=(0, 6))

        # 日志文本框 + 滚动条
        log_area = ttk.Frame(console_frame)
        log_area.pack(fill=X)
        self.log_text = tk.Text(log_area, height=6, font=('Consolas', 10),
                               bg="#1e1e1e", fg="#d4d4d4",
                               insertbackground="white", selectbackground="#264f78")
        self.log_text.pack(side=LEFT, fill=X, expand=YES)

        scrollbar = ttk.Scrollbar(log_area, command=self.log_text.yview, bootstyle="info-round")
        scrollbar.pack(side=RIGHT, fill=Y)
        self.log_text.config(yscrollcommand=scrollbar.set)

        # 添加欢迎信息
        self.log_text.insert(END, "KammaTool for win 初始化完成\n")
        self.log_text.insert(END, "等待操作...\n")

    def _create_file_section(self, parent, title, icon_name=None, padding=15):
        """创建文件工具区块，图标放在标题栏中，避免被边框裁切。"""
        title_holder = ttk.Frame(parent)
        if icon_name in getattr(self, 'file_icons', {}):
            ttk.Label(title_holder, image=self.file_icons[icon_name]).pack(side=LEFT, padx=(0, 6))
        ttk.Label(title_holder, text=title, font=("微软雅黑", 10, "bold")).pack(side=LEFT)
        return ttk.Labelframe(parent, labelwidget=title_holder, padding=padding)

    def _init_file_tab(self, parent):
        """文件工具标签页 - 重构为双栏布局"""
        container = ttk.Frame(parent, padding=15)
        container.pack(fill=BOTH, expand=YES)
        
        # 左右双栏
        left_col = ttk.Frame(container)
        left_col.pack(side=LEFT, fill=BOTH, expand=YES, padx=(0, 10))
        
        right_col = ttk.Frame(container)
        right_col.pack(side=LEFT, fill=BOTH, expand=YES)

        # --- 第一列 (Left) ---

        # 1. 批量重命名
        rename_frame = self._create_file_section(left_col, "批量重命名", "rename")
        rename_frame.pack(fill=X, pady=(0, 15))
        rename_frame.columnconfigure(1, weight=1)
        
        ttk.Label(rename_frame, text="目标目录:").grid(row=0, column=0, padx=5, pady=5, sticky=W)
        self.rename_dir = ttk.Entry(rename_frame, width=40)
        self.rename_dir.grid(row=0, column=1, padx=5, pady=5, sticky=EW)
        ttk.Button(rename_frame, text="浏览", command=lambda: self._select_dir(self.rename_dir), bootstyle="outline").grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(rename_frame, text="名称前缀:").grid(row=1, column=0, padx=5, pady=5, sticky=W)
        self.rename_prefix = ttk.Entry(rename_frame, width=15)
        self.rename_prefix.grid(row=1, column=1, sticky=W, padx=5, pady=5)
        ttk.Button(rename_frame, text="执行重命名", command=self._do_rename, bootstyle="primary").grid(row=1, column=2, padx=5, pady=5)

        ttk.Label(rename_frame, text="删除内容:").grid(row=2, column=0, padx=5, pady=5, sticky=W)
        self.rename_delete_text = ttk.Entry(rename_frame, width=15)
        self.rename_delete_text.grid(row=2, column=1, sticky=W, padx=5, pady=5)
        ttk.Button(rename_frame, text="删除指定内容", command=self._do_delete_filename_content,
                   bootstyle="danger-outline").grid(row=2, column=2, padx=5, pady=5)

        # 2. 批量转PDF
        pdf_frame = self._create_file_section(left_col, "批量转PDF", "pdf")
        pdf_frame.pack(fill=X, pady=15)
        pdf_frame.columnconfigure(1, weight=1)
        
        ttk.Label(pdf_frame, text="目标目录:").grid(row=0, column=0, padx=5, pady=5, sticky=W)
        self.pdf_dir = ttk.Entry(pdf_frame, width=40)
        self.pdf_dir.insert(0, self.config.get('last_pdf_folder', ''))
        self.pdf_dir.grid(row=0, column=1, padx=5, pady=5, sticky=EW)
        ttk.Button(pdf_frame, text="浏览", command=lambda: self._select_dir(self.pdf_dir), bootstyle="outline").grid(row=0, column=2, padx=5, pady=5)
        
        self.pdf_subdirs = tk.BooleanVar(value=self.config.get('pdf_subdirs', True))
        ttk.Checkbutton(pdf_frame, text="包含子文件夹", variable=self.pdf_subdirs).grid(row=1, column=1, sticky=W, padx=5, pady=5)
        ttk.Button(pdf_frame, text="执行转换", command=self._do_convert_pdf, bootstyle="success").grid(row=1, column=2, padx=5, pady=5)

        # 3. 合并PDF
        merge_frame = self._create_file_section(left_col, "合并PDF", "merge")
        merge_frame.pack(fill=X, pady=15)
        merge_frame.columnconfigure(1, weight=1)
        
        ttk.Label(merge_frame, text="目标目录:").grid(row=0, column=0, padx=5, pady=5, sticky=W)
        self.merge_dir = ttk.Entry(merge_frame, width=40)
        self.merge_dir.insert(0, self.config.get('last_merge_folder', ''))
        self.merge_dir.grid(row=0, column=1, padx=5, pady=5, sticky=EW)
        ttk.Button(merge_frame, text="浏览", command=lambda: self._select_dir(self.merge_dir), bootstyle="outline").grid(row=0, column=2, padx=5, pady=5)
        ttk.Button(merge_frame, text="执行合并", command=self._do_merge_pdf, bootstyle="info").grid(row=1, column=2, padx=5, pady=5)

        # --- 第二列 (Right) ---

        # 4. 文件整理归档
        org_frame = self._create_file_section(right_col, "文件整理归档", "organize", padding=16)
        org_frame.pack(fill=BOTH, expand=YES)
        org_frame.columnconfigure(0, weight=1)
        
        # 1. 目录选择
        dir_frame = ttk.Frame(org_frame)
        dir_frame.grid(row=0, column=0, sticky=EW, pady=(0, 12))
        dir_frame.columnconfigure(1, weight=1)
        ttk.Label(dir_frame, text="源目录:").grid(row=0, column=0, padx=(0, 8), pady=4, sticky=W)
        self.org_dir = ttk.Entry(dir_frame, width=35)
        self.org_dir.insert(0, self.config.get('last_organizer_folder', ''))
        self.org_dir.grid(row=0, column=1, padx=(0, 8), pady=4, sticky=EW)
        ttk.Button(dir_frame, text="浏览", command=lambda: self._select_dir(self.org_dir), bootstyle="outline").grid(row=0, column=2, pady=4, sticky=E)
        
        # 2. 筛选模式选择
        mode_frame = ttk.Frame(org_frame)
        mode_frame.grid(row=1, column=0, sticky=EW, pady=(0, 10))
        ttk.Label(mode_frame, text="筛选模式:").pack(side=LEFT, padx=(0, 8))
        self.org_mode = tk.StringVar(value=self.config.get('org_mode', 'days'))
        ttk.Radiobutton(mode_frame, text="按天数", variable=self.org_mode, value='days', 
                        command=self._on_org_mode_change).pack(side=LEFT, padx=10)
        ttk.Radiobutton(mode_frame, text="按日期范围", variable=self.org_mode, value='daterange',
                        command=self._on_org_mode_change).pack(side=LEFT, padx=10)

        # 3. 设置区
        self.org_days_frame = ttk.Frame(org_frame)
        ttk.Label(self.org_days_frame, text="最近天数:").pack(side=LEFT, padx=(0, 8))
        self.org_days = ttk.Entry(self.org_days_frame, width=10)
        self.org_days.insert(0, str(self.config.get('last_organizer_days', 30)))
        self.org_days.pack(side=LEFT)
        
        self.org_daterange_frame = ttk.Frame(org_frame)
        ttk.Label(self.org_daterange_frame, text="日期范围:").pack(side=LEFT, padx=(0, 8))
        self.org_start_date = ttk.Entry(self.org_daterange_frame, width=11)
        self.org_start_date.pack(side=LEFT, padx=2)
        ttk.Button(self.org_daterange_frame, text="📅", width=3, bootstyle="link",
                  command=lambda: self._select_date(self.org_start_date, "开始日期")).pack(side=LEFT)
        ttk.Label(self.org_daterange_frame, text="-").pack(side=LEFT, padx=2)
        self.org_end_date = ttk.Entry(self.org_daterange_frame, width=11)
        self.org_end_date.pack(side=LEFT, padx=2)
        ttk.Button(self.org_daterange_frame, text="📅", width=3, bootstyle="link",
                  command=lambda: self._select_date(self.org_end_date, "结束日期")).pack(side=LEFT)

        if self.org_mode.get() == 'days': self.org_days_frame.grid(row=2, column=0, sticky=W, pady=(0, 12))
        else: self.org_daterange_frame.grid(row=2, column=0, sticky=W, pady=(0, 12))
            
        # 4. 时间类型
        ttk.Separator(org_frame).grid(row=3, column=0, sticky=EW, pady=(2, 12))
        ttk.Label(org_frame, text="判断依据:", font=("微软雅黑", 10, "bold")).grid(row=4, column=0, sticky=W, pady=(0, 6))
        types_frame = ttk.Frame(org_frame)
        types_frame.grid(row=5, column=0, sticky=EW)
        saved_create = bool(self.config.get('org_create_time', True))
        saved_modify = bool(self.config.get('org_modify_time', True))
        saved_access = bool(self.config.get('org_access_time', False))
        if not any([saved_create, saved_modify, saved_access]):
            saved_modify = True
        self.org_create_time = tk.BooleanVar(value=saved_create)
        self.org_modify_time = tk.BooleanVar(value=saved_modify)
        self.org_access_time = tk.BooleanVar(value=saved_access)
        ttk.Checkbutton(types_frame, text="创建时间", variable=self.org_create_time).pack(side=LEFT, padx=(0, 16))
        ttk.Checkbutton(types_frame, text="修改时间", variable=self.org_modify_time).pack(side=LEFT, padx=(0, 16))
        ttk.Checkbutton(types_frame, text="访问时间", variable=self.org_access_time).pack(side=LEFT, padx=(0, 16))

        # 5. 执行
        ttk.Button(org_frame, text="开始整理归档至桌面", command=self._do_organize, bootstyle="warning").grid(row=6, column=0, sticky=EW, pady=(20, 0))

    def _init_excel_tab(self, parent):
        """Excel处理标签页"""
        self._create_workspace_selector(parent)
        
        # 主容器
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=BOTH, expand=YES, pady=5)
        
        # 左中右三列布局
        col1 = ttk.Frame(main_frame)
        col1.pack(side=LEFT, fill=BOTH, expand=YES, padx=(0, 2))
        
        col2 = ttk.Frame(main_frame)
        col2.pack(side=LEFT, fill=BOTH, expand=YES, padx=2)
        
        col3 = ttk.Frame(main_frame)
        col3.pack(side=LEFT, fill=BOTH, expand=YES, padx=(2, 0))
        
        # 第一列：内容编辑
        edit_frame = ttk.Labelframe(col1, text="✍️ 内容编辑", padding=10)
        edit_frame.pack(fill=BOTH, expand=YES)
        
        edit_buttons = [
            ("统一全工作簿单元格样式", self._format_cells),
            ("按A1内容重命名工作表", self._rename_sheets),
            ("取消合并并补齐纵向内容", self._unmerge_cells),
            ("将公式替换为计算结果", self._clear_formulas),
            ("按指定单元格生成归属列", self._set_ownership),
            ("批量删除所有工作表指定行列", self._delete_rows_columns),
            ("批量插入所有工作表指定行列", self._insert_rows_columns),
            ("按名称批量删除工作表", self._delete_sheet),
            ("将A1内容填入新增列", self._add_a1_column),
            ("为单元格多行文本添加序号", self._add_numbering)
        ]
        
        for text, cmd in edit_buttons:
            ttk.Button(edit_frame, text=text, command=cmd, bootstyle="primary-outline").pack(fill=X, pady=4)
            
        # 第二列：文件管理
        mgmt_frame = ttk.Labelframe(col2, text="🗂️ 文件/表管理", padding=10)
        mgmt_frame.pack(fill=BOTH, expand=YES)
        
        mgmt_buttons = [
            ("递归导出各目录文件名清单", self._collect_all_filenames),
            ("导出Excel文件与工作表清单", self._collect_directory),
            ("清洗全角字符并删除全部空格", self._convert_fullwidth),
            ("取消合并、补齐内容并删除重复行", self._split_v2_deduplicate),
            ("按工作表拆分为独立Excel", self._split_excel),
            ("合并目录Excel并按整行去重", self._merge_excel),
            ("合并Excel、转置并按整行去重", self._merge_excel_transpose),
            ("合并Excel工作表为单个PDF", self._export_pdf_merged),
            ("按Excel工作表导出独立PDF", self._export_pdf_split),
        ]
        
        for text, cmd in mgmt_buttons:
            ttk.Button(mgmt_frame, text=text, command=cmd, bootstyle="success-outline").pack(fill=X, pady=4)
            
        # 第三列：高级功能
        ai_frame = ttk.Labelframe(col3, text="🧠 AI协同/高级", padding=10)
        ai_frame.pack(fill=BOTH, expand=YES)
        
        ai_buttons = [
            ("按列名追加数据到模板", self._multi_dimension_import_template),
            ("将超20MB Excel按行拆分", self._split_excel_by_size),
            ("按工作表导出制表符TXT", self._convert_excel_to_txt),
            ("按首列主名称生成字段文本", self._linear_text),
        ]
        
        for text, cmd in ai_buttons:
            ttk.Button(ai_frame, text=text, command=cmd, bootstyle="warning-outline").pack(fill=X, pady=4)
            
    def _init_word_tab(self, parent):
        """Word处理标签页"""
        self._create_workspace_selector(parent)
        
        # 主容器 - 两列布局
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=BOTH, expand=YES, pady=5)
        
        col1 = ttk.Frame(main_frame)
        col1.pack(side=LEFT, fill=BOTH, expand=YES, padx=(0, 2))
        
        col2 = ttk.Frame(main_frame)
        col2.pack(side=LEFT, fill=BOTH, expand=YES, padx=(2, 0))
        
        # 第一列：文档编辑
        edit_frame = ttk.Labelframe(col1, text="✍️ 文档编辑", padding=10)
        edit_frame.pack(fill=BOTH, expand=YES)
        
        edit_buttons = [
            ("统一格式", self._word_format_document),
            ("清除格式", self._word_clear_formatting),
            ("文本批量替换", self._word_replace_text),
            ("统一表格", self._word_unify_tables),
            ("Word转TXT", self._word_to_txt),
        ]
        
        for text, cmd in edit_buttons:
            ttk.Button(edit_frame, text=text, command=cmd, bootstyle="info-outline").pack(fill=X, pady=4)
            
        # 第二列：文件管理
        mgmt_frame = ttk.Labelframe(col2, text="🗂️ 文件管理", padding=10)
        mgmt_frame.pack(fill=BOTH, expand=YES)
        
        mgmt_buttons = [
            ("合并Word文档", self._word_merge_documents),
        ]
        
        for text, cmd in mgmt_buttons:
            ttk.Button(mgmt_frame, text=text, command=cmd, bootstyle="primary-outline").pack(fill=X, pady=4)

    def _create_workspace_selector(self, parent):
        """复用工作区选择器"""
        frame = ttk.Labelframe(parent, text="工作区 Workspace", padding=10)
        frame.pack(fill=X)
        
        # 使用 self.workspace_path 绑定
        ttk.Entry(frame, textvariable=self.workspace_path, state='readonly').pack(side=LEFT, fill=X, expand=YES, padx=(0, 10))
        
        ttk.Button(frame, text="选择文件", command=self._select_workspace_file, bootstyle="outline-secondary").pack(side=LEFT, padx=5)
        ttk.Button(frame, text="选择目录", command=self._select_workspace_dir, bootstyle="outline-secondary").pack(side=LEFT)
        
    # --- 逻辑实现 ---
    
    def _select_dir(self, entry_widget):
        d = filedialog.askdirectory()
        if d:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, d)
            
    def _select_workspace_file(self):
        f = filedialog.askopenfilename(filetypes=[
            ("Office文件", "*.xlsx *.xls *.docx *.doc"),
            ("Excel文件", "*.xlsx *.xls"),
            ("Word文件", "*.docx *.doc"),
            ("所有文件", "*.*")
        ])
        if f:
            self.workspace_path.set(f)
            self.target_path = Path(f)

    def _select_workspace_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.workspace_path.set(d)
            self.target_path = Path(d)

    # --- 文件工具实现 ---
    def _do_rename(self):
        folder = self.rename_dir.get().strip()
        prefix = self.rename_prefix.get().strip()

        if not folder:
            messagebox.showwarning("提示", "请先选择目标目录！")
            return
        if not os.path.isdir(folder):
            messagebox.showerror("错误", f"目录不存在：{folder}")
            return
        if not prefix:
            messagebox.showwarning("提示", "请输入文件前缀！")
            return

        try:
            candidates = []
            for root, dirs, files in os.walk(folder, topdown=True):
                for file_name in files:
                    if not file_name.startswith('~$') and not file_name.startswith(prefix):
                        candidates.append((os.path.join(root, file_name), False))
                for dir_name in dirs:
                    if not dir_name.startswith('~$') and not dir_name.startswith(prefix):
                        candidates.append((os.path.join(root, dir_name), True))
        except Exception:
            candidates = []

        if not candidates:
            messagebox.showinfo("提示",
                "目录中没有需要重命名的文件或子目录\n（名称可能已含该前缀，或目录为空）")
            return

        file_count = sum(1 for _path, is_dir in candidates if not is_dir)
        dir_count = sum(1 for _path, is_dir in candidates if is_dir)

        if not messagebox.askyesno(
            "确认批量重命名",
            f"将对 {len(candidates)} 个项目执行重命名：\n"
            f"文件：{file_count} 个\n"
            f"子目录：{dir_count} 个\n"
            f"前缀：「{prefix}」\n"
            f"目录：{folder}\n\n"
            "此操作会递归修改该目录下的文件名和子目录名，确定继续吗？",
            icon='warning'
        ):
            return

        self.config.set('last_rename_folder', folder)
        self.config.set('last_rename_prefix', prefix)
        self.file_processor.batch_rename(folder, prefix)

    def _do_delete_filename_content(self):
        """递归删除文件名中的指定内容，保留扩展名且不修改目录名。"""
        folder = self.rename_dir.get().strip()
        delete_text = self.rename_delete_text.get()

        if not folder:
            messagebox.showwarning("提示", "请先选择目标目录！")
            return
        if not os.path.isdir(folder):
            messagebox.showerror("错误", f"目录不存在：{folder}")
            return
        if delete_text == "":
            messagebox.showwarning("提示", "请输入需要从文件名中删除的内容！")
            return

        candidates = []
        invalid_count = 0
        try:
            for root, _dirs, files in os.walk(folder, topdown=True):
                for file_name in files:
                    new_name = self.file_processor.preview_filename_deletion(file_name, delete_text)
                    if new_name is None:
                        continue
                    if new_name == "":
                        invalid_count += 1
                        continue
                    candidates.append(os.path.join(root, file_name))
        except Exception as e:
            messagebox.showerror("错误", f"读取目录失败：{e}")
            return

        if not candidates:
            detail = ""
            if invalid_count:
                detail = f"\n另有 {invalid_count} 个文件删除后名称无效，已保护并跳过。"
            messagebox.showinfo(
                "提示",
                "没有找到文件名中包含该内容且可安全改名的文件。" + detail
            )
            return

        if not messagebox.askyesno(
            "确认删除文件名内容",
            f"将修改 {len(candidates)} 个文件名：\n"
            f"删除内容：「{delete_text}」\n"
            f"目录：{folder}\n\n"
            "处理范围包含全部子目录中的文件。\n"
            "文件扩展名会保留，文件夹名称不会修改。\n"
            "所有精确匹配内容都会被删除，并区分大小写。\n\n"
            "这是直接改名操作，确定继续吗？",
            icon='warning'
        ):
            return

        self.config.set('last_rename_folder', folder)
        self.file_processor.delete_filename_content(folder, delete_text)

    def _do_convert_pdf(self):
        folder = self.pdf_dir.get()
        if not folder: return
        self.config.set('last_pdf_folder', folder)
        self.config.set('pdf_subdirs', self.pdf_subdirs.get())
        
        self.progress.start()
        threading.Thread(target=self.file_processor.convert_to_pdf, 
                        args=(folder, self.pdf_subdirs.get(), True, self.log_text, self.progress),
                        daemon=True).start()

    def _do_merge_pdf(self):
        folder = self.merge_dir.get()
        if not folder: return
        self.config.set('last_merge_folder', folder)
        
        self.progress.start()
        threading.Thread(target=self.file_processor.merge_pdfs,
                        args=(folder, True, self.log_text, self.progress),
                        daemon=True).start()
                        
    def _do_organize(self):
        folder = self.org_dir.get().strip()
        mode = self.org_mode.get()
        
        # 输入验证
        if not folder:
            messagebox.showwarning("提示", "请选择源目录！")
            return
        
        if not os.path.exists(folder):
            messagebox.showerror("错误", f"目录不存在：{folder}")
            return
        
        if not os.path.isdir(folder):
            messagebox.showerror("错误", "请选择目录而不是文件！")
            return
        
        # 获取用户选择的时间类型
        time_types = {
            'create_time': self.org_create_time.get(),
            'modify_time': self.org_modify_time.get(),
            'access_time': self.org_access_time.get()
        }
        
        # 检查至少选择了一种时间类型
        if not any(time_types.values()):
            messagebox.showwarning("提示", "请至少选择一种时间类型！")
            return
        
        # 根据模式验证输入并调用
        if mode == 'days':
            # 天数模式
            days_str = self.org_days.get().strip()
            try:
                days = int(days_str)
                if days < 1:
                    messagebox.showwarning("提示", "天数必须大于0！")
                    return
            except ValueError:
                messagebox.showerror("错误", f"无效的天数：{days_str}\n请输入正整数")
                return
            
            # 保存配置
            self.config.set('last_organizer_folder', folder)
            self.config.set('last_organizer_days', days)
            self.config.set('org_mode', 'days')
            self.config.set('org_create_time', self.org_create_time.get())
            self.config.set('org_modify_time', self.org_modify_time.get())
            self.config.set('org_access_time', self.org_access_time.get())
            
            self.progress.start()
            threading.Thread(target=self.file_processor.organize_files,
                            args=(folder,),
                            kwargs={'days': days,
                                   'time_types': time_types,
                                   'file_types_cfg': {'office': True, 'wps': True, 'pdf': True, 'image': True},
                                   'log_widget': self.log_text,
                                   'progress_widget': self.progress},
                            daemon=True).start()
        else:
            # 时间范围模式
            start_str = self.org_start_date.get().strip()
            end_str = self.org_end_date.get().strip()
            
            try:
                # 解析日期（支持多种格式）
                start_date = datetime.strptime(start_str, '%Y-%m-%d')
                end_date = datetime.strptime(end_str, '%Y-%m-%d')
                
                # 设置时间为当天的开始和结束
                start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
                end_date = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)
                
                if start_date > end_date:
                    messagebox.showwarning("提示", "开始时间不能晚于结束时间！")
                    return
                    
            except ValueError as e:
                messagebox.showerror("错误", f"日期格式错误！\n请使用 YYYY-MM-DD 格式\n例如：2025-01-01")
                return
            
            # 保存配置
            self.config.set('last_organizer_folder', folder)
            self.config.set('org_mode', 'daterange')
            self.config.set('org_start_date', start_str)
            self.config.set('org_end_date', end_str)
            self.config.set('org_create_time', self.org_create_time.get())
            self.config.set('org_modify_time', self.org_modify_time.get())
            self.config.set('org_access_time', self.org_access_time.get())
            
            self.progress.start()
            threading.Thread(target=self.file_processor.organize_files,
                            args=(folder,),
                            kwargs={'start_date': start_date,
                                   'end_date': end_date,
                                   'time_types': time_types,
                                   'file_types_cfg': {'office': True, 'wps': True, 'pdf': True, 'image': True},
                                   'log_widget': self.log_text,
                                   'progress_widget': self.progress},
                            daemon=True).start()
    
    def _on_org_mode_change(self):
        """时间筛选模式切换时的UI更新"""
        mode = self.org_mode.get()
        if mode == 'days':
            # 显示天数输入，隐藏日期范围
            self.org_daterange_frame.grid_remove()
            self.org_days_frame.grid(row=2, column=0, sticky=W, pady=(0, 12))
        else:
            # 隐藏天数输入，显示日期范围
            self.org_days_frame.grid_remove()
            self.org_daterange_frame.grid(row=2, column=0, sticky=W, pady=(0, 12))
    
    def _select_date(self, entry_widget, title):
        """简单的日期选择对话框"""
        from tkinter import simpledialog
        date_str = simpledialog.askstring(title, "请输入日期 (YYYY-MM-DD):", 
                                         initialvalue=entry_widget.get())
        if date_str:
            try:
                # 验证日期格式
                datetime.strptime(date_str, '%Y-%m-%d')
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, date_str)
            except ValueError:
                messagebox.showerror("错误", "日期格式错误！\n请使用 YYYY-MM-DD 格式")

    # --- Excel工具实现 (桥接 ExcelProcessor) ---
    def _get_files(self):
        if not hasattr(self, 'target_path') or not self.target_path:
            messagebox.showwarning("提示", "请先选择工作区！")
            return []
        if self.target_path.is_file():
            return [self.target_path]
        else:
            files = self.excel_processor.get_excel_files(str(self.target_path))
            if not files:
                messagebox.showinfo("提示", "当前选择的目录中没有找到Excel文件 (.xlsx, .xls)")
                return []
            return files

    # --- 统一的线程处理 (核心修复) ---
    
    def _process_worker(self, task_func, success_msg="操作完成"):
        """统一的后台工作线程
        Args:
            task_func: 一个无参函数，执行具体逻辑。如果需要参数，使用 data=lambda: func(args)
            success_msg: 完成后的提示标题
        """
        def worker():
            try:
                self.msg_queue.put(('update_progress', [self.progress, 0, 100]))
                self.msg_queue.put(('log', ["--- 开始任务 ---"]))
                
                # 执行任务
                result = task_func()
                
                self.msg_queue.put(('log', ["--- 任务结束 ---"]))
                self.msg_queue.put(('stop_progress', [self.progress]))
                
                # 统一成功提示
                if result is not False: # 允许返回 None 表示无特定返回值，只判断 False 为失败
                     self.msg_queue.put(('show_message', ['info', '成功', f"{success_msg}"]))
                else:
                     self.msg_queue.put(('show_message', ['error', '失败', "操作未能完成，请查看日志或控制台输出。"]))

            except Exception as e:
                import traceback
                err_str = f"发生错误: {str(e)}\n{traceback.format_exc()}"
                print(err_str) # 打印到控制台
                self.msg_queue.put(('log', [str(e)])) # 简短错误到GUI日志
                self.msg_queue.put(('stop_progress', [self.progress]))
                self.msg_queue.put(('show_message', ['error', '错误', f"执行出错:\n{str(e)}"]))

        threading.Thread(target=worker, daemon=True).start()

    def _process_excel_thread(self, func, *args, operation_name="批量处理"):
        """兼容旧接口：处理选中的文件"""
        files = self._get_files()
        if not files: 
            return
        
        # 批量操作确认（超过3个文件时询问）
        if len(files) > 3:
            confirm_msg = f"即将对 {len(files)} 个Excel文件执行【{operation_name}】操作\n\n确定要继续吗？"
            if not messagebox.askyesno("批量操作确认", confirm_msg, icon='question'):
                self.msg_queue.put(('log', [f"用户取消了{operation_name}操作"]))
                return
        
        def task():
            count = 0
            total = len(files)
            for i, f in enumerate(files):
                self.msg_queue.put(('log', [f"正在处理 ({i+1}/{total}): {f.name}"]))
                if func(f, *args): 
                    count += 1
                else:
                    self.msg_queue.put(('log', [f"  x 失败: {f.name}"]))
            self.msg_queue.put(('log', [f"统计: 成功 {count} / 总数 {total}"]))
            return count > 0

        self._process_worker(task, f"{operation_name}完成")

    # --- Excel功能实现 ---
    def _format_cells(self): self._process_excel_thread(self.excel_processor.format_cells, operation_name="统一全工作簿单元格样式")
    def _rename_sheets(self): self._process_excel_thread(self.excel_processor.rename_sheets_by_a1, operation_name="按A1内容重命名工作表")
    def _unmerge_cells(self): self._process_excel_thread(self.excel_processor.unmerge_cells, operation_name="取消合并并补齐纵向内容")
    def _clear_formulas(self): self._process_excel_thread(self.excel_processor.clear_formulas, operation_name="将公式替换为计算结果")
    def _convert_fullwidth(self): self._process_excel_thread(self.excel_processor.convert_fullwidth_to_halfwidth, operation_name="清洗全角字符并删除全部空格")
    def _convert_excel_to_txt(self): self._process_excel_thread(self.excel_processor.convert_excel_to_txt, operation_name="按工作表导出制表符TXT")
    def _add_a1_column(self): self._process_excel_thread(self.excel_processor.add_a1_column, operation_name="将A1内容填入新增列")
    def _add_numbering(self): self._process_excel_thread(self.excel_processor.add_numbering_to_cells, operation_name="为单元格多行文本添加序号")
    def _linear_text(self): self._process_excel_thread(lambda f: self.excel_processor.convert_to_linear_text(f, f.with_suffix('.txt')), operation_name="按首列主名称生成字段文本")
    
    # 复杂包装
    def _delete_rows_columns(self):
        files = self._get_files()
        if not files: return
        dialog = DeleteInsertDialog(self.root, "批量删除所有工作表指定行列")
        if dialog.result:
            r, c = dialog.result
            self._process_excel_thread(
                lambda f: self.excel_processor.delete_rows_or_columns(f, r, c),
                operation_name="批量删除所有工作表指定行列"
            )

    def _insert_rows_columns(self):
        files = self._get_files()
        if not files: return
        dialog = DeleteInsertDialog(self.root, "批量插入所有工作表指定行列")
        if dialog.result:
            r, c = dialog.result
            self._process_excel_thread(
                lambda f: self.excel_processor.insert_rows_or_columns(f, r, c),
                operation_name="批量插入所有工作表指定行列"
            )
            
    def _set_ownership(self):
        files = self._get_files()
        if not files: return
        dialog = OwnershipDialog(self.root)
        if dialog.result:
            self._process_excel_thread(
                lambda f: self.excel_processor.set_ownership(f, dialog.result),
                operation_name="按指定单元格生成归属列"
            )

    def _split_v2_deduplicate(self):
        files = self._get_files()
        if not files: 
            return
        
        def task():
            for f in files:
                self.msg_queue.put(('log', [f"正在取消合并、补齐内容并删除重复行: {f.name}"]))
                res = self.excel_processor.unmerge_v2_deduplicate(f)
                if res: 
                     self.msg_queue.put(('log', [f"  -> 完成"]))
            return True
            
        self._process_worker(task, "取消合并、补齐内容并删除重复行完成")
        
    def _merge_excel(self):
        if not hasattr(self, 'target_path') or not self.target_path.is_dir():
             messagebox.showwarning("提示", "请选择目录以合并Excel并按整行去重")
             return
        output = self.target_path / "merged_all.xlsx"
        self._process_worker(lambda: self.excel_processor.merge_excel_files(self.target_path, output), f"合并目录Excel并按整行去重完成: {output.name}")

    def _merge_excel_transpose(self):
        if not hasattr(self, 'target_path') or not self.target_path.is_dir():
             messagebox.showwarning("提示", "请选择目录")
             return
        output = self.target_path / "merged_transpose.xlsx"
        self._process_worker(lambda: self.excel_processor.merge_excel_transpose(self.target_path, output), f"合并Excel、转置并按整行去重完成: {output.name}")

    def _collect_directory(self):
        if not hasattr(self, 'target_path') or not self.target_path.is_dir():
             messagebox.showwarning("提示", "请选择目录")
             return
        output = filedialog.asksaveasfilename(title="保存Excel文件与工作表清单", defaultextension=".xlsx")
        if not output: return
        self._process_worker(lambda: self.excel_processor.collect_directory(self.target_path, Path(output)), "导出Excel文件与工作表清单完成")
    
    def _collect_all_filenames(self):
        """收集所有文件名功能"""
        if not hasattr(self, 'target_path') or not self.target_path.is_dir():
            messagebox.showwarning("提示", "请选择目录")
            return
        
        # 默认输出文件名为 all_name.xlsx,保存在选中的目录中
        default_output = self.target_path / "all_name.xlsx"
        output = filedialog.asksaveasfilename(
            title="保存各目录文件名清单", 
            defaultextension=".xlsx",
            initialfile="all_name.xlsx",
            initialdir=str(self.target_path)
        )
        if not output: 
            return
        
        self._process_worker(
            lambda: self.excel_processor.collect_all_filenames(self.target_path, Path(output)), 
            "递归导出各目录文件名清单完成"
        )
        
    def _export_pdf_merged(self):
        files = self._get_files()
        if not files: return
        excel_file = self.target_path if self.target_path.is_file() else files[0]
        output_dir = self.target_path.parent if self.target_path.is_file() else self.target_path
        output = output_dir / f"{excel_file.stem}.pdf"
        self._process_worker(lambda: self.pdf_exporter.export_to_pdf_merged(files, output), "合并Excel工作表为单个PDF完成")

    def _split_excel(self):
        files = self._get_files()
        if not files: return
        output_dir = filedialog.askdirectory(title="选择独立Excel输出目录")
        if not output_dir: return
        self._process_worker(lambda: self.excel_splitter.split_to_excel_files(files, Path(output_dir)), "按工作表拆分为独立Excel完成")

    def _export_pdf_split(self):
        files = self._get_files()
        if not files: return
        output_dir = filedialog.askdirectory(title="选择独立PDF输出目录")
        if not output_dir: return
        self._process_worker(lambda: self.pdf_exporter.export_to_pdf_split(files, Path(output_dir)), "按Excel工作表导出独立PDF完成")

    def _multi_dimension_import_template(self):
        template_file = filedialog.askopenfilename(title="选择要追加数据的模板文件", filetypes=[("Excel文件", "*.xlsx *.xls")])
        if not template_file: return
        files = self._get_files()
        if not files: return
        if not messagebox.askyesno("确认", f"确定要处理 {len(files)} 个文件吗？"): return
        self._process_worker(lambda: self.excel_processor.multi_dimension_import_template(Path(template_file), files), "按列名追加数据到模板完成")
    
    def _split_excel_by_size(self):
        """将超20MB Excel按行拆分。"""
        files = self._get_files()
        if not files:
            return
        
        # 批量操作确认
        if len(files) > 1:
            if not messagebox.askyesno("将超20MB Excel按行拆分确认", 
                                       f"即将对 {len(files)} 个Excel文件执行【将超20MB Excel按行拆分】\n\n确定继续吗？",
                                       icon='question'):
                return
        
        self._process_excel_thread(
            self.excel_processor.split_excel_by_size, 
            operation_name="将超20MB Excel按行拆分"
        )
    
    def _delete_sheet(self):
        files = self._get_files()
        if not files: return
        names = simpledialog.askstring("按名称批量删除工作表", "请输入要删除的工作表名称（逗号分隔）：")
        if not names: return
        sheet_names = [n.strip() for n in names.split(',') if n.strip()]
        
        def task():
            for f in files:
                self.excel_processor.delete_sheets(f, sheet_names)
            return True
        self._process_worker(task, "按名称批量删除工作表完成")

    # --- 线程帮助函数 (Deprecated but kep for compatibility if needed inside other methods, but we replaced usages) ---
    def _run_thread(self, func, success_msg, error_msg):
        """Deprecated: redirect to _process_worker"""
        self._process_worker(func, success_msg)

    # --- Word工具实现 ---
    def _get_word_files(self):
        """获取选中的Word文件"""
        if not hasattr(self, 'target_path') or not self.target_path:
            messagebox.showwarning("提示", "请先选择工作区!")
            return []
        if self.target_path.is_file():
            return [self.target_path]
        else:
            files = self.word_processor.get_word_files(str(self.target_path))
            if not files:
                messagebox.showinfo("提示", "当前选择的目录中没有找到Word文件 (.docx, .doc)")
                return []
            return files
    
    def _process_word_thread(self, func, *args, operation_name="批量处理"):
        """处理选中的Word文件"""
        files = self._get_word_files()
        if not files: 
            return
        
        # 批量操作确认(超过3个文件时询问)
        if len(files) > 3:
            confirm_msg = f"即将对 {len(files)} 个Word文件执行【{operation_name}】操作\n\n确定要继续吗?"
            if not messagebox.askyesno("批量操作确认", confirm_msg, icon='question'):
                self.msg_queue.put(('log', [f"用户取消了{operation_name}操作"]))
                return
        
        def task():
            count = 0
            total = len(files)
            for i, f in enumerate(files):
                self.msg_queue.put(('log', [f"正在处理 ({i+1}/{total}): {f.name}"]))
                if func(f, *args): 
                    count += 1
                else:
                    self.msg_queue.put(('log', [f"  x 失败: {f.name}"]))
            self.msg_queue.put(('log', [f"统计: 成功 {count} / 总数 {total}"]))
            return count > 0

        self._process_worker(task, f"{operation_name}完成")
    
    def _word_format_document(self):
        """统一样式化Word文档（应用详细样式配置）"""
        self._process_word_thread(self.word_processor.apply_unified_styles, operation_name="统一样式")
    
    def _word_clear_formatting(self):
        """清除Word文档格式"""
        self._process_word_thread(self.word_processor.clear_formatting, operation_name="清除格式")
    
    def _word_to_txt(self):
        """Word转TXT"""
        self._process_word_thread(self.word_processor.convert_to_txt, operation_name="Word转TXT")
    
    def _word_replace_text(self):
        """批量替换Word文档中的文本"""
        files = self._get_word_files()
        if not files:
            return
        
        # 获取要替换的文本
        old_text = simpledialog.askstring("批量替换", "请输入要查找的文本:")
        if not old_text:
            return
        
        new_text = simpledialog.askstring("批量替换", "请输入替换为:")
        if new_text is None:  # 用户取消
            return
        
        self._process_word_thread(
            lambda f: self.word_processor.replace_text(f, old_text, new_text),
            operation_name="文本批量替换"
        )
    
    def _word_unify_tables(self):
        """统一表格：根据窗口自动调整，支持 .doc 和 .docx"""
        files = self._get_word_files()
        if not files:
            return
        self._process_word_thread(self.word_processor.unify_table_format, operation_name="统一表格")

    def _word_merge_documents(self):
        """合并Word文档"""
        if not hasattr(self, 'target_path') or not self.target_path.is_dir():
            messagebox.showwarning("提示", "请选择目录进行合并")
            return
        
        output = self.target_path / "merged_all.docx"
        self._process_worker(
            lambda: self.word_processor.merge_word_files(self.target_path, output), 
            f"合并完成: {output.name}"
        )

    # --- 基础 ---
    def on_closing(self):
        # 保存安全范围内的窗口几何状态，防止下次启动恢复为过大窗口
        try:
            self.config.set('window_geometry', self._safe_startup_geometry(self.root.geometry()))
        except:
            pass
        self.config.save_config()
        self.root.destroy()

    def process_msg_queue(self):
        """处理消息队列"""
        try:
            while not self.msg_queue.empty():
                msg_type, data = self.msg_queue.get_nowait()
                
                try:
                    if msg_type == 'log':
                        if len(data) > 0:
                            self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {data[0]}\n")
                            self.log_text.see(tk.END)
                    
                    elif msg_type == 'append_text':
                        if len(data) >= 2:
                            widget, text = data[0], data[1]
                            widget.insert(tk.END, text)
                            widget.see(tk.END)
                    
                    elif msg_type == 'update_progress':
                        # data: (widget, val, max_val) 或 (widget, val)
                        if len(data) >= 2:
                            pbar = data[0]
                            pbar['value'] = data[1]
                            # 只有提供了第三个参数才设置maximum
                            if len(data) >= 3 and data[2] is not None:
                                pbar['maximum'] = data[2]
                    
                    elif msg_type == 'stop_progress':
                        if len(data) > 0:
                            data[0].stop()
                    
                    elif msg_type == 'show_message':
                        # data: (type, title, msg)
                        if len(data) >= 3:
                            if data[0] == 'info': 
                                messagebox.showinfo(data[1], data[2])
                            elif data[0] == 'error': 
                                messagebox.showerror(data[1], data[2])
                
                except Exception as e:
                    # 单个消息处理失败不影响其他消息
                    print(f"处理消息失败 [{msg_type}]: {e}")
                    
        except queue.Empty:
            pass
        except Exception as e:
            print(f"消息队列处理异常: {e}")
        
        # 定期调用自己，实现消息队列的持续处理（每100毫秒检查一次）
        self.root.after(100, self.process_msg_queue)
    
class DeleteInsertDialog:
    def __init__(self, parent, title):
        self.result = None
        self.dialog = ttk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("450x350")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=BOTH, expand=YES)
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(side=BOTTOM, pady=20)
        ttk.Button(btn_frame, text="确定", command=self.on_ok, bootstyle="primary").pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=self.dialog.destroy, bootstyle="secondary").pack(side=LEFT)

        ttk.Label(main_frame, text=f"{title}配置", font=("微软雅黑", 14, "bold")).pack(pady=(0, 20))
        
        row_frame = ttk.Labelframe(main_frame, text="行操作", padding="10")
        row_frame.pack(fill=X, pady=5)
        ttk.Label(row_frame, text="起始行:").grid(row=0, column=0, padx=5)
        self.row_start = ttk.Entry(row_frame, width=8)
        self.row_start.grid(row=0, column=1, padx=5)
        ttk.Label(row_frame, text="数量:").grid(row=0, column=2, padx=5)
        self.row_count = ttk.Entry(row_frame, width=8)
        self.row_count.grid(row=0, column=3, padx=5)
        
        col_frame = ttk.Labelframe(main_frame, text="列操作", padding="10")
        col_frame.pack(fill=X, pady=10)
        ttk.Label(col_frame, text="起始列:").grid(row=0, column=0, padx=5)
        self.col_start = ttk.Entry(col_frame, width=8)
        self.col_start.grid(row=0, column=1, padx=5)
        ttk.Label(col_frame, text="数量:").grid(row=0, column=2, padx=5)
        self.col_count = ttk.Entry(col_frame, width=8)
        self.col_count.grid(row=0, column=3, padx=5)
        
        self.dialog.wait_window()
        
    def on_ok(self):
        row_info = self._parse(self.row_start.get(), self.row_count.get())
        col_info = self._parse(self.col_start.get(), self.col_count.get())
        if not row_info and not col_info:
            messagebox.showwarning("提示", "请至少输入行或列的起始位置和数量")
            return
        self.result = (row_info, col_info)
        self.dialog.destroy()
        
    def _parse(self, start, count):
        try:
            s, c = int(start), int(count)
            return (s, c) if s > 0 and c > 0 else None
        except: return None

class OwnershipDialog:
    def __init__(self, parent):
        self.result = None
        self.dialog = ttk.Toplevel(parent)
        self.dialog.title("按指定单元格生成归属列")
        self.dialog.geometry("400x500")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=BOTH, expand=YES)
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(side=BOTTOM, pady=10)
        ttk.Button(btn_frame, text="确定", command=self.on_ok, bootstyle="primary").pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=self.dialog.destroy, bootstyle="secondary").pack(side=LEFT)

        ttk.Label(main_frame, text="输入用于生成归属列的单元格地址", font=("微软雅黑", 12)).pack(pady=(0, 10))
        
        scroll_frame = ttk.Frame(main_frame)
        scroll_frame.pack(fill=BOTH, expand=YES)
        canvas = tk.Canvas(scroll_frame, bg="white", height=200)
        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
        self.entries_frame = ttk.Frame(canvas)
        self.entries_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.entries_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        self.entries = []
        for _ in range(3): self.add_entry()
        
        ttk.Button(main_frame, text="+ 添加输入框", command=self.add_entry, bootstyle="info-outline").pack(fill=X, pady=10)
        self.dialog.wait_window()
        
    def add_entry(self):
        f = ttk.Frame(self.entries_frame, padding="5")
        f.pack(fill=X)
        ttk.Label(f, text=f"地址 {len(self.entries)+1}:").pack(side=LEFT)
        e = ttk.Entry(f, width=15)
        e.pack(side=LEFT, padx=5)
        self.entries.append(e)

    def on_ok(self):
        res = [e.get().strip() for e in self.entries if e.get().strip()]
        if not res:
            messagebox.showwarning("提示", "请输入至少一个地址")
            return
        self.result = res
        self.dialog.destroy()


if __name__ == "__main__":
    from ttkbootstrap import Style
    style = Style(theme="darkly")
    root = style.master
    app = DongToolGUI(root)
    root.mainloop()
