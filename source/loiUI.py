import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import os
import re
from datetime import datetime
import time

class BillEntry:
    def __init__(self, date, name, amount, note=""):
        self.date = date
        self.name = name
        self.amount = amount
        self.note = note

class ElegantBillApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Loi 账单记录")
        self.root.geometry("1000x700")
        self.root.configure(bg="#212529")
        
        # 移除窗口装饰并设置为无边框窗口
        self.root.overrideredirect(True)
        
        # 主题状态：0为浅色模式，1为深色模式
        self.theme_mode = 1
        self.animating = False  # 防止动画冲突
        
        # 初始化数据
        self.current_file = None
        self.bill_data = []  # 原始数据
        self.display_data = []  # 显示数据
        self.selected_items = []
        self.undo_stack = []
        self.font_size = 10  # 默认字体大小
        self.modified = False  # 跟踪是否有未保存的修改
        
        # 排序状态
        self.sort_column = None
        self.sort_reverse = False
        
        # 鼠标拖动相关变量
        self.drag_threshold = 5  # 拖动阈值（像素）
        self.is_dragging = False
        self.mouse_pressed = False
        self.press_start_x = 0
        self.press_start_y = 0
        self.window_start_x = 0
        self.window_start_y = 0
        
        # 标记是否点击了菜单相关组件
        self.clicked_menu_item = False
        
        # 关于窗口实例
        self.about_window = None
        
        # 菜单窗口实例
        self.menu_window = None
        self.menu_open = False
        
        # 主题配置
        self.themes = [
            {  # 浅色主题
                'bg': '#f8f9fa',
                'fg': '#2c3e50',
                'status_fg': '#6c757d',
                'active_fg': '#28a745',
                'pause_fg': '#dc3545',
                'hint_fg': '#adb5bd',
                'separator': '#e9ecef',
                'button_hover': '#e9ecef',
                'window_border': '#dee2e6',
                'confirm_button': '#28a745',
                'scrollbar_bg': '#dee2e6',
                'scrollbar_thumb': '#adb5bd',
                'tree_bg': '#ffffff',
                'tree_fg': '#212529',
                'tree_selected': '#e9ecef',
                'tree_head_bg': '#e9ecef',
                'tree_head_fg': '#495057'
            },
            {  # 深色主题
                'bg': '#212529',
                'fg': '#f8f9fa',
                'status_fg': '#adb5bd',
                'active_fg': '#4caf50',
                'pause_fg': '#f44336',
                'hint_fg': '#6c757d',
                'separator': '#495057',
                'button_hover': '#343a40',
                'window_border': '#343a40',
                'confirm_button': '#4caf50',
                'scrollbar_bg': '#343a40',
                'scrollbar_thumb': '#6c757d',
                'tree_bg': '#2b3035',
                'tree_fg': '#f8f9fa',
                'tree_selected': '#495057',
                'tree_head_bg': '#343a40',
                'tree_head_fg': '#dee2e6'
            }
        ]
        
        # 用于渐变动画的颜色值
        self.current_colors = self.themes[1].copy()
        
        # 创建界面
        self.create_widgets()
        self.load_available_files()
        
        # 绑定事件
        self.root.bind("<Control-n>", lambda e: self.add_item())
        self.root.bind("<Control-N>", lambda e: self.add_item())
        self.root.bind("<Delete>", lambda e: self.delete_item())
        self.root.bind("<Control-u>", lambda e: self.update_item())
        self.root.bind("<Control-U>", lambda e: self.update_item())
        self.root.bind("<Control-f>", lambda e: self.search_item())
        self.root.bind("<Control-F>", lambda e: self.search_item())
        self.root.bind("<Control-s>", lambda e: self.save_file())
        self.root.bind("<Control-S>", lambda e: self.save_file())
        self.root.bind("<Control-z>", lambda e: self.undo())
        self.root.bind("<Control-Z>", lambda e: self.undo())
        self.root.bind("<Control-plus>", self.increase_font)
        self.root.bind("<Control-minus>", self.decrease_font)
        self.root.bind("<Control-MouseWheel>", self.on_mousewheel)
        self.root.bind("<Control-h>", lambda e: self.show_help())
        self.root.bind("<Control-H>", lambda e: self.show_help())
        self.root.bind("<Control-m>", lambda e: self.show_statistics())
        self.root.bind("<Control-M>", lambda e: self.show_statistics())
        self.root.bind("<Control-Up>", lambda e: self.move_up())
        self.root.bind("<Control-Down>", lambda e: self.move_down())
        self.root.bind("<Control-r>", lambda e: self.reset_display())
        self.root.bind("<Control-R>", lambda e: self.reset_display())
        self.root.bind("<r>", self.start_theme_transition)
        self.root.bind("<F4>", self.handle_f4_key)
        
        # 绑定整个窗口的鼠标事件
        self.root.bind("<ButtonPress-1>", self.on_window_press)
        self.root.bind("<ButtonRelease-1>", self.on_window_release)
        self.root.bind("<B1-Motion>", self.on_window_motion)
        
        # 确保窗口能接收键盘事件
        self.root.focus_set()
        
        # 初始更新
        self.center_window(self.root, 1000, 700)
    
    def create_widgets(self):
        # 主容器
        self.main_frame = tk.Frame(
            self.root, 
            bg=self.current_colors['bg'], 
            padx=20, 
            pady=20,
            highlightthickness=1,
            highlightbackground=self.current_colors['window_border']
        )
        self.main_frame.pack(expand=True, fill=tk.BOTH, padx=1, pady=1)
        
        # 标题栏框架（包含菜单和控制按钮）
        self.title_frame = tk.Frame(self.main_frame, bg=self.current_colors['bg'])
        self.title_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 左侧菜单按钮框架
        self.menu_frame = tk.Frame(self.title_frame, bg=self.current_colors['bg'])
        self.menu_frame.pack(side=tk.LEFT)
        
        # 更多按钮
        self.more_button = tk.Label(
            self.menu_frame,
            text="更多",
            font=("Helvetica", 9),
            fg=self.current_colors['hint_fg'],
            bg=self.current_colors['bg'],
            padx=8,
            pady=2,
            cursor="hand2"
        )
        self.more_button.pack(side=tk.LEFT)
        self.more_button.bind("<Button-1>", self.toggle_menu)
        self.more_button.bind("<Enter>", self.on_control_enter)
        self.more_button.bind("<Leave>", self.on_control_leave)
        
        # 标题文本（居中）
        self.title_label = tk.Label(
            self.title_frame,
            text="Loi 账单记录",
            font=("Helvetica", 10),
            fg=self.current_colors['hint_fg'],
            bg=self.current_colors['bg']
        )
        self.title_label.place(relx=0.5, rely=0.5, anchor="center")
        
        # 右侧窗口控制按钮框架
        self.control_frame = tk.Frame(self.title_frame, bg=self.current_colors['bg'])
        self.control_frame.pack(side=tk.RIGHT)
        
        # 最小化按钮
        self.minimize_btn = tk.Label(
            self.control_frame,
            text="—",
            font=("Helvetica", 12),
            fg=self.current_colors['hint_fg'],
            bg=self.current_colors['bg'],
            padx=8,
            pady=2,
            cursor="hand2"
        )
        self.minimize_btn.pack(side=tk.LEFT, padx=2)
        self.minimize_btn.bind("<Button-1>", self.minimize_window)
        self.minimize_btn.bind("<Enter>", self.on_control_enter)
        self.minimize_btn.bind("<Leave>", self.on_control_leave)
        
        # 关闭按钮
        self.close_btn = tk.Label(
            self.control_frame,
            text="×",
            font=("Helvetica", 14),
            fg=self.current_colors['hint_fg'],
            bg=self.current_colors['bg'],
            padx=6,
            pady=0,
            cursor="hand2"
        )
        self.close_btn.pack(side=tk.LEFT, padx=2)
        self.close_btn.bind("<Button-1>", self.close_window)
        self.close_btn.bind("<Enter>", self.on_close_enter)
        self.close_btn.bind("<Leave>", self.on_close_leave)
        
        # 标题栏绑定拖动事件
        self.title_frame.bind("<ButtonPress-1>", self.on_title_press)
        self.title_frame.bind("<ButtonRelease-1>", self.on_title_release)
        self.title_frame.bind("<B1-Motion>", self.on_title_motion)
        self.title_label.bind("<ButtonPress-1>", self.on_title_press)
        self.title_label.bind("<ButtonRelease-1>", self.on_title_release)
        self.title_label.bind("<B1-Motion>", self.on_title_motion)
        
        # 文件选择区域
        file_frame = tk.Frame(self.main_frame, bg=self.current_colors['bg'])
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(
            file_frame, 
            text="选择账单文件:", 
            bg=self.current_colors['bg'],
            fg=self.current_colors['fg']
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        self.file_var = tk.StringVar()
        self.file_combo = ttk.Combobox(file_frame, textvariable=self.file_var, width=25)
        self.file_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.file_combo.bind('<<ComboboxSelected>>', self.on_file_select)
        
        # 新建按钮
        self.new_btn = tk.Label(
            file_frame,
            text="新建",
            font=("Helvetica", 9),
            fg=self.current_colors['hint_fg'],
            bg=self.current_colors['bg'],
            padx=8,
            pady=2,
            cursor="hand2"
        )
        self.new_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.new_btn.bind("<Button-1>", lambda e: self.new_file())
        self.new_btn.bind("<Enter>", self.on_control_enter)
        self.new_btn.bind("<Leave>", self.on_control_leave)
        
        # 刷新按钮
        self.refresh_btn = tk.Label(
            file_frame,
            text="刷新",
            font=("Helvetica", 9),
            fg=self.current_colors['hint_fg'],
            bg=self.current_colors['bg'],
            padx=8,
            pady=2,
            cursor="hand2"
        )
        self.refresh_btn.pack(side=tk.LEFT)
        self.refresh_btn.bind("<Button-1>", lambda e: self.load_available_files())
        self.refresh_btn.bind("<Enter>", self.on_control_enter)
        self.refresh_btn.bind("<Leave>", self.on_control_leave)
        
        # 账单展示区域
        display_frame = tk.Frame(self.main_frame, bg=self.current_colors['bg'])
        display_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 创建Treeview
        columns = ("date", "name", "amount", "note")
        self.tree = ttk.Treeview(display_frame, columns=columns, show="headings", selectmode="extended")
        
        # 定义列
        self.tree.heading("date", text="日期", command=lambda: self.sort_treeview("date"))
        self.tree.heading("name", text="名称", command=lambda: self.sort_treeview("name"))
        self.tree.heading("amount", text="流水", command=lambda: self.sort_treeview("amount"))
        self.tree.heading("note", text="备注", command=lambda: self.sort_treeview("note"))
        
        # 设置列宽
        self.tree.column("date", width=100)
        self.tree.column("name", width=150)
        self.tree.column("amount", width=120)
        self.tree.column("note", width=250)
        
        # 添加滚动条
        v_scrollbar = ttk.Scrollbar(display_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(display_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 绑定选择事件
        self.tree.bind('<<TreeviewSelect>>', self.on_item_select)
        
        # 配置权重
        display_frame.columnconfigure(0, weight=1)
        display_frame.rowconfigure(0, weight=1)
        
        # 工作区域和统计区域
        bottom_frame = tk.Frame(self.main_frame, bg=self.current_colors['bg'])
        bottom_frame.pack(fill=tk.BOTH, pady=(0, 10))
        
        # 左侧工作区域
        work_frame = tk.Frame(bottom_frame, bg=self.current_colors['bg'])
        work_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # 编辑表单
        form_frame = tk.Frame(work_frame, bg=self.current_colors['bg'])
        form_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(
            form_frame, 
            text="日期:", 
            bg=self.current_colors['bg'],
            fg=self.current_colors['fg']
        ).grid(row=0, column=0, sticky=tk.W, pady=2)
        
        self.date_var = tk.StringVar()
        date_entry = tk.Entry(
            form_frame, 
            textvariable=self.date_var, 
            width=15,
            bg=self.current_colors['tree_bg'],
            fg=self.current_colors['tree_fg'],
            insertbackground=self.current_colors['tree_fg'],
            relief=tk.FLAT
        )
        date_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        
        tk.Label(
            form_frame, 
            text="名称:", 
            bg=self.current_colors['bg'],
            fg=self.current_colors['fg']
        ).grid(row=1, column=0, sticky=tk.W, pady=2)
        
        self.name_var = tk.StringVar()
        name_entry = tk.Entry(
            form_frame, 
            textvariable=self.name_var, 
            width=15,
            bg=self.current_colors['tree_bg'],
            fg=self.current_colors['tree_fg'],
            insertbackground=self.current_colors['tree_fg'],
            relief=tk.FLAT
        )
        name_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        
        tk.Label(
            form_frame, 
            text="流水:", 
            bg=self.current_colors['bg'],
            fg=self.current_colors['fg']
        ).grid(row=2, column=0, sticky=tk.W, pady=2)
        
        self.amount_var = tk.StringVar()
        amount_entry = tk.Entry(
            form_frame, 
            textvariable=self.amount_var, 
            width=15,
            bg=self.current_colors['tree_bg'],
            fg=self.current_colors['tree_fg'],
            insertbackground=self.current_colors['tree_fg'],
            relief=tk.FLAT
        )
        amount_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        
        tk.Label(
            form_frame, 
            text="备注:", 
            bg=self.current_colors['bg'],
            fg=self.current_colors['fg']
        ).grid(row=3, column=0, sticky=tk.W, pady=2)
        
        self.note_var = tk.StringVar()
        note_entry = tk.Entry(
            form_frame, 
            textvariable=self.note_var, 
            width=15,
            bg=self.current_colors['tree_bg'],
            fg=self.current_colors['tree_fg'],
            insertbackground=self.current_colors['tree_fg'],
            relief=tk.FLAT
        )
        note_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        
        form_frame.columnconfigure(1, weight=1)
        
        # 按钮区域
        btn_frame = tk.Frame(work_frame, bg=self.current_colors['bg'])
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 第一行按钮
        btn_row1 = tk.Frame(btn_frame, bg=self.current_colors['bg'])
        btn_row1.pack(fill=tk.X, pady=(0, 5))
        
        self.add_btn = self.create_button(btn_row1, "新增", self.add_item)
        self.add_btn.pack(side=tk.LEFT, padx=2)
        
        self.update_btn = self.create_button(btn_row1, "修改", self.update_item)
        self.update_btn.pack(side=tk.LEFT, padx=2)
        
        self.delete_btn = self.create_button(btn_row1, "删除", self.delete_item)
        self.delete_btn.pack(side=tk.LEFT, padx=2)
        
        self.search_btn = self.create_button(btn_row1, "查找", self.search_item)
        self.search_btn.pack(side=tk.LEFT, padx=2)
        
        # 第二行按钮
        btn_row2 = tk.Frame(btn_frame, bg=self.current_colors['bg'])
        btn_row2.pack(fill=tk.X, pady=(0, 5))
        
        self.save_btn = self.create_button(btn_row2, "保存", self.save_file)
        self.save_btn.pack(side=tk.LEFT, padx=2)
        
        self.undo_btn = self.create_button(btn_row2, "撤销", self.undo)
        self.undo_btn.pack(side=tk.LEFT, padx=2)
        
        self.stats_btn = self.create_button(btn_row2, "统计", self.show_statistics)
        self.stats_btn.pack(side=tk.LEFT, padx=2)
        
        self.reset_btn = self.create_button(btn_row2, "重置显示", self.reset_display)
        self.reset_btn.pack(side=tk.LEFT, padx=2)
        
        # 第三行按钮
        btn_row3 = tk.Frame(btn_frame, bg=self.current_colors['bg'])
        btn_row3.pack(fill=tk.X)
        
        self.up_btn = self.create_button(btn_row3, "上移", self.move_up)
        self.up_btn.pack(side=tk.LEFT, padx=2)
        
        self.down_btn = self.create_button(btn_row3, "下移", self.move_down)
        self.down_btn.pack(side=tk.LEFT, padx=2)
        
        # 右侧统计区域
        stats_frame = tk.Frame(bottom_frame, bg=self.current_colors['bg'])
        stats_frame.pack(side=tk.RIGHT, fill=tk.BOTH)
        
        # 统计信息
        self.total_var = tk.StringVar()
        self.total_var.set("总流水: 0.0")
        total_label = tk.Label(
            stats_frame,
            textvariable=self.total_var,
            font=("Helvetica", 10),
            fg=self.current_colors['fg'],
            bg=self.current_colors['bg']
        )
        total_label.pack(anchor=tk.W, pady=(0, 5))
        
        self.selected_var = tk.StringVar()
        self.selected_var.set("选中流水: 0.0")
        selected_label = tk.Label(
            stats_frame,
            textvariable=self.selected_var,
            font=("Helvetica", 10),
            fg=self.current_colors['fg'],
            bg=self.current_colors['bg']
        )
        selected_label.pack(anchor=tk.W, pady=(0, 5))
        
        self.same_type_var = tk.StringVar()
        self.same_type_var.set("同类流水: 0.0")
        same_type_label = tk.Label(
            stats_frame,
            textvariable=self.same_type_var,
            font=("Helvetica", 10),
            fg=self.current_colors['fg'],
            bg=self.current_colors['bg']
        )
        same_type_label.pack(anchor=tk.W, pady=(0, 10))
        
        # 帮助提示
        help_label = tk.Label(
            stats_frame,
            text="Ctrl+加号/减号或Ctrl+鼠标滚轮调整字体大小\n<<按下 Ctrl+H 查阅帮助>>",
            font=("Helvetica", 8),
            fg=self.current_colors['hint_fg'],
            bg=self.current_colors['bg'],
            justify=tk.LEFT
        )
        help_label.pack(anchor=tk.W)
        
        # 日志区域
        log_frame = tk.Frame(self.main_frame, bg=self.current_colors['bg'])
        log_frame.pack(fill=tk.BOTH, pady=(0, 10))
        
        # 日志标签
        log_label = tk.Label(
            log_frame,
            text="日志",
            font=("Helvetica", 10),
            fg=self.current_colors['fg'],
            bg=self.current_colors['bg']
        )
        log_label.pack(anchor=tk.W, pady=(0, 5))
        
        # 创建日志文本框
        self.log_text = tk.Text(
            log_frame, 
            height=4, 
            wrap=tk.WORD,
            bg=self.current_colors['tree_bg'],
            fg=self.current_colors['tree_fg'],
            insertbackground=self.current_colors['tree_fg'],
            relief=tk.FLAT
        )
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 应用当前主题到Treeview
        self.update_treeview_style()
    
    def create_button(self, parent, text, command):
        """创建样式统一的按钮"""
        btn = tk.Label(
            parent,
            text=text,
            font=("Helvetica", 9),
            fg=self.current_colors['hint_fg'],
            bg=self.current_colors['bg'],
            padx=10,
            pady=4,
            cursor="hand2",
            relief=tk.FLAT
        )
        btn.bind("<Button-1>", lambda e: command())
        btn.bind("<Enter>", lambda e: btn.configure(bg=self.current_colors['button_hover']))
        btn.bind("<Leave>", lambda e: btn.configure(bg=self.current_colors['bg']))
        return btn
    
    def update_treeview_style(self):
        """更新Treeview样式以匹配当前主题"""
        style = ttk.Style()
        style.configure("Treeview",
                        background=self.current_colors['tree_bg'],
                        foreground=self.current_colors['tree_fg'],
                        fieldbackground=self.current_colors['tree_bg'],
                        borderwidth=0,
                        font=("Helvetica", self.font_size))
        style.configure("Treeview.Heading",
                        background=self.current_colors['tree_head_bg'],
                        foreground=self.current_colors['tree_head_fg'],
                        borderwidth=0,
                        font=("Helvetica", self.font_size, "bold"))
        style.map('Treeview', background=[('selected', self.current_colors['tree_selected'])])
        
        # 更新日志区域字体大小
        self.log_text.config(font=("Helvetica", self.font_size))
    
    def toggle_menu(self, event=None):
        """切换菜单显示/隐藏"""
        # 如果关于窗口存在，不显示菜单
        if self.about_window is not None:
            return
            
        if self.menu_open:
            self.close_menu()
        else:
            self.open_menu()
    
    def open_menu(self):
        """打开菜单"""
        if self.menu_window is not None:
            return
            
        self.menu_open = True
        self.menu_window = tk.Toplevel(self.root)
        menu_win = self.menu_window
        
        # 设置菜单窗口属性
        menu_win.overrideredirect(True)
        menu_win.configure(bg=self.current_colors['bg'])
        
        # 计算菜单位置（在"更多"按钮下方）
        x = self.more_button.winfo_rootx()
        y = self.more_button.winfo_rooty() + self.more_button.winfo_height() + 2
        
        # 创建菜单框架
        menu_frame = tk.Frame(
            menu_win,
            bg=self.current_colors['bg'],
            highlightthickness=1,
            highlightbackground=self.current_colors['window_border']
        )
        menu_frame.pack(expand=True, fill=tk.BOTH, padx=1, pady=1)
        
        # 字体大小菜单项
        font_menu = tk.Label(
            menu_frame,
            text="字体大小",
            font=("Helvetica", 9),
            fg=self.current_colors['fg'],
            bg=self.current_colors['bg'],
            padx=15,
            pady=8,
            cursor="hand2",
            anchor="w"
        )
        font_menu.pack(fill=tk.X, padx=1, pady=1)
        font_menu.bind("<Enter>", lambda e: font_menu.configure(bg=self.current_colors['button_hover']))
        font_menu.bind("<Leave>", lambda e: font_menu.configure(bg=self.current_colors['bg']))
        
        # 字体大小子菜单
        font_submenu = tk.Frame(
            menu_frame,
            bg=self.current_colors['bg']
        )
        
        def show_font_submenu(e):
            # 计算子菜单位置
            sub_x = font_menu.winfo_rootx() + font_menu.winfo_width()
            sub_y = font_menu.winfo_rooty()
            font_submenu.place(x=font_menu.winfo_width(), y=0)
            font_submenu.tkraise()
        
        def hide_font_submenu(e):
            font_submenu.place_forget()
        
        font_menu.bind("<Enter>", show_font_submenu)
        font_menu.bind("<Leave>", hide_font_submenu)
        
        # 字体大小选项
        small_font = tk.Label(
            font_submenu,
            text="小",
            font=("Helvetica", 9),
            fg=self.current_colors['fg'],
            bg=self.current_colors['bg'],
            padx=15,
            pady=8,
            cursor="hand2",
            anchor="w"
        )
        small_font.pack(fill=tk.X)
        small_font.bind("<Button-1>", lambda e: self.set_font_size(8))
        small_font.bind("<Enter>", lambda e: small_font.configure(bg=self.current_colors['button_hover']))
        small_font.bind("<Leave>", lambda e: small_font.configure(bg=self.current_colors['bg']))
        
        medium_font = tk.Label(
            font_submenu,
            text="中",
            font=("Helvetica", 9),
            fg=self.current_colors['fg'],
            bg=self.current_colors['bg'],
            padx=15,
            pady=8,
            cursor="hand2",
            anchor="w"
        )
        medium_font.pack(fill=tk.X)
        medium_font.bind("<Button-1>", lambda e: self.set_font_size(10))
        medium_font.bind("<Enter>", lambda e: medium_font.configure(bg=self.current_colors['button_hover']))
        medium_font.bind("<Leave>", lambda e: medium_font.configure(bg=self.current_colors['bg']))
        
        large_font = tk.Label(
            font_submenu,
            text="大",
            font=("Helvetica", 9),
            fg=self.current_colors['fg'],
            bg=self.current_colors['bg'],
            padx=15,
            pady=8,
            cursor="hand2",
            anchor="w"
        )
        large_font.pack(fill=tk.X)
        large_font.bind("<Button-1>", lambda e: self.set_font_size(12))
        large_font.bind("<Enter>", lambda e: large_font.configure(bg=self.current_colors['button_hover']))
        large_font.bind("<Leave>", lambda e: large_font.configure(bg=self.current_colors['bg']))
        
        # 主题菜单项
        theme_text = "浅色模式" if self.theme_mode == 1 else "深色模式"
        theme_menu = tk.Label(
            menu_frame,
            text=theme_text,
            font=("Helvetica", 9),
            fg=self.current_colors['fg'],
            bg=self.current_colors['bg'],
            padx=15,
            pady=8,
            cursor="hand2",
            anchor="w"
        )
        theme_menu.pack(fill=tk.X, padx=1, pady=1)
        theme_menu.bind("<Button-1>", self.toggle_theme_from_menu)
        theme_menu.bind("<Enter>", lambda e: theme_menu.configure(bg=self.current_colors['button_hover']))
        theme_menu.bind("<Leave>", lambda e: theme_menu.configure(bg=self.current_colors['bg']))
        
        # 关于菜单项
        about_menu = tk.Label(
            menu_frame,
            text="关于",
            font=("Helvetica", 9),
            fg=self.current_colors['fg'],
            bg=self.current_colors['bg'],
            padx=15,
            pady=8,
            cursor="hand2",
            anchor="w"
        )
        about_menu.pack(fill=tk.X, padx=1, pady=(1, 2))
        about_menu.bind("<Button-1>", self.show_about_from_menu)
        about_menu.bind("<Enter>", lambda e: about_menu.configure(bg=self.current_colors['button_hover']))
        about_menu.bind("<Leave>", lambda e: about_menu.configure(bg=self.current_colors['bg']))
        
        # 设置窗口位置和大小
        menu_win.geometry(f"120x{font_menu.winfo_reqheight() + theme_menu.winfo_reqheight() + about_menu.winfo_reqheight() + 6}+{x}+{y}")
        
        # 点击其他地方关闭菜单
        menu_win.bind("<FocusOut>", self.close_menu_on_focus_out)
        menu_win.focus_set()
    
    def close_menu(self, event=None):
        """关闭菜单"""
        if self.menu_window is not None:
            try:
                self.menu_window.destroy()
            except tk.TclError:
                pass
            self.menu_window = None
        self.menu_open = False
        self.root.focus_set()
    
    def close_menu_on_focus_out(self, event=None):
        """失去焦点时关闭菜单"""
        # 延迟关闭，避免点击菜单项时立即关闭
        self.root.after(100, self.close_menu)
    
    def toggle_theme_from_menu(self, event=None):
        """从菜单切换主题"""
        self.clicked_menu_item = True
        self.close_menu()
        self.start_theme_transition()

    def show_about_from_menu(self, event=None):
        """从菜单显示关于窗口"""
        self.clicked_menu_item = True
        self.close_menu()
        self.show_help()
    
    def show_help(self, event=None):
        """显示帮助窗口"""
        # 如果关于窗口已经存在，则将其置于最前
        if self.about_window is not None:
            try:
                self.about_window.lift()
                self.about_window.focus_force()
                return
            except tk.TclError:
                # 窗口已被销毁
                self.about_window = None
        
        # 创建关于窗口
        self.create_about_window()
    
    def create_about_window(self):
        """创建关于窗口"""
        self.about_window = tk.Toplevel(self.root)
        about_win = self.about_window
        
        # 设置窗口属性
        about_win.title("关于 Loi")
        about_win.geometry("500x500")
        about_win.configure(bg=self.current_colors['bg'])
        about_win.overrideredirect(True)  # 无边框
        
        # 确保关于窗口在主窗口前面
        about_win.attributes('-topmost', True)
        
        # 居中显示
        self.center_window(about_win, 500, 500)
        
        # 创建主容器
        about_main_frame = tk.Frame(
            about_win, 
            bg=self.current_colors['bg'], 
            padx=20, 
            pady=20,
            highlightthickness=1,
            highlightbackground=self.current_colors['window_border']
        )
        about_main_frame.pack(expand=True, fill=tk.BOTH, padx=1, pady=1)
        
        # 标题栏框架
        about_title_frame = tk.Frame(about_main_frame, bg=self.current_colors['bg'])
        about_title_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 标题文本
        about_title_label = tk.Label(
            about_title_frame,
            text="关于 Loi 账单记录",
            font=("Helvetica", 12, "bold"),
            fg=self.current_colors['fg'],
            bg=self.current_colors['bg']
        )
        about_title_label.pack(side=tk.LEFT)
        
        # 关闭按钮
        about_close_btn = tk.Label(
            about_title_frame,
            text="×",
            font=("Helvetica", 14),
            fg=self.current_colors['hint_fg'],
            bg=self.current_colors['bg'],
            padx=8,
            pady=0,
            cursor="hand2"
        )
        about_close_btn.pack(side=tk.RIGHT)
        about_close_btn.bind("<Button-1>", self.close_about_window)
        about_close_btn.bind("<Enter>", lambda e: about_close_btn.configure(bg="#ff4444", fg="white"))
        about_close_btn.bind("<Leave>", lambda e: about_close_btn.configure(bg=self.current_colors['bg'], fg=self.current_colors['hint_fg']))
        
        # 内容框架
        content_frame = tk.Frame(about_main_frame, bg=self.current_colors['bg'])
        content_frame.pack(expand=True, fill=tk.BOTH)
        
        # 创建文本框和自定义滚动条
        text_frame = tk.Frame(content_frame, bg=self.current_colors['bg'])
        text_frame.pack(expand=True, fill=tk.BOTH, pady=(0, 15))
        
        # 自定义滚动条（无边框）
        scrollbar = tk.Canvas(
            text_frame,
            width=12,
            bg=self.current_colors['scrollbar_bg'],
            highlightthickness=0
        )
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(2, 0))
        
        # 文本框
        text_widget = tk.Text(
            text_frame,
            font=("Helvetica", 9),
            fg=self.current_colors['fg'],
            bg=self.current_colors['bg'],
            wrap=tk.WORD,
            padx=10,
            pady=10,
            relief=tk.FLAT,
            highlightthickness=0,
            selectbackground=self.current_colors['button_hover']
        )
        text_widget.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        
        # 插入内容
        help_content = """Loi 账单记录使用说明

第一次使用：
    第一次使用时程序通常没有可用的账单文件，
    您需要通过在文件操作中新建文件后才能够正常使用功能。

基本操作:
- 选择账单文件: 从下拉菜单选择已有的账单文件
- 新建账单: 点击"新建"按钮创建新的账单文件
- 刷新文件列表: 点击"刷新"按钮重新加载文件列表

条目操作:
- 新增条目: 填写表单后点击"新增"按钮或按Ctrl+N
- 修改条目: 选择条目后修改表单内容，点击"修改"按钮或按Ctrl+U
- 删除条目: 选择条目后点击"删除"按钮或按Delete键
- 查找条目: 点击"查找"按钮或按Ctrl+F，输入关键词查找
- 移动条目: 选择条目后点击"上移/下移"按钮或按Ctrl+Up/Down
- 排序显示: 点击列标题进行排序，再次点击切换排序方向
- 重置显示: 点击"重置显示"按钮或按Ctrl+R恢复原始顺序

快捷键:
- Ctrl+N: 新增条目
- Ctrl+U: 修改选中条目
- Delete: 删除选中条目
- Ctrl+F: 查找条目
- Ctrl+S: 保存文件
- Ctrl+Z: 撤销操作
- Ctrl+H: 显示帮助
- Ctrl+M: 高级统计
- Ctrl+上/下: 上下移动选中条目
- Ctrl+R: 重置显示顺序
- Ctrl+加号/减号: 调整字体大小
- Ctrl+鼠标滚轮: 调整字体大小
- R: 切换深色或浅色模式
- F4: 紧急关闭程序

统计功能:
- 总流水: 显示所有条目的流水合计
- 选中流水: 显示选中条目的流水合计
- 同类流水: 显示与选中条目同名的所有条目的流水合计
- 高级统计: 点击"统计"按钮或按Ctrl+M，可以进行多条件筛选统计

数据格式:
- 支出: 直接输入数字(如15、1.5)
- 收入: 以"+"开头的数字(如+3、+1.2)

此软件是由 Momstor 开发，软件不会收集您的任何数据，
所有数据均保存在设备本地，因此对您的数据也不承担任何存储或备份的义务。
此软件是开源的，您使用即须遵守开源协议的所有条款。
        """
        text_widget.insert(tk.END, help_content)
        text_widget.config(state=tk.DISABLED)  # 只读
        
        # 创建滚动条滑块
        self.scrollbar_thumb = scrollbar.create_rectangle(
            2, 0, 10, 20,
            fill=self.current_colors['scrollbar_thumb'],
            outline=""
        )
        
        # 绑定滚动事件
        def on_text_scroll(first, last):
            # 更新滑块位置
            self.update_scrollbar_thumb(scrollbar, text_widget, first, last)
        
        def on_mousewheel(event):
            # 处理鼠标滚轮事件
            text_widget.yview_scroll(int(-1*(event.delta/120)), "units")
            # 手动更新滚动条
            try:
                first, last = text_widget.yview()
                self.update_scrollbar_thumb(scrollbar, text_widget, first, last)
            except:
                pass
        
        # 配置文本框的滚动命令
        text_widget.config(yscrollcommand=on_text_scroll)
        text_widget.bind("<MouseWheel>", on_mousewheel)
        scrollbar.bind("<MouseWheel>", on_mousewheel)
        
        # 初始化滚动条位置
        try:
            first, last = text_widget.yview()
            self.update_scrollbar_thumb(scrollbar, text_widget, first, last)
        except:
            pass
        
        # 确认按钮
        confirm_btn = tk.Label(
            about_main_frame,
            text="确认",
            font=("Helvetica", 10),
            fg="white",
            bg=self.current_colors['confirm_button'],
            padx=20,
            pady=8,
            cursor="hand2"
        )
        confirm_btn.pack(pady=(0, 10))
        confirm_btn.bind("<Button-1>", self.close_about_window)
        confirm_btn.bind("<Enter>", lambda e: confirm_btn.configure(bg=self.darken_color(self.current_colors['confirm_button'])))
        confirm_btn.bind("<Leave>", lambda e: confirm_btn.configure(bg=self.current_colors['confirm_button']))
        
        # 绑定标题栏拖动事件
        about_title_frame.bind("<ButtonPress-1>", lambda e: self.on_about_title_press(e, about_win))
        about_title_frame.bind("<ButtonRelease-1>", lambda e: self.on_about_title_release(e, about_win))
        about_title_frame.bind("<B1-Motion>", lambda e: self.on_about_title_motion(e, about_win))
        about_title_label.bind("<ButtonPress-1>", lambda e: self.on_about_title_press(e, about_win))
        about_title_label.bind("<ButtonRelease-1>", lambda e: self.on_about_title_release(e, about_win))
        about_title_label.bind("<B1-Motion>", lambda e: self.on_about_title_motion(e, about_win))
        
        # 窗口关闭事件
        about_win.protocol("WM_DELETE_WINDOW", self.close_about_window)
    
    def update_scrollbar_thumb(self, scrollbar_canvas, text_widget, first, last):
        """更新滚动条滑块位置和大小"""
        try:
            # 计算滑块位置和大小
            canvas_height = scrollbar_canvas.winfo_height()
            if canvas_height <= 1:  # 防止除零错误
                return
                
            thumb_height = max(20, int(canvas_height * (float(last) - float(first))))
            thumb_y = int(float(first) * canvas_height)
            
            # 更新滑块
            scrollbar_canvas.coords(
                self.scrollbar_thumb,
                2, thumb_y, 10, thumb_y + thumb_height
            )
        except:
            pass
    
    def on_about_title_press(self, event, window):
        """关于窗口标题栏按下事件"""
        self.mouse_pressed = True
        self.press_start_x = event.x_root
        self.press_start_y = event.y_root
        self.is_dragging = False
        self.window_start_x = window.winfo_x()
        self.window_start_y = window.winfo_y()
    
    def on_about_title_motion(self, event, window):
        """关于窗口标题栏拖动事件"""
        if not self.mouse_pressed:
            return
            
        dx = abs(event.x_root - self.press_start_x)
        dy = abs(event.y_root - self.press_start_y)
        
        if dx > self.drag_threshold or dy > self.drag_threshold:
            self.is_dragging = True
            new_x = self.window_start_x + (event.x_root - self.press_start_x)
            new_y = self.window_start_y + (event.y_root - self.press_start_y)
            window.geometry(f"+{int(new_x)}+{int(new_y)}")
    
    def on_about_title_release(self, event, window):
        """关于窗口标题栏释放事件"""
        self.mouse_pressed = False
    
    def close_about_window(self, event=None):
        """关闭关于窗口"""
        if self.about_window is not None:
            try:
                self.about_window.destroy()
            except tk.TclError:
                pass
            self.about_window = None
            # 重新获取主窗口焦点
            self.root.focus_set()
    
    def center_window(self, window, width, height):
        """居中窗口"""
        # 获取屏幕尺寸
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        
        # 计算居中位置
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        window.geometry(f"{width}x{height}+{x}+{y}")
    
    def darken_color(self, hex_color, factor=0.8):
        """加深颜色"""
        hex_color = hex_color.lstrip('#')
        rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        darker_rgb = tuple(max(0, int(c * factor)) for c in rgb)
        return f"#{darker_rgb[0]:02x}{darker_rgb[1]:02x}{darker_rgb[2]:02x}"
    
    def start_theme_transition(self, event=None):
        """开始主题渐变切换"""
        # 如果关于窗口存在，不处理主题切换
        if self.about_window is not None:
            return
            
        if self.animating:
            return
            
        self.animating = True
        target_mode = 1 - self.theme_mode  # 切换到另一个主题
        self.animate_theme_transition(target_mode)
    
    def animate_theme_transition(self, target_mode, step=0):
        """执行主题渐变动画"""
        if step > 10:  # 动画完成
            self.theme_mode = target_mode
            self.animating = False
            return
        
        # 计算渐变颜色
        start_theme = self.themes[1 - target_mode]
        end_theme = self.themes[target_mode]
        
        for key in self.current_colors:
            if key in start_theme and key in end_theme:
                start_color = self.hex_to_rgb(start_theme[key])
                end_color = self.hex_to_rgb(end_theme[key])
                
                # 线性插值计算当前步骤的颜色
                r = int(start_color[0] + (end_color[0] - start_color[0]) * step / 10)
                g = int(start_color[1] + (end_color[1] - start_color[1]) * step / 10)
                b = int(start_color[2] + (end_color[2] - start_color[2]) * step / 10)
                
                self.current_colors[key] = self.rgb_to_hex(r, g, b)
        
        # 应用当前颜色
        self.apply_current_colors()
        
        # 继续动画
        self.root.after(20, lambda: self.animate_theme_transition(target_mode, step + 1))
    
    def hex_to_rgb(self, hex_color):
        """将十六进制颜色转换为RGB元组"""
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    
    def rgb_to_hex(self, r, g, b):
        """将RGB元组转换为十六进制颜色"""
        return f"#{r:02x}{g:02x}{b:02x}"
    
    def apply_current_colors(self):
        """应用当前颜色到所有组件"""
        self.root.configure(bg=self.current_colors['bg'])
        self.main_frame.configure(
            bg=self.current_colors['bg'],
            highlightbackground=self.current_colors['window_border']
        )
        
        # 更新标题栏组件
        self.title_frame.configure(bg=self.current_colors['bg'])
        self.menu_frame.configure(bg=self.current_colors['bg'])
        self.control_frame.configure(bg=self.current_colors['bg'])
        self.title_label.configure(
            fg=self.current_colors['hint_fg'],
            bg=self.current_colors['bg']
        )
        
        # 更新菜单按钮颜色
        self.more_button.configure(
            fg=self.current_colors['hint_fg'],
            bg=self.current_colors['bg']
        )
        
        # 更新窗口控制按钮颜色
        self.minimize_btn.configure(
            fg=self.current_colors['hint_fg'],
            bg=self.current_colors['bg']
        )
        self.close_btn.configure(
            fg=self.current_colors['hint_fg'],
            bg=self.current_colors['bg']
        )
        
        # 更新文件操作按钮
        self.new_btn.configure(
            fg=self.current_colors['hint_fg'],
            bg=self.current_colors['bg']
        )
        self.refresh_btn.configure(
            fg=self.current_colors['hint_fg'],
            bg=self.current_colors['bg']
        )
        
        # 更新操作按钮
        for btn in [self.add_btn, self.update_btn, self.delete_btn, self.search_btn,
                   self.save_btn, self.undo_btn, self.stats_btn, self.reset_btn,
                   self.up_btn, self.down_btn]:
            btn.configure(
                fg=self.current_colors['hint_fg'],
                bg=self.current_colors['bg']
            )
        
        # 更新Treeview样式
        self.update_treeview_style()
        
        # 更新所有其他文本组件
        for widget in self.main_frame.winfo_children():
            if isinstance(widget, tk.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, tk.Label) and hasattr(child, 'cget') and child.cget('bg') == self.themes[1 - self.theme_mode]['bg']:
                        child.configure(
                            fg=self.current_colors['fg'],
                            bg=self.current_colors['bg']
                        )
        
        # 如果菜单窗口存在，也更新其颜色
        if self.menu_window is not None:
            try:
                self.menu_window.configure(bg=self.current_colors['bg'])
                # 更新菜单项颜色
                for child in self.menu_window.winfo_children():
                    if isinstance(child, tk.Frame):
                        child.configure(
                            bg=self.current_colors['bg'],
                            highlightbackground=self.current_colors['window_border']
                        )
                        for grandchild in child.winfo_children():
                            if isinstance(grandchild, tk.Label):
                                grandchild.configure(
                                    fg=self.current_colors['fg'],
                                    bg=self.current_colors['bg']
                                )
            except tk.TclError:
                pass
        
        # 如果关于窗口存在，也更新其颜色
        if self.about_window is not None:
            try:
                self.about_window.configure(bg=self.current_colors['bg'])
            except tk.TclError:
                pass
    
    def on_control_enter(self, event):
        """控制按钮悬停效果（通用）"""
        event.widget.configure(bg=self.current_colors['button_hover'])
    
    def on_control_leave(self, event):
        """控制按钮离开效果（通用）"""
        event.widget.configure(bg=self.current_colors['bg'])
    
    def on_window_press(self, event):
        """整个窗口鼠标按下事件"""
        # 如果关于窗口存在，不处理主窗口的拖动
        if self.about_window is not None:
            return
            
        self.mouse_pressed = True
        self.press_start_x = event.x_root
        self.press_start_y = event.y_root
        self.is_dragging = False
        self.clicked_menu_item = False  # 重置菜单点击标记
        
        # 保存窗口当前位置用于拖动计算
        self.window_start_x = self.root.winfo_x()
        self.window_start_y = self.root.winfo_y()
    
    def on_window_motion(self, event):
        """整个窗口鼠标拖动事件"""
        # 如果关于窗口存在，不处理主窗口的拖动
        if self.about_window is not None:
            return
            
        if not self.mouse_pressed:
            return
            
        # 计算移动距离
        dx = abs(event.x_root - self.press_start_x)
        dy = abs(event.y_root - self.press_start_y)
        
        # 如果移动超过阈值，则认为是拖动
        if dx > self.drag_threshold or dy > self.drag_threshold:
            self.is_dragging = True
            # 执行窗口拖动
            new_x = self.window_start_x + (event.x_root - self.press_start_x)
            new_y = self.window_start_y + (event.y_root - self.press_start_y)
            self.root.geometry(f"+{int(new_x)}+{int(new_y)}")
    
    def on_window_release(self, event):
        """整个窗口鼠标释放事件"""
        # 如果关于窗口存在，不处理主窗口的点击
        if self.about_window is not None:
            return
            
        self.mouse_pressed = False
        
        # 如果不是在拖动，且没有点击菜单项，则检查是否点击了按钮
        if not self.is_dragging and not self.clicked_menu_item:
            widget = event.widget
            # 检查是否点击在控制按钮上
            control_widgets = [self.minimize_btn, self.close_btn, self.more_button,
                              self.new_btn, self.refresh_btn]
            
            # 如果不在控制组件上，则忽略点击
            if widget not in control_widgets:
                # 检查是否点击了操作按钮
                operation_widgets = [self.add_btn, self.update_btn, self.delete_btn, 
                                    self.search_btn, self.save_btn, self.undo_btn, 
                                    self.stats_btn, self.reset_btn, self.up_btn, self.down_btn]
                
                if widget not in operation_widgets:
                    # 如果点击在Treeview上，确保能正常选择
                    if widget != self.tree:
                        # 让Treeview获取焦点
                        self.tree.focus_set()
        
        # 重置菜单点击标记（延迟一点时间，确保所有相关事件都已处理）
        self.root.after(10, lambda: setattr(self, 'clicked_menu_item', False))
    
    def on_title_press(self, event):
        """标题栏按下事件"""
        # 如果关于窗口存在，不处理主窗口的拖动
        if self.about_window is not None:
            return
            
        self.mouse_pressed = True
        self.press_start_x = event.x_root
        self.press_start_y = event.y_root
        self.is_dragging = False
        self.clicked_menu_item = False  # 重置菜单点击标记
        
        # 保存窗口当前位置用于拖动计算
        self.window_start_x = self.root.winfo_x()
        self.window_start_y = self.root.winfo_y()
    
    def on_title_motion(self, event):
        """标题栏拖动事件"""
        # 如果关于窗口存在，不处理主窗口的拖动
        if self.about_window is not None:
            return
            
        if not self.mouse_pressed:
            return
            
        # 计算移动距离
        dx = abs(event.x_root - self.press_start_x)
        dy = abs(event.y_root - self.press_start_y)
        
        # 如果移动超过阈值，则认为是拖动
        if dx > self.drag_threshold or dy > self.drag_threshold:
            self.is_dragging = True
            # 执行窗口拖动
            new_x = self.window_start_x + (event.x_root - self.press_start_x)
            new_y = self.window_start_y + (event.y_root - self.press_start_y)
            self.root.geometry(f"+{int(new_x)}+{int(new_y)}")
    
    def on_title_release(self, event):
        """标题栏释放事件"""
        # 如果关于窗口存在，不处理主窗口的拖动
        if self.about_window is not None:
            return
            
        self.mouse_pressed = False
        # 标题栏不触发其他操作
    
    def on_close_enter(self, event):
        """关闭按钮特殊悬停效果"""
        self.close_btn.configure(bg="#ff4444", fg="white")
    
    def on_close_leave(self, event):
        """关闭按钮特殊离开效果"""
        self.close_btn.configure(bg=self.current_colors['bg'], fg=self.current_colors['hint_fg'])
    
    def minimize_window(self, event=None):
        """最小化窗口 - 使用withdraw替代iconify"""
        # 如果关于窗口存在，不处理最小化
        if self.about_window is not None:
            return
            
        self.root.withdraw()  # 隐藏窗口
        # 创建一个临时的顶层窗口来恢复原窗口
        temp_window = tk.Toplevel()
        temp_window.withdraw()
        temp_window.after(100, lambda: self.restore_window(temp_window))
    
    def restore_window(self, temp_window):
        """恢复窗口显示"""
        temp_window.destroy()
        self.root.deiconify()
        self.root.lift()
    
    def handle_f4_key(self, event=None):
        """处理F4键 - 紧急关闭"""
        self.close_window()
    
    def close_window(self, event=None):
        """关闭窗口"""
        # 先检查是否有未保存的修改
        if self.modified:
            if messagebox.askyesno("保存修改", "当前文件已修改，是否保存？"):
                self.save_file()
        
        # 关闭所有子窗口
        self.close_menu()
        self.close_about_window()
        self.root.destroy()
    
    def set_font_size(self, size):
        self.font_size = size
        self.update_treeview_style()
        self.log_message(f"字体大小已设置为: {size}")
        
    def increase_font(self, event=None):
        self.font_size = min(20, self.font_size + 1)
        self.update_treeview_style()
        self.log_message(f"字体大小已增加至: {self.font_size}")
        
    def decrease_font(self, event=None):
        self.font_size = max(8, self.font_size - 1)
        self.update_treeview_style()
        self.log_message(f"字体大小已减小至: {self.font_size}")
        
    def on_mousewheel(self, event):
        if event.delta > 0:
            self.increase_font()
        else:
            self.decrease_font()
    
    def log_message(self, message):
        """在日志区域添加消息"""
        now = datetime.now().strftime("%H:%M")
        self.log_text.insert(tk.END, f"[{now}] {message}\n")
        self.log_text.see(tk.END)  # 自动滚动到底部
        
    def move_up(self):
        """上移选中条目"""
        if not self.selected_items:
            return
        # 如果当前是排序状态，先重置显示
        if self.sort_column is not None:
            self.reset_display()
        # 保存当前状态以便撤销
        self.save_state()
        self.modified = True
        # 获取所有选中项目的索引
        indices = [self.tree.index(item) for item in self.selected_items]
        # 如果最上面的项目已经是第一个，则不能上移
        if min(indices) == 0:
            return
        # 移动数据
        for index in sorted(indices):
            # 交换数据
            self.bill_data[index], self.bill_data[index-1] = self.bill_data[index-1], self.bill_data[index]
        if self.sort_column is None:
            self.display_data = self.bill_data.copy()
        # 更新Treeview
        self.refresh_treeview()
        # 重新选中移动后的项目
        new_indices = [i-1 for i in indices]
        self.tree.selection_set([self.tree.get_children()[i] for i in new_indices])
        self.log_message("已上移选中条目")
        
    def move_down(self):
        """下移选中条目"""
        if not self.selected_items:
            return
        # 如果当前是排序状态，先重置显示
        if self.sort_column is not None:
            self.reset_display()
        # 保存当前状态以便撤销
        self.save_state()
        self.modified = True
        # 获取所有选中项目的索引
        indices = [self.tree.index(item) for item in self.selected_items]
        # 如果最下面的项目已经是最后一个，则不能下移
        if max(indices) == len(self.bill_data) - 1:
            return
        # 移动数据（从下往上处理）
        for index in sorted(indices, reverse=True):
            # 交换数据
            self.bill_data[index], self.bill_data[index+1] = self.bill_data[index+1], self.bill_data[index]
        if self.sort_column is None:
            self.display_data = self.bill_data.copy()
        # 更新Treeview
        self.refresh_treeview()
        # 重新选中移动后的项目
        new_indices = [i+1 for i in indices]
        self.tree.selection_set([self.tree.get_children()[i] for i in new_indices])
        self.log_message("已下移选中条目")
        
    def refresh_treeview(self):
        """刷新Treeview显示"""
        self.tree.delete(*self.tree.get_children())
        for entry in self.display_data:
            self.tree.insert("", "end", values=(entry.date, entry.name, entry.amount, entry.note))
        
    def sort_treeview(self, column):
        """根据列进行排序"""
        # 如果点击的是当前排序列，则切换排序方向
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False
        
        # 更新表头箭头指示
        for col in ["date", "name", "amount", "note"]:
            if col == column:
                arrow = " ↓" if self.sort_reverse else " ↑"
                self.tree.heading(col, text=self.tree.heading(col)["text"].replace(" ↑", "").replace(" ↓", "") + arrow)
            else:
                self.tree.heading(col, text=self.tree.heading(col)["text"].replace(" ↑", "").replace(" ↓", ""))
        
        # 对显示数据进行排序
        if column == "date":
            self.display_data.sort(key=lambda x: x.date, reverse=self.sort_reverse)
        elif column == "name":
            self.display_data.sort(key=lambda x: x.name, reverse=self.sort_reverse)
        elif column == "amount":
            # 特殊处理：将金额转换为数字进行排序
            def amount_key(x):
                try:
                    if x.amount.startswith('+'):
                        return float(x.amount[1:])
                    else:
                        return float(x.amount)
                except ValueError:
                    return 0
            self.display_data.sort(key=amount_key, reverse=self.sort_reverse)
        elif column == "note":
            self.display_data.sort(key=lambda x: x.note, reverse=self.sort_reverse)
        
        # 刷新显示
        self.refresh_treeview()
        self.log_message(f"已按{column} {'降序' if self.sort_reverse else '升序'}排序")
        
    def reset_display(self):
        """重置显示为原始顺序"""
        self.sort_column = None
        self.sort_reverse = False
        
        # 清除表头箭头
        for col in ["date", "name", "amount", "note"]:
            self.tree.heading(col, text=self.tree.heading(col)["text"].replace(" ↑", "").replace(" ↓", ""))
        
        # 恢复原始显示顺序
        self.display_data = self.bill_data.copy()
        self.refresh_treeview()
        self.log_message("已重置显示顺序")
        
    def load_available_files(self):
        files = [f for f in os.listdir('.') if f.endswith('.md') and re.match(r'^\d{6}\.md$', f)]
        self.file_combo['values'] = files
        if files and not self.file_var.get():
            self.file_var.set(files[0])
            self.load_file(files[0])
            
    def on_file_select(self, event):
        if self.modified:
            if messagebox.askyesno("保存修改", "当前文件已修改，是否保存？"):
                self.save_file()
            self.modified = False
        self.load_file(self.file_var.get())
        
    def load_file(self, filename):
        if not filename:
            return
            
        self.current_file = filename
        self.bill_data = []
        self.display_data = []
        self.tree.delete(*self.tree.get_children())
        
        # 重置排序状态
        self.sort_column = None
        self.sort_reverse = False
        for col in ["date", "name", "amount", "note"]:
            self.tree.heading(col, text=self.tree.heading(col)["text"].replace(" ↑", "").replace(" ↓", ""))
        
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                
            # 跳过文件头
            start_index = 0
            for i, line in enumerate(lines):
                if line.startswith('| 日期'):
                    start_index = i + 2  # 跳过表头和分隔线
                    break
                    
            # 解析数据行
            for i in range(start_index, len(lines)):
                line = lines[i].strip()
                if not line or not line.startswith('|'):
                    continue
                    
                # 解析markdown表格行
                parts = [part.strip() for part in line.split('|')[1:-1]]
                if len(parts) >= 3:
                    date, name, amount = parts[0], parts[1], parts[2]
                    note = parts[3] if len(parts) > 3 else ""
                    self.bill_data.append(BillEntry(date, name, amount, note))
                    self.tree.insert("", "end", values=(date, name, amount, note))
                    
            # 初始化显示数据
            self.display_data = self.bill_data.copy()
            
            self.calculate_totals()
            self.modified = False
            self.log_message(f"已加载文件: {filename}")
            
        except Exception as e:
            messagebox.showerror("错误", f"加载文件时出错: {str(e)}")
            
    def new_file(self):
        """打开年月选择弹窗创建新文件"""
        # 创建年月选择对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("新建账单文件")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()  # 模态窗口
        
        # 设置对话框居中显示
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # 年份选择
        tk.Label(dialog, text="选择年份:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        year_var = tk.StringVar()
        current_year = datetime.now().year
        years = [str(y) for y in range(current_year-10, current_year+11)]  # 当前年份前后10年
        year_combo = ttk.Combobox(dialog, textvariable=year_var, values=years, width=8)
        year_combo.set(str(current_year))
        year_combo.grid(row=0, column=1, padx=10, pady=10)
        
        # 月份选择
        tk.Label(dialog, text="选择月份:").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        month_var = tk.StringVar()
        months = [f"{m:02d}" for m in range(1, 13)]  # 01-12
        month_combo = ttk.Combobox(dialog, textvariable=month_var, values=months, width=8)
        month_combo.set(f"{datetime.now().month:02d}")
        month_combo.grid(row=1, column=1, padx=10, pady=10)
        
        # 按钮区域
        btn_frame = tk.Frame(dialog)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        def create_file():
            year = year_var.get()
            month = month_var.get()
            
            if not year or not month:
                messagebox.showwarning("警告", "请选择完整的年月")
                return
                
            # 生成文件名
            filename = f"{year}{month}.md"
            
            # 检查文件是否存在
            if filename in self.file_combo['values']:
                if messagebox.askyesno("确认", f"文件 {filename} 已存在，是否覆盖?"):
                    self.create_and_load_file(filename)
                    dialog.destroy()
            else:
                self.create_and_load_file(filename)
                dialog.destroy()
        
        ttk.Button(btn_frame, text="确定", command=create_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
    def create_and_load_file(self, filename):
        """创建新的账单文件并加载"""
        # 创建文件头
        content = f"""# {filename[:4]}年{filename[4:6]}月账单

| 日期 | 名称 | 流水 | 备注 |
| ---- | ---- | ---- | ---- |
"""
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(content)
                
            self.log_message(f"已创建新文件: {filename}")
            self.load_available_files()
            self.file_var.set(filename)
            self.load_file(filename)
            
        except Exception as e:
            messagebox.showerror("错误", f"创建文件时出错: {str(e)}")
            
    def save_file(self):
        if not self.current_file:
            messagebox.showwarning("警告", "没有打开的文件")
            return
            
        try:
            # 生成文件内容
            content = f"""# {self.current_file[:4]}年{self.current_file[4:6]}月账单

| 日期 | 名称 | 流水 | 备注 |
| ---- | ---- | ---- | ---- |
"""
            for entry in self.bill_data:
                content += f"| {entry.date} | {entry.name} | {entry.amount} | {entry.note} |\n"
                
            with open(self.current_file, 'w', encoding='utf-8') as f:
                f.write(content)
                
            self.modified = False
            self.log_message(f"已保存文件: {self.current_file}")
            
        except Exception as e:
            messagebox.showerror("错误", f"保存文件时出错: {str(e)}")
            
    def on_item_select(self, event):
        self.selected_items = list(self.tree.selection())
        self.update_form()
        self.calculate_totals()

    def update_form(self):
        """更新表单内容"""
        if len(self.selected_items) == 1:
            item = self.selected_items[0]
            values = self.tree.item(item, 'values')
            self.date_var.set(values[0])
            self.name_var.set(values[1])
            self.amount_var.set(values[2])
            self.note_var.set(values[3] if len(values) > 3 else "")
        else:
            self.clear_form()
            
    def clear_form(self):
        """清空表单"""
        self.date_var.set("")
        self.name_var.set("")
        self.amount_var.set("")
        self.note_var.set("")
        
    def add_item(self):
        """新增条目：插入到选中项之后，或末尾"""
        date = self.date_var.get().strip()
        name = self.name_var.get().strip()
        amount = self.amount_var.get().strip()
        note = self.note_var.get().strip()

        # 如果已有数据但表单为空，说明用户想新增，不报错
        # 但如果是第一次添加，允许用户填写表单后新增
        if not all([date, name, amount]):
            # 如果是占位状态（无数据），允许添加新条目而不报错，但需填写
            if not self.bill_data:
                messagebox.showwarning("警告", "请填写日期、名称和流水")
                return
            # 否则，允许插入空条目？我们不允许
            messagebox.showwarning("警告", "日期、名称和流水不能为空")
            return

        # 验证金额格式
        try:
            if amount.startswith('+'):
                float(amount[1:])
            else:
                float(amount)
        except ValueError:
            messagebox.showwarning("警告", "金额格式不正确")
            return

        # 保存状态用于撤销
        self.save_state()
        self.modified = True

        # 创建新条目
        new_entry = BillEntry(date, name, amount, note)

        # 确定插入位置
        insert_index = len(self.bill_data)  # 默认末尾
        if self.selected_items:
            # 插入到选中项的下一位
            last_selected = self.selected_items[-1]
            display_index = self.tree.index(last_selected)
            if self.sort_column is None:
                insert_index = display_index + 1
            else:
                insert_index = len(self.bill_data)

        # 插入到原始数据
        self.bill_data.insert(insert_index, new_entry)

        # 更新 display_data
        if self.sort_column is not None:
            self.reset_display()
        else:
            self.display_data.insert(insert_index, new_entry)
            self.refresh_treeview()

        # 选中新条目
        new_item = self.tree.get_children()[insert_index]
        self.tree.selection_set(new_item)
        self.tree.focus(new_item)
        self.tree.see(new_item)

        self.log_message(f"已添加条目: {name} {amount}")
        self.calculate_totals()
        self.clear_form()  # 清空表单，准备下一次输入
        
    def update_item(self):
        """修改选中条目，并在刷新后保持选中"""
        if not self.selected_items:
            messagebox.showwarning("警告", "请先选择要修改的条目")
            return

        date = self.date_var.get().strip()
        name = self.name_var.get().strip()
        amount = self.amount_var.get().strip()
        note = self.note_var.get().strip()

        if not all([date, name, amount]):
            messagebox.showwarning("警告", "日期、名称和流水不能为空")
            return

        # 验证金额格式
        try:
            if amount.startswith('+'):
                float(amount[1:])
            else:
                float(amount)
        except ValueError:
            messagebox.showwarning("警告", "金额格式不正确")
            return

        # 保存状态用于撤销
        self.save_state()
        self.modified = True

        # 记录每个选中项在 display_data 中的索引
        indices_to_update = []
        for item in self.selected_items:
            display_index = self.tree.index(item)
            indices_to_update.append(display_index)

        # 获取当前选中项对应的 display_data 条目
        updated_entries = []
        for idx in indices_to_update:
            old_entry = self.display_data[idx]
            # 创建新条目
            new_entry = BillEntry(date, name, amount, note)
            # 更新 display_data
            self.display_data[idx] = new_entry
            # 更新 bill_data：找到原始条目并替换
            try:
                original_index = self.bill_data.index(old_entry)
                self.bill_data[original_index] = new_entry
            except ValueError:
                # 安全查找匹配项
                for i, entry in enumerate(self.bill_data):
                    if entry == old_entry:
                        self.bill_data[i] = new_entry
                        break
            updated_entries.append(new_entry)

        # 刷新界面
        self.refresh_treeview()

        # 恢复选中：根据之前记录的索引重新选中
        children = self.tree.get_children()
        new_selection = []
        for idx in indices_to_update:
            if idx < len(children):
                new_item = children[idx]
                new_selection.append(new_item)

        if new_selection:
            self.tree.selection_set(new_selection)
            self.tree.focus(new_selection[-1])  # 焦点设到最后一个
            self.tree.see(new_selection[-1])    # 滚动到可见

        # 更新 selected_items 和表单
        self.selected_items = new_selection
        self.update_form()  # 更新表单显示新值

        self.log_message(f"已更新 {len(updated_entries)} 个条目")
        self.calculate_totals()
        
    def delete_item(self):
        """删除选中条目，并保持选中状态"""
        if not self.selected_items:
            messagebox.showwarning("警告", "请先选择要删除的条目")
            return
        # 保存状态
        self.save_state()
        self.modified = True
        # 获取要删除的原始索引
        indices_to_delete = []
        for item in self.selected_items:
            display_index = self.tree.index(item)
            old_entry = self.display_data[display_index]
            try:
                original_index = self.bill_data.index(old_entry)
                indices_to_delete.append(original_index)
            except ValueError:
                continue
        indices_to_delete.sort(reverse=True)
        for index in indices_to_delete:
            del self.bill_data[index]
        self.reset_display()
        children = self.tree.get_children()
        if children:
            # 选中原来位置附近的项 (使用新的 children 列表)
            last_idx = max(0, indices_to_delete[-1] - 1)
            last_idx = min(last_idx, len(children) - 1)
            new_item = children[last_idx]
            self.tree.selection_set(new_item)
            self.tree.focus(new_item)
            self.tree.see(new_item)
        else:
            # 表格为空，清空表单
            self.clear_form()
        self.selected_items = self.tree.selection()
        self.log_message(f"已删除 {len(indices_to_delete)} 个条目")
        self.calculate_totals()
        
    def search_item(self):
        keyword = simpledialog.askstring("查找", "请输入要查找的关键词:")
        if not keyword:
            return
            
        # 清除当前选择
        self.tree.selection_remove(self.tree.selection())
        
        # 查找匹配的条目
        found_items = []
        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            if any(keyword.lower() in str(value).lower() for value in values):
                found_items.append(item)
                
        if found_items:
            # 选中所有匹配的条目
            for item in found_items:
                self.tree.selection_add(item)
                self.tree.focus(item)
                self.tree.see(item)  # 滚动到可见位置
                
            self.log_message(f"找到 {len(found_items)} 个匹配的条目")
        else:
            messagebox.showinfo("查找结果", "没有找到匹配的条目")
            self.log_message(f"未找到包含\"{keyword}\"的条目")
            
    def calculate_totals(self):
        """计算总流水、选中流水和同类流水"""
        self.selected_items = [item for item in self.selected_items if self.tree.exists(item)]
        total = 0
        selected_total = 0
        same_type_total = 0
        
        # 计算总流水
        for entry in self.bill_data:
            amount = entry.amount
            if amount.startswith('+'):
                total += float(amount[1:])
            else:
                total -= float(amount)
                
        # 计算选中流水和同类名称
        selected_names = set()
        for item in self.selected_items:
            values = self.tree.item(item, 'values')
            amount = values[2]
            if amount.startswith('+'):
                selected_total += float(amount[1:])
            else:
                selected_total -= float(amount)
            selected_names.add(values[1])
            
        # 计算同类流水
        if selected_names:
            for entry in self.bill_data:
                if entry.name in selected_names:
                    amount = entry.amount
                    if amount.startswith('+'):
                        same_type_total += float(amount[1:])
                    else:
                        same_type_total -= float(amount)
        
        self.total_var.set(f"总流水: {total:.2f}")
        self.selected_var.set(f"选中流水: {selected_total:.2f}")
        self.same_type_var.set(f"同类流水: {same_type_total:.2f}")
        
    def save_state(self):
        """保存当前状态以便撤销"""
        self.undo_stack.append([BillEntry(e.date, e.name, e.amount, e.note) for e in self.bill_data])
        # 限制撤销栈大小
        if len(self.undo_stack) > 50:
            self.undo_stack.pop(0)
            
    def undo(self):
        """撤销操作"""
        if not self.undo_stack:
            messagebox.showinfo("提示", "没有可撤销的操作")
            return
            
        # 恢复上个状态
        self.bill_data = self.undo_stack.pop()
        
        # 刷新显示
        self.reset_display()
        self.calculate_totals()
        self.modified = True
        self.log_message("已撤销上一步操作")
        
    def show_statistics(self):
        """显示高级统计窗口"""
        if not self.bill_data:
            messagebox.showinfo("提示", "没有数据可统计")
            return
            
        stats_window = tk.Toplevel(self.root)
        stats_window.title("高级统计")
        stats_window.geometry("500x400")
        stats_window.transient(self.root)
        stats_window.grab_set()
        
        # 创建统计条件框架
        condition_frame = ttk.LabelFrame(stats_window, text="统计条件", padding="10")
        condition_frame.pack(fill="x", padx=10, pady=5)
        
        # 日期范围
        ttk.Label(condition_frame, text="日期范围:").grid(row=0, column=0, sticky="w", pady=2)
        date_frame = ttk.Frame(condition_frame)
        date_frame.grid(row=0, column=1, sticky="ew", pady=2)
        
        self.start_date_var = tk.StringVar()
        ttk.Entry(date_frame, textvariable=self.start_date_var, width=8).pack(side="left", padx=(0, 5))
        ttk.Label(date_frame, text="至").pack(side="left", padx=5)
        self.end_date_var = tk.StringVar()
        ttk.Entry(date_frame, textvariable=self.end_date_var, width=8).pack(side="left")
        
        # 名称筛选
        ttk.Label(condition_frame, text="名称包含:").grid(row=1, column=0, sticky="w", pady=2)
        self.name_filter_var = tk.StringVar()
        ttk.Entry(condition_frame, textvariable=self.name_filter_var, width=20).grid(row=1, column=1, sticky="w", pady=2)
        
        # 备注筛选
        ttk.Label(condition_frame, text="备注包含:").grid(row=2, column=0, sticky="w", pady=2)
        self.note_filter_var = tk.StringVar()
        ttk.Entry(condition_frame, textvariable=self.note_filter_var, width=20).grid(row=2, column=1, sticky="w", pady=2)
        
        # 金额类型
        ttk.Label(condition_frame, text="金额类型:").grid(row=3, column=0, sticky="w", pady=2)
        self.amount_type_var = tk.StringVar(value="全部")
        amount_frame = ttk.Frame(condition_frame)
        amount_frame.grid(row=3, column=1, sticky="w", pady=2)
        ttk.Radiobutton(amount_frame, text="全部", variable=self.amount_type_var, value="全部").pack(side="left")
        ttk.Radiobutton(amount_frame, text="收入", variable=self.amount_type_var, value="收入").pack(side="left", padx=(10, 0))
        ttk.Radiobutton(amount_frame, text="支出", variable=self.amount_type_var, value="支出").pack(side="left", padx=(10, 0))
        
        # 统计按钮
        button_frame = ttk.Frame(stats_window)
        button_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(button_frame, text="统计", command=lambda: self.calculate_advanced_stats(stats_window)).pack(side="left", padx=(0, 10))
        ttk.Button(button_frame, text="关闭", command=stats_window.destroy).pack(side="left")
        
        # 结果显示区域
        result_frame = ttk.LabelFrame(stats_window, text="统计结果", padding="10")
        result_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.stats_result_var = tk.StringVar()
        self.stats_result_var.set("请设置条件后点击\"统计\"按钮")
        ttk.Label(result_frame, textvariable=self.stats_result_var, wraplength=400).pack(anchor="w")
        
        # 居中显示
        stats_window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - stats_window.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - stats_window.winfo_height()) // 2
        stats_window.geometry(f"+{x}+{y}")
        
    def calculate_advanced_stats(self, stats_window):
        """根据条件计算高级统计"""
        # 获取筛选条件
        start_date = self.start_date_var.get().strip()
        end_date = self.end_date_var.get().strip()
        name_filter = self.name_filter_var.get().strip()
        note_filter = self.note_filter_var.get().strip()
        amount_type = self.amount_type_var.get()
        
        # 筛选数据
        filtered_data = []
        for entry in self.bill_data:
            # 日期筛选
            if start_date and end_date:
                try:
                    date_num = int(entry.date)
                    start_num = int(start_date)
                    end_num = int(end_date)
                    if not (start_num <= date_num <= end_num):
                        continue
                except ValueError:
                    # 如果日期不是数字，使用字符串比较
                    if not (start_date <= entry.date <= end_date):
                        continue
            elif start_date and entry.date != start_date:
                continue
            elif end_date and entry.date != end_date:
                continue
                
            # 名称筛选
            if name_filter and name_filter not in entry.name:
                continue
                
            # 备注筛选
            if note_filter and note_filter not in entry.note:
                continue
                
            # 金额类型筛选
            if amount_type == "收入" and not entry.amount.startswith('+'):
                continue
            if amount_type == "支出" and entry.amount.startswith('+'):
                continue
                
            filtered_data.append(entry)
        
        # 计算统计结果
        total = 0
        income_count = 0
        expense_count = 0
        income_total = 0
        expense_total = 0
        
        for entry in filtered_data:
            amount = entry.amount
            if amount.startswith('+'):
                amount_value = float(amount[1:])
                total += amount_value
                income_count += 1
                income_total += amount_value
            else:
                amount_value = float(amount)
                total -= amount_value
                expense_count += 1
                expense_total += amount_value
        
        # 显示结果
        result_text = f"符合条件的条目数: {len(filtered_data)}\n"
        result_text += f"总收入条目: {income_count}, 总支出条目: {expense_count}\n"
        result_text += f"总收入: {income_total:.2f}, 总支出: {expense_total:.2f}\n"
        result_text += f"净收入: {total:.2f}"
        
        self.stats_result_var.set(result_text)

if __name__ == "__main__":
    root = tk.Tk()
    app = ElegantBillApp(root)
    root.mainloop()