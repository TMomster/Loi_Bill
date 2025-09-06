import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import os
import re
from datetime import datetime

class BillEntry:
    def __init__(self, date, name, amount, note=""):
        self.date = date
        self.name = name
        self.amount = amount
        self.note = note

class BillApp:
    def __init__(self, root):
        self.root = root
        self.root.title("账单记录软件")
        self.root.geometry("1200x800")
        
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
        
        # 创建界面
        self.create_widgets()
        self.create_menu()
        self.load_available_files()
        
        # 绑定快捷键
        self.bind_shortcuts()
        
        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def create_menu(self):
        # 创建菜单栏
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 设置菜单
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="设置", menu=settings_menu)
        
        # 字体大小子菜单
        font_menu = tk.Menu(settings_menu, tearoff=0)
        settings_menu.add_cascade(label="字体大小", menu=font_menu)
        font_menu.add_command(label="小", command=lambda: self.set_font_size(8))
        font_menu.add_command(label="中", command=lambda: self.set_font_size(10))
        font_menu.add_command(label="大", command=lambda: self.set_font_size(12))
        
    def set_font_size(self, size):
        self.font_size = size
        self.update_font_size()
        
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件操作", padding="5")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(file_frame, text="选择账单文件:").grid(row=0, column=0, padx=(0, 5))
        self.file_var = tk.StringVar()
        self.file_combo = ttk.Combobox(file_frame, textvariable=self.file_var, width=25)
        self.file_combo.grid(row=0, column=1, padx=(0, 10))
        self.file_combo.bind('<<ComboboxSelected>>', self.on_file_select)
        
        ttk.Button(file_frame, text="新建", command=self.new_file).grid(row=0, column=2, padx=(0, 5))
        ttk.Button(file_frame, text="刷新", command=self.load_available_files).grid(row=0, column=3)
        
        # 账单展示区域
        display_frame = ttk.LabelFrame(main_frame, text="账单内容", padding="5")
        display_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
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
        
        # 工作区域
        work_frame = ttk.LabelFrame(main_frame, text="工作区域", padding="5")
        work_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        # 编辑表单
        form_frame = ttk.Frame(work_frame)
        form_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Label(form_frame, text="日期:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.date_var = tk.StringVar()
        date_entry = ttk.Entry(form_frame, textvariable=self.date_var, width=15)
        date_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        date_entry.bind('<Alt-Up>', lambda e: self.move_up())
        date_entry.bind('<Alt-Down>', lambda e: self.move_down())

        ttk.Label(form_frame, text="名称:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.name_var = tk.StringVar()
        name_entry = ttk.Entry(form_frame, textvariable=self.name_var, width=15)
        name_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        name_entry.bind('<Alt-Up>', lambda e: self.move_up())
        name_entry.bind('<Alt-Down>', lambda e: self.move_down())

        ttk.Label(form_frame, text="流水:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.amount_var = tk.StringVar()
        amount_entry = ttk.Entry(form_frame, textvariable=self.amount_var, width=15)
        amount_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        amount_entry.bind('<Alt-Up>', lambda e: self.move_up())
        amount_entry.bind('<Alt-Down>', lambda e: self.move_down())

        ttk.Label(form_frame, text="备注:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.note_var = tk.StringVar()
        note_entry = ttk.Entry(form_frame, textvariable=self.note_var, width=15)
        note_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        note_entry.bind('<Alt-Up>', lambda e: self.move_up())
        note_entry.bind('<Alt-Down>', lambda e: self.move_down())
        
        # 按钮区域
        btn_frame = ttk.Frame(work_frame)
        btn_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Button(btn_frame, text="新增", command=self.add_item).grid(row=0, column=0, padx=2, pady=5)
        ttk.Button(btn_frame, text="修改", command=self.update_item).grid(row=0, column=1, padx=2, pady=5)
        ttk.Button(btn_frame, text="删除", command=self.delete_item).grid(row=0, column=2, padx=2, pady=5)
        ttk.Button(btn_frame, text="查找", command=self.search_item).grid(row=1, column=0, padx=2, pady=5)
        ttk.Button(btn_frame, text="保存", command=self.save_file).grid(row=1, column=1, padx=2, pady=5)
        ttk.Button(btn_frame, text="撤销", command=self.undo).grid(row=1, column=2, padx=2, pady=5)
        ttk.Button(btn_frame, text="统计", command=self.show_statistics).grid(row=2, column=0, columnspan=3, pady=5, sticky=tk.W+tk.E)
        ttk.Button(btn_frame, text="上移", command=self.move_up).grid(row=3, column=0, padx=2, pady=5)
        ttk.Button(btn_frame, text="下移", command=self.move_down).grid(row=3, column=1, padx=2, pady=5)
        ttk.Button(btn_frame, text="重置显示", command=self.reset_display).grid(row=3, column=2, padx=2, pady=5)
        
        # 统计区域
        stats_frame = ttk.LabelFrame(main_frame, text="统计信息", padding="5")
        stats_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        self.total_var = tk.StringVar()
        self.total_var.set("总流水: 0.0")
        ttk.Label(stats_frame, textvariable=self.total_var, font=("Arial", 12)).grid(row=0, column=0, padx=(0, 20))
        
        self.selected_var = tk.StringVar()
        self.selected_var.set("选中流水: 0.0")
        ttk.Label(stats_frame, textvariable=self.selected_var, font=("Arial", 12)).grid(row=0, column=1, padx=(0, 20))
        
        self.same_type_var = tk.StringVar()
        self.same_type_var.set("同类流水: 0.0")
        ttk.Label(stats_frame, textvariable=self.same_type_var, font=("Arial", 12)).grid(row=0, column=2)
        
        # 字体大小提示
        ttk.Label(stats_frame, text="Ctrl+加号/减号或Ctrl+鼠标滚轮调整字体大小  <<按下 ctrl+H 查阅帮助>>", font=("Arial", 9)).grid(row=1, column=0, columnspan=3, pady=(5, 0))
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="日志", padding="5")
        log_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建日志文本框
        self.log_text = tk.Text(log_frame, height=4, wrap=tk.WORD)
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 配置权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        main_frame.columnconfigure(0, weight=3)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=0)  # 日志区域不需要太多空间
        
        display_frame.columnconfigure(0, weight=1)
        display_frame.rowconfigure(0, weight=1)
        
        work_frame.columnconfigure(0, weight=1)
        form_frame.columnconfigure(1, weight=1)
        
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
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
        
    def bind_shortcuts(self):
        self.root.bind('<Control-n>', lambda e: self.add_item())
        self.root.bind('<Control-N>', lambda e: self.add_item())
        self.root.bind('<Delete>', lambda e: self.delete_item())
        self.root.bind('<Control-u>', lambda e: self.update_item())
        self.root.bind('<Control-U>', lambda e: self.update_item())
        self.root.bind('<Control-f>', lambda e: self.search_item())
        self.root.bind('<Control-F>', lambda e: self.search_item())
        self.root.bind('<Control-s>', lambda e: self.save_file())
        self.root.bind('<Control-S>', lambda e: self.save_file())
        self.root.bind('<Control-z>', lambda e: self.undo())
        self.root.bind('<Control-Z>', lambda e: self.undo())
        self.root.bind('<Control-plus>', self.increase_font)
        self.root.bind('<Control-minus>', self.decrease_font)
        self.root.bind('<Control-MouseWheel>', self.on_mousewheel)
        self.root.bind('<Control-h>', lambda e: self.show_help())
        self.root.bind('<Control-H>', lambda e: self.show_help())
        self.root.bind('<Control-m>', lambda e: self.show_statistics())
        self.root.bind('<Control-M>', lambda e: self.show_statistics())
        self.root.bind('<Control-Up>', lambda e: self.move_up())
        self.root.bind('<Control-Down>', lambda e: self.move_down())
        self.root.bind('<Control-r>', lambda e: self.reset_display())
        self.root.bind('<Control-R>', lambda e: self.reset_display())
        
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
        
    def increase_font(self, event=None):
        self.font_size = min(20, self.font_size + 1)
        self.update_font_size()
        
    def decrease_font(self, event=None):
        self.font_size = max(8, self.font_size - 1)
        self.update_font_size()
        
    def on_mousewheel(self, event):
        if event.delta > 0:
            self.increase_font()
        else:
            self.decrease_font()
            
    def update_font_size(self):
        # 更新Treeview的字体大小
        style = ttk.Style()
        style.configure("Treeview", font=("Arial", self.font_size))
        style.configure("Treeview.Heading", font=("Arial", self.font_size, "bold"))
        
        # 更新日志区域的字体大小
        self.log_text.config(font=("Arial", self.font_size))
        
    def show_help(self):
        """显示帮助窗口"""
        help_window = tk.Toplevel(self.root)
        help_window.title("帮助")
        help_window.geometry("600x400")
        help_window.transient(self.root)  # 设置为 transient 窗口
        help_window.grab_set()  # 模态窗口，必须先关闭才能操作主窗口
        
        # 帮助内容
        help_text = """账单记录软件使用说明

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
- Alt+上/下: 切换条目选择

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
        
        text_widget = tk.Text(help_window, wrap="word", padx=10, pady=10)
        text_widget.insert("1.0", help_text)
        text_widget.config(state="disabled")
        
        scrollbar = ttk.Scrollbar(help_window, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 关闭按钮
        close_button = ttk.Button(help_window, text="关闭", command=help_window.destroy)
        close_button.pack(pady=10)
        
        # 居中显示
        help_window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - help_window.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - help_window.winfo_height()) // 2
        help_window.geometry(f"+{x}+{y}")
        
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
        ttk.Label(dialog, text="选择年份:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        year_var = tk.StringVar()
        current_year = datetime.now().year
        years = [str(y) for y in range(current_year-10, current_year+11)]  # 当前年份前后10年
        year_combo = ttk.Combobox(dialog, textvariable=year_var, values=years, width=8)
        year_combo.set(str(current_year))
        year_combo.grid(row=0, column=1, padx=10, pady=10)
        
        # 月份选择
        ttk.Label(dialog, text="选择月份:").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        month_var = tk.StringVar()
        months = [f"{m:02d}" for m in range(1, 13)]  # 01-12
        month_combo = ttk.Combobox(dialog, textvariable=month_var, values=months, width=8)
        month_combo.set(f"{datetime.now().month:02d}")
        month_combo.grid(row=1, column=1, padx=10, pady=10)
        
        # 按钮区域
        btn_frame = ttk.Frame(dialog)
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
        
        ttk.Button(btn_frame, text="确定", command=create_file).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side="left", padx=5)
        
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
        
    def on_closing(self):
        """处理窗口关闭事件"""
        if self.modified:
            if messagebox.askyesno("保存修改", "当前文件已修改，是否保存？"):
                self.save_file()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = BillApp(root)
    root.mainloop()
