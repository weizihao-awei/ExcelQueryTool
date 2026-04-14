#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
UI 组件模块
负责用户界面的创建和管理
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from .utils import calc_text_width


class ExcelFilterUI:
    """Excel 筛选工具 UI 类"""
    
    def __init__(self, root, callbacks):
        self.root = root
        self.callbacks = callbacks  # 回调函数字典
        self.root.title("Excel 筛选处理工具")
        # 设置默认全屏显示
        self.root.state('zoomed')
        self.root.minsize(1000, 600)
        
        # 数据存储
        self.filter_widgets = {}  # 筛选控件字典
        self.sheet_names = []  # 工作表名称列表
        
        # 设置样式
        self.setup_styles()
        
        # 创建界面
        self.create_ui()
    
    def setup_styles(self):
        """设置界面样式"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # 配置 Treeview 样式
        style.configure("Custom.Treeview", 
                       rowheight=25,
                       font=('微软雅黑', 10))
        style.configure("Custom.Treeview.Heading",
                       font=('微软雅黑', 10, 'bold'),
                       background='#f0f0f0')
        
        # 配置筛选卡片样式 - 轻量级边框
        style.configure("FilterCard.TFrame",
                       background='white')
        
        # 配置筛选标签样式
        style.configure("FilterLabel.TLabel",
                       font=('微软雅黑', 9, 'bold'),
                       foreground='#333333',
                       background='white')
        
        # 配置筛选区头部样式
        style.configure("FilterHeader.TFrame",
                       background='#f8f9fa')
        
        # 配置折叠按钮样式
        style.configure("Toggle.TButton",
                       font=('微软雅黑', 9),
                       padding=2)
    
    def create_ui(self):
        """创建用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # ===== 顶部工具栏 =====
        toolbar = ttk.Frame(main_frame)
        toolbar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 打开文件按钮
        self.open_btn = ttk.Button(
            toolbar, 
            text="打开 Excel 文件", 
            command=self.callbacks.get('open_file'),
            width=18
        )
        self.open_btn.pack(side=tk.LEFT, padx=(0, 10), ipady=3)
        
        # 文件路径显示
        self.file_label = ttk.Label(toolbar, text="未选择文件", foreground="gray")
        self.file_label.pack(side=tk.LEFT, padx=(0, 20))
        
        # 数据信息
        self.info_label = ttk.Label(toolbar, text="")
        self.info_label.pack(side=tk.LEFT)

        # 工作表选择
        self.sheet_frame = ttk.Frame(toolbar)
        self.sheet_frame.pack(side=tk.LEFT, padx=(20, 0))

        ttk.Label(self.sheet_frame, text="工作表:").pack(side=tk.LEFT)
        self.sheet_combo = ttk.Combobox(
            self.sheet_frame,
            state='readonly',
            width=15
        )
        self.sheet_combo.pack(side=tk.LEFT, padx=(5, 0))
        self.sheet_combo.bind('<<ComboboxSelected>>', self.on_sheet_selected)
        
        # 禁用鼠标滚轮滚动
        self.sheet_combo.bind('<MouseWheel>', lambda e: 'break')
        self.sheet_combo.bind('<Button-4>', lambda e: 'break')  # Linux
        self.sheet_combo.bind('<Button-5>', lambda e: 'break')  # Linux

        # 表头行设置
        self.header_frame = ttk.Frame(toolbar)
        self.header_frame.pack(side=tk.LEFT, padx=(15, 0))

        ttk.Label(self.header_frame, text="表头行:").pack(side=tk.LEFT)
        self.header_var = tk.StringVar(value='1')
        self.header_spin = tk.Spinbox(
            self.header_frame,
            from_=1,
            to=10,
            width=5,
            textvariable=self.header_var
        )
        self.header_spin.pack(side=tk.LEFT, padx=(5, 0))

        # 应用表头按钮
        self.apply_header_btn = ttk.Button(
            self.header_frame,
            text="应用",
            command=self.callbacks.get('reload_current_sheet'),
            width=6
        )
        self.apply_header_btn.pack(side=tk.LEFT, padx=(5, 0))
        
        # 导出按钮框架
        export_frame = ttk.Frame(toolbar)
        export_frame.pack(side=tk.RIGHT)
        
        # 导出筛选结果按钮
        self.export_btn = ttk.Button(
            export_frame,
            text="导出筛选结果",
            command=self.callbacks.get('export_filtered_data'),
            state=tk.DISABLED,
            width=15
        )
        self.export_btn.pack(side=tk.LEFT, ipady=3)
        
        # ===== 筛选区域 =====
        self.filter_container = ttk.Frame(main_frame, style="FilterCard.TFrame")
        self.filter_container.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        self.filter_container.grid_remove()  # 初始隐藏
        
        # 配置筛选容器列权重
        self.filter_container.columnconfigure(0, weight=1)
        
        # 筛选区头部 - 包含标题和折叠按钮
        self.filter_header = tk.Frame(self.filter_container, bg='#f8f9fa', 
                                      highlightbackground='#dee2e6', highlightthickness=1)
        self.filter_header.grid(row=0, column=0, sticky=(tk.W, tk.E))
        self.filter_header.columnconfigure(0, weight=1)
        
        # 筛选区标题
        self.filter_title = tk.Label(
            self.filter_header,
            text="📋 筛选条件",
            font=('微软雅黑', 10, 'bold'),
            bg='#f8f9fa',
            fg='#495057',
            padx=10,
            pady=8
        )
        self.filter_title.grid(row=0, column=0, sticky=tk.W)
        
        # 折叠/展开按钮
        self.toggle_btn = tk.Button(
            self.filter_header,
            text="▼ 收起",
            font=('微软雅黑', 9),
            bg='#f8f9fa',
            fg='#6c757d',
            relief=tk.FLAT,
            cursor='hand2',
            command=self.toggle_filter_panel,
            padx=10,
            pady=5
        )
        self.toggle_btn.grid(row=0, column=1, sticky=tk.E, padx=5)
        
        # 筛选内容区域（可折叠）
        self.filter_content = tk.Frame(self.filter_container, bg='white',
                                       highlightbackground='#dee2e6', 
                                       highlightthickness=1)
        self.filter_content.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.filter_content.columnconfigure(0, weight=1)
        
        # 筛选控件框架
        self.filter_frame = tk.Frame(self.filter_content, bg='white', padx=8, pady=8)
        self.filter_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))

        # 筛选操作按钮框架
        self.filter_btn_frame = tk.Frame(self.filter_content, bg='white', padx=10, pady=10)
        self.filter_btn_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))

        # 搜索按钮 - 使用更醒目的样式
        self.search_btn = tk.Button(
            self.filter_btn_frame,
            text="🔍 搜索",
            font=('微软雅黑', 9),
            bg='#0d6efd',
            fg='white',
            activebackground='#0b5ed7',
            activeforeground='white',
            relief=tk.FLAT,
            cursor='hand2',
            command=self.callbacks.get('apply_filters'),
            padx=15,
            pady=5
        )
        self.search_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 重置按钮
        self.reset_btn = tk.Button(
            self.filter_btn_frame,
            text="↺ 重置",
            font=('微软雅黑', 9),
            bg='#6c757d',
            fg='white',
            activebackground='#5c636a',
            activeforeground='white',
            relief=tk.FLAT,
            cursor='hand2',
            command=self.callbacks.get('reset_filters'),
            padx=15,
            pady=5
        )
        self.reset_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # ===== 数据表格区域 =====
        table_frame = ttk.Frame(main_frame)
        table_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        
        # 创建 Treeview
        self.tree = ttk.Treeview(
            table_frame,
            style="Custom.Treeview",
            show='headings',
            selectmode='extended'
        )
        
        # 滚动条
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # 放置组件
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 状态栏
        self.status_label = ttk.Label(
            main_frame, 
            text="就绪 - 请打开 Excel 文件",
            anchor=tk.W,
            relief=tk.SUNKEN,
            padding=(5, 2)
        )
        self.status_label.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
    
    def update_file_info(self, file_name):
        """更新文件信息显示"""
        self.file_label.config(text=file_name, foreground="black")
    
    def update_data_info(self, rows, columns):
        """更新数据信息显示"""
        self.info_label.config(text=f"共 {rows} 行 × {columns} 列")
    
    def update_sheet_combo(self, sheet_names):
        """更新工作表下拉框"""
        self.sheet_names = sheet_names
        self.sheet_combo['values'] = sheet_names
        if sheet_names:
            self.sheet_combo.set(sheet_names[0])
    
    def get_selected_sheet(self):
        """获取当前选择的工作表"""
        return self.sheet_combo.get()
    
    def get_header_row(self):
        """获取表头行设置"""
        try:
            return int(self.header_var.get()) - 1  # 转换为 0-based
        except:
            return 0
    
    def on_sheet_selected(self, event=None):
        """工作表选择事件"""
        selected_sheet = self.sheet_combo.get()
        if selected_sheet and 'on_sheet_selected' in self.callbacks:
            self.callbacks['on_sheet_selected'](selected_sheet)
    
    def ask_header_row(self, sheet_name):
        """弹窗询问用户表头行设置"""
        dialog = tk.Toplevel(self.root)
        dialog.title("设置表头行")
        dialog.geometry("350x180")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (350 // 2)
        y = (dialog.winfo_screenheight() // 2) - (180 // 2)
        dialog.geometry(f"+{x}+{y}")

        # 提示文本
        tk.Label(
            dialog,
            text=f"工作表: {sheet_name}",
            font=('微软雅黑', 10, 'bold')
        ).place(x=20, y=15)

        tk.Label(
            dialog,
            text="请选择哪一行作为列名（表头）:",
            font=('微软雅黑', 10)
        ).place(x=20, y=45)

        # 表头行选择 - 使用绝对定位
        tk.Label(dialog, text="第", font=('微软雅黑', 10)).place(x=80, y=75)

        header_var = tk.StringVar(value='1')
        spin = tk.Spinbox(
            dialog,
            from_=1,
            to=10,
            width=6,
            textvariable=header_var,
            font=('微软雅黑', 10),
            justify=tk.CENTER
        )
        spin.place(x=100, y=75)

        tk.Label(dialog, text="行", font=('微软雅黑', 10)).place(x=165, y=75)

        result = [1]  # 使用列表存储结果

        def on_ok():
            try:
                result[0] = int(header_var.get())
            except:
                result[0] = 1
            dialog.destroy()

        def on_cancel():
            result[0] = None
            dialog.destroy()

        # 按钮 - 使用绝对定位，固定大小
        ok_btn = tk.Button(
            dialog,
            text="确定",
            command=on_ok,
            font=('微软雅黑', 10),
            width=8,
            height=1
        )
        ok_btn.place(x=70, y=120)

        cancel_btn = tk.Button(
            dialog,
            text="取消",
            command=on_cancel,
            font=('微软雅黑', 10),
            width=8,
            height=1
        )
        cancel_btn.place(x=190, y=120)

        # 等待对话框关闭
        self.root.wait_window(dialog)
        return result[0]
    
    def create_filter_widgets(self, columns, get_unique_values_callback):
        """创建筛选控件"""
        # 清除旧的筛选控件
        for widget in self.filter_frame.winfo_children():
            widget.destroy()
        self.filter_widgets = {}

        if not columns:
            return

        self.filter_container.grid()

        # 每行显示的列数 - 改为5列
        cols_per_row = 5

        # 配置grid权重，使各列均匀分布
        for c in range(cols_per_row):
            self.filter_frame.columnconfigure(c, weight=1)

        # 为每列创建筛选控件
        for col_idx, col_name in enumerate(columns):
            row = col_idx // cols_per_row
            col = col_idx % cols_per_row

            # 列框架 - 使用轻量级卡片设计，无边框
            col_frame = tk.Frame(
                self.filter_frame,
                bg='white',
                highlightthickness=0,
                padx=8,
                pady=6
            )
            col_frame.grid(row=row, column=col, padx=4, pady=3, sticky=(tk.W, tk.E))

            # 列名标签 - 更紧凑的设计
            col_label = tk.Label(
                col_frame,
                text=str(col_name)[:18],  # 列名显示限制
                font=('微软雅黑', 9, 'bold'),
                bg='white',
                fg='#495057',
                anchor='w'
            )
            col_label.pack(fill=tk.X, pady=(0, 4))

            # 下拉选择框 - 可编辑，支持输入搜索和下拉选择
            all_unique_values = get_unique_values_callback(col_name)

            combo = ttk.Combobox(
                col_frame,
                values=all_unique_values,
                width=16,
                state='normal',  # 可编辑状态
                font=('微软雅黑', 9)
            )
            combo.set('')  # 默认为空
            combo.pack(fill=tk.X, pady=(0, 4))

            # 禁用鼠标滚轮滚动
            combo.bind('<MouseWheel>', lambda e: 'break')
            combo.bind('<Button-4>', lambda e: 'break')  # Linux
            combo.bind('<Button-5>', lambda e: 'break')  # Linux

            # 绑定事件：输入时过滤下拉选项
            combo.bind('<KeyRelease>', lambda e, c=col_name, cb=combo: self.on_combo_key_release(c, cb))
            # 绑定事件：选择时应用筛选
            combo.bind('<<ComboboxSelected>>', lambda e, c=col_name: self.on_combo_selected(c))

            # 保存控件引用
            self.filter_widgets[col_name] = {
                'frame': col_frame,
                'combo': combo,
                'all_values': all_unique_values  # 保存所有原始值用于过滤
            }

        # 确保筛选区域可见
        self.filter_frame.update_idletasks()
    
    def on_combo_key_release(self, col_name, combo):
        """下拉框输入时 - 动态过滤选项"""
        widgets = self.filter_widgets.get(col_name)
        if not widgets:
            return

        all_values = widgets.get('all_values', [])
        current_text = combo.get()

        # 过滤匹配的值
        if current_text.strip():
            filtered_values = [v for v in all_values if current_text.lower() in v.lower()]
            combo['values'] = filtered_values
        else:
            # 输入为空时，恢复所有选项
            combo['values'] = all_values

    def on_combo_selected(self, col_name):
        """下拉框被选择时 - 自动应用筛选"""
        if 'apply_filters' in self.callbacks:
            self.callbacks['apply_filters']()
    
    def get_filter_criteria(self):
        """获取筛选条件"""
        criteria = {}
        for col_name, widgets in self.filter_widgets.items():
            combo_value = widgets['combo'].get().strip()
            if combo_value:
                criteria[col_name] = combo_value
        return criteria
    
    def reset_filter_widgets(self):
        """重置筛选控件"""
        for col_name, widgets in self.filter_widgets.items():
            all_values = widgets.get('all_values', [])

            # 重置下拉框 - 恢复所有选项并清空输入
            widgets['combo'].config(state='normal')
            widgets['combo']['values'] = all_values
            widgets['combo'].set('')
    
    def display_data(self, columns, data):
        """在表格中显示数据"""
        # 清除现有数据
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 清除列
        self.tree['columns'] = ()

        if not data:
            return

        # 设置列
        self.tree['columns'] = columns

        # 获取表格显示区域的宽度
        self.tree.update_idletasks()
        tree_width = self.tree.winfo_width()
        if tree_width < 100:  # 如果还未渲染，使用默认值
            tree_width = 800

        # 计算每列的权重（基于列名长度）
        col_weights = []
        for col in columns:
            col_name_width = calc_text_width(str(col))
            col_weights.append(max(col_name_width, 5))  # 最小权重为5

        total_weight = sum(col_weights)
        num_cols = len(columns)

        # 配置列 - 自适应宽度，刚好占满显示区域
        for i, col in enumerate(columns):
            self.tree.heading(col, text=col, anchor=tk.W)
            # 计算该列应占的比例
            weight_ratio = col_weights[i] / total_weight if total_weight > 0 else 1 / num_cols
            # 计算列宽
            col_width = int(tree_width * weight_ratio)
            # 最小宽度确保能显示列名
            min_width = int(col_weights[i] * 9 + 20)
            col_width = max(col_width, min_width)
            # 最后一列stretch=True以填充剩余空间
            is_last = (i == num_cols - 1)
            self.tree.column(col, width=col_width, minwidth=min_width, anchor=tk.W, stretch=is_last)
        
        # 插入数据
        for row_idx, row in enumerate(data):
            # 使用行号作为 iid，避免索引类型问题
            self.tree.insert('', tk.END, iid=str(row_idx), values=row)
    
    def show_export_dialog(self, default_filename):
        """显示导出格式选择对话框"""
        # 创建导出格式选择对话框
        export_dialog = tk.Toplevel(self.root)
        export_dialog.title("导出筛选结果")
        export_dialog.geometry("420x240")
        export_dialog.transient(self.root)
        export_dialog.grab_set()
        export_dialog.resizable(False, False)
        
        # 居中显示
        export_dialog.update_idletasks()
        x = (export_dialog.winfo_screenwidth() // 2) - (400 // 2)
        y = (export_dialog.winfo_screenheight() // 2) - (200 // 2)
        export_dialog.geometry(f"+{x}+{y}")
        
        # 内容框架
        content_frame = tk.Frame(export_dialog)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # 提示文本
        tk.Label(
            content_frame,
            text="请选择导出格式",
            font=('微软雅黑', 12, 'bold')
        ).pack(pady=(0, 15))
        
        # 格式选择
        format_var = tk.StringVar(value='excel')
        
        excel_radio = tk.Radiobutton(
            content_frame,
            text="Excel 文件 (.xlsx)",
            variable=format_var,
            value='excel',
            font=('微软雅黑', 10),
            anchor=tk.W
        )
        excel_radio.pack(fill=tk.X, pady=5)
        
        md_radio = tk.Radiobutton(
            content_frame,
            text="Markdown 文档 (.md)",
            variable=format_var,
            value='markdown',
            font=('微软雅黑', 10),
            anchor=tk.W
        )
        md_radio.pack(fill=tk.X, pady=5)
        
        # 文件名显示
        tk.Label(
            content_frame,
            text=f"默认文件名：{default_filename}",
            font=('微软雅黑', 8),
            fg='gray',
            wraplength=360,
            justify=tk.LEFT
        ).pack(pady=(10, 0))
        
        result = {'format': None, 'confirmed': False}
        
        def on_export():
            result['format'] = format_var.get()
            result['confirmed'] = True
            export_dialog.destroy()
        
        def on_cancel():
            result['confirmed'] = False
            export_dialog.destroy()
        
        # 按钮框架
        btn_frame = tk.Frame(export_dialog)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=20, pady=15)
        
        tk.Button(
            btn_frame,
            text="导出",
            command=on_export,
            font=('微软雅黑', 10),
            width=12,
            bg='#0d6efd',
            fg='white',
            activebackground='#0b5ed7',
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        tk.Button(
            btn_frame,
            text="取消",
            command=on_cancel,
            font=('微软雅黑', 10),
            width=12,
            cursor='hand2'
        ).pack(side=tk.RIGHT)
        
        # 等待对话框关闭
        self.root.wait_window(export_dialog)
        
        return result
    
    def ask_save_file(self, defaultextension, filetypes, initialfile):
        """询问保存文件路径"""
        return filedialog.asksaveasfilename(
            title="导出筛选结果",
            defaultextension=defaultextension,
            filetypes=filetypes,
            initialfile=initialfile
        )
    
    def update_status(self, message):
        """更新状态栏"""
        self.status_label.config(text=message)
    
    def toggle_filter_panel(self):
        """折叠/展开筛选区"""
        if self.filter_content.winfo_viewable():
            # 当前可见，执行折叠
            self.filter_content.grid_remove()
            self.toggle_btn.config(text="▶ 展开")
            self.filter_title.config(text=f"📋 筛选条件 ({len(self.filter_widgets)}个字段)")
        else:
            # 当前隐藏，执行展开
            self.filter_content.grid()
            self.toggle_btn.config(text="▼ 收起")
            self.filter_title.config(text="📋 筛选条件")
    
    def enable_export_button(self, enable):
        """启用或禁用导出按钮"""
        self.export_btn.config(state=tk.NORMAL if enable else tk.DISABLED)
    
    def show_error(self, title, message):
        """显示错误消息"""
        messagebox.showerror(title, message)
    
    def show_info(self, title, message):
        """显示信息消息"""
        messagebox.showinfo(title, message)
    
    def show_warning(self, title, message):
        """显示警告消息"""
        messagebox.showwarning(title, message)
