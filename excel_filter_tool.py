#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 筛选处理工具 - GUI 应用程序
基于 Tkinter 开发，支持 Excel 文件导入、多维度筛选和多格式导出
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from tkinter.scrolledtext import ScrolledText


class ExcelFilterTool:
    """Excel 筛选处理工具主类"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 筛选处理工具")
        # 设置默认全屏显示
        self.root.state('zoomed')
        self.root.minsize(1000, 600)
        
        # 数据存储
        self.df = None  # 原始数据
        self.filtered_df = None  # 筛选后的数据
        self.columns = []  # 列名列表
        self.filter_widgets = {}  # 筛选控件字典
        self.excel_file_path = None  # Excel 文件路径
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
            command=self.open_file,
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
            command=self.reload_current_sheet,
            width=6
        )
        self.apply_header_btn.pack(side=tk.LEFT, padx=(5, 0))
        
        # 导出按钮框架
        export_frame = ttk.Frame(toolbar)
        export_frame.pack(side=tk.RIGHT)
        
        # 导出全部按钮
        self.export_all_btn = ttk.Button(
            export_frame,
            text="导出全部结果",
            command=lambda: self.export_data(export_selected=False),
            state=tk.DISABLED,
            width=15
        )
        self.export_all_btn.pack(side=tk.LEFT, padx=(0, 5), ipady=3)

        # 导出选中按钮
        self.export_sel_btn = ttk.Button(
            export_frame,
            text="导出选中行",
            command=lambda: self.export_data(export_selected=True),
            state=tk.DISABLED,
            width=15
        )
        self.export_sel_btn.pack(side=tk.LEFT, ipady=3)

        # 清除筛选按钮
        self.clear_btn = ttk.Button(
            toolbar,
            text="清除筛选",
            command=self.clear_filters,
            state=tk.DISABLED,
            width=12
        )
        self.clear_btn.pack(side=tk.RIGHT, padx=(0, 10), ipady=3)
        
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
            command=self.apply_filters,
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
            command=self.reset_filters,
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
        
    def open_file(self):
        """打开 Excel 文件"""
        file_path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xls"),
                ("Excel 2007+", "*.xlsx"),
                ("Excel 97-2003", "*.xls"),
                ("所有文件", "*.*")
            ]
        )

        if not file_path:
            return

        self.excel_file_path = file_path

        try:
            # 获取所有工作表名称
            xl = pd.ExcelFile(file_path)
            self.sheet_names = xl.sheet_names

            # 更新工作表选择下拉框
            self.sheet_combo['values'] = self.sheet_names
            if self.sheet_names:
                self.sheet_combo.set(self.sheet_names[0])
                # 加载第一个工作表
                self.load_sheet(self.sheet_names[0])

            # 更新界面
            self.file_label.config(text=file_path.split('/')[-1].split('\\')[-1], foreground="black")

        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            messagebox.showerror("错误", f"无法读取文件：\n{str(e)}\n\n详细信息：\n{error_detail}")

    def on_sheet_selected(self, event=None):
        """工作表选择事件"""
        selected_sheet = self.sheet_combo.get()
        if selected_sheet:
            # 切换工作表时询问表头行
            self.load_sheet(selected_sheet, ask_header=True)

    def reload_current_sheet(self):
        """重新加载当前工作表（用于表头行变更）"""
        current_sheet = self.sheet_combo.get()
        if current_sheet:
            # 点击应用按钮时不询问，直接使用当前设置的表头行
            self.load_sheet(current_sheet, ask_header=False)

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

    def load_sheet(self, sheet_name, ask_header=True):
        """加载指定工作表"""
        try:
            # 如果需要询问表头行
            if ask_header:
                header_row_input = self.ask_header_row(sheet_name)
                if header_row_input is None:
                    return  # 用户取消
                self.header_var.set(str(header_row_input))

            # 获取表头行设置（用户输入是1-based，pandas需要0-based）
            try:
                header_row = int(self.header_var.get()) - 1
            except:
                header_row = 0

            # 读取指定工作表，指定表头行
            self.df = pd.read_excel(
                self.excel_file_path,
                sheet_name=sheet_name,
                header=header_row
            )

            # 重置索引为整数，避免浮点数索引问题
            self.df = self.df.reset_index(drop=True)

            self.filtered_df = self.df.copy()
            self.columns = list(self.df.columns)

            # 更新界面
            self.info_label.config(text=f"共 {len(self.df)} 行 × {len(self.columns)} 列")

            # 创建筛选控件
            self.create_filter_widgets()

            # 显示数据
            self.display_data()

            # 启用按钮
            self.export_all_btn.config(state=tk.NORMAL)
            self.export_sel_btn.config(state=tk.NORMAL)
            self.clear_btn.config(state=tk.NORMAL)

            header_info = f"第 {header_row + 1} 行作为表头" if header_row >= 0 else "无表头"
            self.update_status(f"已加载工作表 '{sheet_name}'，{header_info}，共 {len(self.df)} 行数据")

        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            messagebox.showerror("错误", f"无法加载工作表：\n{str(e)}\n\n详细信息：\n{error_detail}")
            
    def create_filter_widgets(self):
        """创建筛选控件 - 每行5列，每列有下拉框和搜索框（互斥使用）"""
        # 清除旧的筛选控件
        for widget in self.filter_frame.winfo_children():
            widget.destroy()
        self.filter_widgets = {}

        if not self.columns:
            return

        self.filter_container.grid()

        # 每行显示的列数 - 改为5列
        cols_per_row = 5

        # 配置grid权重，使各列均匀分布
        for c in range(cols_per_row):
            self.filter_frame.columnconfigure(c, weight=1)

        # 为每列创建筛选控件
        for col_idx, col_name in enumerate(self.columns):
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
            try:
                unique_vals = self.df[col_name].dropna().astype(str).unique()
                all_unique_values = sorted(unique_vals, key=str)
            except:
                all_unique_values = []

            combo = ttk.Combobox(
                col_frame,
                values=all_unique_values,
                width=16,
                state='normal',  # 可编辑状态
                font=('微软雅黑', 9)
            )
            combo.set('')  # 默认为空
            combo.pack(fill=tk.X, pady=(0, 4))

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
        # 选择后自动应用筛选
        self.apply_filters()
            
    def apply_filters(self):
        """应用筛选条件 - 点击搜索按钮时调用"""
        if self.df is None:
            return

        # 从原始数据开始筛选
        mask = pd.Series([True] * len(self.df), index=self.df.index)
        active_filters = 0

        for col_name, widgets in self.filter_widgets.items():
            # 检查下拉框（如果有输入值）
            combo_value = widgets['combo'].get().strip()
            if combo_value:
                mask &= self.df[col_name].astype(str) == combo_value
                active_filters += 1

        # 应用筛选
        self.filtered_df = self.df[mask].copy()

        # 刷新显示
        self.display_data()

        # 更新状态
        filter_info = f"（{active_filters}个筛选条件）" if active_filters > 0 else ""
        self.update_status(f"筛选结果：{len(self.filtered_df)} / {len(self.df)} 行{filter_info}")

    def reset_filters(self):
        """重置所有筛选条件到初始状态"""
        if not self.filter_widgets:
            return

        for col_name, widgets in self.filter_widgets.items():
            all_values = widgets.get('all_values', [])

            # 重置下拉框 - 恢复所有选项并清空输入
            widgets['combo'].config(state='normal')
            widgets['combo']['values'] = all_values
            widgets['combo'].set('')

        # 重置数据为全部
        self.filtered_df = self.df.copy()
        self.display_data()
        self.update_status("已重置所有筛选条件，显示全部数据")

    def clear_filters(self):
        """清除所有筛选条件（兼容旧方法，调用reset_filters）"""
        self.reset_filters()
        
    def display_data(self):
        """在表格中显示数据"""
        # 清除现有数据
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 清除列
        self.tree['columns'] = ()

        if self.filtered_df is None or self.filtered_df.empty:
            return

        # 设置列
        self.tree['columns'] = self.columns

        # 获取表格显示区域的宽度
        self.tree.update_idletasks()
        tree_width = self.tree.winfo_width()
        if tree_width < 100:  # 如果还未渲染，使用默认值
            tree_width = 800

        # 计算每列的权重（基于列名长度）
        col_weights = []
        for col in self.columns:
            col_name_width = self._calc_text_width(str(col))
            col_weights.append(max(col_name_width, 5))  # 最小权重为5

        total_weight = sum(col_weights)
        num_cols = len(self.columns)

        # 配置列 - 自适应宽度，刚好占满显示区域
        for i, col in enumerate(self.columns):
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
        
        # 插入数据（分批加载以提高性能）
        batch_size = 100
        data_to_show = self.filtered_df.head(1000)  # 最多显示1000行
        
        for row_idx, (idx, row) in enumerate(data_to_show.iterrows()):
            values = [str(v) if pd.notna(v) else '' for v in row.values]
            # 使用行号作为 iid，避免索引类型问题
            self.tree.insert('', tk.END, iid=str(row_idx), values=values)
            
        if len(self.filtered_df) > 1000:
            self.update_status(f"显示前 1000 行（共 {len(self.filtered_df)} 行）")
            
    def export_data(self, export_selected=False):
        """导出数据"""
        if self.filtered_df is None or self.filtered_df.empty:
            messagebox.showwarning("警告", "没有可导出的数据")
            return
            
        # 确定要导出的数据
        if export_selected:
            selected_items = self.tree.selection()
            if not selected_items:
                messagebox.showwarning("警告", "请先选择要导出的行")
                return
            # 获取选中行的数据（通过行号）
            selected_row_indices = [int(iid) for iid in selected_items]
            data_to_show = self.filtered_df.head(1000)
            selected_data = []
            for row_idx, (idx, row) in enumerate(data_to_show.iterrows()):
                if row_idx in selected_row_indices:
                    selected_data.append(row)
            export_df = pd.DataFrame(selected_data).copy()
        else:
            export_df = self.filtered_df.copy()
            
        # 选择导出格式
        file_types = [
            ("Excel 文件", "*.xlsx"),
            ("Markdown 文件", "*.md"),
            ("所有文件", "*.*")
        ]
        
        file_path = filedialog.asksaveasfilename(
            title="导出数据",
            defaultextension=".xlsx",
            filetypes=file_types
        )
        
        if not file_path:
            return
            
        try:
            if file_path.endswith('.md'):
                # 导出为 Markdown
                self.export_to_markdown(export_df, file_path)
            else:
                # 导出为 Excel
                export_df.to_excel(file_path, index=False, engine='openpyxl')
                
            messagebox.showinfo(
                "成功", 
                f"已成功导出 {len(export_df)} 行数据到：\n{file_path}"
            )
            self.update_status(f"已导出 {len(export_df)} 行数据")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：\n{str(e)}")
            
    def export_to_markdown(self, df, file_path):
        """导出为 Markdown 表格格式"""
        lines = []
        
        # 表头
        headers = [str(col) for col in df.columns]
        lines.append('| ' + ' | '.join(headers) + ' |')
        lines.append('|' + '|'.join(['---' for _ in headers]) + '|')
        
        # 数据行
        for _, row in df.iterrows():
            values = [str(v) if pd.notna(v) else '' for v in row.values]
            # 处理包含 | 的单元格
            values = [v.replace('|', '\\|') for v in values]
            lines.append('| ' + ' | '.join(values) + ' |')
        
        # 写入文件
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))
            
    def update_status(self, message):
        """更新状态栏"""
        self.status_label.config(text=message)

    def _calc_text_width(self, text):
        """计算文本宽度，中文字符按2个单位宽度计算"""
        width = 0
        for char in str(text):
            if '\u4e00' <= char <= '\u9fff':  # 中文字符范围
                width += 2
            elif char.isupper():
                width += 1.2  # 大写字母稍宽
            else:
                width += 1
        return width

    def toggle_filter_panel(self):
        """折叠/展开筛选区"""
        if self.filter_content.winfo_viewable():
            # 当前可见，执行折叠
            self.filter_content.grid_remove()
            self.toggle_btn.config(text="▶ 展开")
            self.filter_title.config(text=f"📋 筛选条件 ({len(self.columns)}个字段)")
        else:
            # 当前隐藏，执行展开
            self.filter_content.grid()
            self.toggle_btn.config(text="▼ 收起")
            self.filter_title.config(text="📋 筛选条件")


def main():
    """主函数"""
    root = tk.Tk()
    
    # 设置 DPI 感知（Windows）
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    
    app = ExcelFilterTool(root)
    root.mainloop()


if __name__ == "__main__":
    main()
