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
        self.root.geometry("1200x800")
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
            width=20
        )
        self.sheet_combo.pack(side=tk.LEFT, padx=(5, 0))
        self.sheet_combo.bind('<<ComboboxSelected>>', self.on_sheet_selected)
        
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
        self.filter_container = ttk.LabelFrame(main_frame, text="筛选条件", padding="5")
        self.filter_container.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        self.filter_container.grid_remove()  # 初始隐藏
        
        # 创建 Canvas 用于水平滚动
        self.filter_canvas = tk.Canvas(self.filter_container, height=100, highlightthickness=0)
        self.filter_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 水平滚动条
        filter_hsb = ttk.Scrollbar(self.filter_container, orient="horizontal", command=self.filter_canvas.xview)
        filter_hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.filter_canvas.configure(xscrollcommand=filter_hsb.set)
        
        # 筛选框架放在 Canvas 内
        self.filter_frame = ttk.Frame(self.filter_canvas)
        self.filter_canvas_window = self.filter_canvas.create_window((0, 0), window=self.filter_frame, anchor=tk.NW)
        
        # 绑定事件更新滚动区域
        self.filter_frame.bind("<Configure>", self.on_filter_frame_configure)
        self.filter_canvas.bind("<Configure>", self.on_filter_canvas_configure)
        
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
            self.load_sheet(selected_sheet)

    def load_sheet(self, sheet_name):
        """加载指定工作表"""
        try:
            # 读取指定工作表
            self.df = pd.read_excel(self.excel_file_path, sheet_name=sheet_name)

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

            self.update_status(f"已加载工作表 '{sheet_name}'，共 {len(self.df)} 行数据")

        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            messagebox.showerror("错误", f"无法加载工作表：\n{str(e)}\n\n详细信息：\n{error_detail}")
            
    def on_filter_frame_configure(self, event=None):
        """更新 Canvas 滚动区域"""
        self.filter_canvas.configure(scrollregion=self.filter_canvas.bbox("all"))

    def on_filter_canvas_configure(self, event=None):
        """更新 Canvas 窗口大小"""
        canvas_width = event.width
        self.filter_canvas.itemconfig(self.filter_canvas_window, width=canvas_width)

    def create_filter_widgets(self):
        """创建筛选控件"""
        # 清除旧的筛选控件
        for widget in self.filter_frame.winfo_children():
            widget.destroy()
        self.filter_widgets = {}

        if not self.columns:
            return

        self.filter_container.grid()
        
        # 为每列创建筛选控件
        for col_idx, col_name in enumerate(self.columns):
            # 列框架
            col_frame = ttk.Frame(self.filter_frame)
            col_frame.pack(side=tk.LEFT, padx=3, pady=2, fill=tk.Y)

            # 列标题
            display_name = str(col_name)[:12] + ('..' if len(str(col_name)) > 12 else '')
            ttk.Label(
                col_frame,
                text=display_name,
                font=('微软雅黑', 9, 'bold'),
                width=14,
                anchor=tk.CENTER
            ).pack(anchor=tk.N, pady=(0, 2))

            # 下拉选择框
            try:
                unique_vals = self.df[col_name].dropna().astype(str).unique()
                unique_values = ['全部'] + sorted(unique_vals, key=str)[:30]
            except:
                unique_values = ['全部']

            combo = ttk.Combobox(
                col_frame,
                values=unique_values,
                width=13,
                state='readonly',
                font=('微软雅黑', 9)
            )
            combo.set('全部')
            combo.pack(anchor=tk.N, pady=(0, 2))
            combo.bind('<<ComboboxSelected>>', lambda e: self.apply_filters())

            # 关键字搜索框
            entry = ttk.Entry(col_frame, width=13, font=('微软雅黑', 9))
            entry.pack(anchor=tk.N)
            entry.bind('<KeyRelease>', lambda e: self.apply_filters())

            # 保存控件引用
            self.filter_widgets[col_name] = {
                'combo': combo,
                'entry': entry
            }
            
    def apply_filters(self):
        """应用筛选条件"""
        if self.df is None:
            return
            
        # 从原始数据开始筛选
        mask = pd.Series([True] * len(self.df), index=self.df.index)
        
        for col_name, widgets in self.filter_widgets.items():
            combo_value = widgets['combo'].get()
            entry_value = widgets['entry'].get().strip()
            
            # 下拉筛选
            if combo_value != '全部':
                mask &= self.df[col_name].astype(str) == combo_value
                
            # 关键字筛选
            if entry_value:
                mask &= self.df[col_name].astype(str).str.contains(
                    entry_value, 
                    case=False, 
                    na=False
                )
        
        # 应用筛选
        self.filtered_df = self.df[mask].copy()
        
        # 刷新显示
        self.display_data()
        
        # 更新状态
        self.update_status(f"筛选结果：{len(self.filtered_df)} / {len(self.df)} 行")
        
    def clear_filters(self):
        """清除所有筛选条件"""
        if not self.filter_widgets:
            return
            
        for widgets in self.filter_widgets.values():
            widgets['combo'].set('全部')
            widgets['entry'].delete(0, tk.END)
            
        self.apply_filters()
        self.update_status("已清除所有筛选条件")
        
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
        
        # 配置列
        for col in self.columns:
            self.tree.heading(col, text=col, anchor=tk.W)
            # 根据内容长度设置列宽
            max_len = max(
                len(str(col)),
                int(self.filtered_df[col].astype(str).str.len().max()) if len(self.filtered_df) > 0 else 0
            )
            width = int(min(max(max_len * 10, 80), 300))
            self.tree.column(col, width=width, anchor=tk.W)
        
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
