#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主应用模块
负责应用的初始化和运行，整合数据处理模块和 UI 模块
"""

import tkinter as tk
from tkinter import filedialog
import traceback
from .data_handler import ExcelDataHandler
from .ui_components import ExcelFilterUI
from .utils import get_file_name_from_path, get_default_export_filename


class ExcelFilterApp:
    """Excel 筛选工具应用类"""
    
    def __init__(self, root):
        self.root = root
        self.data_handler = ExcelDataHandler()
        
        # 注册回调函数
        callbacks = {
            'open_file': self.open_file,
            'on_sheet_selected': self.on_sheet_selected,
            'reload_current_sheet': self.reload_current_sheet,
            'apply_filters': self.apply_filters,
            'reset_filters': self.reset_filters,
            'export_filtered_data': self.export_filtered_data
        }
        
        self.ui = ExcelFilterUI(root, callbacks)
    
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

        try:
            # 打开文件并获取工作表名称
            self.data_handler.open_file(file_path)
            
            # 更新界面
            file_name = get_file_name_from_path(file_path)
            self.ui.update_file_info(file_name)
            
            # 更新工作表下拉框
            self.ui.update_sheet_combo(self.data_handler.sheet_names)
            
            # 加载第一个工作表
            if self.data_handler.sheet_names:
                self.load_sheet(self.data_handler.sheet_names[0], ask_header=True)

        except Exception as e:
            error_detail = traceback.format_exc()
            self.ui.show_error("错误", f"无法读取文件：\n{str(e)}\n\n详细信息：\n{error_detail}")
    
    def on_sheet_selected(self, sheet_name):
        """工作表选择事件"""
        if sheet_name:
            # 切换工作表时询问表头行
            self.load_sheet(sheet_name, ask_header=True)
    
    def reload_current_sheet(self):
        """重新加载当前工作表（用于表头行变更）"""
        current_sheet = self.ui.get_selected_sheet()
        if current_sheet:
            # 点击应用按钮时不询问，直接使用当前设置的表头行
            self.load_sheet(current_sheet, ask_header=False)
    
    def load_sheet(self, sheet_name, ask_header=True):
        """加载指定工作表"""
        try:
            # 如果需要询问表头行
            if ask_header:
                header_row_input = self.ui.ask_header_row(sheet_name)
                if header_row_input is None:
                    return  # 用户取消
                header_row = header_row_input - 1  # 转换为 0-based
            else:
                header_row = self.ui.get_header_row()

            # 加载工作表
            self.data_handler.load_sheet(sheet_name, header_row)

            # 更新界面
            data_info = self.data_handler.get_data_info()
            self.ui.update_data_info(data_info["rows"], data_info["columns"])

            # 创建筛选控件
            self.ui.create_filter_widgets(
                self.data_handler.columns,
                self.data_handler.get_unique_values
            )

            # 显示数据
            self.display_data()

            # 启用按钮
            self.ui.enable_export_button(True)

            header_info = f"第 {header_row + 1} 行作为表头" if header_row >= 0 else "无表头"
            self.ui.update_status(f"已加载工作表 '{sheet_name}'，{header_info}，共 {data_info['rows']} 行数据")

        except Exception as e:
            error_detail = traceback.format_exc()
            self.ui.show_error("错误", f"无法加载工作表：\n{str(e)}\n\n详细信息：\n{error_detail}")
    
    def apply_filters(self):
        """应用筛选条件"""
        if self.data_handler.df is None:
            return

        # 获取筛选条件
        filter_criteria = self.ui.get_filter_criteria()
        
        # 应用筛选
        filtered_count = self.data_handler.apply_filters(filter_criteria)
        
        # 显示筛选后的数据
        self.display_data()
        
        # 更新状态
        active_filters = len(filter_criteria)
        filter_info = f"（{active_filters}个筛选条件）" if active_filters > 0 else ""
        total_count = len(self.data_handler.df)
        self.ui.update_status(f"筛选结果：{filtered_count} / {total_count} 行{filter_info}")
    
    def reset_filters(self):
        """重置所有筛选条件"""
        if not self.data_handler.df:
            return

        # 重置数据
        self.data_handler.reset_filters()
        
        # 重置筛选控件
        self.ui.reset_filter_widgets()
        
        # 显示全部数据
        self.display_data()
        
        # 更新状态
        self.ui.update_status("已重置所有筛选条件，显示全部数据")
    
    def display_data(self):
        """显示数据"""
        if self.data_handler.filtered_df is None or self.data_handler.filtered_df.empty:
            self.ui.display_data([], [])
            return
        
        # 准备数据
        columns = self.data_handler.columns
        data_to_show = self.data_handler.filtered_df.head(1000)  # 最多显示1000行
        
        # 转换数据格式
        data = []
        for _, row in data_to_show.iterrows():
            values = [str(v) if v is not None else '' for v in row.values]
            data.append(values)
        
        # 显示数据
        self.ui.display_data(columns, data)
        
        # 更新状态
        if len(self.data_handler.filtered_df) > 1000:
            self.ui.update_status(f"显示前 1000 行（共 {len(self.data_handler.filtered_df)} 行）")
    
    def export_filtered_data(self):
        """导出筛选结果"""
        if self.data_handler.filtered_df is None or self.data_handler.filtered_df.empty:
            self.ui.show_warning("警告", "没有可导出的数据")
            return
        
        # 获取默认文件名
        original_file_name = get_file_name_from_path(self.data_handler.excel_file_path)
        default_filename = get_default_export_filename(original_file_name)
        
        # 显示导出格式选择对话框
        result = self.ui.show_export_dialog(default_filename)
        
        if not result['confirmed']:
            return
        
        # 根据选择的格式设置默认扩展名
        if result['format'] == 'markdown':
            defaultextension = ".md"
            filetypes = [("Markdown 文件", "*.md"), ("所有文件", "*.*")]
        else:
            defaultextension = ".xlsx"
            filetypes = [("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        
        # 默认文件名
        default_name = f"{default_filename}{defaultextension}"
        
        # 询问保存路径
        file_path = self.ui.ask_save_file(defaultextension, filetypes, default_name)
        
        if not file_path:
            return
            
        try:
            if result['format'] == 'markdown':
                # 导出为 Markdown
                self.data_handler.export_to_markdown(file_path)
            else:
                # 导出为 Excel
                self.data_handler.export_to_excel(file_path)
                
            self.ui.show_info(
                "成功", 
                f"已成功导出 {len(self.data_handler.filtered_df)} 行数据到：\n{file_path}"
            )
            self.ui.update_status(f"已导出 {len(self.data_handler.filtered_df)} 行数据")
            
        except Exception as e:
            self.ui.show_error("错误", f"导出失败：\n{str(e)}")


def main():
    """主函数"""
    root = tk.Tk()
    
    # 设置 DPI 感知（Windows）
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    
    app = ExcelFilterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
