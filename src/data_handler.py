#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 数据处理模块
负责 Excel 文件的读取、筛选和导出
"""

import pandas as pd
import re


class ExcelDataHandler:
    """Excel 数据处理类"""
    
    def __init__(self):
        self.df = None  # 原始数据
        self.filtered_df = None  # 筛选后的数据
        self.columns = []  # 列名列表
        self.excel_file_path = None  # Excel 文件路径
        self.sheet_names = []  # 工作表名称列表
    
    def open_file(self, file_path):
        """打开 Excel 文件并获取工作表名称"""
        self.excel_file_path = file_path
        
        try:
            # 获取所有工作表名称
            xl = pd.ExcelFile(file_path)
            self.sheet_names = xl.sheet_names
            return True
        except Exception as e:
            raise Exception(f"无法读取文件：{str(e)}")
    
    def load_sheet(self, sheet_name, header_row=0, fill_nan=False):
        """加载指定工作表"""
        try:
            # 读取指定工作表，指定表头行
            self.df = pd.read_excel(
                self.excel_file_path,
                sheet_name=sheet_name,
                header=header_row
            )

            # 重置索引为整数，避免浮点数索引问题
            self.df = self.df.reset_index(drop=True)

            # 如果需要填充 nan 值
            if fill_nan:
                self.df = self.df.fillna('')

            self.filtered_df = self.df.copy()
            self.columns = list(self.df.columns)

            return True
        except Exception as e:
            raise Exception(f"无法加载工作表：{str(e)}")
    
    def get_unique_values(self, column_name):
        """获取指定列的唯一值"""
        if self.df is None:
            return []
        
        try:
            unique_vals = self.df[column_name].dropna().astype(str).unique()
            return sorted(unique_vals, key=str)
        except:
            return []
    
    def apply_filters(self, filter_criteria):
        """应用筛选条件"""
        if self.df is None:
            return

        # 从原始数据开始筛选
        mask = pd.Series([True] * len(self.df), index=self.df.index)
        
        for col_name, value in filter_criteria.items():
            if value.strip():
                # 转义正则表达式特殊字符，避免特殊字符导致匹配失败
                escaped_value = re.escape(value.strip())
                # 只匹配非 nan 值且包含筛选值的行（模糊匹配）
                mask &= (pd.notna(self.df[col_name])) & (self.df[col_name].astype(str).str.contains(escaped_value, case=False))

        # 应用筛选
        self.filtered_df = self.df[mask].copy()
        
        # 筛选后重新编号序号列
        if '序号' in self.filtered_df.columns:
            self.filtered_df['序号'] = range(1, len(self.filtered_df) + 1)
        
        return len(self.filtered_df)
    
    def reset_filters(self):
        """重置所有筛选条件"""
        if self.df is not None:
            self.filtered_df = self.df.copy()
            # 重置后重新编号序号列
            if '序号' in self.filtered_df.columns:
                self.filtered_df['序号'] = range(1, len(self.filtered_df) + 1)
    
    def export_to_excel(self, file_path):
        """导出为 Excel 文件"""
        if self.filtered_df is None or self.filtered_df.empty:
            raise Exception("没有可导出的数据")
        
        try:
            self.filtered_df.to_excel(file_path, index=False, engine='openpyxl')
            return True
        except Exception as e:
            raise Exception(f"导出失败：{str(e)}")
    
    def export_to_markdown(self, file_path):
        """导出为 Markdown 表格格式"""
        if self.filtered_df is None or self.filtered_df.empty:
            raise Exception("没有可导出的数据")
        
        try:
            lines = []
            
            # 表头
            headers = [str(col) for col in self.filtered_df.columns]
            lines.append('| ' + ' | '.join(headers) + ' |')
            lines.append('|' + '|'.join(['---' for _ in headers]) + '|')
            
            # 数据行
            for _, row in self.filtered_df.iterrows():
                values = [str(v) if pd.notna(v) else '' for v in row.values]
                # 处理包含 | 的单元格
                values = [v.replace('|', '\\|') for v in values]
                lines.append('| ' + ' | '.join(values) + ' |')
            
            # 写入文件
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(lines))
            
            return True
        except Exception as e:
            raise Exception(f"导出失败：{str(e)}")
    
    def get_data_info(self):
        """获取数据信息"""
        if self.df is None:
            return {"rows": 0, "columns": 0}
        
        return {
            "rows": len(self.df),
            "columns": len(self.columns)
        }
    
    def get_filtered_data_info(self):
        """获取筛选后的数据信息"""
        if self.filtered_df is None:
            return 0
        
        return len(self.filtered_df)
    
    def add_serial_number(self, column_name="序号"):
        """为数据添加序号列
        
        Args:
            column_name: 序号列的名称，默认为"序号"
        """
        if self.df is None:
            return False
        
        try:
            # 为原始数据添加序号列
            if column_name in self.df.columns:
                # 如果序号列已存在，先删除
                self.df = self.df.drop(columns=[column_name])
            
            # 在最前面添加序号列
            self.df.insert(0, column_name, range(1, len(self.df) + 1))
            
            # 同步更新筛选后的数据
            if self.filtered_df is not None:
                if column_name in self.filtered_df.columns:
                    self.filtered_df = self.filtered_df.drop(columns=[column_name])
                self.filtered_df.insert(0, column_name, range(1, len(self.filtered_df) + 1))
            
            # 更新列名列表
            self.columns = list(self.df.columns)
            
            return True
        except Exception as e:
            raise Exception(f"添加序号列失败：{str(e)}")
