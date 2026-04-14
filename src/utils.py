#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工具模块
包含辅助函数和工具方法
"""


def calc_text_width(text):
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


def get_file_name_from_path(file_path):
    """从文件路径中提取文件名"""
    import os
    return os.path.basename(file_path)


def get_default_export_filename(original_file_name, suffix="筛选结果"):
    """生成默认导出文件名"""
    import os
    base_name = os.path.splitext(original_file_name)[0]
    return f"{base_name}-{suffix}"
