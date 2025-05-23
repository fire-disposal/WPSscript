#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
获取Excel文件中所有工作表的名称
以数组形式输出所有工作表名称
"""

import os
import openpyxl

# 文件读取部分，便于修改需读取文件名
input_file = "example.xlsx"  # 请修改为实际的文件名

def get_worksheet_names(file_path):
    """
    获取Excel文件中所有工作表的名称
    
    Args:
        file_path: Excel文件路径
    
    Returns:
        工作表名称列表
    """
    # 检查文件是否存在
    if not os.path.exists(file_path):
        print(f"错误: 文件 '{file_path}' 不存在!")
        return []
    
    try:
        # 打开工作簿
        print(f"正在打开文件: {file_path}")
        wb = openpyxl.load_workbook(file_path, read_only=True)
        
        # 获取所有工作表名称
        worksheet_names = wb.sheetnames
        
        # 关闭工作簿
        wb.close()
        
        return worksheet_names
    except Exception as e:
        print(f"处理文件时出错: {str(e)}")
        return []

def main():
    # 获取工作表名称
    worksheet_names = get_worksheet_names(input_file)
    
    if worksheet_names:
        # 打印工作表名称数组
        print("\n工作表名称列表:")
        print(worksheet_names)
        
        # 打印格式化的工作表名称
        print("\n工作表详情:")
        for i, name in enumerate(worksheet_names, 1):
            print(f"{i}. {name}")
        
        # 打印工作表数量
        print(f"\n共有 {len(worksheet_names)} 个工作表")
    else:
        print("未能获取工作表名称，请检查文件是否存在或是否为有效的Excel文件。")

if __name__ == "__main__":
    main()