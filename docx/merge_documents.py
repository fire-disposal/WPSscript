#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
合并多个Word文档
"""

import os
from docx import Document
from datetime import datetime

# 文件读取部分，便于修改需读取文件名
input_files = [
    "document1.docx",
    "document2.docx",
    "document3.docx"
]  # 请修改为实际的文件名列表

# 输出文件名
output_file = f"合并文档（{datetime.now().strftime('%Y%m%d_%H%M%S')}）.docx"

def merge_documents(file_paths):
    """
    合并多个Word文档
    
    Args:
        file_paths: Word文档路径列表
    
    Returns:
        合并后的Document对象
    """
    # 检查文件是否都存在
    missing_files = [f for f in file_paths if not os.path.exists(f)]
    if missing_files:
        print(f"错误: 以下文件不存在: {', '.join(missing_files)}")
        return None
    
    # 创建一个新文档作为合并的目标
    merged_doc = Document()
    
    # 遍历所有输入文件
    for i, file_path in enumerate(file_paths):
        print(f"正在处理文件 {i+1}/{len(file_paths)}: {file_path}")
        
        # 打开当前文档
        doc = Document(file_path)
        
        # 如果不是第一个文档，添加分节符
        if i > 0:
            merged_doc.add_page_break()
            
        # 添加文件名作为标题
        file_name = os.path.basename(file_path)
        merged_doc.add_heading(f"文件: {file_name}", level=1)
        merged_doc.add_paragraph()  # 添加空行
        
        # 复制所有段落
        for para in doc.paragraphs:
            new_para = merged_doc.add_paragraph()
            # 复制文本和格式
            for run in para.runs:
                new_run = new_para.add_run(run.text)
                # 复制格式
                new_run.bold = run.bold
                new_run.italic = run.italic
                new_run.underline = run.underline
                new_run.font.size = run.font.size
                if run.font.color.rgb:
                    new_run.font.color.rgb = run.font.color.rgb
        
        # 复制所有表格
        for table in doc.tables:
            # 获取表格行数和列数
            rows = len(table.rows)
            cols = len(table.rows[0].cells) if rows > 0 else 0
            
            # 创建新表格
            new_table = merged_doc.add_table(rows=rows, cols=cols)
            new_table.style = table.style
            
            # 复制单元格内容
            for i, row in enumerate(table.rows):
                for j, cell in enumerate(row.cells):
                    # 复制单元格文本
                    for para in cell.paragraphs:
                        if para.text:
                            new_table.rows[i].cells[j].text = para.text
                            break
        
        print(f"已合并文件: {file_path}")
    
    return merged_doc

def main():
    # 执行文档合并
    merged_doc = merge_documents(input_files)
    
    if merged_doc:
        # 保存合并后的文档
        merged_doc.save(output_file)
        print(f"合并完成! 文档已保存为: {output_file}")
        print(f"共合并了 {len(input_files)} 个文档")

if __name__ == "__main__":
    main()