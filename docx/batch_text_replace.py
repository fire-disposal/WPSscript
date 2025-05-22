#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
批量替换Word文档中的文本
"""

import os
from docx import Document
import re

# 文件读取部分，便于修改需读取文件名
input_file = "example.docx"  # 请修改为实际的文件名
replace_dict = {
    "原文本1": "替换文本1",
    "原文本2": "替换文本2",
}

def batch_replace_text(doc_path, replacements):
    """
    批量替换Word文档中的文本
    
    Args:
        doc_path: Word文档路径
        replacements: 替换字典，格式为 {原文本: 替换文本}
    
    Returns:
        Document对象
    """
    print(f"正在处理文件: {doc_path}")
    
    # 检查文件是否存在
    if not os.path.exists(doc_path):
        print(f"错误: 文件 '{doc_path}' 不存在!")
        return None
    
    # 打开文档
    doc = Document(doc_path)
    
    # 替换计数器
    replace_count = 0
    
    # 遍历所有段落
    for para in doc.paragraphs:
        for old_text, new_text in replacements.items():
            if old_text in para.text:
                # 使用正则表达式替换文本，保留原格式
                inline_elements = []
                for run in para.runs:
                    inline_elements.append((run.text, run))
                
                # 合并所有文本
                full_text = para.text
                
                # 替换文本
                new_full_text = full_text.replace(old_text, new_text)
                
                if full_text != new_full_text:
                    # 文本已更改，更新第一个run并清除其他run
                    if para.runs:
                        para.runs[0].text = new_full_text
                        for run in para.runs[1:]:
                            run.text = ""
                        replace_count += 1
                        print(f"替换: '{old_text}' -> '{new_text}'")
    
    # 遍历所有表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for old_text, new_text in replacements.items():
                        if old_text in para.text:
                            # 使用相同的替换逻辑
                            full_text = para.text
                            new_full_text = full_text.replace(old_text, new_text)
                            
                            if full_text != new_full_text:
                                if para.runs:
                                    para.runs[0].text = new_full_text
                                    for run in para.runs[1:]:
                                        run.text = ""
                                    replace_count += 1
                                    print(f"表格中替换: '{old_text}' -> '{new_text}'") 
    print(f"共完成 {replace_count} 处替换")
    return doc

def main():
    # 构建输出文件名
    file_name, file_ext = os.path.splitext(input_file)
    output_file = f"{file_name}（已修改）{file_ext}"
    
    # 执行批量替换
    doc = batch_replace_text(input_file, replace_dict)
    
    if doc:
        # 保存修改后的文档
        doc.save(output_file)
        print(f"文档已保存为: {output_file}")

if __name__ == "__main__":
    main()