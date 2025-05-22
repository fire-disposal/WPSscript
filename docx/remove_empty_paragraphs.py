#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
清除Word文档中的多个连续空段落，保留一个空段落
"""

import os
from docx import Document

# 文件读取部分，便于修改需读取文件名
input_file = "探索知识海洋.docx"  # 请修改为实际的文件名

def is_paragraph_empty(paragraph):
    """
    判断段落是否为空
    
    Args:
        paragraph: 段落对象
    
    Returns:
        布尔值，表示段落是否为空
    """
    # 检查段落文本是否为空或只包含空白字符
    return not paragraph.text.strip()

def remove_empty_paragraphs(doc_path):
    """
    清除Word文档中的多个连续空段落，保留一个空段落
    
    Args:
        doc_path: Word文档路径
    
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
    
    # 创建一个新文档来存储处理后的内容
    new_doc = Document()
    
    # 复制原文档的样式
    for style in doc.styles:
        if style.name not in new_doc.styles:
            try:
                new_doc.styles.add_style(style.name, style.type)
            except:
                # 如果样式已存在或无法添加，则跳过
                pass
    
    # 跟踪连续空段落的计数
    empty_count = 0
    removed_count = 0
    
    # 遍历所有段落
    for para in doc.paragraphs:
        if is_paragraph_empty(para):
            empty_count += 1
            
            # 如果是第一个空段落，则保留
            if empty_count == 1:
                new_para = new_doc.add_paragraph()
                # 尝试复制段落格式
                try:
                    if para.style:
                        new_para.style = para.style
                except:
                    pass
        else:
            # 非空段落，重置计数器
            if empty_count > 1:
                removed_count += (empty_count - 1)
                print(f"移除了 {empty_count - 1} 个连续空段落")
            empty_count = 0
            
            # 复制非空段落
            new_para = new_doc.add_paragraph(para.text)
            # 复制段落格式
            try:
                if para.style:
                    new_para.style = para.style
            except:
                pass
            
            # 复制段落中的格式
            for i, run in enumerate(para.runs):
                if i < len(new_para.runs):
                    new_run = new_para.runs[i]
                    # 复制格式
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    new_run.underline = run.underline
                    new_run.font.size = run.font.size
                    if run.font.color.rgb:
                        new_run.font.color.rgb = run.font.color.rgb
    
    # 处理文档末尾的连续空段落
    if empty_count > 1:
        removed_count += (empty_count - 1)
        print(f"移除了 {empty_count - 1} 个连续空段落")
    
    # 复制表格
    for table in doc.tables:
        new_table = new_doc.add_table(rows=len(table.rows), cols=len(table.columns))
        # 复制表格样式
        if table.style:
            new_table.style = table.style
        
        # 复制单元格内容
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                if i < len(new_table.rows) and j < len(new_table.rows[i].cells):
                    new_cell = new_table.rows[i].cells[j]
                    new_cell.text = cell.text
    
    print(f"处理完成，共移除了 {removed_count} 个多余的空段落")
    return new_doc

def main():
    # 构建输出文件名
    file_name, file_ext = os.path.splitext(input_file)
    output_file = f"{file_name}（已修改）{file_ext}"
    
    # 执行空段落清除
    doc = remove_empty_paragraphs(input_file)
    
    if doc:
        # 保存修改后的文档
        doc.save(output_file)
        print(f"文档已保存为: {output_file}")

if __name__ == "__main__":
    main()