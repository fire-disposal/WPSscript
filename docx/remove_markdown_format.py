#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
去除Word文档中的Markdown格式标记：
1. 移除两个及以上连续的***（星号）
2. 移除#后带空格的标题记号
"""

import os
import re
from docx import Document

# 文件读取部分，便于修改需读取文件名
input_file = "探索知识海洋.docx"  # 请修改为实际的文件名

def remove_markdown_format(doc_path):
    """
    去除Word文档中的Markdown格式标记
    
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
    
    # 统计替换次数
    asterisk_count = 0
    heading_count = 0
    
    # 遍历所有段落
    for para in doc.paragraphs:
        # *必须死
        if '*' in para.text:
            original_text = para.text
            # 使用正则表达式匹配两个及以上连续的*
            new_text = re.sub(r'\*{2,}', '', original_text)
            
            if original_text != new_text:
                # 计算替换次数（匹配的模式数量）
                matches = re.findall(r'\*{2,}', original_text)
                asterisk_count += len(matches)
                
                # 更新段落文本
                for i, run in enumerate(para.runs):
                    if i == 0:
                        # 将所有文本放在第一个run中
                        run.text = new_text
                    else:
                        # 清空其他run
                        run.text = ""
        
        # 检查是否有#后带空格的标题记号
        if re.search(r'#\s+', para.text):
            original_text = para.text
            # 使用正则表达式匹配#后带空格的标题记号
            new_text = re.sub(r'#\s+', '', original_text)
            
            if original_text != new_text:
                # 计算替换次数（匹配的模式数量）
                matches = re.findall(r'#\s+', original_text)
                heading_count += len(matches)
                
                # 更新段落文本
                for i, run in enumerate(para.runs):
                    if i == 0:
                        # 将所有文本放在第一个run中
                        run.text = new_text
                    else:
                        # 清空其他run
                        run.text = ""
    
    # 遍历所有表格中的文本
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    # 检查是否有连续的***（两个及以上）
                    if '**' in para.text:
                        original_text = para.text
                        # 使用正则表达式匹配两个及以上连续的*
                        new_text = re.sub(r'\*{2,}', '', original_text)
                        
                        if original_text != new_text:
                            # 计算替换次数
                            matches = re.findall(r'\*{2,}', original_text)
                            asterisk_count += len(matches)
                            
                            # 更新段落文本
                            for i, run in enumerate(para.runs):
                                if i == 0:
                                    run.text = new_text
                                else:
                                    run.text = ""
                    
                    # 检查是否有#后带空格的标题记号
                    if re.search(r'#\s+', para.text):
                        original_text = para.text
                        # 使用正则表达式匹配#后带空格的标题记号
                        new_text = re.sub(r'#\s+', '', original_text)
                        
                        if original_text != new_text:
                            # 计算替换次数
                            matches = re.findall(r'#\s+', original_text)
                            heading_count += len(matches)
                            
                            # 更新段落文本
                            for i, run in enumerate(para.runs):
                                if i == 0:
                                    run.text = new_text
                                else:
                                    run.text = ""
    
    print(f"处理完成，共移除了 {asterisk_count} 处连续星号(***)和 {heading_count} 处标题记号(#)")
    return doc

def main():
    # 构建输出文件名
    file_name, file_ext = os.path.splitext(input_file)
    output_file = f"{file_name}（已修改）{file_ext}"
    
    # 执行Markdown格式移除
    doc = remove_markdown_format(input_file)
    
    if doc:
        # 保存修改后的文档
        doc.save(output_file)
        print(f"文档已保存为: {output_file}")

if __name__ == "__main__":
    main()