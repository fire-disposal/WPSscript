#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
将Word文档中的Markdown格式应用为DOCX格式：
1. 将#、##、###等标题格式应用为对应级别的标题样式
2. 将**文本**格式应用为加粗文本
3. 将*文本*格式应用为斜体文本
4. 将~~文本~~格式应用为删除线文本
5. 应用格式后清除Markdown格式标记
"""

import os
import re
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, RGBColor
from docx.enum.style import WD_STYLE_TYPE

# 文件读取部分，便于修改需读取文件名
input_file = "探索知识海洋.docx"  # 请修改为实际的文件名

def clean_remaining_markdown_marks(doc):
    """
    清除文档中所有剩余的Markdown格式标记
    
    Args:
        doc: Document对象
    
    Returns:
        清除的标记数量
    """
    asterisk_count = 0
    tilde_count = 0
    hash_count = 0
    
    # 清除段落中的Markdown标记
    for para in doc.paragraphs:
        # 检查是否有星号(*)
        if '*' in para.text:
            original_text = para.text
            new_text = original_text.replace('*', '')
            if original_text != new_text:
                para.text = new_text
                asterisk_count += original_text.count('*')
        
        # 检查是否有波浪线(~)
        if '~' in para.text:
            original_text = para.text
            new_text = original_text.replace('~', '')
            if original_text != new_text:
                para.text = new_text
                tilde_count += original_text.count('~')
        
        # 检查是否有#后跟空格的标题标记
        if re.search(r'#\s+', para.text):
            original_text = para.text
            new_text = re.sub(r'#\s+', '', original_text)
            if original_text != new_text:
                para.text = new_text
                hash_count += len(re.findall(r'#\s+', original_text))
    
    # 清除表格中的Markdown标记
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    # 检查是否有星号(*)
                    if '*' in para.text:
                        original_text = para.text
                        new_text = original_text.replace('*', '')
                        if original_text != new_text:
                            para.text = new_text
                            asterisk_count += original_text.count('*')
                    
                    # 检查是否有波浪线(~)
                    if '~' in para.text:
                        original_text = para.text
                        new_text = original_text.replace('~', '')
                        if original_text != new_text:
                            para.text = new_text
                            tilde_count += original_text.count('~')
                    
                    # 检查是否有#后跟空格的标题标记
                    if re.search(r'#\s+', para.text):
                        original_text = para.text
                        new_text = re.sub(r'#\s+', '', original_text)
                        if original_text != new_text:
                            para.text = new_text
                            hash_count += len(re.findall(r'#\s+', original_text))
    
    print(f"清除了 {asterisk_count} 个星号(*)、{tilde_count} 个波浪线(~)和 {hash_count} 个标题标记(#)")
    return asterisk_count + tilde_count + hash_count

def apply_markdown_styles(doc_path):
    """
    将Word文档中的Markdown格式应用为DOCX格式
    
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
    
    # 统计处理次数
    heading_count = 0
    bold_count = 0
    italic_count = 0
    strikethrough_count = 0
    
    # 处理段落中的Markdown格式
    for para in doc.paragraphs:
        # 处理标题格式
        if re.match(r'^#+\s+', para.text):
            # 获取标题级别（#的数量）
            heading_level = len(re.match(r'^(#+)\s+', para.text).group(1))
            if heading_level > 9:  # Word只支持到9级标题
                heading_level = 9
            
            # 移除Markdown标记，保留标题文本
            heading_text = re.sub(r'^#+\s+', '', para.text)
            
            # 应用标题样式
            para.text = heading_text
            para.style = f'Heading {heading_level}'
            
            heading_count += 1
            print(f"应用了{heading_level}级标题样式: {heading_text[:30]}...")
        else:
            # 处理段落内的格式（加粗、斜体、删除线）
            # 由于python-docx的限制，我们需要处理每个run
            
            # 先检查是否有Markdown格式
            has_bold = '**' in para.text
            has_italic = re.search(r'(?<!\*)\*(?!\*).*?(?<!\*)\*(?!\*)', para.text) is not None
            has_strikethrough = '~~' in para.text
            
            if has_bold or has_italic or has_strikethrough:
                # 保存原始文本
                original_text = para.text
                
                # 创建一个新的段落来替换原始段落
                new_para = doc.add_paragraph()
                new_para.style = para.style
                
                # 处理加粗格式 **text**
                text = original_text
                bold_matches = re.finditer(r'\*\*(.*?)\*\*', text)
                last_end = 0
                has_matches = False
                
                for match in bold_matches:
                    has_matches = True
                    # 添加匹配前的文本
                    if match.start() > last_end:
                        new_para.add_run(text[last_end:match.start()])
                    
                    # 添加加粗文本
                    bold_run = new_para.add_run(match.group(1))
                    bold_run.bold = True
                    
                    last_end = match.end()
                    bold_count += 1
                
                # 如果没有匹配项，或者有未处理的文本
                if not has_matches:
                    text_to_process = text
                else:
                    text_to_process = text[last_end:] if last_end < len(text) else ""
                
                # 处理斜体格式 *text*（确保不是**的一部分）
                italic_matches = re.finditer(r'(?<!\*)\*(?!\*)(.*?)(?<!\*)\*(?!\*)', text_to_process)
                last_end = 0
                has_matches = False
                
                for match in italic_matches:
                    has_matches = True
                    # 添加匹配前的文本
                    if match.start() > last_end:
                        new_para.add_run(text_to_process[last_end:match.start()])
                    
                    # 添加斜体文本
                    italic_run = new_para.add_run(match.group(1))
                    italic_run.italic = True
                    
                    last_end = match.end()
                    italic_count += 1
                
                # 如果没有匹配项，或者有未处理的文本
                if not has_matches:
                    text_to_process_2 = text_to_process
                else:
                    text_to_process_2 = text_to_process[last_end:] if last_end < len(text_to_process) else ""
                
                # 处理删除线格式 ~~text~~
                strikethrough_matches = re.finditer(r'~~(.*?)~~', text_to_process_2)
                last_end = 0
                has_matches = False
                
                for match in strikethrough_matches:
                    has_matches = True
                    # 添加匹配前的文本
                    if match.start() > last_end:
                        new_para.add_run(text_to_process_2[last_end:match.start()])
                    
                    # 添加删除线文本
                    strikethrough_run = new_para.add_run(match.group(1))
                    strikethrough_run.font.strike = True
                    
                    last_end = match.end()
                    strikethrough_count += 1
                
                # 添加剩余的文本
                if last_end < len(text_to_process_2):
                    new_para.add_run(text_to_process_2[last_end:])
                
                # 删除原始段落
                p = para._p
                p.getparent().remove(p)
    
    # 处理表格中的Markdown格式
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    # 处理标题格式
                    if re.match(r'^#+\s+', para.text):
                        # 获取标题级别（#的数量）
                        heading_level = len(re.match(r'^(#+)\s+', para.text).group(1))
                        if heading_level > 9:  # Word只支持到9级标题
                            heading_level = 9
                        
                        # 移除Markdown标记，保留标题文本
                        heading_text = re.sub(r'^#+\s+', '', para.text)
                        
                        # 应用标题样式
                        para.text = heading_text
                        para.style = f'Heading {heading_level}'
                        
                        heading_count += 1
                    else:
                        # 处理段落内的格式（加粗、斜体、删除线）
                        # 由于python-docx的限制，我们需要处理每个run
                        
                        # 先检查是否有Markdown格式
                        has_bold = '**' in para.text
                        has_italic = re.search(r'(?<!\*)\*(?!\*).*?(?<!\*)\*(?!\*)', para.text) is not None
                        has_strikethrough = '~~' in para.text
                        
                        if has_bold or has_italic or has_strikethrough:
                            # 保存原始文本
                            original_text = para.text
                            
                            # 清空段落
                            para.clear()
                            
                            # 处理加粗格式 **text**
                            text = original_text
                            bold_matches = re.finditer(r'\*\*(.*?)\*\*', text)
                            last_end = 0
                            has_matches = False
                            
                            for match in bold_matches:
                                has_matches = True
                                # 添加匹配前的文本
                                if match.start() > last_end:
                                    para.add_run(text[last_end:match.start()])
                                
                                # 添加加粗文本
                                bold_run = para.add_run(match.group(1))
                                bold_run.bold = True
                                
                                last_end = match.end()
                                bold_count += 1
                            
                            # 如果没有匹配项，或者有未处理的文本
                            if not has_matches:
                                text_to_process = text
                            else:
                                text_to_process = text[last_end:] if last_end < len(text) else ""
                            
                            # 处理斜体格式 *text*（确保不是**的一部分）
                            italic_matches = re.finditer(r'(?<!\*)\*(?!\*)(.*?)(?<!\*)\*(?!\*)', text_to_process)
                            last_end = 0
                            has_matches = False
                            
                            for match in italic_matches:
                                has_matches = True
                                # 添加匹配前的文本
                                if match.start() > last_end:
                                    para.add_run(text_to_process[last_end:match.start()])
                                
                                # 添加斜体文本
                                italic_run = para.add_run(match.group(1))
                                italic_run.italic = True
                                
                                last_end = match.end()
                                italic_count += 1
                            
                            # 如果没有匹配项，或者有未处理的文本
                            if not has_matches:
                                text_to_process_2 = text_to_process
                            else:
                                text_to_process_2 = text_to_process[last_end:] if last_end < len(text_to_process) else ""
                            
                            # 处理删除线格式 ~~text~~
                            strikethrough_matches = re.finditer(r'~~(.*?)~~', text_to_process_2)
                            last_end = 0
                            has_matches = False
                            
                            for match in strikethrough_matches:
                                has_matches = True
                                # 添加匹配前的文本
                                if match.start() > last_end:
                                    para.add_run(text_to_process_2[last_end:match.start()])
                                
                                # 添加删除线文本
                                strikethrough_run = para.add_run(match.group(1))
                                strikethrough_run.font.strike = True
                                
                                last_end = match.end()
                                strikethrough_count += 1
                            
                            # 添加剩余的文本
                            if last_end < len(text_to_process_2):
                                para.add_run(text_to_process_2[last_end:])
    
    print(f"处理完成，共应用了 {heading_count} 处标题样式、{bold_count} 处加粗格式、{italic_count} 处斜体格式和 {strikethrough_count} 处删除线格式")
    return doc

def main():
    # 构建输出文件名
    file_name, file_ext = os.path.splitext(input_file)
    output_file = f"{file_name}（已修改）{file_ext}"
    
    # 执行Markdown格式应用
    doc = apply_markdown_styles(input_file)
    
    if doc:
        # 清除所有剩余的Markdown标记
        print("正在清除剩余的Markdown格式标记...")
        clean_remaining_markdown_marks(doc)
        
        # 保存修改后的文档
        doc.save(output_file)
        print(f"文档已保存为: {output_file}")

if __name__ == "__main__":
    main()