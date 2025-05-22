#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
提取Word文档中的所有样式信息并输出为JSON文件
"""

import os
import json
from docx import Document
from datetime import datetime

# 文件读取部分，便于修改需读取文件名
input_file = "探索知识海洋.docx"  # 请修改为实际的文件名

# 输出文件名
output_file = f"{os.path.splitext(input_file)[0]}_styles.json"

def extract_styles(file_path):
    """
    提取Word文档中的所有样式信息
    
    Args:
        file_path: Word文档路径
    
    Returns:
        包含样式信息的字典
    """
    # 检查文件是否存在
    if not os.path.exists(file_path):
        print(f"错误: 文件不存在: {file_path}")
        return None
    
    print(f"正在处理文件: {file_path}")
    
    # 打开文档
    doc = Document(file_path)
    
    # 创建样式信息字典
    styles_info = {}
    
    # 获取文档中的所有样式
    for style in doc.styles:
        # 只处理段落样式
        if style.type == 1:  # WD_STYLE_TYPE.PARAGRAPH
            style_info = {}
            
            # 获取样式名称
            style_name = style.name
            print(f"正在提取样式: {style_name}")
            
            # 获取样式基本信息
            style_info["style_id"] = style.style_id
            style_info["type"] = "paragraph"
            
            # 获取字体信息
            if style.font:
                font_info = {}
                if hasattr(style.font, "name") and style.font.name:
                    font_info["name"] = style.font.name
                if hasattr(style.font, "size") and style.font.size:
                    font_info["size"] = style.font.size.pt if hasattr(style.font.size, "pt") else None
                if hasattr(style.font, "bold") and style.font.bold is not None:
                    font_info["bold"] = style.font.bold
                if hasattr(style.font, "italic") and style.font.italic is not None:
                    font_info["italic"] = style.font.italic
                if hasattr(style.font, "underline") and style.font.underline is not None:
                    font_info["underline"] = style.font.underline
                if hasattr(style.font, "color") and style.font.color and style.font.color.rgb:
                    font_info["color"] = style.font.color.rgb
                
                style_info["font"] = font_info
            
            # 获取段落格式信息
            if style.paragraph_format:
                para_format = {}
                if hasattr(style.paragraph_format, "alignment") and style.paragraph_format.alignment:
                    para_format["alignment"] = style.paragraph_format.alignment
                if hasattr(style.paragraph_format, "first_line_indent") and style.paragraph_format.first_line_indent:
                    para_format["first_line_indent"] = style.paragraph_format.first_line_indent.pt if hasattr(style.paragraph_format.first_line_indent, "pt") else None
                if hasattr(style.paragraph_format, "left_indent") and style.paragraph_format.left_indent:
                    para_format["left_indent"] = style.paragraph_format.left_indent.pt if hasattr(style.paragraph_format.left_indent, "pt") else None
                if hasattr(style.paragraph_format, "right_indent") and style.paragraph_format.right_indent:
                    para_format["right_indent"] = style.paragraph_format.right_indent.pt if hasattr(style.paragraph_format.right_indent, "pt") else None
                if hasattr(style.paragraph_format, "line_spacing") and style.paragraph_format.line_spacing:
                    para_format["line_spacing"] = style.paragraph_format.line_spacing
                if hasattr(style.paragraph_format, "space_before") and style.paragraph_format.space_before:
                    para_format["space_before"] = style.paragraph_format.space_before.pt if hasattr(style.paragraph_format.space_before, "pt") else None
                if hasattr(style.paragraph_format, "space_after") and style.paragraph_format.space_after:
                    para_format["space_after"] = style.paragraph_format.space_after.pt if hasattr(style.paragraph_format.space_after, "pt") else None
                
                style_info["paragraph_format"] = para_format
            
            # 获取基础样式
            if style.base_style:
                style_info["base_style"] = style.base_style.name
            
            # 添加到样式信息字典
            styles_info[style_name] = style_info
    
    # 获取标题样式级别信息
    for i in range(1, 10):  # 通常标题级别从1到9
        heading_style_name = f"Heading {i}"
        if heading_style_name in styles_info:
            styles_info[heading_style_name]["heading_level"] = i
    
    print(f"样式提取完成，共提取了 {len(styles_info)} 个样式")
    return styles_info

def main():
    # 执行样式提取
    styles_info = extract_styles(input_file)
    
    if styles_info:
        # 保存为JSON文件
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(styles_info, f, ensure_ascii=False, indent=4)
        
        print(f"样式信息已保存为: {output_file}")

if __name__ == "__main__":
    main()