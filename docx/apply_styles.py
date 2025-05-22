#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
根据JSON文件中的样式信息修改Word文档中的样式
"""

import os
import json
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE

# 文件读取部分，便于修改需读取文件名
input_file = "document.docx"  # 请修改为实际的文件名
styles_json_file = "document_styles.json"  # 请修改为实际的样式JSON文件名

# 输出文件名
output_file = f"{os.path.splitext(input_file)[0]}（已修改）.docx"

def apply_styles(file_path, styles_info):
    """
    根据样式信息修改Word文档中的样式
    
    Args:
        file_path: Word文档路径
        styles_info: 包含样式信息的字典
    
    Returns:
        修改后的Document对象
    """
    # 检查文件是否存在
    if not os.path.exists(file_path):
        print(f"错误: 文件不存在: {file_path}")
        return None
    
    print(f"正在处理文件: {file_path}")
    
    # 打开文档
    doc = Document(file_path)
    
    # 遍历样式信息字典
    for style_name, style_info in styles_info.items():
        print(f"正在应用样式: {style_name}")
        
        # 检查样式是否存在于文档中
        try:
            style = doc.styles[style_name]
        except KeyError:
            print(f"警告: 文档中不存在样式 '{style_name}'，将创建新样式")
            # 创建新样式
            style_type = WD_STYLE_TYPE.PARAGRAPH
            style = doc.styles.add_style(style_name, style_type)
        
        # 应用字体信息
        if "font" in style_info:
            font_info = style_info["font"]
            
            if "name" in font_info and font_info["name"]:
                style.font.name = font_info["name"]
            
            if "size" in font_info and font_info["size"]:
                style.font.size = Pt(font_info["size"])
            
            if "bold" in font_info:
                style.font.bold = font_info["bold"]
            
            if "italic" in font_info:
                style.font.italic = font_info["italic"]
            
            if "underline" in font_info:
                style.font.underline = font_info["underline"]
            
            if "color" in font_info and font_info["color"]:
                # 处理颜色值
                color_str = font_info["color"]
                if isinstance(color_str, str) and color_str.startswith('#'):
                    # 如果是十六进制颜色值
                    color_str = color_str.lstrip('#')
                    r, g, b = int(color_str[0:2], 16), int(color_str[2:4], 16), int(color_str[4:6], 16)
                    style.font.color.rgb = RGBColor(r, g, b)
                elif isinstance(color_str, list) and len(color_str) == 3:
                    # 如果是RGB列表
                    r, g, b = color_str
                    style.font.color.rgb = RGBColor(r, g, b)
        
        # 应用段落格式信息
        if "paragraph_format" in style_info:
            para_format = style_info["paragraph_format"]
            
            if "alignment" in para_format:
                alignment_value = para_format["alignment"]
                # 处理对齐方式
                if isinstance(alignment_value, int):
                    style.paragraph_format.alignment = alignment_value
                elif isinstance(alignment_value, str):
                    alignment_map = {
                        "left": WD_PARAGRAPH_ALIGNMENT.LEFT,
                        "center": WD_PARAGRAPH_ALIGNMENT.CENTER,
                        "right": WD_PARAGRAPH_ALIGNMENT.RIGHT,
                        "justify": WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                    }
                    if alignment_value.lower() in alignment_map:
                        style.paragraph_format.alignment = alignment_map[alignment_value.lower()]
            
            if "first_line_indent" in para_format and para_format["first_line_indent"] is not None:
                style.paragraph_format.first_line_indent = Pt(para_format["first_line_indent"])
            
            if "left_indent" in para_format and para_format["left_indent"] is not None:
                style.paragraph_format.left_indent = Pt(para_format["left_indent"])
            
            if "right_indent" in para_format and para_format["right_indent"] is not None:
                style.paragraph_format.right_indent = Pt(para_format["right_indent"])
            
            if "line_spacing" in para_format and para_format["line_spacing"] is not None:
                style.paragraph_format.line_spacing = para_format["line_spacing"]
            
            if "space_before" in para_format and para_format["space_before"] is not None:
                style.paragraph_format.space_before = Pt(para_format["space_before"])
            
            if "space_after" in para_format and para_format["space_after"] is not None:
                style.paragraph_format.space_after = Pt(para_format["space_after"])
    
    print(f"样式应用完成，共应用了 {len(styles_info)} 个样式")
    return doc

def main():
    # 读取样式信息JSON文件
    if not os.path.exists(styles_json_file):
        print(f"错误: 样式JSON文件不存在: {styles_json_file}")
        return
    
    with open(styles_json_file, 'r', encoding='utf-8') as f:
        styles_info = json.load(f)
    
    print(f"已读取样式信息，共 {len(styles_info)} 个样式")
    
    # 应用样式
    modified_doc = apply_styles(input_file, styles_info)
    
    if modified_doc:
        # 保存修改后的文档
        modified_doc.save(output_file)
        print(f"文档样式已修改，保存为: {output_file}")

if __name__ == "__main__":
    main()