#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
提取PowerPoint文件中的所有文本
"""

import os
from pptx import Presentation

# 文件读取部分，便于修改需读取文件名
input_file = "example.pptx"  # 请修改为实际的文件名

def extract_text_from_shape(shape):
    """
    从形状中提取文本
    
    Args:
        shape: PowerPoint形状对象
    
    Returns:
        提取的文本
    """
    text = ""
    
    # 如果形状有文本框
    if shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                text += run.text + " "
            text += "\n"
    
    # 如果形状是表格
    elif shape.has_table:
        for row in shape.table.rows:
            row_text = []
            for cell in row.cells:
                if cell.text:
                    row_text.append(cell.text.strip())
            text += " | ".join(row_text) + "\n"
    
    # 如果形状是图表
    elif shape.has_chart:
        text += "[图表数据]\n"
    
    # 递归处理组合形状
    elif shape.shape_type == 6:  # 6 表示组合形状
        for subshape in shape.shapes:
            text += extract_text_from_shape(subshape)
    
    return text

def extract_text_from_pptx(pptx_path):
    """
    提取PowerPoint文件中的所有文本
    
    Args:
        pptx_path: PowerPoint文件路径
    
    Returns:
        提取的文本字典，格式为 {幻灯片索引: 文本内容}
    """
    print(f"正在处理文件: {pptx_path}")
    
    # 检查文件是否存在
    if not os.path.exists(pptx_path):
        print(f"错误: 文件 '{pptx_path}' 不存在!")
        return {}
    
    # 打开演示文稿
    prs = Presentation(pptx_path)
    
    # 提取文本
    slides_text = {}
    
    for i, slide in enumerate(prs.slides):
        slide_text = f"--- 幻灯片 {i+1} ---\n"
        
        # 提取幻灯片标题
        if slide.shapes.title:
            slide_text += f"标题: {slide.shapes.title.text}\n\n"
        
        # 提取所有形状中的文本
        for shape in slide.shapes:
            shape_text = extract_text_from_shape(shape)
            if shape_text.strip():
                slide_text += shape_text + "\n"
        
        slides_text[i+1] = slide_text
        print(f"已提取幻灯片 {i+1} 的文本")
    
    return slides_text

def main():
    # 执行文本提取
    slides_text = extract_text_from_pptx(input_file)
    
    if slides_text:
        # 构建输出文件名
        file_name = os.path.splitext(os.path.basename(input_file))[0]
        output_file = f"{file_name}_文本提取.txt"
        
        # 保存提取的文本
        with open(output_file, "w", encoding="utf-8") as f:
            for slide_num, text in slides_text.items():
                f.write(text)
                f.write("\n" + "="*50 + "\n\n")
        
        print(f"文本提取完成! 共提取了 {len(slides_text)} 张幻灯片的文本")
        print(f"文本已保存到: {output_file}")
    else:
        print("未能提取任何文本或处理过程中出现错误")

if __name__ == "__main__":
    main()