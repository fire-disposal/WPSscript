#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
提取PowerPoint文件中的所有文本
"""

import os
from pptx import Presentation

# 文件读取部分，便于修改需读取文件名
input_file = "学习社团.pptx"  # 请修改为实际的文件名

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
        paragraphs_text = []
        for paragraph in shape.text_frame.paragraphs:
            if paragraph.runs:  # 只处理有内容的段落
                paragraph_text = ""
                for run in paragraph.runs:
                    paragraph_text += run.text + " "
                paragraphs_text.append(paragraph_text.strip())
        # 只在有内容的段落之间添加换行符
        text += "\n".join(filter(None, paragraphs_text))
    
    # 如果形状是表格
    elif shape.has_table:
        table_rows = []
        for row in shape.table.rows:
            row_text = []
            for cell in row.cells:
                if cell.text:
                    row_text.append(cell.text.strip())
            if row_text:  # 只添加非空行
                table_rows.append(" | ".join(row_text))
        if table_rows:
            text += "\n".join(table_rows)
    
    # 如果形状是图表
    elif shape.has_chart:
        text += "[图表数据]"
    
    # 递归处理组合形状
    elif shape.shape_type == 6:  # 6 表示组合形状
        subshape_texts = []
        for subshape in shape.shapes:
            subshape_text = extract_text_from_shape(subshape)
            if subshape_text.strip():
                subshape_texts.append(subshape_text)
        if subshape_texts:
            text += "\n".join(subshape_texts)
    
    return text

def extract_text_from_pptx(pptx_path):
    """
    提取PowerPoint文件中的所有文本
    
    Args:
        pptx_path: PowerPoint文件路径
    
    Returns:
        提取的文本字典，格式为 {幻灯片索引: 文本内容}
    """
    # 不再输出处理进度
    
    # 检查文件是否存在
    if not os.path.exists(pptx_path):
        print(f"错误: 文件 '{pptx_path}' 不存在!")
        return {}
    
    # 打开演示文稿
    prs = Presentation(pptx_path)
    
    # 提取文本
    slides_text = {}
    
    for i, slide in enumerate(prs.slides):
        slide_text = f"======幻灯片 {i+1}======\n"
        
        # 提取幻灯片标题
        if slide.shapes.title and slide.shapes.title.text.strip():
            slide_text += f"标题: {slide.shapes.title.text.strip()}\n"
        
        # 提取所有形状中的文本
        shape_texts = []
        for shape in slide.shapes:
            # 跳过标题形状，因为已经单独处理过了
            if shape == slide.shapes.title:
                continue
                
            shape_text = extract_text_from_shape(shape)
            if shape_text.strip():
                shape_texts.append(shape_text)
        
        # 只在有内容的形状文本之间添加换行符
        if shape_texts:
            slide_text += "\n".join(shape_texts)
        
        # 移除多余的空行，但保留基本结构
        lines = [line for line in slide_text.splitlines() if line.strip()]
        
        # 重新构建文本，保持结构清晰
        slide_text = lines[0]  # 幻灯片标题行
        
        # 如果有标题和内容，在标题和内容之间添加一个空行
        if len(lines) > 1 and lines[1].startswith("标题:"):
            slide_text += "\n" + lines[1]
            content_start = 2
        else:
            content_start = 1
            
        # 添加内容，避免连续的空行
        if len(lines) > content_start:
            slide_text += "\n\n" + "\n".join(lines[content_start:])
        slides_text[i+1] = slide_text
        # 不再输出每张幻灯片的处理进度
    
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
                # 确保每张幻灯片的文本格式正确
                if not text.endswith("\n"):
                    text += "\n"
                f.write(text)
                print(text)
                # 不再需要额外的分隔符，因为每个幻灯片标题已经有了明显的格式
        
        print(f"文本提取完成! 共提取了 {len(slides_text)} 张幻灯片的文本")
        
        print(f"文本已保存到: {output_file}")
    else:
        print("未能提取任何文本或处理过程中出现错误")

if __name__ == "__main__":
    main()