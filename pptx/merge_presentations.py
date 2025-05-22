#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
合并多个PowerPoint文件
"""

import os
from pptx import Presentation
from datetime import datetime

# 文件读取部分，便于修改需读取文件名
input_files = [
    "presentation1.pptx",
    "presentation2.pptx",
    "presentation3.pptx"
]  # 请修改为实际的文件名列表

# 输出文件名
output_file = f"合并演示文稿（{datetime.now().strftime('%Y%m%d_%H%M%S')}）.pptx"

def merge_presentations(file_paths):
    """
    合并多个PowerPoint文件
    
    Args:
        file_paths: PowerPoint文件路径列表
    
    Returns:
        合并后的Presentation对象
    """
    # 检查文件是否都存在
    missing_files = [f for f in file_paths if not os.path.exists(f)]
    if missing_files:
        print(f"错误: 以下文件不存在: {', '.join(missing_files)}")
        return None
    
    # 创建一个新的演示文稿作为合并的目标
    merged_prs = Presentation()
    
    # 记录每个文件的幻灯片数量
    slide_counts = {}
    
    # 遍历所有输入文件
    for i, file_path in enumerate(file_paths):
        print(f"正在处理文件 {i+1}/{len(file_paths)}: {file_path}")
        
        # 打开当前演示文稿
        prs = Presentation(file_path)
        
        # 记录幻灯片数量
        slide_counts[file_path] = len(prs.slides)
        
        # 添加分隔幻灯片，显示文件名（对第一个文件除外）
        if i > 0:
            # 创建分隔幻灯片
            slide_layout = merged_prs.slide_layouts[0]  # 使用标题幻灯片布局
            slide = merged_prs.slides.add_slide(slide_layout)
            
            # 设置标题
            title = slide.shapes.title
            title.text = f"文件: {os.path.basename(file_path)}"
            
            # 如果有副标题占位符，设置副标题
            for shape in slide.placeholders:
                if shape.placeholder_format.type == 2:  # 2 表示副标题
                    shape.text = f"包含 {len(prs.slides)} 张幻灯片"
        
        # 复制所有幻灯片
        for slide in prs.slides:
            # 获取幻灯片布局
            # 注意：我们尝试使用相同的布局，但如果不存在，则使用默认布局
            try:
                slide_layout = merged_prs.slide_layouts[slide.slide_layout.index]
            except:
                slide_layout = merged_prs.slide_layouts[6]  # 6 通常是空白布局
            
            # 创建新幻灯片
            new_slide = merged_prs.slides.add_slide(slide_layout)
            
            # 复制所有形状
            for shape in slide.shapes:
                # 对于文本框，复制文本内容
                if shape.has_text_frame:
                    # 查找目标幻灯片中的对应形状
                    for target_shape in new_slide.shapes:
                        if (hasattr(target_shape, 'placeholder_format') and 
                            hasattr(shape, 'placeholder_format') and
                            target_shape.placeholder_format.idx == shape.placeholder_format.idx):
                            
                            # 复制文本
                            if shape.text:
                                target_shape.text = shape.text
                            break
            
            # 注意：完全复制所有元素（包括图片、表格等）需要更复杂的代码
            # 这里只是一个简化版本，主要复制文本内容
        
        print(f"已合并文件: {file_path}")
    
    # 打印合并统计信息
    print("\n合并统计:")
    total_slides = 0
    for file_path, count in slide_counts.items():
        print(f"  - {os.path.basename(file_path)}: {count} 张幻灯片")
        total_slides += count
    
    # 添加分隔幻灯片的数量
    separator_slides = len(file_paths) - 1
    if separator_slides > 0:
        print(f"  - 分隔幻灯片: {separator_slides} 张")
    
    print(f"  - 总计: {total_slides + separator_slides} 张幻灯片")
    
    return merged_prs

def main():
    # 执行演示文稿合并
    merged_prs = merge_presentations(input_files)
    
    if merged_prs:
        # 保存合并后的演示文稿
        merged_prs.save(output_file)
        print(f"合并完成! 演示文稿已保存为: {output_file}")
        print(f"共合并了 {len(input_files)} 个演示文稿")

if __name__ == "__main__":
    main()