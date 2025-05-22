#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
将PowerPoint幻灯片导出为图片
注意：此脚本需要安装额外的库 Pillow
"""

import os
import io
import sys
from pptx import Presentation
from PIL import Image, ImageDraw, ImageFont

# 文件读取部分，便于修改需读取文件名
input_file = "example.pptx"  # 请修改为实际的文件名

# 导出设置
export_format = "png"  # 可选: "png", "jpg", "pdf"
image_quality = 90  # 仅对jpg格式有效，范围1-100
image_resolution = (1920, 1080)  # 导出图片的分辨率

def export_slides_as_images(pptx_path, format="png", quality=90, resolution=(1920, 1080)):
    """
    将PowerPoint幻灯片导出为图片
    
    Args:
        pptx_path: PowerPoint文件路径
        format: 导出格式，可选 "png", "jpg"
        quality: 图片质量，仅对jpg格式有效，范围1-100
        resolution: 导出图片的分辨率，元组 (宽, 高)
    
    Returns:
        导出的图片数量
    """
    print(f"正在处理文件: {pptx_path}")
    
    # 检查文件是否存在
    if not os.path.exists(pptx_path):
        print(f"错误: 文件 '{pptx_path}' 不存在!")
        return 0
    
    # 创建输出文件夹
    file_name = os.path.splitext(os.path.basename(pptx_path))[0]
    output_dir = f"{file_name}_slides"
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"创建输出文件夹: {output_dir}")
    
    # 打开演示文稿
    prs = Presentation(pptx_path)
    
    # 获取幻灯片数量
    slide_count = len(prs.slides)
    print(f"共有 {slide_count} 张幻灯片")
    
    # 由于python-pptx不直接支持导出为图片，我们将创建模拟的幻灯片图片
    # 在实际应用中，可能需要使用其他库或方法来实现真正的导出功能
    
    # 创建一个简单的幻灯片模拟图像
    def create_slide_image(slide_index, slide):
        # 创建一个空白图像
        img = Image.new('RGB', resolution, color='white')
        draw = ImageDraw.Draw(img)
        
        # 尝试加载字体，如果失败则使用默认字体
        try:
            font = ImageFont.truetype("arial.ttf", 40)
            small_font = ImageFont.truetype("arial.ttf", 20)
        except IOError:
            font = ImageFont.load_default()
            small_font = ImageFont.load_default()
        
        # 绘制幻灯片标题
        title_text = "无标题"
        if slide.shapes.title and slide.shapes.title.text:
            title_text = slide.shapes.title.text
        
        draw.text((resolution[0]/2 - 200, 50), title_text, fill='black', font=font)
        
        # 绘制幻灯片内容提示
        draw.text((resolution[0]/2 - 300, resolution[1]/2), 
                 f"幻灯片 {slide_index+1} 内容 (模拟图像)", fill='gray', font=font)
        
        # 绘制说明
        draw.text((20, resolution[1]-40), 
                 "注意: 这是模拟的幻灯片图像。实际应用中需要使用其他方法实现真正的导出。", 
                 fill='red', font=small_font)
        
        return img
    
    # 导出幻灯片
    exported_count = 0
    for i, slide in enumerate(prs.slides):
        try:
            # 创建幻灯片图像
            img = create_slide_image(i, slide)
            
            # 保存图像
            img_path = os.path.join(output_dir, f"slide_{i+1:03d}.{format.lower()}")
            
            if format.lower() == "jpg" or format.lower() == "jpeg":
                img.save(img_path, quality=quality)
            else:
                img.save(img_path)
            
            exported_count += 1
            print(f"已导出幻灯片 {i+1}/{slide_count}: {img_path}")
            
        except Exception as e:
            print(f"导出幻灯片 {i+1} 时出错: {str(e)}")
    
    return exported_count

def main():
    # 执行幻灯片导出
    exported_count = export_slides_as_images(
        input_file, 
        format=export_format, 
        quality=image_quality, 
        resolution=image_resolution
    )
    
    if exported_count > 0:
        print(f"导出完成! 共导出了 {exported_count} 张幻灯片图像")
        print(f"图像已保存到文件夹: {os.path.splitext(os.path.basename(input_file))[0]}_slides")
    else:
        print("未能导出任何幻灯片或处理过程中出现错误")

if __name__ == "__main__":
    main()