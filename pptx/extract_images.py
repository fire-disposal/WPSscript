#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
提取PowerPoint文件中的所有图片（包括背景图片）并保存为PNG格式
"""

import os
import io
from pptx import Presentation
from PIL import Image

# 文件读取部分，便于修改需读取文件名
input_file = "学习社团.pptx"  # 请修改为实际的文件名

def extract_images_from_pptx(pptx_path):
    """
    从PowerPoint文件中提取所有图片（包括背景图片）并保存为PNG格式
    
    Args:
        pptx_path: PowerPoint文件路径
    
    Returns:
        提取的图片数量
    """
    print(f"正在处理文件: {pptx_path}")
    
    # 检查文件是否存在
    if not os.path.exists(pptx_path):
        print(f"错误: 文件 '{pptx_path}' 不存在!")
        return 0
    
    # 创建输出文件夹
    file_name = os.path.splitext(os.path.basename(pptx_path))[0]
    output_dir = f"{file_name}_images"
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"创建输出文件夹: {output_dir}")
    
    # 打开演示文稿
    prs = Presentation(pptx_path)
    
    # 获取幻灯片数量
    slide_count = len(prs.slides)
    print(f"共有 {slide_count} 张幻灯片")
    
    # 提取图片
    extracted_count = 0
    background_count = 0
    
    # 遍历所有幻灯片
    for i, slide in enumerate(prs.slides):
        print(f"正在处理幻灯片 {i+1}/{slide_count}")
        
        # 提取普通图片
        for j, shape in enumerate(slide.shapes):
            if shape.shape_type == 13:  # 13 表示图片
                try:
                    # 获取图片数据
                    image = shape.image
                    image_bytes = image.blob
                    
                    # 使用PIL打开图片
                    img = Image.open(io.BytesIO(image_bytes))
                    
                    # 保存为PNG格式
                    img_path = os.path.join(output_dir, f"slide_{i+1:03d}_image_{j+1:03d}.png")
                    img.save(img_path, "PNG")
                    
                    extracted_count += 1
                    print(f"  - 已提取图片: {img_path}")
                except Exception as e:
                    print(f"  - 提取图片时出错: {str(e)}")
        
        # 尝试提取背景图片
        try:
            # 获取幻灯片背景
            background = slide.background
            
            # 检查是否有填充
            if hasattr(background, 'fill') and background.fill:
                # 检查是否是图片填充
                if hasattr(background.fill, 'type') and background.fill.type == 2:  # 2 表示图片填充
                    try:
                        # 获取背景图片
                        image = background.fill.image
                        image_bytes = image.blob
                        
                        # 使用PIL打开图片
                        img = Image.open(io.BytesIO(image_bytes))
                        
                        # 保存为PNG格式
                        img_path = os.path.join(output_dir, f"slide_{i+1:03d}_background.png")
                        img.save(img_path, "PNG")
                        
                        background_count += 1
                        print(f"  - 已提取背景图片: {img_path}")
                    except Exception as e:
                        print(f"  - 提取背景图片时出错: {str(e)}")
        except Exception as e:
            print(f"  - 处理背景时出错: {str(e)}")
        
        # 尝试提取幻灯片母版中的背景图片
        try:
            # 获取幻灯片母版
            slide_layout = slide.slide_layout
            master = slide_layout.slide_master
            
            # 检查母版背景
            if hasattr(master, 'background') and master.background:
                # 检查是否有填充
                if hasattr(master.background, 'fill') and master.background.fill:
                    # 检查是否是图片填充
                    if hasattr(master.background.fill, 'type') and master.background.fill.type == 2:  # 2 表示图片填充
                        try:
                            # 获取背景图片
                            image = master.background.fill.image
                            image_bytes = image.blob
                            
                            # 使用PIL打开图片
                            img = Image.open(io.BytesIO(image_bytes))
                            
                            # 保存为PNG格式
                            img_path = os.path.join(output_dir, f"slide_{i+1:03d}_master_background.png")
                            img.save(img_path, "PNG")
                            
                            background_count += 1
                            print(f"  - 已提取母版背景图片: {img_path}")
                        except Exception as e:
                            print(f"  - 提取母版背景图片时出错: {str(e)}")
        except Exception as e:
            print(f"  - 处理母版背景时出错: {str(e)}")
    
    return extracted_count, background_count

def main():
    # 执行图片提取
    extracted_count, background_count = extract_images_from_pptx(input_file)
    
    total_count = extracted_count + background_count
    
    if total_count > 0:
        print(f"\n提取完成! 共提取了 {total_count} 张图片")
        print(f"  - 普通图片: {extracted_count} 张")
        print(f"  - 背景图片: {background_count} 张")
        print(f"图片已保存到文件夹: {os.path.splitext(os.path.basename(input_file))[0]}_images")
    else:
        print("未能提取任何图片或处理过程中出现错误")

if __name__ == "__main__":
    main()