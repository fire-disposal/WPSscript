#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
提取Word文档中的所有图片
"""

import os
import io
from docx import Document
from PIL import Image

# 文件读取部分，便于修改需读取文件名
input_file = "example.docx"  # 请修改为实际的文件名

def extract_images(doc_path):
    """
    提取Word文档中的所有图片
    
    Args:
        doc_path: Word文档路径
    
    Returns:
        提取的图片数量
    """
    print(f"正在处理文件: {doc_path}")
    
    # 检查文件是否存在
    if not os.path.exists(doc_path):
        print(f"错误: 文件 '{doc_path}' 不存在!")
        return 0
    
    # 创建输出文件夹
    file_name = os.path.splitext(os.path.basename(doc_path))[0]
    output_dir = f"{file_name}_images"
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"创建输出文件夹: {output_dir}")
    
    # 打开文档
    doc = Document(doc_path)
    
    # 提取图片
    image_count = 0
    
    # 遍历文档中的关系
    for rel in doc.part.rels.values():
        # 检查关系类型是否为图片
        if "image" in rel.reltype:
            # 获取图片数据
            image_data = rel.target_part.blob
            
            # 尝试确定图片格式
            try:
                img = Image.open(io.BytesIO(image_data))
                img_format = img.format.lower()
                
                # 保存图片
                image_count += 1
                img_filename = f"{output_dir}/image_{image_count:03d}.{img_format}"
                with open(img_filename, "wb") as img_file:
                    img_file.write(image_data)
                
                print(f"已提取图片 {image_count}: {img_filename} ({img.width}x{img.height}, {img_format})")
                
            except Exception as e:
                print(f"警告: 无法处理图片 {image_count + 1}: {str(e)}")
    
    print(f"共提取了 {image_count} 张图片到文件夹: {output_dir}")
    return image_count

def main():
    # 执行图片提取
    image_count = extract_images(input_file)
    
    if image_count > 0:
        print(f"图片提取完成! 共提取了 {image_count} 张图片")
    else:
        print("未找到任何图片或处理过程中出现错误")

if __name__ == "__main__":
    main()