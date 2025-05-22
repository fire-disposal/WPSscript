#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
移除Word文档中的所有批注
"""

import os
import zipfile
import shutil
import tempfile
import xml.etree.ElementTree as ET
from docx import Document

# 文件读取部分，便于修改需读取文件名
input_file = "探索知识海洋.docx"  # 请修改为实际的文件名

def remove_comments(doc_path):
    """
    移除Word文档中的所有批注
    
    Args:
        doc_path: Word文档路径
    
    Returns:
        处理后的文档路径
    """
    print(f"正在处理文件: {doc_path}")
    
    # 检查文件是否存在
    if not os.path.exists(doc_path):
        print(f"错误: 文件 '{doc_path}' 不存在!")
        return None
    
    # 创建临时目录
    temp_dir = tempfile.mkdtemp()
    
    try:
        # 构建输出文件名
        file_name, file_ext = os.path.splitext(doc_path)
        output_file = f"{file_name}（已修改）{file_ext}"
        
        # 复制原始文件到临时文件
        temp_file = os.path.join(temp_dir, "temp.docx")
        shutil.copy2(doc_path, temp_file)
        
        # 解压docx文件（实际上是一个zip文件）
        extract_dir = os.path.join(temp_dir, "extracted")
        with zipfile.ZipFile(temp_file, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
        
        # 处理评论
        comments_removed = False
        
        # 1. 删除comments.xml文件（如果存在）
        comments_path = os.path.join(extract_dir, "word", "comments.xml")
        if os.path.exists(comments_path):
            os.remove(comments_path)
            comments_removed = True
            print("已删除comments.xml文件")
        
        # 2. 删除commentsExtended.xml文件（如果存在，Office 2013+）
        comments_extended_path = os.path.join(extract_dir, "word", "commentsExtended.xml")
        if os.path.exists(comments_extended_path):
            os.remove(comments_extended_path)
            comments_removed = True
            print("已删除commentsExtended.xml文件")
        
        # 3. 删除WPS格式的评论文件（如果存在）
        wps_comments_path = os.path.join(extract_dir, "word", "wpsComments.xml")
        if os.path.exists(wps_comments_path):
            os.remove(wps_comments_path)
            comments_removed = True
            print("已删除wpsComments.xml文件")
        
        # 4. 修改document.xml，移除评论引用
        document_path = os.path.join(extract_dir, "word", "document.xml")
        if os.path.exists(document_path):
            # 解析XML
            ET.register_namespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            ET.register_namespace('w14', 'http://schemas.microsoft.com/office/word/2010/wordml')
            ET.register_namespace('w15', 'http://schemas.microsoft.com/office/word/2012/wordml')
            
            tree = ET.parse(document_path)
            root = tree.getroot()
            
            # 定义命名空间
            namespaces = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
                'w15': 'http://schemas.microsoft.com/office/word/2012/wordml'
            }
            
            # 查找并移除评论引用标记
            comment_references = 0
            
            # 移除评论引用 (w:commentReference)
            for elem in root.findall('.//w:commentReference', namespaces):
                parent = elem.getparent() if hasattr(elem, 'getparent') else None
                if parent is not None:
                    parent.remove(elem)
                    comment_references += 1
            
            # 移除评论范围开始标记 (w:commentRangeStart)
            for elem in root.findall('.//w:commentRangeStart', namespaces):
                parent = elem.getparent() if hasattr(elem, 'getparent') else None
                if parent is not None:
                    parent.remove(elem)
                    comment_references += 1
            
            # 移除评论范围结束标记 (w:commentRangeEnd)
            for elem in root.findall('.//w:commentRangeEnd', namespaces):
                parent = elem.getparent() if hasattr(elem, 'getparent') else None
                if parent is not None:
                    parent.remove(elem)
                    comment_references += 1
            
            # 移除WPS格式的评论标记
            for elem in root.findall('.//*[@wpscomment]', namespaces):
                if 'wpscomment' in elem.attrib:
                    del elem.attrib['wpscomment']
                    comment_references += 1
            
            # 查找可能的自定义评论标记
            custom_comment_patterns = [
                './/*[contains(text(), "[批注")]',
                './/*[contains(text(), "【批注】")]',
                './/*[contains(text(), "（批注）")]',
                './/*[contains(text(), "(批注)")]'
            ]
            
            for pattern in custom_comment_patterns:
                try:
                    for elem in root.xpath(pattern):
                        text = elem.text
                        if text:
                            # 移除批注标记
                            import re
                            new_text = re.sub(r'\[批注.*?\]|\【批注.*?\】|\（批注.*?\）|\(批注.*?\)', '', text)
                            elem.text = new_text
                            comment_references += 1
                except:
                    # xpath可能不可用，跳过
                    pass
            
            if comment_references > 0:
                # 保存修改后的XML
                tree.write(document_path, encoding='UTF-8', xml_declaration=True)
                comments_removed = True
                print(f"已从文档中移除 {comment_references} 处评论引用")
        
        # 5. 修改[Content_Types].xml，移除评论内容类型
        content_types_path = os.path.join(extract_dir, "[Content_Types].xml")
        if os.path.exists(content_types_path):
            # 解析XML
            ET.register_namespace('', 'http://schemas.openxmlformats.org/package/2006/content-types')
            
            tree = ET.parse(content_types_path)
            root = tree.getroot()
            
            # 查找并移除评论相关的内容类型
            comment_content_types = 0
            
            for elem in root.findall(".//{http://schemas.openxmlformats.org/package/2006/content-types}Override"):
                part_name = elem.get('PartName', '')
                if 'comments.xml' in part_name or 'commentsExtended.xml' in part_name or 'wpsComments.xml' in part_name:
                    root.remove(elem)
                    comment_content_types += 1
            
            if comment_content_types > 0:
                # 保存修改后的XML
                tree.write(content_types_path, encoding='UTF-8', xml_declaration=True)
                comments_removed = True
                print(f"已从内容类型中移除 {comment_content_types} 个评论相关条目")
        
        # 6. 修改document.xml.rels，移除评论关系
        rels_path = os.path.join(extract_dir, "word", "_rels", "document.xml.rels")
        if os.path.exists(rels_path):
            # 解析XML
            ET.register_namespace('', 'http://schemas.openxmlformats.org/package/2006/relationships')
            
            tree = ET.parse(rels_path)
            root = tree.getroot()
            
            # 查找并移除评论相关的关系
            comment_relationships = 0
            
            for elem in root.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                target = elem.get('Target', '')
                rel_type = elem.get('Type', '')
                if 'comments.xml' in target or 'commentsExtended.xml' in target or 'wpsComments.xml' in target or 'comment' in rel_type.lower():
                    root.remove(elem)
                    comment_relationships += 1
            
            if comment_relationships > 0:
                # 保存修改后的XML
                tree.write(rels_path, encoding='UTF-8', xml_declaration=True)
                comments_removed = True
                print(f"已从关系文件中移除 {comment_relationships} 个评论相关关系")
        
        # 重新打包docx文件
        with zipfile.ZipFile(output_file, 'w') as zip_out:
            for root_dir, dirs, files in os.walk(extract_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, extract_dir)
                    zip_out.write(file_path, arcname)
        
        if comments_removed:
            print(f"已成功移除所有评论并保存到: {output_file}")
        else:
            print("未找到任何评论，文档已复制到输出路径")
        
        return output_file
    
    except Exception as e:
        print(f"移除评论时出错: {str(e)}")
        return None
    
    finally:
        # 清理临时文件
        try:
            shutil.rmtree(temp_dir)
        except:
            print(f"警告: 无法删除临时目录 {temp_dir}")

def main():
    # 移除评论
    output_file = remove_comments(input_file)
    
    if output_file:
        print(f"处理完成，输出文件: {output_file}")
    else:
        print("处理失败，未生成输出文件")

if __name__ == "__main__":
    main()