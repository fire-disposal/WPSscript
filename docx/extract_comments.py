#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
提取Word文档中的所有批注信息并以JSON格式输出
"""

import os
import json
import datetime
import re
from docx2python import docx2python

# 文件读取部分，便于修改需读取文件名
input_file = "探索知识海洋.docx"  # 请修改为实际的文件名

def extract_comments(doc_path):
    """
    使用docx2python库提取Word文档中的所有批注信息
    
    Args:
        doc_path: Word文档路径
    
    Returns:
        批注信息列表，每个批注包含id、作者、日期、内容和引用文本等信息
    """
    print(f"正在处理文件: {doc_path}")
    
    # 检查文件是否存在
    if not os.path.exists(doc_path):
        print(f"错误: 文件 '{doc_path}' 不存在!")
        return None
    
    comments = []
    
    try:
        # 使用docx2python提取文档内容和批注
        doc = docx2python(doc_path)
        
        # 获取批注信息
        comment_dict = doc.comments
        
        # 如果没有批注，尝试使用WPS格式的批注
        if not comment_dict and hasattr(doc, 'wps_comments'):
            comment_dict = doc.wps_comments
        
        # 如果仍然没有批注，尝试直接从XML提取
        if not comment_dict:
            print("尝试从文档XML中直接提取批注...")
            # 获取文档内容作为文本
            doc_text = ""
            for paragraph in doc.document_runs:
                for run in paragraph:
                    if isinstance(run, str):
                        doc_text += run
            
            # 尝试查找批注标记
            comment_pattern = r'\[批注(\d+)\](.*?)\[/批注\]'
            comment_matches = re.findall(comment_pattern, doc_text)
            
            for i, (comment_id, content) in enumerate(comment_matches):
                comment = {
                    "id": comment_id if comment_id else str(i+1),
                    "author": "未知作者",
                    "date": "",
                    "content": content.strip(),
                    "reference_text": "无法获取引用文本"
                }
                comments.append(comment)
                print(f"找到批注 #{comment_id} - 内容: {content[:30]}...")
        else:
            # 处理标准格式的批注
            for comment_id, comment_info in comment_dict.items():
                # 提取批注内容
                content = ""
                if isinstance(comment_info, dict) and 'text' in comment_info:
                    content = comment_info['text']
                elif isinstance(comment_info, str):
                    content = comment_info
                elif isinstance(comment_info, list):
                    for item in comment_info:
                        if isinstance(item, str):
                            content += item
                        elif isinstance(item, dict) and 'text' in item:
                            content += item['text']
                
                # 提取作者和日期
                author = "未知作者"
                date_str = ""
                if isinstance(comment_info, dict):
                    author = comment_info.get('author', "未知作者")
                    date_str = comment_info.get('date', "")
                
                # 尝试格式化日期
                formatted_date = ""
                if date_str:
                    try:
                        date_obj = datetime.datetime.fromisoformat(date_str.replace('Z', '+00:00'))
                        formatted_date = date_obj.strftime('%Y-%m-%d %H:%M:%S')
                    except ValueError:
                        formatted_date = date_str
                
                # 尝试获取引用文本
                reference_text = "无法获取引用文本"
                if isinstance(comment_info, dict) and 'reference_text' in comment_info:
                    reference_text = comment_info['reference_text']
                
                # 创建批注对象
                comment = {
                    "id": str(comment_id),
                    "author": author,
                    "date": formatted_date,
                    "content": content,
                    "reference_text": reference_text
                }
                
                comments.append(comment)
                print(f"找到批注 #{comment_id} - 作者: {author}, 内容: {content[:30] if content else ''}...")
        
        # 如果仍然没有找到批注，尝试检查文档中的特殊标记
        if not comments:
            print("尝试查找特殊批注标记...")
            # 遍历文档中的所有文本
            all_text = ""
            for paragraph in doc.document_runs:
                for run in paragraph:
                    if isinstance(run, str):
                        all_text += run
            
            # 查找可能的批注标记
            possible_patterns = [
                r'【批注[:：]?(.*?)】',  # 中文方括号批注
                r'（批注[:：]?(.*?)）',  # 中文圆括号批注
                r'\(批注[:：]?(.*?)\)',  # 英文圆括号批注
                r'\[批注[:：]?(.*?)\]',  # 英文方括号批注
                r'批注[:：]?(.*?)$'      # 行末批注
            ]
            
            comment_count = 0
            for pattern in possible_patterns:
                matches = re.findall(pattern, all_text)
                for i, content in enumerate(matches):
                    comment_count += 1
                    comment = {
                        "id": str(comment_count),
                        "author": "未知作者",
                        "date": "",
                        "content": content.strip(),
                        "reference_text": "无法获取引用文本"
                    }
                    comments.append(comment)
                    print(f"找到特殊批注标记 #{comment_count} - 内容: {content[:30]}...")
    
    except Exception as e:
        print(f"提取批注时出错: {str(e)}")
    
    print(f"共提取到 {len(comments)} 条批注")
    return comments

def save_comments_to_json(comments, output_path):
    """
    将批注信息保存为JSON文件
    
    Args:
        comments: 批注信息列表
        output_path: 输出文件路径
    """
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(comments, f, ensure_ascii=False, indent=4)
        print(f"批注信息已保存到: {output_path}")
    except Exception as e:
        print(f"保存JSON文件时出错: {str(e)}")

def main():
    # 构建输出文件名
    file_name, _ = os.path.splitext(input_file)
    output_file = f"{file_name}_comments.json"
    
    # 提取批注信息
    comments = extract_comments(input_file)
    
    # 保存为JSON文件，即使没有找到批注也创建空的JSON文件
    save_comments_to_json(comments, output_file)
    print(f"批注提取完成，共 {len(comments)} 条批注")

if __name__ == "__main__":
    main()