#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
提取Word文档中的所有修订内容并以JSON格式输出
"""

import os
import json
import datetime
from docx import Document
from lxml import etree as ET

from docx.opc.constants import RELATIONSHIP_TYPE as RT

# 文件读取部分，便于修改需读取文件名
input_file = "探索知识海洋.docx"  # 请修改为实际的文件名

def extract_revisions(doc_path):
    """
    提取Word文档中的所有修订内容
    
    Args:
        doc_path: Word文档路径
    
    Returns:
        修订内容列表，每个修订包含id、作者、日期、类型、原始内容和修改后内容等信息
        同时支持修订组的概念，相关的修订会被分到同一组中
    """
    print(f"正在处理文件: {doc_path}")
    
    # 检查文件是否存在
    if not os.path.exists(doc_path):
        print(f"错误: 文件 '{doc_path}' 不存在!")
        return []
    
    revisions = []
    revision_groups = []
    current_group = None
    
    try:
        # 加载Word文档
        doc = Document(doc_path)
        
        # 获取文档的主要部分
        document_part = doc.part
        
        # 获取文档的XML内容
        xml_content = document_part.blob
        
        # 解析XML
        # 需要定义命名空间以正确解析XML
        namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'
        }
        
        # 创建XML根元素
        root = ET.fromstring(xml_content)
        
        # 查找所有修订标记
        # 1. 插入的内容 (w:ins)
        # 2. 删除的内容 (w:del)
        # 3. 格式修改 (w:rPrChange)
        # 4. 段落属性修改 (w:pPrChange)
        
        revision_id = 0
        
        # 查找插入的内容
        for ins_elem in root.findall('.//w:ins', namespaces):
            revision_id += 1
            
            # 获取作者和日期
            author = ins_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', '未知作者')
            date_str = ins_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', '')
            
            # 获取插入的文本内容
            text_content = ""
            try:
                # 直接查找所有文本元素
                for r in ins_elem.findall('.//w:t', namespaces):
                    if r.text:
                        text_content += r.text
                
                # 如果没有找到文本，尝试查找运行元素
                if not text_content:
                    for r in ins_elem.findall('.//w:r', namespaces):
                        for t in r.findall('.//w:t', namespaces):
                            if t.text:
                                text_content += t.text
            except Exception as e:
                print(f"提取插入文本时出错: {str(e)}")
            
            # 格式化日期
            formatted_date = format_date(date_str)
            
            # 尝试获取修订组ID
            group_id = ins_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', '')
            
            revision = {
                "id": str(revision_id),
                "type": "插入",
                "author": author,
                "date": formatted_date,
                "original_content": "",
                "revised_content": text_content,
                "group_id": group_id
            }
            
            # 检查是否属于现有组
            if current_group and current_group["author"] == author and abs(parse_date(current_group["date"]) - parse_date(formatted_date)) < datetime.timedelta(seconds=5):
                # 如果作者相同且时间相近（5秒内），认为属于同一组
                current_group["revisions"].append(revision)
            else:
                # 创建新组
                current_group = {
                    "group_id": len(revision_groups) + 1,
                    "author": author,
                    "date": formatted_date,
                    "revisions": [revision]
                }
                revision_groups.append(current_group)
            
            revisions.append(revision)
            # 输出信息将在分组后统一打印
        
        # 查找删除的内容
        for del_elem in root.findall('.//w:del', namespaces):
            revision_id += 1
            
            # 获取作者和日期
            author = del_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', '未知作者')
            date_str = del_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', '')
            
            # 获取删除的文本内容
            text_content = ""
            for r in del_elem.findall('.//w:delText', namespaces):
                if r.text:
                    text_content += r.text
            
            # 格式化日期
            formatted_date = format_date(date_str)
            
            # 尝试获取修订组ID
            group_id = del_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', '')
            
            revision = {
                "id": str(revision_id),
                "type": "删除",
                "author": author,
                "date": formatted_date,
                "original_content": text_content,
                "revised_content": "",
                "group_id": group_id
            }
            
            # 检查是否属于现有组
            if current_group and current_group["author"] == author and abs(parse_date(current_group["date"]) - parse_date(formatted_date)) < datetime.timedelta(seconds=5):
                # 如果作者相同且时间相近（5秒内），认为属于同一组
                current_group["revisions"].append(revision)
            else:
                # 创建新组
                current_group = {
                    "group_id": len(revision_groups) + 1,
                    "author": author,
                    "date": formatted_date,
                    "revisions": [revision]
                }
                revision_groups.append(current_group)
            
            revisions.append(revision)
            # 输出信息将在分组后统一打印
        
        # 查找格式修改
        for rpr_change in root.findall('.//w:rPrChange', namespaces):
            revision_id += 1
            
            # 获取作者和日期
            author = rpr_change.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', '未知作者')
            date_str = rpr_change.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', '')
            
            # 尝试获取相关文本
            text_content = ""
            try:
                if hasattr(rpr_change, 'getparent'):
                    parent = rpr_change.getparent()
                    if parent is not None:
                        for t in parent.findall('.//w:t', namespaces):
                            if t.text:
                                text_content += t.text
                else:
                    # 如果没有getparent方法，尝试在周围查找文本
                    for run in root.findall('.//w:r', namespaces):
                        if rpr_change in list(run):
                            for t in run.findall('.//w:t', namespaces):
                                if t.text:
                                    text_content += t.text
                            break
            except Exception as e:
                print(f"获取格式修改相关文本时出错: {str(e)}")
            
            # 格式化日期
            formatted_date = format_date(date_str)
            
            # 分析格式变化
            format_changes = []
            format_details = {}
            
            for elem in rpr_change:
                tag = elem.tag.split('}')[-1]
                format_changes.append(tag)
                
                # 提取更详细的格式信息
                if tag == 'b':  # 粗体
                    format_details['粗体'] = '开启' if elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '1') != '0' else '关闭'
                elif tag == 'i':  # 斜体
                    format_details['斜体'] = '开启' if elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '1') != '0' else '关闭'
                elif tag == 'u':  # 下划线
                    format_details['下划线'] = elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '单线')
                elif tag == 'color':  # 颜色
                    format_details['颜色'] = elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '自动')
                elif tag == 'sz':  # 字号
                    size_val = elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '')
                    if size_val:
                        format_details['字号'] = f"{int(size_val) / 2}磅"
                elif tag == 'highlight':  # 突出显示
                    format_details['突出显示'] = elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '无')
            
            # 构建格式描述
            format_description = "、".join(format_changes) if format_changes else "未知格式变化"
            
            # 添加详细信息
            if format_details:
                details_str = "; ".join([f"{k}: {v}" for k, v in format_details.items()])
                format_description += f" ({details_str})"
            
            # 尝试获取修订组ID
            group_id = rpr_change.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', '')
            
            revision = {
                "id": str(revision_id),
                "type": "格式修改",
                "author": author,
                "date": formatted_date,
                "original_content": text_content,
                "revised_content": f"格式变化: {format_description}",
                "group_id": group_id
            }
            
            # 检查是否属于现有组
            if current_group and current_group["author"] == author and abs(parse_date(current_group["date"]) - parse_date(formatted_date)) < datetime.timedelta(seconds=5):
                # 如果作者相同且时间相近（5秒内），认为属于同一组
                current_group["revisions"].append(revision)
            else:
                # 创建新组
                current_group = {
                    "group_id": len(revision_groups) + 1,
                    "author": author,
                    "date": formatted_date,
                    "revisions": [revision]
                }
                revision_groups.append(current_group)
            
            revisions.append(revision)
            # 输出信息将在分组后统一打印
        
        # 查找段落属性修改
        for ppr_change in root.findall('.//w:pPrChange', namespaces):
            revision_id += 1
            
            # 获取作者和日期
            author = ppr_change.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', '未知作者')
            date_str = ppr_change.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', '')
            
            # 尝试获取相关段落文本
            text_content = ""
            try:
                if hasattr(ppr_change, 'getparent'):
                    parent = ppr_change.getparent()
                    if parent is not None and hasattr(parent, 'getparent'):
                        parent = parent.getparent()  # 获取段落元素
                        if parent is not None:
                            for t in parent.findall('.//w:t', namespaces):
                                if t.text:
                                    text_content += t.text
                else:
                    # 如果没有getparent方法，尝试在周围查找文本
                    for para in root.findall('.//w:p', namespaces):
                        if ppr_change in list(para.findall('.//w:pPrChange', namespaces)):
                            for t in para.findall('.//w:t', namespaces):
                                if t.text:
                                    text_content += t.text
                            break
            except Exception as e:
                print(f"获取段落属性修改相关文本时出错: {str(e)}")
            
            # 格式化日期
            formatted_date = format_date(date_str)
            
            # 分析段落属性变化
            para_changes = []
            para_details = {}
            
            for elem in ppr_change:
                tag = elem.tag.split('}')[-1]
                para_changes.append(tag)
                
                # 提取更详细的段落属性信息
                if tag == 'jc':  # 对齐方式
                    para_details['对齐方式'] = elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '未知')
                elif tag == 'spacing':  # 间距
                    before = elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}before', '')
                    after = elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}after', '')
                    line = elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}line', '')
                    
                    if before:
                        para_details['段前间距'] = f"{int(before) / 20}磅"
                    if after:
                        para_details['段后间距'] = f"{int(after) / 20}磅"
                    if line:
                        para_details['行距'] = f"{int(line) / 240}倍"
                elif tag == 'ind':  # 缩进
                    left = elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}left', '')
                    right = elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}right', '')
                    firstLine = elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLine', '')
                    
                    if left:
                        para_details['左缩进'] = f"{int(left) / 20}字符"
                    if right:
                        para_details['右缩进'] = f"{int(right) / 20}字符"
                    if firstLine:
                        para_details['首行缩进'] = f"{int(firstLine) / 20}字符"
            
            # 构建段落属性描述
            para_description = "、".join(para_changes) if para_changes else "未知段落属性变化"
            
            # 添加详细信息
            if para_details:
                details_str = "; ".join([f"{k}: {v}" for k, v in para_details.items()])
                para_description += f" ({details_str})"
            
            # 尝试获取修订组ID
            group_id = ppr_change.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', '')
            
            revision = {
                "id": str(revision_id),
                "type": "段落属性修改",
                "author": author,
                "date": formatted_date,
                "original_content": text_content,
                "revised_content": f"段落属性变化: {para_description}",
                "group_id": group_id
            }
            
            # 检查是否属于现有组
            if current_group and current_group["author"] == author and abs(parse_date(current_group["date"]) - parse_date(formatted_date)) < datetime.timedelta(seconds=5):
                # 如果作者相同且时间相近（5秒内），认为属于同一组
                current_group["revisions"].append(revision)
            else:
                # 创建新组
                current_group = {
                    "group_id": len(revision_groups) + 1,
                    "author": author,
                    "date": formatted_date,
                    "revisions": [revision]
                }
                revision_groups.append(current_group)
            
            revisions.append(revision)
            # 输出信息将在分组后统一打印
        
        # 如果没有找到修订，尝试使用其他方法
        if not revisions:
            print("尝试使用替代方法查找修订...")
            
            # 尝试使用文档关系查找修订历史
            try:
                # 检查文档是否有修订历史关系
                revision_parts = []
                for rel_id, rel in document_part.rels.items():
                    if rel.reltype == RT.REVISION_LOG:
                        revision_parts.append(rel.target_part)
                
                if revision_parts:
                    print(f"找到 {len(revision_parts)} 个修订历史部分")
                    
                    for i, part in enumerate(revision_parts):
                        revision_id += 1
                        revision = {
                            "id": str(revision_id),
                            "type": "修订历史",
                            "author": "未知作者",
                            "date": "",
                            "original_content": "无法获取具体内容",
                            "revised_content": f"修订历史部分 #{i+1}"
                        }
                        revisions.append(revision)
            except Exception as e:
                print(f"查找修订历史时出错: {str(e)}")
        
    except Exception as e:
        print(f"提取修订内容时出错: {str(e)}")
    
    # 按组打印修订信息，更加直观简洁
    if revisions:
        print("\n" + "="*50)
        print(f"文档修订内容摘要 - 共 {len(revisions)} 条修订，分为 {len(revision_groups)} 个修订组")
        print("="*50)
        
        for i, group in enumerate(revision_groups):
            group_revisions = group["revisions"]
            revision_types = {}
            
            # 统计组内各类型修订的数量
            for rev in group_revisions:
                if rev["type"] not in revision_types:
                    revision_types[rev["type"]] = 0
                revision_types[rev["type"]] += 1
            
            # 格式化类型统计信息
            type_info = ", ".join([f"{count}个{type_name}" for type_name, count in revision_types.items()])
            
            # 获取组内第一个修订的内容作为示例
            sample_content = ""
            for rev in group_revisions:
                if rev["original_content"] or rev["revised_content"]:
                    content = rev["original_content"] if rev["type"] == "删除" else rev["revised_content"]
                    if content:
                        sample_content = content[:50] + ("..." if len(content) > 50 else "")
                        break
            
            print(f"\n修订组 #{i+1}")
            print(f"  作者: {group['author']}")
            print(f"  日期: {group['date']}")
            print(f"  内容: {type_info}")
            if sample_content:
                print(f"  示例: {sample_content}")
        
        print("\n" + "="*50)
    else:
        print("未找到任何修订内容")
    
    return revisions, revision_groups

def parse_date(date_str):
    """
    解析日期字符串为datetime对象
    
    Args:
        date_str: 日期字符串
    
    Returns:
        datetime对象
    """
    if not date_str:
        return datetime.datetime.now()
    
    try:
        # 处理ISO格式的日期
        return datetime.datetime.fromisoformat(date_str.replace('Z', '+00:00'))
    except ValueError:
        # 如果无法解析，则返回当前时间
        return datetime.datetime.now()

def format_date(date_str):
    """
    格式化日期字符串
    
    Args:
        date_str: ISO格式的日期字符串
    
    Returns:
        格式化后的日期字符串
    """
    if not date_str:
        return ""
    
    try:
        # 处理ISO格式的日期
        date_obj = datetime.datetime.fromisoformat(date_str.replace('Z', '+00:00'))
        return date_obj.strftime('%Y-%m-%d %H:%M:%S')
    except ValueError:
        # 如果无法解析，则返回原始字符串
        return date_str

def save_revisions_to_json(revisions, revision_groups, output_path):
    """
    将修订内容保存为JSON文件
    
    Args:
        revisions: 修订内容列表
        revision_groups: 修订组列表
        output_path: 输出文件路径
    """
    try:
        # 创建包含修订和修订组的结构
        result = {
            "revisions": revisions,
            "revision_groups": revision_groups
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=4)
        # 不在这里输出保存信息，而是在main函数中统一输出
    except Exception as e:
        print(f"保存JSON文件时出错: {str(e)}")

def main():
    # 构建输出文件名
    file_name, _ = os.path.splitext(input_file)
    output_file = f"{file_name}_revisions.json"
    
    # 提取修订内容
    revisions, revision_groups = extract_revisions(input_file)
    
    # 保存为JSON文件，即使没有找到修订也创建空的JSON文件
    save_revisions_to_json(revisions or [], revision_groups or [], output_file)
    print(f"\n修订内容已保存到: {output_file}")

if __name__ == "__main__":
    main()