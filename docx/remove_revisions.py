#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
移除Word文档中的所有修订内容
"""

import os
import zipfile
import shutil
import tempfile
import re
from lxml import etree as ET

# 文件读取部分，便于修改需读取文件名
input_file = "探索知识海洋.docx"  # 请修改为实际的文件名

def remove_revisions(doc_path, accept_all=True):
    """
    移除Word文档中的所有修订内容
    
    Args:
        doc_path: Word文档路径
        accept_all: 是否接受所有修订（True）或拒绝所有修订（False）
    
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
        
        # 处理document.xml文件，移除修订标记
        document_path = os.path.join(extract_dir, "word", "document.xml")
        if os.path.exists(document_path):
            # 读取XML内容
            with open(document_path, 'r', encoding='utf-8') as f:
                xml_content = f.read()
            
            # 定义命名空间
            ns = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
                'w15': 'http://schemas.microsoft.com/office/word/2012/wordml'
            }
            
            # 使用lxml解析XML
            parser = ET.XMLParser(recover=True)
            root = ET.fromstring(xml_content.encode('utf-8'), parser)
            
            # 查找所有修订标记
            ins_elements = root.xpath('.//w:ins', namespaces=ns)
            del_elements = root.xpath('.//w:del', namespaces=ns)
            rpr_change_elements = root.xpath('.//w:rPrChange', namespaces=ns)
            ppr_change_elements = root.xpath('.//w:pPrChange', namespaces=ns)
            
            # 查找表格相关的修订标记
            tbl_ins_elements = root.xpath('.//w:tblIns', namespaces=ns)  # 表格插入
            tbl_del_elements = root.xpath('.//w:tblDel', namespaces=ns)  # 表格删除
            tbl_pr_change_elements = root.xpath('.//w:tblPrChange', namespaces=ns)  # 表格属性修改
            tr_ins_elements = root.xpath('.//w:trIns', namespaces=ns)  # 表格行插入
            tr_del_elements = root.xpath('.//w:trDel', namespaces=ns)  # 表格行删除
            tc_ins_elements = root.xpath('.//w:tcIns', namespaces=ns)  # 表格单元格插入
            tc_del_elements = root.xpath('.//w:tcDel', namespaces=ns)  # 表格单元格删除
            tc_pr_change_elements = root.xpath('.//w:tcPrChange', namespaces=ns)  # 表格单元格属性修改
            
            # 统计修订数量
            total_revisions = (len(ins_elements) + len(del_elements) +
                              len(rpr_change_elements) + len(ppr_change_elements) +
                              len(tbl_ins_elements) + len(tbl_del_elements) +
                              len(tbl_pr_change_elements) + len(tr_ins_elements) +
                              len(tr_del_elements) + len(tc_ins_elements) +
                              len(tc_del_elements) + len(tc_pr_change_elements))
            print(f"找到 {total_revisions} 个修订标记")
            
            if total_revisions == 0:
                print("文档中没有修订内容，无需处理")
                # 清理临时文件
                shutil.rmtree(temp_dir)
                # 复制原始文件到输出路径
                shutil.copy2(doc_path, output_file)
                return output_file
            
            # 处理修订
            processed_count = 0
            
            if accept_all:
                # 接受所有修订
                
                # 处理插入的内容：保留内容但移除修订标记
                for ins_elem in ins_elements:
                    parent = ins_elem.getparent()
                    if parent is not None:
                        # 获取ins元素的索引
                        index = parent.index(ins_elem)
                        
                        # 将ins元素的所有子元素移动到父元素中
                        for child in list(ins_elem):
                            parent.insert(index, child)
                            index += 1
                        
                        # 移除ins元素
                        parent.remove(ins_elem)
                        processed_count += 1
                
                # 处理删除的内容：直接移除
                for del_elem in del_elements:
                    parent = del_elem.getparent()
                    if parent is not None:
                        parent.remove(del_elem)
                        processed_count += 1
                
                # 处理格式修改：应用新格式（移除修订标记）
                for rpr_change in rpr_change_elements:
                    parent = rpr_change.getparent()
                    if parent is not None:
                        parent.remove(rpr_change)
                        processed_count += 1
                
                # 处理段落属性修改：应用新属性（移除修订标记）
                for ppr_change in ppr_change_elements:
                    parent = ppr_change.getparent()
                    if parent is not None:
                        parent.remove(ppr_change)
                        processed_count += 1
                
                # 处理表格相关修订
                
                # 处理表格插入：保留表格但移除修订标记
                for tbl_ins in tbl_ins_elements:
                    parent = tbl_ins.getparent()
                    if parent is not None:
                        # 移除表格插入标记，保留表格内容
                        parent.remove(tbl_ins)
                        processed_count += 1
                
                # 处理表格删除：移除表格删除标记和表格
                for tbl_del in tbl_del_elements:
                    parent = tbl_del.getparent()
                    if parent is not None:
                        # 查找关联的表格元素
                        tbl_elem = parent.getparent()
                        if tbl_elem is not None and tbl_elem.tag == "{%s}tbl" % ns['w']:
                            # 移除整个表格
                            tbl_parent = tbl_elem.getparent()
                            if tbl_parent is not None:
                                tbl_parent.remove(tbl_elem)
                        else:
                            # 如果找不到表格，只移除删除标记
                            parent.remove(tbl_del)
                        processed_count += 1
                
                # 处理表格属性修改：应用新属性
                for tbl_pr_change in tbl_pr_change_elements:
                    parent = tbl_pr_change.getparent()
                    if parent is not None:
                        parent.remove(tbl_pr_change)
                        processed_count += 1
                
                # 处理表格行插入：保留行但移除修订标记
                for tr_ins in tr_ins_elements:
                    parent = tr_ins.getparent()
                    if parent is not None:
                        parent.remove(tr_ins)
                        processed_count += 1
                
                # 处理表格行删除：移除行删除标记和行
                for tr_del in tr_del_elements:
                    parent = tr_del.getparent()
                    if parent is not None:
                        # 查找关联的行元素
                        tr_elem = parent.getparent()
                        if tr_elem is not None and tr_elem.tag == "{%s}tr" % ns['w']:
                            # 移除整个行
                            tr_parent = tr_elem.getparent()
                            if tr_parent is not None:
                                tr_parent.remove(tr_elem)
                        else:
                            # 如果找不到行，只移除删除标记
                            parent.remove(tr_del)
                        processed_count += 1
                
                # 处理表格单元格插入：保留单元格但移除修订标记
                for tc_ins in tc_ins_elements:
                    parent = tc_ins.getparent()
                    if parent is not None:
                        parent.remove(tc_ins)
                        processed_count += 1
                
                # 处理表格单元格删除：移除单元格删除标记和单元格
                for tc_del in tc_del_elements:
                    parent = tc_del.getparent()
                    if parent is not None:
                        # 查找关联的单元格元素
                        tc_elem = parent.getparent()
                        if tc_elem is not None and tc_elem.tag == "{%s}tc" % ns['w']:
                            # 移除整个单元格
                            tc_parent = tc_elem.getparent()
                            if tc_parent is not None:
                                tc_parent.remove(tc_elem)
                        else:
                            # 如果找不到单元格，只移除删除标记
                            parent.remove(tc_del)
                        processed_count += 1
                
                # 处理表格单元格属性修改：应用新属性
                for tc_pr_change in tc_pr_change_elements:
                    parent = tc_pr_change.getparent()
                    if parent is not None:
                        parent.remove(tc_pr_change)
                        processed_count += 1
            else:
                # 拒绝所有修订
                
                # 处理插入的内容：直接移除
                for ins_elem in ins_elements:
                    parent = ins_elem.getparent()
                    if parent is not None:
                        parent.remove(ins_elem)
                        processed_count += 1
                
                # 处理删除的内容：保留被删除的内容
                for del_elem in del_elements:
                    parent = del_elem.getparent()
                    if parent is not None:
                        # 获取del元素的索引
                        index = parent.index(del_elem)
                        
                        # 将del元素的所有子元素移动到父元素中
                        for child in list(del_elem):
                            if child.tag == "{%s}delText" % ns['w']:
                                # 创建一个新的文本运行元素
                                run = ET.Element("{%s}r" % ns['w'])
                                text = ET.Element("{%s}t" % ns['w'])
                                text.text = child.text
                                run.append(text)
                                parent.insert(index, run)
                                index += 1
                        
                        # 移除del元素
                        parent.remove(del_elem)
                        processed_count += 1
                
                # 处理表格相关修订
                
                # 处理表格插入：移除表格插入标记和表格
                for tbl_ins in tbl_ins_elements:
                    parent = tbl_ins.getparent()
                    if parent is not None:
                        # 查找关联的表格元素
                        tbl_elem = parent.getparent()
                        if tbl_elem is not None and tbl_elem.tag == "{%s}tbl" % ns['w']:
                            # 移除整个表格
                            tbl_parent = tbl_elem.getparent()
                            if tbl_parent is not None:
                                tbl_parent.remove(tbl_elem)
                        else:
                            # 如果找不到表格，只移除插入标记
                            parent.remove(tbl_ins)
                        processed_count += 1
                
                # 处理表格删除：保留表格但移除删除标记
                for tbl_del in tbl_del_elements:
                    parent = tbl_del.getparent()
                    if parent is not None:
                        parent.remove(tbl_del)
                        processed_count += 1
                
                # 处理表格属性修改：保留原属性
                for tbl_pr_change in tbl_pr_change_elements:
                    parent = tbl_pr_change.getparent()
                    if parent is not None:
                        # 获取原始属性并应用
                        original_props = list(tbl_pr_change)
                        if original_props and parent.tag == "{%s}tblPr" % ns['w']:
                            # 清除当前属性
                            for child in list(parent):
                                if child != tbl_pr_change:
                                    parent.remove(child)
                            
                            # 应用原始属性
                            for prop_elem in original_props:
                                parent.append(prop_elem)
                        
                        # 移除属性修改标记
                        parent.remove(tbl_pr_change)
                        processed_count += 1
                
                # 处理表格行插入：移除行插入标记和行
                for tr_ins in tr_ins_elements:
                    parent = tr_ins.getparent()
                    if parent is not None:
                        # 查找关联的行元素
                        tr_elem = parent.getparent()
                        if tr_elem is not None and tr_elem.tag == "{%s}tr" % ns['w']:
                            # 移除整个行
                            tr_parent = tr_elem.getparent()
                            if tr_parent is not None:
                                tr_parent.remove(tr_elem)
                        else:
                            # 如果找不到行，只移除插入标记
                            parent.remove(tr_ins)
                        processed_count += 1
                
                # 处理表格行删除：保留行但移除删除标记
                for tr_del in tr_del_elements:
                    parent = tr_del.getparent()
                    if parent is not None:
                        parent.remove(tr_del)
                        processed_count += 1
                
                # 处理表格单元格插入：移除单元格插入标记和单元格
                for tc_ins in tc_ins_elements:
                    parent = tc_ins.getparent()
                    if parent is not None:
                        # 查找关联的单元格元素
                        tc_elem = parent.getparent()
                        if tc_elem is not None and tc_elem.tag == "{%s}tc" % ns['w']:
                            # 移除整个单元格
                            tc_parent = tc_elem.getparent()
                            if tc_parent is not None:
                                tc_parent.remove(tc_elem)
                        else:
                            # 如果找不到单元格，只移除插入标记
                            parent.remove(tc_ins)
                        processed_count += 1
                
                # 处理表格单元格删除：保留单元格但移除删除标记
                for tc_del in tc_del_elements:
                    parent = tc_del.getparent()
                    if parent is not None:
                        parent.remove(tc_del)
                        processed_count += 1
                
                # 处理表格单元格属性修改：保留原属性
                for tc_pr_change in tc_pr_change_elements:
                    parent = tc_pr_change.getparent()
                    if parent is not None:
                        # 获取原始属性并应用
                        original_props = list(tc_pr_change)
                        if original_props and parent.tag == "{%s}tcPr" % ns['w']:
                            # 清除当前属性
                            for child in list(parent):
                                if child != tc_pr_change:
                                    parent.remove(child)
                            
                            # 应用原始属性
                            for prop_elem in original_props:
                                parent.append(prop_elem)
                        
                        # 移除属性修改标记
                        parent.remove(tc_pr_change)
                        processed_count += 1
            
            # 1. 修改文档设置，关闭修订模式
            settings_section = root.xpath('.//w:settings', namespaces=ns)
            if settings_section:
                settings = settings_section[0]
            else:
                # 如果没有settings节点，创建一个
                body = root.xpath('.//w:body', namespaces=ns)[0]
                settings = ET.SubElement(body, "{%s}settings" % ns['w'])
            
            # 删除现有的trackRevisions元素
            track_revisions = root.xpath('.//w:trackRevisions', namespaces=ns)
            for elem in track_revisions:
                parent = elem.getparent()
                if parent is not None:
                    parent.remove(elem)
            
            # 添加关闭修订模式的设置
            track_revisions_elem = ET.Element("{%s}trackRevisions" % ns['w'])
            track_revisions_elem.set("{%s}val" % ns['w'], "false")
            settings.append(track_revisions_elem)
            
            # 2. 移除所有修订ID属性
            for elem in root.xpath('.//*[@w:id]', namespaces=ns):
                if '{%s}id' % ns['w'] in elem.attrib:
                    del elem.attrib['{%s}id' % ns['w']]
            
            # 保存修改后的XML
            xml_str = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            with open(document_path, 'wb') as f:
                f.write(xml_str)
            
            print(f"成功处理 {processed_count} 个修订标记")
            print(f"{'接受' if accept_all else '拒绝'}了所有修订内容")
            
            # 3. 修改settings.xml文件，确保修订模式关闭
            settings_path = os.path.join(extract_dir, "word", "settings.xml")
            if os.path.exists(settings_path):
                try:
                    # 读取settings.xml
                    settings_tree = ET.parse(settings_path)
                    settings_root = settings_tree.getroot()
                    
                    # 删除trackRevisions元素
                    for elem in settings_root.xpath('.//w:trackRevisions', namespaces=ns):
                        parent = elem.getparent()
                        if parent is not None:
                            parent.remove(elem)
                    
                    # 添加关闭修订模式的设置
                    track_elem = ET.Element("{%s}trackRevisions" % ns['w'])
                    track_elem.set("{%s}val" % ns['w'], "false")
                    settings_root.append(track_elem)
                    
                    # 保存修改后的settings.xml
                    settings_tree.write(settings_path, encoding='UTF-8', xml_declaration=True)
                    print("已更新settings.xml，关闭修订模式")
                except Exception as e:
                    print(f"更新settings.xml时出错: {str(e)}")
        
        # 重新打包docx文件
        with zipfile.ZipFile(output_file, 'w') as zip_out:
            for root_dir, dirs, files in os.walk(extract_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, extract_dir)
                    zip_out.write(file_path, arcname)
        
        return output_file
    
    except Exception as e:
        print(f"移除修订内容时出错: {str(e)}")
        return None
    
    finally:
        # 清理临时文件
        try:
            shutil.rmtree(temp_dir)
        except:
            print(f"警告: 无法删除临时目录 {temp_dir}")

def main():
    # 构建输出文件名
    file_name, file_ext = os.path.splitext(input_file)
    output_file = f"{file_name}（已修改）{file_ext}"
    
    # 询问用户是接受还是拒绝修订
    accept_all = True  # 默认接受所有修订
    try:
        choice = input("是否接受所有修订？(Y/N，默认Y): ").strip().upper()
        if choice == 'N':
            accept_all = False
    except:
        # 如果在非交互环境中运行，使用默认值
        pass
    
    # 移除修订内容
    output_file = remove_revisions(input_file, accept_all)
    
    if output_file:
        print(f"修订内容已{'接受' if accept_all else '拒绝'}并保存到: {output_file}")
    else:
        print("处理失败，未生成输出文件")

if __name__ == "__main__":
    main()

