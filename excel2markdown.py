import os
import zipfile
import xml.dom.minidom
import shutil
from PyQt6.QtWidgets import QTreeWidgetItem
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

def update_excel_sheets_tree(excel_sheets_tree, filenames):
    excel_sheets_tree.clear()
    for filename in filenames:
        if filename.lower().endswith('.xlsx'):
            file_item = QTreeWidgetItem(excel_sheets_tree)
            file_item.setText(0, os.path.basename(filename))
            file_item.setData(0, Qt.ItemDataRole.UserRole, filename)
            try:
                with zipfile.ZipFile(filename, 'r') as zip_ref:
                    with zip_ref.open('xl/workbook.xml') as f:
                        workbook = xml.dom.minidom.parse(f)
                        sheets = workbook.getElementsByTagName('sheet')
                        for sheet in sheets:
                            sheet_name = sheet.getAttribute('name')
                            sheet_item = QTreeWidgetItem(file_item)
                            sheet_item.setText(0, sheet_name)
                            sheet_item.setData(0, Qt.ItemDataRole.UserRole, sheet_name)
                            sheet_item.setForeground(0, QColor(0, 0, 0))  # 黑色文本
            except Exception as e:
                print(f"无法读取Excel工作表: {str(e)}")
    excel_sheets_tree.expandAll()

def excel_to_markdown(file_path, output_path, sheet_name, has_header=True):
    """
    将Excel文件转换为Markdown格式的表格。

    :param file_path: Excel文件路径
    :param output_path: 输出Markdown文件路径
    :param sheet_name: 要转换的工作表名称
    :param has_header: 是否将第一行视为表头
    """
    temp_dir = 'temp_excel_data'
    
    try:
        # 解压Excel文件
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # 读取共享字符串
        strings = read_shared_strings(temp_dir)

        # 读取工作表数据
        sheet_data = read_sheet_data(temp_dir, sheet_name, strings)

        # 生成Markdown表格
        markdown_table = generate_markdown_table(sheet_data, has_header)

        # 写入Markdown文件
        with open(output_path, 'w', encoding='utf-8') as md_file:
            md_file.write(markdown_table)

        print(f"Markdown文件已生成: {output_path}")

    finally:
        # 清理临时文件
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

def read_shared_strings(temp_dir):
    """读取共享字符串"""
    strings = []
    shared_strings_path = os.path.join(temp_dir, "xl", "sharedStrings.xml")
    if os.path.exists(shared_strings_path):
        with open(shared_strings_path, 'r', encoding='utf-8') as data:
            dom = xml.dom.minidom.parse(data)
            for string in dom.getElementsByTagName('t'):
                strings.append(string.childNodes[0].nodeValue)
    return strings

def read_sheet_data(temp_dir, sheet_name, strings):
    """读取指定工作表的数据"""
    sheet_id = get_sheet_id(temp_dir, sheet_name)
    if sheet_id is None:
        raise ValueError(f"找不到工作表: {sheet_name}")

    sheet_path = os.path.join(temp_dir, f"xl/worksheets/sheet{sheet_id}.xml")
    result = []
    if os.path.exists(sheet_path):
        with open(sheet_path, 'r', encoding='utf-8') as data:
            dom = xml.dom.minidom.parse(data)
            for row in dom.getElementsByTagName('row'):
                row_data = []
                for cell in row.getElementsByTagName('c'):
                    value = get_cell_value(cell, strings)
                    row_data.append(value)
                result.append(row_data)
    return result

def get_sheet_id(temp_dir, sheet_name):
    """获取工作表ID"""
    workbook_path = os.path.join(temp_dir, "xl", "workbook.xml")
    with open(workbook_path, 'r', encoding='utf-8') as f:
        workbook = xml.dom.minidom.parse(f)
        sheets = workbook.getElementsByTagName('sheet')
        for sheet in sheets:
            if sheet.getAttribute('name') == sheet_name:
                return sheet.getAttribute('r:id').replace('rId', '')
    return None

def get_cell_value(cell, strings):
    """获取单元格的值"""
    if cell.getAttribute('t') == 's':
        shared_string_index = int(cell.getElementsByTagName('v')[0].childNodes[0].nodeValue)
        return strings[shared_string_index]
    elif cell.getElementsByTagName('v'):
        return cell.getElementsByTagName('v')[0].childNodes[0].nodeValue
    return ''

def generate_markdown_table(data, has_header):
    """生成Markdown格式的表格"""
    if not data:
        return "| 空表格 |\n|-|\n"

    markdown_table = ""
    if has_header:
        markdown_table += "|" + "|".join(data[0]) + "|\n"
        markdown_table += "|" + "|".join(["-" for _ in data[0]]) + "|\n"
        start_row = 1
    else:
        markdown_table += "|" + "|".join([f"Column {i+1}" for i in range(len(data[0]))]) + "|\n"
        markdown_table += "|" + "|".join(["-" for _ in data[0]]) + "|\n"
        start_row = 0

    for row in data[start_row:]:
        markdown_table += "|" + "|".join(row) + "|\n"

    return markdown_table