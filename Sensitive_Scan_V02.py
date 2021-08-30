#!/usr/bin/env python3
# -*- coding: utf-8 -*-


# @date:   2021/7/29 20:24 
# @Version: V01
# @Revise time:   2021/8/27 16:12
# @Version: V02
# 修改执行多次结果不一致，排序问题

import os
import re
import sys
from openpyxl.workbook.workbook import Workbook
import xml.etree.ElementTree as ET


def write_to_excel(path: str, head, data):
    """
    保存为xlsx文件
    :param path:
    :param head:
    :param data:
    :return:
    """
    wb = Workbook()
    sheet = wb.active

    data.insert(0, list(head))

    for row_index, row_item in enumerate(data):

        for col_index, col_item in enumerate(row_item):
            sheet.cell(row=row_index + 1, column=col_index + 1, value=col_item)

    wb.save(path)
    print('写入成功')


def get_files(path=None):
    """
    获取文件列表
    :param path:
    :return:
    """
    path = path if path else os.getcwd()
    files_path = []

    for root, dirs, files in os.walk(path):

        for file in files:
            file_path = os.path.join(root, file)

            if os.path.splitext(file_path)[-1] in ['.xlsx', '.doc', '.docx', '.xls' ,'.xml','.so.1','.so.2','.so.5','.so.7','.so.0','.so.8','.so','.ko']:
                continue

            if file in ['QuectelWordGroups.xml','Sensitive_Scan_V02.py']:
                continue

            if file_path == sys.argv[0]:
                continue

            if os.path.isfile(file_path) and os.path.getsize(file_path):
                files_path.append(file_path)

    return files_path


def get_values_by_xml(path):
    """
    获取xml规则
    :param path:
    :return:
    """
    document = ET.parse(path)

    texts = []
    root = document.getroot()
    for group in root:
        for word in group:
            texts.append(word.text)

    return texts


def read_file(path):
    """
    读文件
    :param path:
    :return:
    """
    data = []
    with open(path, 'r', encoding='utf-8') as f:
        for line in f:
            data.append(line.strip())
    return data


def main():
    files = get_files()
    rules = get_values_by_xml('QuectelWordGroups.xml')
    head = ['file', 'line_num', 'matched', 'rule']
    data = []

    for rule in rules:
        try:
            ptn = re.compile(r'{}'.format(rule))

            for file in files:
                try:
                    count = 0
                    for line in read_file(file):
                        try:
                            count += 1

                            result = ptn.findall(line)
                            if result:
                                # print(result)
                                matched = str(result)
                                data.append([file, count, matched, rule])
                        except Exception as e:
                            print('ERROR last: ', e)
                            continue
                except Exception as e:
                    #print('FILE: ', file)
                    #print('ERROR second: ', e)
                    continue
        except Exception as e:
            #print('RULE', rule)
            #print('ERROR first: ', e)
            continue

    write_to_excel('result.xlsx', head, data)


main()
