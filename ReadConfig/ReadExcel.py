#!/usr/bin/env pytho    
# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
import xlrd
import json
import loadConfig
import UIWindow
from datetime import date, datetime
import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QTextEdit, QAction, QFileDialog
from PyQt5.QtGui import QIcon


def read_excel():
    # 打开文件
    workbook = xlrd.open_workbook(r'E:\H5Client\BlockHexa\config\test.xlsx')
    # 获取所有sheet
    print(workbook.sheet_names())  # [u'sheet1', u'sheet2']
    list = workbook.sheet_names()
    for book in list:
        if "|" in book:
            print("bookName -- ", book)
            sheet2 = workbook.sheet_by_name(book)
            # sheet的名称，行数，列数
            print(sheet2.name, sheet2.nrows, sheet2.ncols)
            dict = {}
            arr = []
            if "Global" in sheet2.name:
                print()
            else:
                # 需要提取第2行作为类型
                # 需要提取第3行作为key
                book = sheet2.name.split("|")[1];
                keys = sheet2.row_values(2)
                typeList = sheet2.row_values(1)

                for index in range(sheet2.nrows):
                    if index < 4:
                        continue
                    print(index)
                    rows = sheet2.row_values(index)  # 获取第四行内容
                    if rows[0] == "":
                        break
                    obj = {}
                    for keyIndex in range(len(keys)):
                        if typeList[keyIndex] == "":
                            break
                        print(rows[keyIndex])
                        if rows[keyIndex] == "":
                            continue
                        obj[keys[keyIndex]] = dealDataType(typeList[keyIndex], rows[keyIndex])
                    print(obj)
                    arr.append(obj)
                    print(arr)
                jsonData = json.dumps(arr, ensure_ascii=False)
                print(jsonData)
                with open(book + ".json", "w", encoding="utf-8")as f:
                    f.write(jsonData)


# 获取单元格内容
# print(sheet2.cell(1, 0).value.encode('utf-8'))
# print(sheet2.cell_value(1, 0).encode('utf-8'))
# print(sheet2.row(1)[0].value.encode('utf-8'))

# 获取单元格内容的数据类型
# print(sheet2.cell(1, 0).ctype)


def excelToJson(name, ):
    print("")


def dealDataType(typeName, val):
    if typeName == "int":
        return int(val)
    elif typeName == "float":
        return float(val)
    elif typeName == "string":
        return str(val)
    elif "[]" in typeName:
        print(list(val))
        return list(val)
    elif typeName == "object":
        print(val)
        return val
    else:
        print("发现未知类型请进行处理 ==================" + typeName)
        return val


if __name__ == '__main__':
    # root = Tk()
    # path = StringVar()
    #
    # Label(root, text="目标路径").grid(row=0, column=0)
    # Entry(root, textvariable=path).grid()
    # path = loadConfig.loadConfig("config.json")
    app = QApplication(sys.argv)
    ex = UIWindow.Example()
    sys.exit(app.exec_())
