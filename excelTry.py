#!/usr/bin/env python
# coding=utf-8

import xlrd
import json
from tkinter import *
from tkinter import filedialog


class GuiLable():
    def __init__(self):
        self.file = ""
        self.textList = []
        self.root = Tk()
        self.root.title("导表工具")
        self.root.geometry("600x300")
        self.GuiLayout()
        self.root.mainloop()

    def GuiLayout(self):
        self.frame1 = Frame(self.root)
        self.frame1.pack(padx=30, pady=10, fill="x")
        self.fileButton = Button(self.frame1, text="导入文件", padx=50, command=self.openXLS)
        self.fileButton.pack(side="left", fill="x")
        self.frame2 = Frame(self.root)
        self.frame2.pack(padx=30, pady=10, fill="x")
        self.scrollbar = Scrollbar(self.frame2)
        self.scrollbar.pack(side="right", fill="y")
        self.listbox = Listbox(self.frame2, yscrollcommand=self.scrollbar.set)
        self.listbox.pack(side="left", fill="x", expand=True)

    def openXLS(self):
        self.textList = []
        self.file = filedialog.askopenfilename()
        if self.file != "":
            fileType = self.file.split(".")[1]
        if fileType != "xlsx":
            self.textList.append("文件不合法")
            for item in self.textList:
                self.listbox.insert(END, item)
            return
        else:
            self.textList.append("开始转换文件")
            excelToJson = ExcelToJson()
            OutText = self.file + "导出表.txt"
            result = excelToJson.ReadToJson(self.file, OutText, self.textList)
            if result:
                self.textList.append("JSON文件导出成功，位置为：" + OutText)
            elif not result:
                self.textList.append("JSON文件导出失败")
            for item in self.textList:
                self.listbox.insert(END, item)


class ExcelToJson(object):
    def __init__(self):
        pass

    def ReadToJson(self, path, name, text):
        x1 = xlrd.open_workbook(path)
        sheetNames = x1.sheet_names()
        sheet1 = x1.sheet_by_name(sheetNames[0])
        lkey = []  # key值
        # 打开并清空原始文件
        fp = open(name, 'w')
        fp.truncate()
        key_check = False
        end_check = False
        row_count = sheet1.nrows
        if row_count > 2:
            for i in range(0, row_count):
                line = sheet1.row_values(i)
                # 开头判断  id值不为空
                if not key_check:
                    if line[0] != "":
                        lkey = line
                        key_check = True
                        continue
                    else:
                        continue
                # 结尾判断
                if key_check:
                    if line[0] == "":
                        end_check = True
                        break
                print(line)
                # 每一行初始字典清空
                dict = {}
                for j in range(sheet1.ncols):
                    key = lkey[j]
                    if lkey[j] == "":
                        break
                    if line[j] != "":
                        dict[key] = line[j]
                # 每读完一行  生成一组json记录
                dstr = json.dumps(dict)
                print(dstr)
                text.append(dstr)
                fp.write(dstr + "\n")
            return True
        else:
            return False


if __name__ == '__main__':
    GuiLable()