# -*- coding: utf-8 -*-
# @Time    : 18-4-3 下午5:17
# @Author  : wangge
# @Email   : ge.wang@easytransfer.cn
# @File    : 03.简单样式处理.py
# @Software: PyCharm

import openpyxl

from openpyxl.styles import Alignment, Font

# 工作薄
wb = openpyxl.load_workbook(
    filename="1.xlsx",
    read_only=False
)

# 获取一个 sheet
ws = wb["工作页04"]


# 文本对齐方式
# 居中对齐
align = Alignment(horizontal='center', vertical='center')
ws.cell(row=1, column=1).alignment = align

# 字体大小
font = Font(size=10)
ws.cell(row=1, column=1).font = font


wb.save(filename="1.xlsx")
