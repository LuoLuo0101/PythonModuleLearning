# -*- coding: utf-8 -*-
# @Time    : 18-4-4 上午10:19
# @Author  : luo
# @File    : 04.Styles模块.py
# @Software: PyCharm

from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.styles.colors import RED, Color

"""
Styles能够提供的功能有

1、font to set font size, color, underlining, etc.(能够设定字体的大小、颜色、下划线等属性)
2、fill to set a pattern or color gradient(能够设置单元格的填充样式或颜色渐变)
3、border to set borders on a cell(能够设置单元格的边框)
4、cell alignment(能够设置单元格的对齐)
5、protection(能够设置访问限制)

"""

# 字体一旦修改就不可更改，如果需要修改，需要重新创建一个Font，或者被复制font.copy(name="xxx")
font = Font(
    # name = '微软雅黑',
    size=15,
    bold=True,
    italic=False,
    vertAlign=None,
    underline='none',
    strike=False,
    color='FF008B00'
)

# 设置单元格的填充样式或者渐变颜色
fill = PatternFill(
    fill_type=None,
    start_color='FFFF3030',
    end_color='FF000000'
)

# 设置单元格的对齐
alignment = Alignment(
    horizontal='general',
    vertical='bottom',
    text_rotation=0,
    wrap_text=False,
    shrink_to_fit=False,
    indent=0
)
# 设置访问限制
protection = Protection(
    locked=True,
    hidden=False
)

