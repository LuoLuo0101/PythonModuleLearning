# -*- coding: utf-8 -*-
# @Time    : 18-3-22 下午6:15
# @Author  : luo
# @File    : t1.py
# @Software: PyCharm

from pypinyin import pinyin, lazy_pinyin, Style, slug


# ---------   示例使用   ---------
p1 = pinyin("中信")
print(p1)

p2 = pinyin('中心', heteronym=True)   # 多音字模式
print(p2)

p3 = pinyin('中心', style=Style.FIRST_LETTER)  # 设置拼音风格
print(p3)

# ---------   处理不包含拼音的字符   ---------

print("----------------------------------------")
state = slug("州州", separator=' ')
city = slug("", separator=' ')
print(state.title())
print(city.title())