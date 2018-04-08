# -*- coding: utf-8 -*-
# @Time    : 18-4-4 下午2:13
# @Author  : wangge
# @Email   : ge.wang@easytransfer.cn
# @File    : 05.测试用例.py
# @Software: PyCharm

import openpyxl
from datetime import datetime

from openpyxl.styles import Alignment

"""
B2-Q13：第一个表格区域

B2-L2：左边标题（交易审批—日数据）
M2-Q2：右边标题（交易审批—累计数据）
业务类型(B3)	渠道(C3)	企业名称(D3)	申请件数(E3)	通过件数(F3)	通过率(G3)	申请金额(H3)	通过金额(I3)	免面签率(J3)
首次进件企业数	(K3) 企业是否首次进件(L3)	累计通过金额(M3)	期末贷款余额(N3)	企业授信余额(O3)	逾期率(P)	不良率(Q3)
"""


class ExcelCreater(object):

    def __init__(self, file_name, suffix="xlsx"):
        self.file_name = file_name
        self.suffix = suffix
        self.wb = openpyxl.Workbook(write_only=False, iso_dates=True)
        self.ws = self.wb.active

    def change_file_name(self, file_name=None):
        self.file_name = file_name

    def change_suffix(self, suffix="xlsx"):
        self.suffix = suffix

    def set_title(self, title=None):
        self.ws.title = title

    def create_sheet(self, sheet_name=None, index=None):
        self.ws = self.wb.create_sheet(title=sheet_name, index=index)

    def merge_cells(self, range_string=None, start_row=None, start_column=None, end_row=None, end_column=None):
        self.ws.merge_cells(
            range_string=range_string,
            start_row=start_row,
            start_column=start_column,
            end_row=end_row,
            end_column=end_column
        )

    def set_value(self, cell_name=None, row=None, column=None, value=None):
        align = Alignment(horizontal='center', vertical='center')
        if cell_name is not None:
            self.ws[cell_name] = value
            self.ws[cell_name].alignment = align
        elif isinstance(row, int) and isinstance(column, int):
            self.ws.cell(row=row, column=column, value=value)
            self.ws.cell(row=row, column=column).alignment = align
        else:
            return False
        return True

    def set_center_alignment(self, cell_name=None, row=None, column=None):
        align = Alignment(horizontal='center', vertical='center')
        if cell_name is not None:
            self.ws[cell_name].alignment = align
        elif isinstance(row, int) and isinstance(column, int):
            self.ws.cell(row=row, column=column).alignment = align
        else:
            return False
        return True

    def save(self):
        file_name = "%s.%s" % (self.file_name, self.suffix)
        self.wb.save(filename=file_name)


if __name__ == '__main__':
    c_chr = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"]
    date_now = datetime.now().date()
    title = date_now.strftime("%Y-%m-%d")
    excel = ExcelCreater(file_name=title)   # 设置
    excel.set_title(title=title)    # 设置sheet标题

    excel.merge_cells(range_string="B2:L2")
    excel.merge_cells(range_string="M2:Q2")
    # 两个标题
    title_dict = {
        "B2": "交易审批—日数据",
        "M2": "交易审批—累计数据"
    }
    for k, v in title_dict.items():
        excel.set_value(cell_name=k, value=v)

    # 每一个列标题
    set_value_dict = {
        "B3": "业务类型",
        "C3": "渠道",
        "D3": "企业名称",
        "E3": "申请件数",
        "F3": "通过件数",
        "G3": "通过率",
        "H3": "申请金额",
        "I3": "通过金额",
        "J3": "免面签率",
        "K3": "首次进件企业数",
        "L3": "企业是否首次进件",
        "M3": "累计通过金额",
        "N3": "期末贷款余额",
        "O3": "企业授信余额",
        "P3": "逾期率",
        "Q3": "不良率",
    }
    for k, v in set_value_dict.items():
        excel.set_value(cell_name=k, value=v)

    enterprise_datas = [
        {
            "enterprise_name": "北京中诺口腔医院",
            "apply_piece": "1",
            "pass_piece": "0",
            "pass_rate": "=%s%d/%s%d"
        }
    ]
    r, c = 4, 4
    for item in enterprise_datas:
        enterprise_name = item.get("enterprise_name")
        excel.set_value(cell_name="%s%d" % (r, c_chr[c-1]), value=enterprise_name)
        c += 1

        apply_piece = item.get("apply_piece")
        excel.set_value(cell_name="%s%d" % (r, c_chr[c - 1]), value=apply_piece)
        c += 1

        pass_piece = item.get("pass_piece")
        excel.set_value(cell_name="%s%d" % (r, c_chr[c - 1]), value=pass_piece)
        c += 1

        pass_rate = item.get("pass_rate") % ()
        excel.set_value(cell_name="%s%d" % (r, c_chr[c - 1]), value=pass_rate)
        c += 1

    excel.save()
