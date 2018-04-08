# coding:utf8
__author__ = 'dimples'
__date__ = '2018/4/2 18:14'


from openpyxl.workbook import Workbook
from openpyxl.styles import Color, Fill, Font, Alignment
from openpyxl.cell import Cell

demo = ['业务类型', '渠道', '企业名称', '申请件数', '通过件数',
        '通过率', '申请金额', '通过金额', '免面签率', '首次进件企业数',
        '企业是否首次进件', '累计通过金额', '期末贷款余额', '企业授信余额',
        '逾期率', '不良率', '']

# 企业名称
company_list = ['企业名称', '北京中诺口腔医院', '北京倪氏威尔默口腔', '北京昌通军程医药科技发展天通苑口腔诊所',
                '北京维尔口腔医院上地口腔', '太原市恒伦口腔医院', '大众版小计',  '北京恒惠科达医疗器械', '奥齿泰（北京）商贸']
# 申请件数
applyNumber_list = ['申请件数', 1, 1, 1, 1, 1, '=SUM(E4:E8)', 1, 1, '=SUM(E10:E11)', '=SUM(E9,E12)']
# 通过件数
adoptNumber_list = ['通过件数', 0, 1, 0, 1, 1, '=SUM(F4:F8)', 1, 1, '=sum(F10:F11)', '=SUM(F9, F12)']
# 通过率
adoptRate_list = ['通过率', 0+'%', 100+'%', 0+'%', '100%', '100%', '=F9/E9', '100%', '100%', '=F12/E12' ]
# 申请金额
applyMoney_list = ['申请金额', 12697.3, 7220, 19000, 18400, 32200, '=SUM(H4:H8)', 51425, 100000,
                   '=SUM(H10:H11)', '=SUM(H9,H12)']
# 免面签率
noInterview_list = ['免面签率', '0.00%', '0.00%', '0.00%', '100.00%', '0.00%', '=1/F9']
# 企业是否首次进件
firstInfo_list = ['企业是否首次进件', '否', '否', '是', '否', '否', ' ', '否', '否']
# 累计通过余额
ljtgye_list = ['累计通过余额', 361063.78, 145236, 0, 111440, 501150, ' ', 14254300, 4395000]
# 期末贷款余额
qmdkye_list = ['期末贷款余额', 109815.78, 113963, 0, 79937, 440625, '', 5993146, 2529675]
# 企业授信余额
qysxye_list = ['企业授信余额', 390184.22, -113963, 0, 420063, -140625, '', 24006854, 27470325]
# 逾期率
yyl_list = ['逾期率', 0, 6.14, 0, 0, 0, '', 4.13, 3.95]
# 不良率

#

wb = Workbook()
ws = wb.worksheets[0]
ws.title = u"2018-03-29"
sheet = wb.active


# 合并单元格
def merge(str, str2, str3):
    ws.merge_cells(str)
    sheet[str2].alignment = Alignment(horizontal='center', vertical='center')
    sheet[str2] = str3


# 循环--列
def fill_column(lists, str):
    for i in range(len(lists)):
        sheet[str % (i+3)].alignment = Alignment(horizontal='center', vertical='center')
        sheet[str % (i+3)].value = lists[i]

merge('B2:L2', 'B2', '交易审批_日数据')
merge('M2:Q2', 'M2', '交易审批_累计数据')
merge('B4:B8', 'B4', '大众版')
merge('C4:C8', 'C4', '金服侠')
merge('B10:B11', 'B10', '行业版')
merge('B12:D12', 'B12', '大众版小计')
merge('B13:D13', 'B13', '总计')
sheet['C10'] = '金服侠'


merge('B17:F17', 'B17', '企业审批_日数据')
merge('B19:B23', 'B19', '大众版')
merge('B24:C24', 'B24', '大众版小计')
merge('B25:C25', 'B25', '行业版小计')
merge('B26:D26', 'B26', '总计')
channel_list = ['渠道', '正雅', '牙医管家', '茄子云', '茄子云-牙分期', '金服侠']
spNumber_list1 = ['审批件数', 0, 0, 0, 0, 1, '=SUM(D19:D23)']
tgNumber_list = ['通过件数', 0, 0, 0, 0, 1, '=SUM(E19:E23)']
createNumber_list = ['创建件数', 0, 0, 0, 0, 1, '=SUM(F19:F23)']

for i in range(len(company_list)):
    if i == 6:
        merge('B9:D9', 'B%d' % (i+3), '大众版小计')
    else:
        # print('company_list=', company_list[i])
        sheet["D%d" % (i+3)].value = company_list[i]

fill_column(applyNumber_list, 'E%d')
fill_column(adoptNumber_list, "F%d")
fill_column(adoptRate_list, "G%d")
fill_column(applyMoney_list, "H%d")
fill_column(noInterview_list, "J%d")
fill_column(firstInfo_list, "L%d")
fill_column(ljtgye_list, "M%d")
fill_column(qmdkye_list, "N%d")
fill_column(qysxye_list, "O%d")
fill_column(yyl_list, "P%d")
# fill_column(firstInfo_list, "Q%d")



















wb.save('work.xlsx')









