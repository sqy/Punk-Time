import xlrd  # 引入Excel读取模块（读excel库-xlrd、写excel库-xlwt）
from mailmerge import MailMerge  # 引用邮件处理模块

datafile_path = '合同Test.xlsx'  # 表格位置
data = xlrd.open_workbook(datafile_path)  # 获取数据
table = data.sheet_by_name('合同数据（勿动）')  # 表格内工作表
ncols = table.ncols  #定义列数
nrows = table.nrows  #定义行数

zh = table.col_values(0)
document = MailMerge(template)
document.merge(
wordname = float_to_str(table.col_values(i)[2]) + float_to_str(table.col_values(i)[3]) + float_to_str(table.col_values(i)[4]) + "分包" + '合同.docx'  # 甲方作为文件名
            document.write(wordname)
print(zh)  # 手填
