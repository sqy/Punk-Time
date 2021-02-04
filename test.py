import xlrd  # 引入Excel读取模块（读excel库-xlrd、写excel库-xlwt）
from mailmerge import MailMerge  # 引用邮件处理模块

datafile_path = '合同Test.xlsx'  # 表格位置
data = xlrd.open_workbook(datafile_path)  # 获取数据
table = data.sheet_by_name('合同数据（勿动）')  # 表格内工作表
ncols = table.ncols  #定义列数
nrows = table.nrows  #定义行数

def float_to_str(float_value):   #定义浮点转字符函数
    try:
        float_value = int(float_value)    #尝试转换为整数
    except ValueError as e:    #文本或文本含数字情况赋值失败
        return float_value     #即无需转换，直接返回
    else:
        if (float_value == int(float_value)):  #如果值为整数
            return str(int(float_value))       #值转换为整数后转换为字符
        else :                                 #如果值为浮点数
            return str(float_value)            #值转换为字符

def generate_file(template):
    for i in range(ncols):  # 循环逐行打印
        if i == 1:  # 选择第2列，即B列
            document = MailMerge(template)
            document.merge(
                float_to_str(table.col_values(0)[2])=float_to_str(table.col_values(i)[2])  # 手填
            )
            wordname = float_to_str(table.col_values(i)[2]) + float_to_str(table.col_values(i)[3]) + float_to_str(table.col_values(i)[4]) + "分包" + '合同.docx'  # 甲方作为文件名
            document.write(wordname)  # 创建新文件
#def generate_file(template):
    #for i in range(2,10): # 循环逐行打印
        #document = MailMerge(template)
        #document.merge(
            #'str(table.col_values(0)[i])'=float_to_str(table.col_values(1)[i]),  # 手填
        #)
        #wordname = float_to_str(table.col_values(0)[2]) + float_to_str(table.col_values(0)[3]) + float_to_str(table.col_values(0)[4]) + "分包" + '合同.docx'  # 甲方作为文件名
        #document.write(wordname)  # 创建新文件

generate_file('基础设施项目分包合同（范本）2020自动版.docx')

#for i in range(2,10):  # 循环逐行打印
    #print(float_to_str(table.col_values(0)[i]))
    #if i == 1:
        #for j in range(2,nrows):
            #print(float_to_str(table.col_values(i)[j]))