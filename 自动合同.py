import xlrd  # 引入Excel读取模块（读excel库-xlrd、写excel库-xlwt）
from mailmerge import MailMerge  # 引用邮件处理模块

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
      
datafile_path = r'合同Test.xlsx'  # 表格位置
data = xlrd.open_workbook(datafile_path)  # 获取数据
table = data.sheet_by_name(r'合同数据（勿动）')  # 表格内工作表
ncols = table.ncols  #定义列数
nrows = table.nrows  #定义行数
template = r'模板\基础设施项目分包合同（范本）2020自动版.docx'  # 模版位置
document = MailMerge(template)

for i in range(ncols):  # 循环逐行打印
  if i == 1:  # 选择第2列，即B列
    document = MailMerge(template)
    document.merge(
      项目全称=float_to_str(table.col_values(i)[2]),  # 手填 
      工程名称=float_to_str(table.col_values(i)[3]),  # 手填
      分包模式=float_to_str(table.col_values(i)[4]),  # 手填
      合同编号=float_to_str(table.col_values(i)[5]),
      承包人名称=float_to_str(table.col_values(i)[6]),  # 选择
      承包人法定代表人=float_to_str(table.col_values(i)[7]),  # 选择
      承包人住所=float_to_str(table.col_values(i)[8]),  # 选择
      承包人增值税资质=float_to_str(table.col_values(i)[9]),  # 选择
      承包人纳税人识别号=float_to_str(table.col_values(i)[10]),  # 选择
      承包人开户行=float_to_str(table.col_values(i)[11]),  # 选择
      承包人账户=float_to_str(table.col_values(i)[12]),  # 选择
      分包人名称=float_to_str(table.col_values(i)[13]),  # 第二部分
      分包人法定代表人=float_to_str(table.col_values(i)[14]),  # 第二部分
      分包人住所=float_to_str(table.col_values(i)[15]),  # 第二部分
      分包人电子邮件地址=float_to_str(table.col_values(i)[16]),  # 第二部分
      分包人邮政编码=float_to_str(table.col_values(i)[17]),  # 第二部分
      分包人资质证书编号=float_to_str(table.col_values(i)[18]),  # 第二部分
      分包人资质类别及等级=float_to_str(table.col_values(i)[19]),  # 第二部分
      分包人复审时间及有效期=float_to_str(table.col_values(i)[20]),  # 第二部分
      分包人统一社会信用代码=float_to_str(table.col_values(i)[21]),  # 第二部分
      分包人安全生产许可证号码=float_to_str(table.col_values(i)[22]),  # 第二部分
      分包人安全生产许可证有效期=float_to_str(table.col_values(i)[23]),  # 第二部分
      分包人增值税资质=float_to_str(table.col_values(i)[24]),  # 第二部分
      分包人纳税人识别号=float_to_str(table.col_values(i)[25]),  # 第二部分
      分包人开户行=float_to_str(table.col_values(i)[26]),  # 第二部分
      分包人开户行账户=float_to_str(table.col_values(i)[27]),  # 第二部分
      分包人联系地址=float_to_str(table.col_values(i)[28]),  # 第二部分
      分包人联系电话=float_to_str(table.col_values(i)[29]),  # 第二部分
      工程地点=float_to_str(table.col_values(i)[30]),  # 手填
      省=float_to_str(table.col_values(i)[31]),  # 手填
      市=float_to_str(table.col_values(i)[32]),  # 手填
      项目规模=float_to_str(table.col_values(i)[33]),  # 手填
      业主全称=float_to_str(table.col_values(i)[34]),  # 手填
      分包内容=float_to_str(table.col_values(i)[35]),  # 手填
      承包方式=float_to_str(table.col_values(i)[36]),  # 手填
      开工日期年=float_to_str(table.col_values(i)[37]),  # 手填
      开工日期月=float_to_str(table.col_values(i)[38]),  # 手填
      开工日期日=float_to_str(table.col_values(i)[39]),  # 手填
      完工日期年=float_to_str(table.col_values(i)[40]),  # 计算
      完工日期月=float_to_str(table.col_values(i)[41]),  # 计算
      完工日期日=float_to_str(table.col_values(i)[42]),  # 计算
      工期=float_to_str(table.col_values(i)[43]),  # 手填
      质量验收标准等级=float_to_str(table.col_values(i)[44]),  # 模板
      达到质量奖项=float_to_str(table.col_values(i)[45]),  # 模板
      质量整改违约罚款=float_to_str(table.col_values(i)[46]),  # 模板
      质量返工违约罚款=float_to_str(table.col_values(i)[47]),  # 模板
      质量三检违约罚款=float_to_str(table.col_values(i)[48]),  # 模板
      合同额=float_to_str(table.col_values(i)[49]),  # 手填
      大写合同额=float_to_str(table.col_values(i)[50]),  # 计算
      不含税合同额=float_to_str(table.col_values(i)[51]),  # 计算
      不含税大写合同额=float_to_str(table.col_values(i)[52]),  # 计算
      增值税率=float_to_str(table.col_values(i)[53]),  # 计算
      税额=float_to_str(table.col_values(i)[54]),  # 手填
      大写税额=float_to_str(table.col_values(i)[55]),  # 计算
      人工费=float_to_str(table.col_values(i)[56]),  # 手填
      大写人工费=float_to_str(table.col_values(i)[57]),  # 计算
      人工费比例=float_to_str(table.col_values(i)[58]),  # 计算
      试验费用选择=float_to_str(table.col_values(i)[59]),  # 选择一、二
      水电接口提供人=float_to_str(table.col_values(i)[60]),
      水电费用承担人=float_to_str(table.col_values(i)[61]),
      水电费用选择=float_to_str(table.col_values(i)[62]),  # 选择一、二
      水费单价=float_to_str(table.col_values(i)[63]),  # 手填 水电费用选择二无
      电费单价=float_to_str(table.col_values(i)[64]),  # 手填 水电费用选择二无
      乙方驻地提供人=float_to_str(table.col_values(i)[65]),
      乙方驻地费用承担人=float_to_str(table.col_values(i)[66]),
      履保比例=float_to_str(table.col_values(i)[67]),
      履保额=float_to_str(table.col_values(i)[68]),  # 计算
      履保额1=float_to_str(table.col_values(i)[69]),  # 选择后计算
      履保额2=float_to_str(table.col_values(i)[70]),  # 选择后计算
      履保额3=float_to_str(table.col_values(i)[71]),  # 选择后计算
      履保返还时间=float_to_str(table.col_values(i)[72]),
      零星用工勾=float_to_str(table.col_values(i)[73]),
      普工单价=float_to_str(table.col_values(i)[74]),
      技工单价=float_to_str(table.col_values(i)[75]),
      综合零星用工机械勾=float_to_str(table.col_values(i)[76]),
      综合零工单价=float_to_str(table.col_values(i)[77]),
      窝工工人补偿标准=float_to_str(table.col_values(i)[78]),
      窝工管理人员补偿标准=float_to_str(table.col_values(i)[79]),
      中期结算报审日=float_to_str(table.col_values(i)[80]),
      云筑网产值单未录罚款=float_to_str(table.col_values(i)[81]),
      最终结算方式选择=float_to_str(table.col_values(i)[82]),
      中期付款比例=float_to_str(table.col_values(i)[83]),
      乙方收款办理人=float_to_str(table.col_values(i)[84]),
      收普票=float_to_str(table.col_values(i)[85]),
      收专票=float_to_str(table.col_values(i)[86]),
      农民工工资支付日期=float_to_str(table.col_values(i)[87]),
      甲方项目经理姓名=float_to_str(table.col_values(i)[88]),
      甲方项目经理身份证=float_to_str(table.col_values(i)[89]),
      甲方项目经理联系方式=float_to_str(table.col_values(i)[90]),
      乙方项目经理姓名=float_to_str(table.col_values(i)[91]),
      乙方项目经理身份证=float_to_str(table.col_values(i)[92]),
      乙方项目经理联系方式=float_to_str(table.col_values(i)[93]),
      项目经理兼任罚款=float_to_str(table.col_values(i)[94]),
      项目经理擅自更换罚款=float_to_str(table.col_values(i)[95]),
      管理人员擅自更换罚款=float_to_str(table.col_values(i)[96]),
      甲方的其它权利=float_to_str(table.col_values(i)[97]),
      甲方的其它义务=float_to_str(table.col_values(i)[98]),
      乙方冒名罚款=float_to_str(table.col_values(i)[99]),
      乙方材料罚款=float_to_str(table.col_values(i)[100]),
      其他相关专业约束条款=float_to_str(table.col_values(i)[101]),
      材料需用计划日期=float_to_str(table.col_values(i)[102]),
      材料领用人姓名=float_to_str(table.col_values(i)[103]),
      材料领用人身份证=float_to_str(table.col_values(i)[104]),
      擅用甲供材罚款=float_to_str(table.col_values(i)[105]),
      甲供机械机具=float_to_str(table.col_values(i)[106]),
      燃料动力费承担方=float_to_str(table.col_values(i)[107]),
      甲供周转材=float_to_str(table.col_values(i)[108]),
      周转材归还完好率=float_to_str(table.col_values(i)[109]),
      质保金比例=float_to_str(table.col_values(i)[110]),
      文明施工标准=float_to_str(table.col_values(i)[111]),
      文明施工扣款比例=float_to_str(table.col_values(i)[112]),
      检查罚款=float_to_str(table.col_values(i)[113]),
      主合同条款摘录=float_to_str(table.col_values(i)[114])
    )
    wordname = float_to_str(table.col_values(i)[2]) + float_to_str(table.col_values(i)[3]) + float_to_str(table.col_values(i)[4]) + "分包" + '合同.docx'  # 甲方作为文件名
    document.write(wordname)  # 创建新文件
