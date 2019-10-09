import pandas as pd
import numpy as np

#导入三张表格（人员清单，就诊清单，理赔清单）
person = pd.read_excel('/Users/Durantfang/Downloads/0人员清单.xlsx')
hospital = pd.read_excel('/Users/Durantfang/Downloads/2就诊清单.xlsx')
settlement = pd.read_excel('/Users/Durantfang/Downloads/理赔清单.xlsx')

#删除空白字段
for sheet in [person, hospital, settlement]:
    for i in sheet.columns:
        if sheet[i].count() == 0:
            sheet.drop(labels = i, axis =1, inplace = True)
           
 #删除清单无用字段，添加后续分析需要字段
del_person_col = ['保单生效日','保单终止日','险种代码','险种类型','子分支机构名称2','分计划名称1','被保险人保单生效日','折合人数',
                  '被保险人保单终止日','被保险人状态','保单经过天数','被保险人保单经过天数']
person.drop(labels =del_person_col,axis = 1, inplace = True)
person['实际保费'] = person['分险种保费-期初'] + person['加保保费'] - person['减保保费']

del_hospital_col = ['就诊起始日期','就诊终止日期','就诊医院','保单生效日','保单终止日',
                 '子分支机构名称2','分计划名称1']
hospital.drop(labels =del_hospital_col,axis = 1, inplace = True)

del_settlement_col = ['险种代码', '险种名称', '险种类型', '保单生效日', '保单终止日', '子分支机构名称2',
       '分计划名称1']
settlement.drop(labels =del_settlement_col,axis = 1, inplace = True)

settlement['理赔时效'] = settlement['结案日期'] - settlement['受理日期']

#导出CSV文件
document = [person, hospital, settlement]
target_location = ['person', 'hospital', 'settlement']
for a,b in zip(document, target_location):  
    a.to_csv('/Users/Durantfang/Desktop/Data/%s'%b, encoding = 'utf-8', index = False)
