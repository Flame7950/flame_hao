# 本程序用于检查济南热力集团数据库数据合法性
# 2023-6-3  By:黄浩   Tel:18660160658
import os

import pyodbc

path=os.getcwd()
a=os.listdir(path)
for i in a:
    if '.mdb' in i:
        name=i
wjmc='DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+'.\\'+name

cwxx=[]
cnxn = pyodbc.connect(wjmc)   #为什么相对路径可以绝对路径就不行
# cursor()使用该连接创建（并返回）一个游标或类游标的对象
crsr = cnxn.cursor()


sshl=('热供水','热回水')
msfs=('直埋','架空','管廊')
cz=('钢','PERT')
gwjb=('一次网','二次网')
bwfs=('聚氨酯','岩棉')
bwtgxz=('塑套钢','塑套')
tzfsw=('阀门','焊口','弯头','供热交换站','入户','三通','变径','直线点','固支','补偿器','封头','检查井')
gj=('40','50','65','80','100','125','150','200','250','300','350','400','450','500','600','700','800','900','1000','1200','1400','1500')
yl=('1.6','2.0','2.5')
glbm=('和道','和光','和礼','和智','和茂','和安','和忠','和勇','和义')
tcfs=('竣工','普查')
# 管线类型
i = 1
for row in crsr.execute("SELECT GXLX FROM RSLINE"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row != '热水':
        ms="线表第"+ str(i) +'行，管线类型填写有误，请检查！'
        cwxx.append(ms)
    i+=1
# 分类代码
i = 1
for row in crsr.execute("SELECT FLDM FROM RSLINE"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row != '50602001000':
        ms="线表第"+ str(i) +'行，分类代码填写有误，请检查！'
        cwxx.append(ms)
    i+=1
# 起点埋深
i = 1
for row in crsr.execute("SELECT QDMS FROM RSLINE"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row == 'None':
        ms="线表第"+ str(i) +'行，起点埋深未填，请检查！'
        cwxx.append(ms)
    i+=1
# 终点埋深
i = 1
for row in crsr.execute("SELECT ZDMS FROM RSLINE"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row == 'None':
        ms="线表第"+ str(i) +'行，终点埋深未填，请检查！'
        cwxx.append(ms)
    i+=1
# 起点高程
i = 1
for row in crsr.execute("SELECT QDGC FROM RSLINE"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row == 'None':
        ms="线表第"+ str(i) +'行，起点高程未填，请检查！'
        cwxx.append(ms)
    i+=1
# 终点高程
i = 1
for row in crsr.execute("SELECT ZDGC FROM RSLINE"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row == 'None':
        ms="线表第"+ str(i) +'行，终点高程未填，请检查！'
        cwxx.append(ms)
    i+=1
# 埋设方式
i = 1
for row in crsr.execute("SELECT MSFS FROM RSLINE"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row not in msfs:
        ms="线表第"+ str(i) +'行，埋设方式填写有误，请检查！'
        cwxx.append(ms)
    i+=1 
# 埋设日期
i = 1
for row in crsr.execute("SELECT MSRQ FROM RSLINE"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row == 'None':
        ms="线表第"+ str(i) +'行，埋设日期未填，请检查！'
        cwxx.append(ms)
    i+=1
# 管径
i = 1
for row in crsr.execute("SELECT GJ FROM RSLINE"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row not in gj:
        ms="线表第"+ str(i) +'行，管径填写有误，请检查！'
        cwxx.append(ms)
    i+=1
#材质
i = 1
for row in crsr.execute("SELECT CZ FROM RSLINE"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row not in cz:
        ms="线表第"+ str(i) +'行，材质填写有误，请检查！'
        cwxx.append(ms)
    i+=1
# 所在位置
i = 1
for row in crsr.execute("SELECT SZWZ FROM RSLINE"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row =='':
        ms="线表第"+ str(i) +'行，所在位置未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 管网级别        
i = 1
for row in crsr.execute("SELECT GWJB FROM RSLINE"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row not in gwjb:
        ms="线表第"+ str(i) +'行，管网级别填写有误，请检查！'
        cwxx.append(ms)
    i+=1
# 所属回路 
i = 1
for row in crsr.execute("SELECT SSHL FROM RSLINE"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row not in sshl:
        ms="线表第"+ str(i) +'行，所属回路填写有误，请检查！'
        cwxx.append(ms)
    i+=1 
# 压力
i = 1
for row in crsr.execute("SELECT YL FROM RSLINE"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row not in yl:
        ms="线表第"+ str(i) +'行，压力填写有误，请检查！'
        cwxx.append(ms)
    i+=1
# 保温方式
i = 1
for row in crsr.execute("SELECT BWFS FROM RSLINE"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row not in bwfs:
        ms="线表第"+ str(i) +'行，保温方式填写有误，请检查！'
        cwxx.append(ms)
    i+=1
# 保温套管性质
i = 1
for row in crsr.execute("SELECT BWTGXZ FROM RSLINE"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row not in bwtgxz:
        ms="线表第"+ str(i) +'行，保温套管性质填写有误，请检查！'
        cwxx.append(ms)
    i+=1
# 管材供应商
i = 1
for row in crsr.execute("SELECT GCGYS FROM RSLINE"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row =='None':
        ms="线表第"+ str(i) +'行，管材供应商未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 建设日期
i = 1
for row in crsr.execute("SELECT JSRQ FROM RSLINE"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row == 'None':
        ms="线表第"+ str(i) +'行，建设日期未填，请检查！'
        cwxx.append(ms)
    i+=1
# 权属单位
i = 1
for row in crsr.execute("SELECT QSDW FROM RSLINE"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="线表第"+ str(i) +'行，权属单位未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 施工单位
i = 1
for row in crsr.execute("SELECT SGDW FROM RSLINE"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='':
        ms="线表第"+ str(i) +'行，施工单位未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 管理部门
i = 1
for row in crsr.execute("SELECT GLBM FROM RSLINE"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row not in glbm:
        ms="线表第"+ str(i) +'行，管理部门填写有误，请检查！'
        cwxx.append(ms)
    i+=1 
# 设计单位
i = 1
for row in crsr.execute("SELECT SJDW FROM RSLINE"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="线表第"+ str(i) +'行，设计单位未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 监理单位
i = 1
for row in crsr.execute("SELECT JLDW FROM RSLINE"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="线表第"+ str(i) +'行，监理单位未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 无损检测单位
i = 1
for row in crsr.execute("SELECT WSJCDW FROM RSLINE"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="线表第"+ str(i) +'行，无损检测单位未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 探测日期
i = 1
for row in crsr.execute("SELECT TCRQ FROM RSLINE"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row == 'None':
        ms="线表第"+ str(i) +'行，探测日期未填，请检查！'
        cwxx.append(ms)
    i+=1
# 探测单位
i = 1
for row in crsr.execute("SELECT TCDW FROM RSLINE"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="线表第"+ str(i) +'行，探测单位未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 探测方式
i = 1
for row in crsr.execute("SELECT TCFS FROM RSLINE"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row not in tcfs:
        ms="线表第"+ str(i) +'行，探测方式填写有误，请检查！'
        cwxx.append(ms)
    i+=1 





# 点表
# 管线类型
i = 1
for row in crsr.execute("SELECT GXLX FROM RSPOINT"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row != '热水':
        ms="点表第"+ str(i) +'行，管线类型填写有误，请检查！'
        cwxx.append(ms)
    i+=1
# 分类代码
i = 1
for row in crsr.execute("SELECT FLDM FROM RSPOINT"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="点表第"+ str(i) +'行，分类代码未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 特征附属物
i = 1
for row in crsr.execute("SELECT TZFSW FROM RSPOINT"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row not in tzfsw:
        ms="点表第"+ str(i) +'行，特征附属物填写有误，请检查！'
        cwxx.append(ms)
    i+=1
# 权属单位
i = 1
for row in crsr.execute("SELECT QSDW FROM RSPOINT"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="点表第"+ str(i) +'行，权属单位未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 建设日期
i = 1
for row in crsr.execute("SELECT JSRQ FROM RSPOINT"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="点表第"+ str(i) +'行，建设日期未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 所在位置
i = 1
for row in crsr.execute("SELECT SZWZ FROM RSPOINT"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="点表第"+ str(i) +'行，所在位置未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 探测日期
i = 1
for row in crsr.execute("SELECT TCRQ FROM RSPOINT"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="点表第"+ str(i) +'行，探测日期未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 探测单位
i = 1
for row in crsr.execute("SELECT TCDW FROM RSPOINT"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="点表第"+ str(i) +'行，探测单位未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 探测方式
i = 1
for row in crsr.execute("SELECT TCFS FROM RSPOINT"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row not in tcfs:
        ms="点表第"+ str(i) +'行，探测方式填写有误，请检查！'
        cwxx.append(ms)
    i+=1 
# 管网级别
i = 1
for row in crsr.execute("SELECT GWJB FROM RSPOINT"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row not in gwjb:
        ms="点表第"+ str(i) +'行，管网级别填写有误，请检查！'
        cwxx.append(ms)
    i+=1
# 热源名称
i = 1
for row in crsr.execute("SELECT RYMC FROM RSPOINT"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row =='None':
        ms="点表第"+ str(i) +'行，热源名称填写有误，请检查！'
        cwxx.append(ms)
    i+=1
# 管理部门
i = 1
for row in crsr.execute("SELECT GLBM FROM RSPOINT"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row not in glbm:
        ms="点表第"+ str(i) +'行，管理部门填写有误，请检查！'
        cwxx.append(ms)
    i+=1 
# 换热方式
i = 1
for row in crsr.execute("SELECT HRFS FROM RSPOINT"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row != '水-水':
        ms="点表第"+ str(i) +'行，换热方式填写有误，请检查！'
        cwxx.append(ms)
    i+=1
# 设计单位
i = 1
for row in crsr.execute("SELECT SJDW FROM RSPOINT"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="点表第"+ str(i) +'行，设计单位未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 类型
i=1
for row in crsr.execute("SELECT LX FROM RSPOINT WHERE TZFSW='阀门' or TZFSW='补偿器'"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="点表第"+ str(i) +'行，阀门类型未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 规格
i=1
for row in crsr.execute("SELECT GG FROM RSPOINT WHERE TZFSW='阀门' or TZFSW='补偿器'"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="点表第"+ str(i) +'行，规格未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 设计压力
i=1
for row in crsr.execute("SELECT SJYL FROM RSPOINT WHERE TZFSW='阀门' or TZFSW='补偿器'"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="点表第"+ str(i) +'行，设计压力未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 补偿量
i=1
for row in crsr.execute("SELECT BCL FROM RSPOINT WHERE TZFSW='补偿器'"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="点表第"+ str(i) +'行，补偿量未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 曲率
i=1
for row in crsr.execute("SELECT QL FROM RSPOINT WHERE TZFSW='弯头'"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="点表第"+ str(i) +'行，曲率未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 生产厂家
i=1
for row in crsr.execute("SELECT SCCJ FROM RSPOINT WHERE TZFSW='弯头' or TZFSW='补偿器' or TZFSW='阀门' or TZFSW='三通' or TZFSW='变径'"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="点表第"+ str(i) +'行，生产厂家未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 施工单位
i=1
for row in crsr.execute("SELECT SGDW FROM RSPOINT"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="点表第"+ str(i) +'行，施工单位未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 监理单位
i=1
for row in crsr.execute("SELECT JLDW FROM RSPOINT"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="点表第"+ str(i) +'行，监理单位未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 运行主人
i = 1
for row in crsr.execute("SELECT GLBM FROM RSLINE"):
    row=str(row).replace('\',)','')
    row=row.replace('(\'','')
    if row not in glbm:
        ms="点表第"+ str(i) +'行，运行主人填写有误，请检查！'
        cwxx.append(ms)
    i+=1 
# 井底深
i=1
for row in crsr.execute("SELECT JDS FROM RSPOINT WHERE TZFSW='检修井'"):
    row=str(row).replace(',)','')
    row=row.replace('(','')
    if row =='None':
        ms="点表第"+ str(i) +'行，井底深未填写，请检查！'
        cwxx.append(ms)
    i+=1
# 保存错误信息
str = '\n'
txt = open('错误信息.txt','w')
txt.write(str.join(cwxx))
txt.close()
