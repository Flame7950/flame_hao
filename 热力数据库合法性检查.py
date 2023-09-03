# 本程序用于检查济南热力集团数据库数据合法性
# 2023-6-3  By:黄浩   Tel:18660160658
# 2023-6-11  增加函数，对前期重复代码进行了合并    By:黄浩  
# 2023-9-3   修正提示文本中行错误问题 
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

# 函数查询线表填写错误
def cxxb(cxnr,cxb,shnr,ts1,ts2):
    i = 1
    sql="SELECT ID,"+cxnr+" FROM "+cxb
    for row in crsr.execute(sql):
        h=int(row[0])
        row=row[1]
        if row not in shnr:
            ms=ts1+"第"+ str(h) +"行，" + ts2 + "填写有误，请检查！"
            cwxx.append(ms)
        i+=1
# 函数查询线表未填
def cxxbwt(cxnr,cxb,shnr,ts1,ts2):
    i = 1
    sql="SELECT ID,"+cxnr+" FROM "+cxb
    for row in crsr.execute(sql):
        h=int(row[0])
        row=row[1]
        if str(row) == shnr:
            ms=ts1+"第"+ str(h) +"行，" + ts2 + "未填，请检查！"
            cwxx.append(ms)
        i+=1

sshl=('热供水','热回水')
msfs=('直埋','架空','管廊')
cz=('钢','PERT')
gwjb=('一次网','二次网')
bwfs=('聚氨酯','岩棉')
bwtgxz=('塑套钢','塑套')
tzfsw=('阀门','焊口','弯头','供热交换站','入户','三通','变径','直线点','固支','补偿器','盖堵','出地','检查井')
gj=('25','32','40','50','65','80','100','125','150','200','250','300','350','400','450','500','600','700','800','900','1000','1200','1400','1500')
yl=('1.6','2.0','2.5')
glbm=('和道','和光','和礼','和智','和茂','和安','和忠','和勇','和义','和康','起步区项目部','燃机项目部')
tcfs=('竣工','探测')
gxlx=("热水",)
gxfldm=("50602001000",)
# 管线类型
cxxb("GXLX","RSLINE",gxlx,"线表","管线类型")
# 分类代码
cxxb("FLDM","RSLINE",gxfldm,"线表","分类代码")
# 起点埋深
cxxbwt("QDMS","RSLINE","None","线表","起点埋深")
# 终点埋深
cxxbwt("ZDMS","RSLINE","None","线表","终点埋深")
# 起点高程
cxxbwt("QDGC","RSLINE","None","线表","起点高程")
# 终点高程
cxxbwt("ZDGC","RSLINE","None","线表","终点高程")
# 埋设方式
cxxb("MSFS","RSLINE",msfs,"线表","埋设方式")
# 埋设日期
cxxbwt("MSRQ","RSLINE","None","线表","埋设日期")
# 管径
cxxb("GJ","RSLINE",gj,"线表","管径")
#材质
cxxb("CZ","RSLINE",cz,"线表","材质")
# 所在位置
cxxbwt("SZWZ","RSLINE","None","线表","所在位置")
# 管网级别
cxxb("GWJB","RSLINE",gwjb,"线表","管网级别")     
# 所属回路 
cxxb("SSHL","RSLINE",sshl,"线表","所属回路") 
# 压力
cxxb("YL","RSLINE",yl,"线表","压力") 
# 保温方式
cxxb("BWFS","RSLINE",bwfs,"线表","保温方式") 
# 保温套管性质
cxxb("BWTGXZ","RSLINE",bwtgxz,"线表","保温套管性质") 
# 管材供应商
cxxbwt("GCGYS","RSLINE","None","线表","管材供应商")
# 建设日期
cxxbwt("JSRQ","RSLINE","None","线表","建设日期")
# 权属单位
cxxbwt("QSDW","RSLINE","None","线表","权属单位")
# 施工单位
cxxbwt("SGDW","RSLINE","None","线表","施工单位")
# 管理部门
cxxb("GLBM","RSLINE",glbm,"线表","管理部门") 
# 设计单位
cxxbwt("SJDW","RSLINE","None","线表","设计单位")
# 监理单位
cxxbwt("JLDW","RSLINE","None","线表","监理单位")
# 无损检测单位
cxxbwt("WSJCDW","RSLINE","None","线表","无损检测单位")
# 探测日期
cxxbwt("TCRQ","RSLINE","None","线表","探测日期")
# 探测单位
cxxbwt("TCDW","RSLINE","None","线表","探测单位")
# 探测方式
cxxb("TCFS","RSLINE",tcfs,"线表","探测方式") 





# 点表
# 管线类型
cxxb("GXLX","RSPOINT",gxlx,"点表","管线类型")
# 分类代码
cxxbwt("FLDM","RSPOINT","None","点表","分类代码")
# 特征附属物
cxxb("TZFSW","RSPOINT",tzfsw,"点表","特征附属物")
# 权属单位
cxxbwt("QSDW","RSPOINT","None","点表","权属单位")
# 建设日期
cxxbwt("JSRQ","RSPOINT","None","点表","建设日期")
# 所在位置
cxxbwt("SZWZ","RSPOINT","None","点表","所在位置")
# 探测日期
cxxbwt("TCRQ","RSPOINT","None","点表","探测日期")
# 探测单位
cxxbwt("TCDW","RSPOINT","None","点表","探测单位")
# 探测方式
cxxb("TCFS","RSPOINT",tcfs,"点表","探测方式")
# 管网级别
cxxb("GWJB","RSPOINT",gwjb,"点表","管网级别")
# 热源名称
cxxbwt("RYMC","RSPOINT","None","点表","热源名称")
# 管理部门
cxxb("GLBM","RSPOINT",glbm,"点表","管理部门")
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
cxxbwt("SJDW","RSPOINT","None","点表","设计单位")
# 类型
cxxbwt("LX","RSPOINT WHERE TZFSW='阀门' or TZFSW='补偿器'","None","点表","类型")
# 规格
cxxbwt("GG","RSPOINT WHERE TZFSW='阀门' or TZFSW='补偿器'","None","点表","规格")
# 设计压力
cxxbwt("SJYL","RSPOINT WHERE TZFSW='阀门' or TZFSW='补偿器'","None","点表","设计压力")
# 补偿量
cxxbwt("BCL","RSPOINT WHERE TZFSW='补偿器'","None","点表","补偿量")
# 曲率
cxxbwt("QL","RSPOINT WHERE TZFSW='弯头'","None","点表","曲率")
# 生产厂家
cxxbwt("SCCJ","RSPOINT WHERE TZFSW='弯头' or TZFSW='补偿器' or TZFSW='阀门' or TZFSW='三通' or TZFSW='变径'","None","点表","生产厂家")
# 施工单位
cxxbwt("SGDW","RSPOINT","None","点表","施工单位")
# 监理单位
cxxbwt("JLDW","RSPOINT","None","点表","监理单位")
# 运行主人
cxxb("YXZR","RSPOINT",glbm,"点表","运行主人")
# 井底深
cxxbwt("JDS","RSPOINT WHERE TZFSW='检修井'","None","点表","井底深")
# 保存错误信息

cwxx.sort()
str = '\n'
txt = open('错误信息.txt','w')
txt.write(str.join(cwxx))
txt.close()
