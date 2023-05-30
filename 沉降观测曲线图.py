from datetime import datetime

import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import pandas as pd

mk=['o','v','*','d','d']    #列表点类型
df=pd.read_excel(io='./test.xlsx',sheet_name='01汇总表')
dfn=df.set_index('点名')
dfn=dfn.T
a = dfn['日期'].apply(lambda x:datetime.strftime(x,'%Y年%m月%d日'))
x=dfn['日期']

#每张图片绘制五条曲线
ts=len(dfn.columns)//5
for i in range(0,ts):
    fig=plt.figure(figsize=(20,10),dpi=120)
    plt.rcParams['font.sans-serif'] = ['SimSun']
    plt.rcParams['axes.unicode_minus'] = False
    plt.title("沉降观测曲线图",fontsize=25)     #设置标题
    plt.xlabel('日期',fontsize=12)             #设置X轴名称，字体
    plt.xticks(x,a,rotation=45)               #设置X轴刻度，并用a列表代替原X显示
    plt.margins(x=0)                          #设置X轴坐标从0开始
    plt.ylabel('沉降量(mm)',fontsize=16)      #设置Y轴名称，字体
    plt.ylim(-7,1)                            #设置Y轴坐标区间
    plt.axhline(-5,color='red',linestyle='--')
    plt.annotate('报警值',xycoords='figure fraction',xy=(0.15,0.31),fontsize=15,color='red')
    for xhh in range(1,6):
        xh='DB'+str(5*i+xhh)
        xhz=dfn[xh]
        plt.plot(x,xhz,label=xh,marker=mk[xhh-1])
    plt.legend(fontsize=16)
    tpm='曲线图'+str(i+1)
    plt.savefig(tpm)

#处理最后不足5条的曲线
tssy=len(dfn.columns)%5
fig=plt.figure(figsize=(20,10),dpi=120)
plt.rcParams['font.sans-serif'] = ['SimSun']
plt.rcParams['axes.unicode_minus'] = False
plt.title("沉降观测曲线图",fontsize=25)     #设置标题
plt.xlabel('日期',fontsize=12)           #设置X轴名称，字体
plt.xticks(x,a,rotation=45)              #设置X轴刻度，并用a列表代替原X显示
plt.margins(x=0)                         #设置X轴坐标从0开始
plt.ylabel('沉降量(mm)',fontsize=16)      #设置Y轴名称，字体
plt.ylim(-7,1)                           #设置Y轴坐标区间
plt.axhline(-5,color='red',linestyle='--')
plt.annotate('报警值',xycoords='figure fraction',xy=(0.15,0.31),fontsize=15,color='red')
for yl in range(1,tssy):
    xhl='DB'+str(5*ts+yl)
    xhzl=dfn[xhl]
    plt.plot(x,xhzl,label=xhl,marker=mk[yl-1])
plt.legend(fontsize=16)
tpml='曲线图'+str(ts+1)
plt.savefig(tpml)