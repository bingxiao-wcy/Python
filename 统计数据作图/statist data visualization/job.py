#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import matplotlib.pyplot as plt
import xlsxwriter

df_row_data = pd.read_excel("数据.xlsx")
df_data = pd.DataFrame(df_row_data,columns = ['业务主管','日期','目标达成率','时间进度','昨日业绩','在库库存数量','在库库存数量金额'])

#昨日业绩图
def seven_yesterday_performance():
    plt.rcParams['font.sans-serif']=['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    plt.figure(figsize=(6,6))
    for name in df_data['业务主管'].drop_duplicates():
        df_temp = df_data.loc[df_data['业务主管'] == name]
        y = df_temp['昨日业绩'].tail(7)
        x = df_temp['日期'].tail(7) 
        plt.plot(x,y)
        plt.scatter(x,y)
        plt.title("昨日业绩", fontsize = 15)
        plt.xlabel('日期', fontsize = 13)
        plt.ylabel('业绩值', fontsize = 13)
        plt.xticks(fontsize = 10,rotation = 60)
        plt.yticks(fontsize = 10)
        for a, b in zip(x, y):
            plt.text(a, b, b, ha='center', va='bottom', fontsize=15)
        plt.savefig("seven_昨日业绩_%s.jpg" % name)

def fifteen_yesterday_performance():
    plt.rcParams['font.sans-serif']=['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    plt.figure(figsize=(8,6))
    for name in df_data['业务主管'].drop_duplicates():
        df_temp = df_data.loc[df_data['业务主管'] == name]
        y = df_temp['昨日业绩'].tail(15)
        x = df_temp['日期'].tail(15) 
        plt.plot(x,y)
        plt.scatter(x,y)
        plt.title("昨日业绩", fontsize = 15)
        plt.xlabel('日期', fontsize = 13)
        plt.ylabel('业绩值', fontsize = 13)
        plt.xticks(fontsize = 10,rotation = 60)
        plt.yticks(fontsize = 10)
        i = -1
        for a, b in zip(x, y):
            i = i + 1
            if i % 2 == 0: 
                plt.text(a,b,b,ha='center',va='bottom')
            else:
                continue
        plt.savefig("fifteen_昨日业绩_%s.jpg" % name)

#目标达成率/时间进度
def seven_TargetAchievementRatio():
    plt.rcParams['font.sans-serif']=['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    plt.figure(figsize = (6,6))
    for name in df_data['业务主管'].drop_duplicates():
        df_temp = df_data.loc[df_data['业务主管'] == name]
        y1 = df_temp['时间进度'].tail(7)
        y2 = df_temp['目标达成率'].tail(7)
        x = df_temp['日期'].tail(7) 
        plt.plot(x, y1, color='cyan', label='sinx')
        plt.plot(x, y2, 'b', label='cosx')
        plt.scatter(x,y1)
        plt.scatter(x,y2)
        plt.title("目标达成率/时间进度", fontsize = 15)
        plt.xlabel('日期', fontsize = 13)
        plt.ylabel('目标达成率/时间进度', fontsize = 13)
        plt.legend(['时间进度', '目标达成率'], loc='lower right', scatterpoints=1)
        plt.xticks(fontsize = 10,rotation = 60)
        plt.yticks(fontsize = 10)
        for a, b in zip(x, y1):
            plt.text(a,b,format(b,'.2f'),ha='left',va='bottom')
        for a, b in zip(x, y2):
            plt.text(a,b,format(b, '.2f'),ha='left',va='top')
        plt.savefig("seven_目标达成率and时间进度_%s.jpg" % name)

def fifteen_TargetAchievementRatio():
    plt.rcParams['font.sans-serif']=['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    plt.figure(figsize = (8,6))
    for name in df_data['业务主管'].drop_duplicates():
        df_temp = df_data.loc[df_data['业务主管'] == name]
        y1 = df_temp['时间进度'].tail(15)
        y2 = df_temp['目标达成率'].tail(15)
        x = df_temp['日期'].tail(15) 
        plt.plot(x, y1, color='cyan', label='sinx')
        plt.plot(x, y2, 'b', label='cosx')
        plt.scatter(x,y1)
        plt.scatter(x,y2)
        plt.title("目标达成率/时间进度", fontsize = 15)
        plt.xlabel('日期', fontsize = 13)
        plt.ylabel('目标达成率/时间进度', fontsize = 13)
        plt.legend(['时间进度', '目标达成率'], loc='lower right', scatterpoints=1)
        plt.xticks(fontsize = 10,rotation = 60)
        plt.yticks(fontsize = 10)
        i = -1
        for a, b in zip(x, y1):
            i = i + 1
            if i % 2 == 0: 
                plt.text(a,b,format(b,'.2f'),ha='left',va='bottom')
            else:
                continue
        i = -1  
        for a, b in zip(x, y2):
            i = i + 1
            if i % 2 == 0: 
                plt.text(a,b,format(b,'.2f'),ha='left',va='top')
            else:
                continue
        plt.savefig("fifteen_目标达成率and时间进度_%s.jpg" % name)


#在库库存数量和在库库存数量金额
def seven_StockQuantity():
    plt.rcParams['font.sans-serif']=['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    plt.figure(figsize = (6,6))
    for name in df_data['业务主管'].drop_duplicates():
        df_temp = df_data.loc[df_data['业务主管'] == name]
        y1 = df_temp['在库库存数量'].tail(7)
        y2 = df_temp['在库库存数量金额'].tail(7)
        x = df_temp['日期'].tail(7) 
        plt.plot(x, y1, color='cyan', label='sinx')
        plt.plot(x, y2, 'b', label='cosx')
        plt.scatter(x,y1)
        plt.scatter(x,y2)
        plt.title("在库库存数量和在库库存数量金额", fontsize = 15)
        plt.xlabel('日期', fontsize = 13)
        plt.ylabel('在库库存数量/在库库存数量金额', fontsize = 13)
        plt.legend(['在库库存数量', '在库库存数量金额'], loc='right', scatterpoints=1)
        plt.xticks(fontsize = 10,rotation = 60)
        plt.yticks(fontsize = 10)
        for a, b in zip(x, y1):
            plt.text(a,b,format(b,'.2f'),ha='center',va='bottom')
        for a, b in zip(x, y2):
            plt.text(a,b,format(b, '.2f'),ha='center',va='bottom')
        plt.savefig("seven_在库库存数量和在库库存数量金额_%s.jpg" % name)

def fifteen_StockQuantity():
    plt.rcParams['font.sans-serif']=['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    plt.figure(figsize = (8,6))
    for name in df_data['业务主管'].drop_duplicates():
        df_temp = df_data.loc[df_data['业务主管'] == name]
        y1 = df_temp['在库库存数量'].tail(15)
        y2 = df_temp['在库库存数量金额'].tail(15)
        x = df_temp['日期'].tail(15) 
        plt.plot(x, y1, color='cyan', label='sinx')
        plt.plot(x, y2, 'b', label='cosx')
        plt.scatter(x,y1)
        plt.scatter(x,y2)
        plt.title("在库库存数量和在库库存数量金额", fontsize = 15)
        plt.xlabel('日期', fontsize = 13)
        plt.ylabel('在库库存数量/在库库存数量金额', fontsize = 13)
        plt.legend(['在库库存数量', '在库库存数量金额'], loc='right', scatterpoints=1)
        plt.xticks(fontsize = 10,rotation = 60)
        plt.yticks(fontsize = 10)
        i = -1
        for a, b in zip(x, y1):
            i = i + 1
            if i % 2 == 0: 
                plt.text(a,b,format(b,'.2f'),ha='center',va='bottom')
            else:
                continue
        i = -1  
        for a, b in zip(x, y2):
            i = i + 1
            if i % 2 == 0: 
                plt.text(a,b,format(b,'.2f'),ha='center',va='bottom')
            else:
                continue
        plt.savefig("fifteen_在库库存数量和在库库存数量金额_%s.jpg" % name)

def creatExcel():
    workbook = xlsxwriter.Workbook('Excel.xlsx')
    for name in df_data['业务主管'].drop_duplicates():
        sheet = workbook.add_worksheet(name)
        #sheet.write(0,0,name)
        sheet.write(1,2,"目标")
        sheet.write(1,5,"7天趋势图")
        sheet.write(1,15,"15天趋势图")
        sheet.write(5,2,"目标达成率/时间进度")
        sheet.write(35,2,"昨日业绩")
        sheet.write(65,2,"在库库存数量和在库库存数量金额")
        sheet.insert_image('F5','seven_目标达成率and时间进度_%s.jpg' % name)
        sheet.insert_image('F35','seven_昨日业绩_%s.jpg' % name)
        sheet.insert_image('F65','seven_在库库存数量和在库库存数量金额_%s.jpg' % name)
        sheet.insert_image('P5','fifteen_目标达成率and时间进度_%s.jpg' % name)
        sheet.insert_image('P35','fifteen_昨日业绩_%s.jpg' % name)
        sheet.insert_image('P65','fifteen_在库库存数量和在库库存数量金额_%s.jpg' % name)
        merge_format = workbook.add_format({
        'align':    'center',#水平居中
        'valign':   'vcenter',#垂直居中
        })
        sheet.merge_range('A1:B5', name,merge_format)
    workbook.close()

seven_StockQuantity()
seven_TargetAchievementRatio()
seven_yesterday_performance()
fifteen_StockQuantity()
fifteen_TargetAchievementRatio()
fifteen_yesterday_performance()
creatExcel()

