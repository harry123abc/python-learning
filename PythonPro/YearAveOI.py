# -*- coding: utf-8 -*-
"""
Created on Mon Dec 17 10:47:28 2018

@author: Administrator
"""

import xlrd
import xlsxwriter
import re
from datetime import date,datetime
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import math as mt
import scipy as sp
import statsmodels.api as sm
from statsmodels import regression 
from scipy.stats import norm
from functools import reduce
import os
from xlrd import xldate_as_tuple


def getAStockCodesFromExcel(file_name,sheet_name):
        try:
           file_path=os.path.join(os.getcwd(),file_name) #os.getcwd()返回当前进程的工作目录
           #stock_code = pd.read_csv(filepath_or_buffer=file_path, encoding='gbk')
           workbook1 = xlrd.open_workbook (file_path)
           worksheet1 = workbook1.sheet_names()
           print('worksheet1 is %s' %worksheet1)
        except :
            print ('Some bad happened')
        #定位到sheet1
        worksheet1 = workbook1.sheet_by_name(sheet_name)
        Code=worksheet1.col_values(1)[:]
        Name=worksheet1.col_values(0)[:]
        num_rows = worksheet1.nrows
        return Code,Name,num_rows
    
def path_change():
        os.chdir('D:\\Python 工程\\主力合约数据excel')
        
path_change()

index_Code,index_name ,num_rows1 = getAStockCodesFromExcel('指数名称列表.xlsx','Sheet1')
contract_Code,contract_name, num_rows2 = getAStockCodesFromExcel('合约列表.xlsx','Sheet1')


path = os.path.join(os.getcwd()) + '\\'
suffix1 = '指数.xlsx'


suffix2 = '_main.xlsx'


main_contract_Close = []
main_contract_TradingDate = []
#main_contract_TradingDate_1 = []
spread = []

def find_repeat(source,elmt): # The source may be a list or string.
        elmt_index = []
        s_index = 0
        e_index = len(source)
        
        while(s_index < e_index):
                try:
                    temp = source.index(elmt,s_index,e_index)
                    elmt_index = temp
                    s_index = temp + 1
                except ValueError:
                    break
 
        return elmt_index

elmt_index = []
elmt_index_max = [] 
elmt_index1= []
elmt_index1_max =[]
min_date1 = []
min_date2 =[]
max_date1 = []
max_date2 = []
max_date = []
Close = []
Date = []
TradingDate = []
main_contract_Vol = []
main_contract_OI = []
main_contract_Close = []
main_contract_TradingDate = []
suffix = 'result.xlsx'
col_close = []
col_date = []

#读取中证指数数据，包括日期和成交量信息
for m in range(num_rows1):
    workbook1 = xlrd.open_workbook (path + index_name[m] + suffix1)
    worksheet1 = workbook1.sheet_names()
    ##print('worksheet is %s' %worksheet)
    worksheet1 = workbook1.sheet_by_name(u'sheet1')
    Close.append(worksheet1.col_values(1)[:])
    TradingDate.append(worksheet1.col_values(0)[:])
    #main_contract_TradingDate[n][:] = int (main_contract_TradingDate[n][:])
    #main_contract_TradingDate[n] = int (main_contract_TradingDate[n])
    min_date1.append(min(TradingDate[m]))
    max_date1.append(max(TradingDate[m]))
    
min_date1 = max(min_date1)
max_date1 = min(max_date1)


#读取参与计算合约品种数据，包括成交量和交易日期
for n in range(num_rows2):
    workbook2 = xlrd.open_workbook (path + contract_name[n] + suffix2)
    worksheet2 = workbook2.sheet_names()
    ##print('worksheet is %s' %worksheet)
    worksheet2 = workbook2.sheet_by_name(u'VOL')
    worksheet_OI = workbook2.sheet_by_name(u'OI')
    main_contract_Vol.append(worksheet2.col_values(4)[1:])
    main_contract_OI.append(worksheet_OI.col_values(4)[1:])
    Date.append(worksheet_OI.col_values(0)[1:])
    for t in range(len(main_contract_Vol[n])):  #处理数据当中的空值
        if main_contract_Vol[n][t]=='':
           main_contract_Vol[n][t]=0.0
    worksheet3 = workbook2.sheet_by_name(u'CLOSE')       
    main_contract_Close.append(worksheet3.col_values(4)[1:])
    main_contract_TradingDate.append(worksheet2.col_values(0)[1:])
    for c in range(len(main_contract_TradingDate[n])):
        if main_contract_TradingDate[n][c] == '':
           main_contract_TradingDate[n][c] = int(main_contract_TradingDate[n][c-1])+1
    min_date2.append(min(main_contract_TradingDate[n]))
    max_date2.append(max(main_contract_TradingDate[n]))
min_date2 = max(min_date2)
max_date2 = min(max_date2)

min_date =max(min_date1,min_date2)
max_date =min(max_date1,max_date2)    
    
 
#将导入指数的长度进行裁剪
for n in range(num_rows1):
    if min_date <39903:
        elmt_index.append(find_repeat(TradingDate[n][:], 39903.0))
    else:
        elmt_index.append(find_repeat(TradingDate[n][:], min_date))
    
    if TradingDate[n][1]<39903:
        print ('%s :开始时间早于中证指数' %index_name[n])
    else:
        print ('%s :开始时间晚于中证指数' %index_name[n])
    if elmt_index[n] != []:
    #main_contract_TradingDate_1.append(main_contract_TradingDate)
       del TradingDate[n][0:elmt_index[n]]
       del Close[n][0:elmt_index[n]]
       
    if max_date >43395:
        elmt_index_max.append(find_repeat(TradingDate[n][:],43395.0))
    else:
        elmt_index_max.append(find_repeat(TradingDate[n][:],max_date))
    
    if elmt_index_max[n] != []:
        del TradingDate[n][(elmt_index_max[n]):]
        del Close[n][(elmt_index_max[n]):]    
       
    workbook = xlsxwriter.Workbook(path + index_name[n] + 'result.xlsx')  #生成表格   
    
    worksheet = workbook.add_worksheet(u'sheet1')   #在文件中创建一个名为TEST的sheet,不加名字默认为sheet1
    for i in range(len(TradingDate[n])):
               worksheet.write(i,0,TradingDate[n][i])
               worksheet.write(i,1,Close[n][i])#循环写处理后的数据生成的列表
    print('The %s has been written in the Excel' %index_name[n])
    workbook.close()
    
    
for index in range(len(elmt_index)):
    print ('%s:最早上市时间位置信息: %d'%(index_name[index], elmt_index[index-1]))    
        
        
#将导入的品种的长度进行裁剪：        
for n in range(num_rows2):
    if min_date <39903:
        elmt_index1.append(find_repeat(main_contract_TradingDate[n][:], 39903.0))
    else:
        elmt_index1.append(find_repeat(main_contract_TradingDate[n][:], min_date))
    
    if main_contract_TradingDate[n][1]<39903.0:
        print ('%s :开始时间早于中证指数' %contract_name[n])
    else:
        print ('%s :开始时间晚于中证指数' %contract_name[n])
                
    #把小于时间最小值的部分删掉
    if elmt_index1[n] != []:
    #main_contract_TradingDate_1.append(main_contract_TradingDate)
        del main_contract_TradingDate[n][0:elmt_index1[n]]
        del main_contract_Vol[n][0:elmt_index1[n]]
        del main_contract_Close[n][0:elmt_index1[n]]
       
    if max_date > 43395:
        elmt_index1_max.append(find_repeat(main_contract_TradingDate[n][:],43395.0))
    else:
        elmt_index1_max.append(find_repeat(main_contract_TradingDate[n][:],max_date))
    #把大于时间最大值的部分删掉
    if elmt_index1_max[n] != []:
        del main_contract_TradingDate[n][elmt_index1_max[n]:]
        del main_contract_Vol[n][elmt_index1_max[n]:]   
        del main_contract_Close[n][elmt_index1_max[n]:]
        
    workbook = xlsxwriter.Workbook(path + contract_name[n] + 'result.xlsx')  #生成表格
    worksheet = workbook.add_worksheet(u'sheet1')   #在文件中创建一个名为TEST的sheet,不加名字默认为sheet1
    for i in range(len(main_contract_TradingDate[n])):
               worksheet.write(i,0,main_contract_TradingDate[n][i])
               worksheet.write(i,1,main_contract_Vol[n][i])#循环写处理后的数据生成的列表
               worksheet.write(i,2,main_contract_Close[n][i])
    workbook.close()
for index in range(len(elmt_index1)):
    print ('%s:最早上市时间位置信息: %d' %(contract_name[index], elmt_index1[index-1]))

#计算个品种的3个月日均收益率
cont_ret = []
cont_Date = []
cont_Date_temp = []
count_ = []

#因为数据里面有空值，所以要把数据空值补全
for n in range(num_rows2):
   count=0
   for i in range(len(main_contract_Close[n])):
       try:
          if main_contract_Close[n][i]=='' and main_contract_Close[n][i+1]!='':
             count = count + 1
             main_contract_Close[n][i] =(main_contract_Close[n][i-1]+main_contract_Close[n][i+1])/2
             count_.append(count)
       except: 
             print('i==1 or there is no Nan')

#先把主力合约日均收益率求出来，然后再用DataFrame转成月均收益率，然后加三个月再除以3
for n in range(num_rows2):
    cont_ret.append(np.log(np.divide(main_contract_Close[n][1:],main_contract_Close[n][:-1])))

#将从Excel当中导入的数字日期转换成%y,%m,%d的格式,设置一个temp变量防止更新的时候被覆盖掉
for i in range(len(main_contract_TradingDate[n])):
    cont_Date_temp.append(datetime(*xldate_as_tuple(int(main_contract_TradingDate[n][i]),0)))#把excel中按照数字提取出来的时间，转化成时间格式
    cont_Date.append(cont_Date_temp[i].strftime("%y-%m"))

Date_OI = []
for n in range(num_rows2):
    Date_OI_ = []
    Date_OI_temp = []
    for i in range(len(Date[n])):
        #如果日期里面有空值，则加1
        if Date[n][i]=='':
           Date[n][i]=Date[n][i-1]+1
        Date_OI_temp.append(datetime(*xldate_as_tuple(int(Date[n][i]),0)))
        Date_OI_ .append(Date_OI_temp[i].strftime("%y-%m"))
    Date_OI.append(Date_OI_)
    
    
#将一个矩阵进行转置
def transpose(matrix):
    return zip(*matrix)

#计算每个月的月均持仓量
#Step1:先把OI和日期放到一个DataFrame下面，日期作为Index
OI_ = []
OI = []
Date_ = []
grouped_OI_temp = []
grouped_OI = []
Base_date = cont_Date[0]
Last_date = cont_Date[-1]
Base_position = []
Last_position = []
Base_Base = []
Data = [] 
OI_twl = []
OI_acc = []
OI_acc_ = []
length = []
for n in range(num_rows2):  
    OI.append(pd.DataFrame(main_contract_OI[n],index = (Date_OI[n]), columns = [contract_name[n]]))
    #Date_.append(pd.DataFrame(Date_OI[n], columns = ['Date']))
    #OI_.append(pd.merge(Date_[n], OI[n], left_index=True, right_index=True))
    OI[n].index.name = 'Date'
    OI[n].fillna(method='pad')
    OI[n]=OI[n].convert_objects(convert_numeric = True)
    grouped_OI_temp = OI[n].groupby('Date').mean()
    grouped_OI_temp.fillna(method='pad')
    grouped_OI.append(grouped_OI_temp)
    Base_position.append(find_repeat(list(grouped_OI[n].index),Base_date))
    Last_position.append(find_repeat(list(grouped_OI[n].index),Last_date))
    #要确认是否有品种合约的开始日是在基期前一年以后才上市
    if (Base_position[n] - 12 >= 0):
       Base_Base.append(Base_position[n] -12)
    else:
       for i in range(0,12):
           if(Base_position[n]-i) == 0:
              Base_Base.append(Base_position[n]-i)
    Data.append(grouped_OI[n].iloc[Base_Base[n]:Last_position[n]])
    Data[n] = Data[n].convert_objects(convert_numeric = True)
    length.append(len(Data[n]))
    
ll = max(length)
mm = min(length)
for n in range(num_rows2):
    OI_twl_ = []
    OI_acc_ = []
    #处理好的数据，如果开头就是基期往前一年的话，每一个季度即三个月计算一次12个月的平均值，
    #
    if Data[n].index[0] == '14-03':
       for m in range(len(Data[n])):
           if m % 3 == 0 and (m+12)<=len(Data[n]):
              OI_twl_.append(Data[n][m:m+12].mean())
    else:
        #这个地方不应该是从0：ll-len(Data[n])
       for j in range(int(np.ceil(len(Data[n])/3))):
          if (ll-len(Data[n]) + 3*j)<=12 and ll != len(Data[n]):
             OI_twl_.append(Data[n][0:ll-len(Data[n])+3*j].mean())
          else:
             OI_twl_.append(Data[n][ll-len(Data[n])+3*j-12:ll-len(Data[n])+3*j].mean())
    OI_twl.append(OI_twl_)
    for t in range(len(OI_twl[n])):
        OI_acc_.append(OI_twl[n][t].tolist()[0])
    OI_acc.append(OI_acc_)      
OI_result = pd.DataFrame(list(transpose(OI_acc)),columns = (contract_name))
OI_result.to_csv("年日均持仓量")
    #Last_postion
    #Base_position = Base_position()
#计算前十二个月的和和每十二个月的和的均值



#计算每个品种的月均收益率
cont_ret_ = pd.DataFrame(list(transpose(cont_ret)),columns=(contract_name))
cont_Date_ = pd.DataFrame(cont_Date, columns = ['Date'])
cont = pd.merge(cont_Date_,cont_ret_, left_index=True, right_index=True)
grouped_cont = cont.groupby('Date').sum()


#计算每个品种的3个月月均收益率，每三个月加一次和
ret_three = []
for n in range(len(grouped_cont)):
    if n % 3 == 0 and (n+2)< len(grouped_cont):
       ret_three.append(np.divide((grouped_cont.iloc[n]+grouped_cont.iloc[n+1]+grouped_cont.iloc[n+2]),3))



