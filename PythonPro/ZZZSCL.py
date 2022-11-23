# -*- coding: utf-8 -*-
"""
Created on Wed Oct 31 14:01:32 2018

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
        os.chdir('D:\\Python 工程\\主力合约数据excel\\')
        
path_change()

index_Code,index_name ,num_rows1 = getAStockCodesFromExcel('指数名称列表.xlsx','Sheet1')
contract_Code,contract_name, num_rows2 = getAStockCodesFromExcel('合约列表（较少品种版）.xlsx','Sheet1')


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
TradingDate = []
main_contract_Vol = []
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
    main_contract_Vol.append(worksheet2.col_values(4)[1:])
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

#将一个矩阵进行转置
def transpose(matrix):
    return zip(*matrix)

#计算每个品种的月均收益率
cont_ret_ = pd.DataFrame(list(transpose(cont_ret)),columns=(contract_name))
cont_Date_ = pd.DataFrame(cont_Date,columns=['Date'])
cont = pd.merge(cont_Date_,cont_ret_, left_index=True, right_index=True)
grouped_cont = cont.groupby('Date').sum()

#计算每个品种的3个月月均收益率，每三个月加一次和
ret_three = []
for n in range(len(grouped_cont)):
    if n % 3 == 0 and (n+2)< len(grouped_cont):
       ret_three.append(np.divide((grouped_cont.iloc[n]+grouped_cont.iloc[n+1]+grouped_cont.iloc[n+2]),3))

#对每个品种按照3个月收益率（ret_three）进行排序，并将排序结果写在每个品种Dataframe结构的后面
#冒泡排序法
def bubblesort(arr):
    length= len(arr)
    for i in range(length-1):
        for j in range(length-1):
            if arr[j] > arr[j+1]:
                temp = arr[j+1]
                arr[j+1]=arr[j]
                arr[j]=temp
    return arr

#以下函数是用来找一个数在上一个list或什么的位置的,第一个参数是目标，第二个参数是要找的位置
def findposition_les (matrix1,matrix2):
    position = []
    for i in range(int(np.ceil(len(matrix1)/2))):
        for j in range(len(matrix2)):
            if matrix1[i] == matrix2[j]:
               position.append(j)
    return position
                
def findposition_lar (matrix1,matrix2):
    position = []
    for i in range(int(np.ceil(len(matrix1)/2)),len(matrix1)):
        for j in range(len(matrix2)):
            if matrix1[i] == matrix2[j]:
               position.append(j)
    return position

#找到本次调权过程中，低于中位数的品种在上次调权过程当中的位置。
def int_ceil(num):
    num1 = int(np.ceil(len(num))/2)
    return num1
pos_les = []
pos_lar = []
sort_result1 = []
sort_result2 = []   
sort_result_temp =[]
sort_result_ = []
median = []
w_intial = []
sort_result_less_med = []
sort_result_lar_med = []
for n in range(len(ret_three)):
        sort_result1.append(pd.DataFrame(ret_three[n][0:int_ceil(ret_three[n])].sort_values(),columns = ['ret']))
        sort_result2.append(pd.DataFrame(ret_three[n][int_ceil(ret_three[n]):].sort_values(ascending=False),columns = ['ret']))
        sort_result_temp.append(pd.concat([sort_result1[n],sort_result2[n]]))
        w = np.ones(len(contract_name))/len(contract_name)
        #sort_result_temp.append(pd.DataFrame(sort_result[n],columns=['ret']))
        w = pd.DataFrame(w,columns=['weight'],index=sort_result_temp[n].index)
        w_intial.append(w)
        sort_result_.append(pd.merge(sort_result_temp[n],w,left_index=True, right_index=True))
        
for n in range(1,len(ret_three)):
       pos_les.append(findposition_les(list(sort_result_[n].index),list(sort_result_[n-1].index))) 
       pos_lar.append(findposition_lar(list(sort_result_[n].index),list(sort_result_[n-1].index)))
        #median.append(sort_result[n].median())
        #sort_result_less_med.append((sort_result_[n][sort_result_[n]['ret'][:]<median[n]]))
        #sort_result_lar_med.append((sort_result_[n][sort_result_[n]['ret'][:]>median[n]]))
        #sort_result_less_med.append((sort_result_[n][sort_result_[n][:]>median[n]]).dropna())
        #sort_result_less_med[n]['ret'].fillna('null')
        #cols=[x for i,x in enumerate(sort_result_less_med[n].columns) if sort_result_less_med[n].iat[0,i]=='null']
        #sort_result_less_med[n]=sort_result_less_med[n].drop(cols,axis=1)
        #sort_resort_less_med.append(sort_result_less_med[n].dropna())
        #index = np.linspace(1,len(sort_result_less_med[n]),len(sort_result_less_med[n]))       #产生1:15个数
        #index_ = pd.Series(index)

#计算每一期的权重
p = 0.5
w_record = []
mm = int(np.ceil(len(sort_result_[0])/2))
N = len(sort_result_[0])
num = []
w_temp_ = []
den1_ = []
num_ = []
rank = []
for n in range(1,len(ret_three)):#每一期算一个权重
    den1 = 0
    w_temp = 0
    num = []
    rank = []
    for i in range(1,int_ceil(sort_result_[n])+1):
        den1 = den1 + (int_ceil(sort_result_[n]) - i + 1)
        w_temp = w_temp + sort_result_[n-1]['weight'][pos_les[n-1][i-1]]# w_temp 是一堆权重的和应该是一个数
    w_temp_.append(w_temp)
    print ('第 %d 期上半部分 %f' %(n,w_temp_[n-1]))
    for i in range(1,int_ceil(sort_result_[n])+1):
           rank.append((int_ceil(sort_result_[n]) - i + 1))
           num.append(p * w_temp_[n-1] * rank[i-1])
           #sort_result_[n]['weight'][i-1] = sort_result_[n-1]['weight'][pos_les[n-1][i-1]]-num[i-1]/den1
    num_.append(num)
    
    for i in range(1,int_ceil(sort_result_[n])+1):
        sort_result_[n]['weight'][i-1] = sort_result_[n-1]['weight'][pos_les[n-1][i-1]]-num_[n-1][i-1]/den1
        
    
           #wadj = wadj + sort_result_[n]['weight'][i]
           #w_record.append(w[i])
    
for n in range(1,len(ret_three)):
    den2 = 0
    wadj = 0
    for i in range(1,int_ceil(sort_result_[n])+1):
        den2 = den2 + (N-mm-i+1)
        wadj = wadj + sort_result_[n]['weight'][i-1]
    print('第 %d 期上半部分 %f' %(n,wadj))
    for i in range(int_ceil(sort_result_[n]),len(sort_result_[n])):
            sort_result_[n]['weight'][i] = sort_result_[n-1]['weight'][pos_lar[n-1][i-13]] + (w_temp_[n-1]-wadj)*(N-mm-i+12+1)/den2
        

#结果验证
for n in range(len(ret_three)):
    print('第 %d 期权重：' %n)
    print(sort_result_[n]['weight'].sum())
    for i in range(len(sort_result_[n])):
        if sort_result_[n]['weight'][i]<0:
            print('%s is wrong' %contract_name[i])




#还得计算调权上一年度的日均持仓规模
#伪代码：
'''
每个品种读取对应日期的持仓量数据
'''



for m in range(num_rows1):
    close,date,suibian= getAStockCodesFromExcel(index_name[m]+suffix,'sheet1')
    col_close.append(close)
    col_date.append(date)
    ##print('worksheet is %s' %worksheet)

#遍历sheet1中所有单元格ceil

#num_rows = worksheet3.nrows
#num_cols = worksheet1.ncols
#col_close = worksheet1.col_values(1)[:]
#col_date  = worksheet1.col_values(1)[:]
col_close_bench = col_close[0]
col_date_bench = col_date[0]

#print (col_close[:]) 

ret = []
Mrk_Rf = []
alpha = []
beta = []
def linreg(x,y):
    x = sm.add_constant(x)
    model = regression.linear_model.OLS(y, x).fit()
    x = x[:, 1]
    return model.params[0], model.params[1]

file_name='合约列表.xlsx'
file_path=os.path.join(os.getcwd(),file_name)
workbook1 = xlrd.open_workbook (file_path)
worksheet1 = workbook1.sheet_names()
#定位到sheet1
worksheet1 = workbook1.sheet_by_name(u'Sheet1')
#遍历sheet1中所有行row
contract_name = worksheet1.col_values(0)[1:]
path =  os.getcwd()
suffix = 'result.xlsx'
num_row1 = worksheet1.nrows

Mrk_Rf.append(list(map(mt.log,np.divide(np.array(col_close_bench[1:]),np.array(col_close_bench[:-1])))))

def add(x , y):
    return x + y

#计算收益率，夏普比率，以及alpha,beta值等信息
for mm in range(0,len(col_close)):
    ret.append(list(map(mt.log,np.divide(np.array(col_close[mm][1:]),np.array(col_close[mm][:-1])))))

    #print (data_return)
    plt.plot(ret[mm])
    narray=np.array(ret[mm])
    sharpration=mt.sqrt(252)*(narray).mean()/narray.std()
    #Sp = pd.DataFrame(sharpration ,  col_date, columns=['A','B'])
    print('SharpRation is %f' %sharpration) 

    condifence_level=0.99
    z=norm.ppf(condifence_level)
    np.mean(ret[mm])
    n_days=10
    VaR=z *np.std(ret)*np.sqrt(n_days)


    alpha1 ,beta1 =linreg(Mrk_Rf[0][:],ret[mm][:])
    alpha.append(alpha1)
    beta.append(beta1)
    print ("alpha: ", str(alpha))
    print ("beta: ", str(beta))

    X2 = np.linspace(min(Mrk_Rf[0][:]), max(Mrk_Rf[0][:]), 100)
    Y = X2*beta[mm]+alpha[mm]
    '''
    plt.figure(figsize = (10,7))
    plt.scatter(Mrk_Rf[0][:], ret[mm][:] , alpha=0.3)
    plt.xlabel("Standerd Commodity Index Daily Return %d" %mm)
    plt.ylabel("Enhanced Commodity Index Daily Retrun %d" %mm)

    plt.plot(X2, Y, 'r' , alpha = 0.9)
    plt.show()
    '''


    dollar_vol = []
    lamda = []
    lamda_all =[]
    for n in range(num_rows1-1):
        workbook = xlrd.open_workbook (path+'\\' + contract_name[n] + suffix)
        worksheet = workbook.sheet_names()
        ##print('worksheet is %s' %worksheet)
        worksheet = workbook.sheet_by_name(u'sheet1')
        #print ('此次计算的品种 %s:' %contract_name[n])
        dollar_vol = (worksheet.col_values(1)[:-1])
        date_1=(worksheet.col_values(0)[:-1])
        tt=pd.DataFrame(ret[mm][:],np.array(date_1),columns=['ret'])
        tt2=pd.DataFrame(dollar_vol,np.array(date_1),columns=['dollar_vol'])
        tt2.fillna(0)
        ff=pd.DataFrame(Mrk_Rf[0][:],np.array(date_1),columns=['Mrk_Rf'])
        tt3=pd.merge(tt,tt2,left_index=True,right_index=True)
        final=pd.merge(tt3,ff,left_index=True,right_index=True)
        y=final.ret
        x1=final.Mrk_Rf
        x2=np.sign(np.array(final.ret-final.Mrk_Rf))*np.array(final.dollar_vol)
        x3=[x1,x2]
        m=np.size(x3)
        x=np.reshape(x3,[int(m/2),2])
        x=sm.add_constant(x)
        results=sm.OLS(y,x).fit()
        #print ('%s' %contract_name[n])
        #print (results.params)
        lamda.append(results.params[2])

    lamda_all.append(reduce(add, lamda))
    print ('%s指数流动性 %f' %(index_name[mm],lamda_all[0]))
    with open(path+'\\'+'result.txt',"r+") as f:
        f.read()
        f.write(index_name[mm] + ':'+ str(lamda_all[0])+ '\n')
    f.close()
    


    
  
'''
for curr_row in range(num_rows):
    row = worksheet1.row_values(curr_row)
print('row%s is %s' %(curr_row,row))
#遍历sheet1中所有列col
'''


#遍历sheet1中所有单元格ceil
'''
for rown in range(num_rows):
    for coln in range(num_cols):
        ceil = worksheet1.ceil_value(rown,coln)
    print (ceil)
'''



