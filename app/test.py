import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import xlrd
import hashlib
from pandas import DataFrame
from collections import Counter
from openpyxl import load_workbook
from openpyxl.styles  import Font
bins_listy=[100,200,300,400,500]

    #應該是建立新檔案 問
    #path = 'Q.txt'
    f = open(app.config["NEW_EXCEL"] + '/new_excel_trademark.xlsx', 'w',encoding='UTF-8')

    a=pd.read_excel(app.config["EXCEL_UPLOADS"] + '/ori.xlsx', sheet_name='Sheet1')
    df = pd.DataFrame(a)

    sp = request.args.get('trademark')
    List= df[sp].tolist()

    df1=df[sp].apply(str)
    #print(df1)
    u=[]
    for i in df1:
        u.append(i)
    #print(u)

    T="".join(u)
    #print(T)

    sha = hashlib.md5(T.encode("utf-8")).hexdigest()
    print(sha)
    print(sha,file=f)  


    bin_str = ""
    for n in sha:
        bin_str += bin(int(n,16))[2:].zfill(4)

    print(bin_str)  #藏入資料
    print(bin_str,file=f)  


    recounted = Counter(List)   #統計資料出現次數
    #print(recounted)     #檢查

    b1=df.groupby(sp).size()
    #print(b1)   #檢查
    A=max(b1)
    #print("最大次數",A)  #找出最大值的點為多少

    a=np.array(List)   #將資料改成陣列  (分數)
    a1=np.unique(a)    
    #print(a1)

    b=np.array(b1)     #將資料轉成陣列  (次數)

    c=np.vstack((b,a1))   #合併變成二維陣列
    print(c) #檢查

    MAX=max(List)
    #print(MAX)

    n=-1
    p=[]

    for row in c:    
        for col in row: 
            n=n+1
            if(col==A):     #判斷次數如果是跟最大一樣
                #print(c[1][n])
                p.append(c[1][n])

    M=p[0]     #找第一個出現的最大值
    print("最高點:",M) 

    ori_file = app.config["EXCEL_UPLOADS"] + '/ori.xlsx'
    new_file = app.config["NEW_EXCEL"] + '/new_excel_trademark.xlsx'

    m=-1
    o=[]
    for i in a:
        m=m+1  
        if i < M:  
            i=i-1
            df.at[m,sp] = i
            print(m,i)
            DataFrame(df).to_excel(ori_file,sheet_name='Sheet1', index=False, header=True)
            if(i<0): 
                o.append(m)
        else:     
            if(i==M):
                df.at[m,sp] = i
                DataFrame(df).to_excel(ori_file,sheet_name='Sheet1', index=False, header=True)
            else:
                i=i+1
                df.at[m,sp] = i
                DataFrame(df).to_excel(ori_file,sheet_name='Sheet1', index=False, header=True)  
            
    print("溢位:",o)

    def union_without_repetition(list1,list2):
        result = list(set(list1) | set(list2))
        return result

    r=pd.read_excel(ori_file, sheet_name='Sheet1')
    sr= df[sp].tolist()
    ar=np.array(sr)

    bins_listx=union_without_repetition(List,sr)
    plt.hist(sr,range(0,MAX+2),align='left', edgecolor='#000000',linewidth=2)
    plt.xlabel('get')
    plt.ylabel('count')
    plt.xticks(bins_listx)
    plt.yticks(bins_listy)
    #plt.show()
    plt.savefig("s1.png")
    plt.close()

    #print(bin_str)  #藏入資料
    tc=len(bin_str)
    print("長度:",tc) #128bits

    n=-1
    t=-1
    for i in ar:
        n=n+1
        if(i==M):
            t=t+1    #計算第在第幾個資料
            if(t<tc):
                if(bin_str[t]=='1'):
                    i=i-1
                    df.at[n,sp] = i
                    DataFrame(df).to_excel(new_file,sheet_name='Sheet1', index=False, header=True)
                if(bin_str[t]=='0'):
                    DataFrame(df).to_excel(new_file,sheet_name='Sheet1', index=False, header=True)
                if(bin_str[t]==' '):
                    i=i+1
                    df.at[n,sp] = i
                    DataFrame(df).to_excel(new_file,sheet_name='Sheet1', index=False, header=True)
        else:
            df.at[n,sp]=i
            DataFrame(df).to_excel(new_file,sheet_name='Sheet1', index=False, header=True)



    def union_without_repetition(list1,list2):
        result = list(set(list1) | set(list2))
        return result

    r=pd.read_excel(new_file)
    sx= df[sp].tolist()  

    bins_listx=union_without_repetition(List,sr)
    plt.hist(sx,range(0,MAX+2),align='left', edgecolor='#000000',linewidth=2)
    plt.xlabel('get')
    plt.ylabel('count')
    plt.xticks(bins_listx)
    plt.yticks(bins_listy)
    #plt.show()
    plt.savefig("s2.png")

    wb=load_workbook(filename = 's98.xlsx')
    sheet_ranges = wb['Sheet1']

    k = pd.read_excel(new_file,nrows=0)
    L = k.columns.tolist()
    print(L) 
    r=L.index(sp)+1
    #print(r)  #第幾列

    font = Font(size=15)

    for i in o:
        #print(i) #第幾行
        i=i+2 #PYTHON讀的格子0 在EXCEL是2
        f = sheet_ranges.cell(row=i, column=r)
        f.font = font
        f.value=0
        wb.save('Q3.xlsx')