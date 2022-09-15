from app import app
from flask import Flask, render_template, request, redirect, send_from_directory, abort, flash, session, Blueprint
import os
from os import listdir
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.styles import Font
from werkzeug.utils import secure_filename
#修改excel
import pandas as pd 
from pandas import DataFrame 
import numpy as np 
import matplotlib.pyplot as plt 
from collections import Counter
import xlrd
import hashlib
#用sqlite上傳檔案
#from . import db
#用MongoDB上傳檔案
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles  import Font
#import pymongo
#from pymongo import MongoClient #使用mongodb
#import certifi #為了解決連線到SSL的問題
import pathlib #分割副檔名
import csv
import shutil #移動檔案 覆蓋原檔案
#import pandas as pd 重複了
import json
#set FLASK_ENV=development

@app.route("/") #主頁
def index():
    return render_template("public/index.html")

def allowed_excel(filename):

    #上傳的文件要有副檔名
    if not "." in filename:
        return False

    #將.從副檔名中分割出來
    ext = filename.rsplit(".", 1)[1]

    #確認副檔名和ALLOWED_EXCEL_EXTENSIONS中的一樣
    if ext.upper() in app.config["ALLOWED_EXCEL_EXTENSIONS"]:
        return True
    else:
        return False                

app.config["EXCEL_UPLOADS"] = "/app/app/static/excel" #儲存位置
app.config["ALLOWED_EXCEL_EXTENSIONS"] = ["XLSX"] #允許的副檔名
app.config["SECRET_KEY"] = "OCML3BOswQEUeaxcuKHLpw" #隨機產生的SECRET_KEY，有這個才能跑flash

@app.route("/upload-excel", methods=["GET", "POST"]) #上傳excel檔
def upload_excel():

    if request.method == "POST":

        if request.files:

            excel = request.files["excel"]

            if excel.filename == "":
                flash('未選取檔案', 'warning')
                return redirect(request.url)

            if allowed_excel(excel.filename):
                filename = secure_filename(excel.filename)
                
                #如果檔名已經存在，則刪除舊檔，建立新檔
                if os.path.isfile(app.config["EXCEL_UPLOADS"] + excel.filename):
                    os.remove(app.config["EXCEL_UPLOADS"] + excel.filename)
                    excel.save(os.path.join(app.config["EXCEL_UPLOADS"], excel.filename))
                else:
                    excel.save(os.path.join(app.config["EXCEL_UPLOADS"], excel.filename))

                #如果ori.xlsx存在，則刪除舊檔，建立新檔
                str_upload_path = str(app.config["EXCEL_UPLOADS"])
                if os.path.isfile(app.config["EXCEL_UPLOADS"] + "/ori.xlsx"):
                    os.remove(app.config["EXCEL_UPLOADS"] + "/ori.xlsx")
                    os.rename(str_upload_path + "/" + excel.filename,str_upload_path + "/" + "ori.xlsx")
                else:
                    os.rename(str_upload_path + "/" + excel.filename,str_upload_path + "/" + "ori.xlsx")

                flash('Excel saved', 'success')
                return redirect(request.url)
                #return redirect("/download/"+filename) #會下載剛剛上傳的檔案

            else:
                flash('請上傳附檔名為 .xlsx 的檔案', 'warning')
                return redirect(request.url)

    return render_template("public/upload_excel.html")

app.config["NEW_EXCEL"] = "/app/app/static/new_excel" #新excel檔(總成績歸零)的儲存位置
#將總成績清空
@app.route("/test", methods=["GET", "POST"])
def test():

    # 使用openpyxl建立新活頁簿wb_new
    wb_new = Workbook()
    wb_new.save(app.config["NEW_EXCEL"] + '/new_excel_test.xlsx')

    # 使用openpyxl讀取原始檔案
    wb = load_workbook(app.config["EXCEL_UPLOADS"] + '/ori.xlsx')
    ws = wb.worksheets[0]

    # 使用openpyxl讀取new_excel
    wb_new = load_workbook(app.config["NEW_EXCEL"] + '/new_excel_test.xlsx')
    ws_new = wb_new.active

    a = pd.read_excel(app.config["EXCEL_UPLOADS"] + '/ori.xlsx')
    df = pd.DataFrame(a)
    List= df['總成績'].tolist()  
    #print(List) #列出總成績那列的數字

    n=-1
    for i in List:
        n=n+1  
        df.at[n, "總成績"] = 0 
        df = DataFrame(df) 
        DataFrame(df).to_excel(app.config["NEW_EXCEL"] + "/" + 'new_excel_test.xlsx', sheet_name='Sheet1', index=False, header=True)
    return redirect("/download/"+'new_excel_test.xlsx') 

#下載檔案，用from flask import send_from_directory, abort
#app.config["CLIENT_EXCELS"] = "/app/app/static/excel" #要從哪裡下載

app.config["EXCEL_TM"] = "/app/app/static/excel_TM" #藏入商標後的檔案存檔路徑
app.config["EXCEL_SHA"] = "/app/app/static/excel_sha" #雜湊值存檔路徑
app.config["EXCEL_HIST"] = "/app/app/static/excel_hist" #excel產生的直方圖存檔路徑

#輸入欄位名稱_藏入商標
@app.route("/add-trademark", methods=["POST"])
def add_trademark():
    return render_template("public/addTM.html")

#藏入商標
@app.route("/trademark", methods=["GET", "POST"])
def trademark():

    # 使用openpyxl建立新活頁簿wb_new
    wb_new = Workbook()
    wb_new.save(app.config["EXCEL_TM"] + '/new_excel_TM.xlsx')

    # 使用openpyxl讀取原始檔案
    wb = load_workbook(app.config["EXCEL_UPLOADS"] + '/ori.xlsx')
    ws = wb.worksheets[0]

    # 使用openpyxl讀取new_excel
    wb_new = load_workbook(app.config["EXCEL_TM"] + '/new_excel_TM.xlsx')
    ws_new = wb_new.active

    #雜湊值的檔案
    path = app.config["EXCEL_SHA"] + '/sha.txt'
    f = open(path, 'w',encoding='UTF-8')

    #讀舊檔，輸入欄位名稱
    a=pd.read_excel(app.config["EXCEL_UPLOADS"] + '/ori.xlsx')
    df = pd.DataFrame(a)
    sp = request.args.get('TMcolname') #request.args.get('TMcolname')是由使用者輸入的欄位名稱
    List= df[sp].tolist()

    df1=df[sp].apply(str)
    u=[]
    for i in df1:
        u.append(i)

    T="".join(u)

    #生成雜湊值
    sha = hashlib.md5(T.encode("utf-8")).hexdigest()
    print(sha,file=f)  #要留 存檔

    bin_str = ""
    for n in sha:
        bin_str += bin(int(n,16))[2:].zfill(4)
    print(bin_str,file=f)  #要留 存檔

    recounted = Counter(List)   #統計資料出現次數
    b1=df.groupby(sp).size()
    A=max(b1)  #找出最大值的點為多少

    a=np.array(List)   #將資料改成陣列  (分數)
    a1=np.unique(a)    

    b=np.array(b1)     #將資料轉成陣列  (次數)

    c=np.vstack((b,a1))   #合併變成二維陣列

    MAX=max(List)

    plt.hist(List,range(0,MAX+2),align='left', edgecolor='#000000',linewidth=2)
    plt.xticks(List)
    plt.savefig(app.config["EXCEL_HIST"] + "/hist_ori.png") #原始資料的直方圖
    plt.close()

    n=-1
    p=[]

    for row in c:    
        for col in row: 
            n=n+1
            if(col==A):     #判斷次數如果是跟最大一樣
                p.append(c[1][n])

    M=p[0]     #找第一個出現的最大值

    m=-1
    o=[]
    for i in a:
        m=m+1  
        if i < M:  
            i=i-1
            df.at[m,sp] = i
            DataFrame(df).to_excel(app.config["EXCEL_TM"] + '/new_excel_TM.xlsx',sheet_name='Sheet1', index=False, header=True)
            if(i<0): 
                o.append(m)
        else:     
            if(i==M):
                df.at[m,sp] = i
                DataFrame(df).to_excel(app.config["EXCEL_TM"] + '/new_excel_TM.xlsx',sheet_name='Sheet1', index=False, header=True)
            else:
                i=i+1
                df.at[m,sp] = i
                DataFrame(df).to_excel(app.config["EXCEL_TM"] + '/new_excel_TM.xlsx',sheet_name='Sheet1', index=False, header=True)  

    def union_without_repetition(list1,list2):
        result = list(set(list1) | set(list2))
        return result

    r=pd.read_excel(app.config["EXCEL_TM"] + '/new_excel_TM.xlsx',sheet_name='Sheet1')
    sr= df[sp].tolist()
    ar=np.array(sr)

    bins_listx=union_without_repetition(List,sr)
    plt.hist(sr,range(-1,MAX+2),align='left', edgecolor='#000000',linewidth=2)
    plt.xticks(bins_listx)
    plt.savefig(app.config["EXCEL_HIST"] + "/hist_shifting.png")
    plt.close()

    tc=len(bin_str)

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
                    DataFrame(df).to_excel(app.config["EXCEL_TM"] + '/new_excel_TM.xlsx',sheet_name='Sheet1', index=False, header=True)
                if(bin_str[t]=='0'):
                    DataFrame(df).to_excel(app.config["EXCEL_TM"] + '/new_excel_TM.xlsx',sheet_name='Sheet1', index=False, header=True)
                if(bin_str[t]==' '):
                    i=i+1
                    df.at[n,sp] = i
                    DataFrame(df).to_excel(app.config["EXCEL_TM"] + '/new_excel_TM.xlsx',sheet_name='Sheet1', index=False, header=True)
        else:
            df.at[n,sp]=i
            DataFrame(df).to_excel(app.config["EXCEL_TM"] + '/new_excel_TM.xlsx',sheet_name='Sheet1', index=False, header=True)



    def union_without_repetition(list1,list2):
        result = list(set(list1) | set(list2))
        return result

    fn = app.config["EXCEL_TM"] + '/new_excel_TM.xlsx'
    wb=load_workbook(fn)
    sheet_ranges = wb['Sheet1']

    k = pd.read_excel(app.config["EXCEL_TM"] + '/new_excel_TM.xlsx',nrows=0)
    L = k.columns.tolist()
    r=L.index(sp)+1

    font = Font(size=15)

    for i in o:
        i=i+2 #PYTHON讀的格子0 在EXCEL是2
        f = sheet_ranges.cell(row=i, column=r)
        f.font = font
        f.value=0
        wb.save(app.config["EXCEL_TM"] + '/new_excel_TM.xlsx')

    r=pd.read_excel(app.config["EXCEL_TM"] + '/new_excel_TM.xlsx')
    sx= df[sp].tolist()
    bins_listx2=union_without_repetition(bins_listx,sx)
    plt.hist(sx,range(0,MAX+2),align='left', edgecolor='#000000',linewidth=2)
    plt.xticks(bins_listx2)
    plt.savefig(app.config["EXCEL_HIST"] + "/hist_hiding.png")

    return redirect("/download2/"+'new_excel_TM.xlsx')

@app.route("/download/<excel_name>")
def downloadfile(excel_name):
    try:
        return send_from_directory(app.config["NEW_EXCEL"], path=excel_name, as_attachment=True)
    except FileNotFoundError:
        abort(404)
#原本的程式碼return send_from_directory(app.config["CLIENT_EXCELS"], filename=excel_name, as_attachment=True)，現在filename要改成path

@app.route("/download2/<excel_name>")
def downloadfile2(excel_name):
    try:
        return send_from_directory(app.config["EXCEL_TM"], path=excel_name, as_attachment=True)
    except FileNotFoundError:
        abort(404)
#原本的程式碼return send_from_directory(app.config["CLIENT_EXCELS"], filename=excel_name, as_attachment=True)，現在filename要改成path