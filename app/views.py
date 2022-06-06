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
#用sqlite上傳檔案
#from . import db
#用MongoDB上傳檔案
import openpyxl
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

app.config["NEW_EXCEL"] = "/app/app/static/new_excel" #新excel檔的儲存位置

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

@app.route("/download/<excel_name>")
def downloadfile(excel_name):
    try:
        return send_from_directory(app.config["NEW_EXCEL"], path=excel_name, as_attachment=True)
    except FileNotFoundError:
        abort(404)
#原本的程式碼return send_from_directory(app.config["CLIENT_EXCELS"], filename=excel_name, as_attachment=True)，現在filename要改成path