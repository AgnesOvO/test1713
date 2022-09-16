from app import app

from flask import Flask, render_template, request, redirect, send_from_directory, abort, flash, session, Blueprint

app.config["EXCEL_TM"] = "/app/app/static/excel_TM" #藏入商標後的檔案存檔路徑
app.config["EXCEL_SHA"] = "/app/app/static/excel_sha" #雜湊值存檔路徑
app.config["EXCEL_HIST"] = "/app/app/static/excel_hist" #excel產生的直方圖存檔路徑

#@app.route("/download_HIST/<excel_name>")
#def downloadfile_SHA(excel_name):
#    try:
#        return send_from_directory(app.config["EXCEL_HIST"], path=excel_name, as_attachment=True)
#    except FileNotFoundError:
#        abort(404)