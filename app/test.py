from pathlib import Path
import os
import time

upload_path = os.path.join(os.path.expanduser("~"), 'Downloads')
os.mkdir(upload_path + '\\' + time.strftime('%Y-%m-%d %H%M', time.localtime())) #在下載裡建立新資料夾

str_upload_path = str(upload_path)
print(str_upload_path + "\ori.xlsx")
print(upload_path)