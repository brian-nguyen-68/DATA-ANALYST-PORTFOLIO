import pandas as pd
import xlwings as xw
import os
from datetime import date, datetime,timedelta

import win32com.client as win32
import openpyxl as xl
from pathlib import Path
import xlwings as xw
win32c =  win32.constants
# create excel object
excel = win32.gencache.EnsureDispatch('Excel.Application')

# excel can be visible or not
excel.Visible = True  # False
# Lấy ngày hiện tại
today = datetime.today()

# Tạo một khoảng thời gian để lùi lại một ngày
one_day = timedelta(days=1)

# Lấy ngày hôm qua bằng cách trừ khoảng thời gian một ngày từ ngày hiện tại
yesterday = today - one_day

# Lấy phần ngày của ngày hôm qua
day_of_month_yesterday = yesterday.day
today = date.today()
t = str(today.month)

bc = fr"G:\\.shortcut-targets-by-id\\12I_9JRkJRh5iv-g-rvkTZu6vgx0hZc2Y\\03.2023\\04. BAO CAO NGAY\\BC TON KHO + TOC DO RA HANG\\Tốc độ ra hàng_sample.xlsx"
tam1 = fr"G:\\.shortcut-targets-by-id\\12I_9JRkJRh5iv-g-rvkTZu6vgx0hZc2Y\\03.2023\\04. BAO CAO NGAY\\BC TON KHO + TOC DO RA HANG\\BC TON KHO\\THANG {t}.2023\\Báo cáo tồn kho theo kho {day_of_month_yesterday}.{t}.xlsx"
#lay data raw



wb1 =xw.Book(tam1)
template = xw.Book(bc)
sheettonun= template.sheets["tồn kho UN"]
sheettonul = template.sheets["tồn kho UL"]
sheetstt= template.sheets["data"]

sheettonun.range("A5:AV10000").clear_contents()
sheettonul.range("A5:Z5000").clear_contents()
sheetstt.range("A5:AN30000").clear_contents()


dfshapetonun = pd.read_excel(tam1, sheet_name="BC TỒN KHO THEO KHO UNIS")
dfshapetonul  = pd.read_excel(tam1, sheet_name="BC TỒN KHO THEO KHO UNILUX")
dfshapestt = pd.read_excel(tam1, sheet_name="so chi tiet")



shapetonun  = dfshapetonun["Unnamed: 0"].count()
shapetonul  = dfshapetonul["Unnamed: 0"].count()
shapestt = dfshapestt["Ngày"].count()


sheettonun1 = wb1.sheets["BC TỒN KHO THEO KHO UNIS"]
sheettonul1 = wb1.sheets["BC TỒN KHO THEO KHO UNILUX"]
sheetstt1 = wb1.sheets["so chi tiet"]

#copy data
rngtonun = sheettonun1.range(fr"A9:AU{shapetonun+7}").copy()
sheettonun.range("A4").paste('values')

rngtonul= sheettonul1.range(fr"A8:Y{shapetonul+6}").copy()
sheettonul.range("A4").paste('values')

rngstt= sheetstt1.range(fr"A2:AB{shapestt+1}").copy()
sheetstt.range("A4").paste('values')


# save lai file
template.save()


