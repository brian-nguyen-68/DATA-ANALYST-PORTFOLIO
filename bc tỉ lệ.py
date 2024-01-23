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

stt = fr"D:\\.shortcut-targets-by-id\\17iVJ0hhkx3ZIc5d3mWBvN1A4R-QCnV4r\\06.2024\\04.BAOCAONGAY\\BC KEO HANG\\bc.xlsx"
bc = fr"D:\\.shortcut-targets-by-id\\17iVJ0hhkx3ZIc5d3mWBvN1A4R-QCnV4r\\06.2024\\04.BAOCAONGAY\\BC KEO HANG\\BC TỈ LỆ\\BC TỈ LỆ.xlsx"
tam = fr"D:\\.shortcut-targets-by-id\\17iVJ0hhkx3ZIc5d3mWBvN1A4R-QCnV4r\\06.2024\\04.BAOCAONGAY\\BC KEO HANG\\BC TỈ LỆ\\tam.xlsx"
#lay data raw
loccn = ["222", "000"]
locloai = ["PALLET","VANCHUYEN"]


dfshape5 = pd.read_excel(stt, sheet_name="BÁO CÁO DOANH THU THỰC")
shape5 = dfshape5.shape[0]
# loc cac dk

dfstt = dfshape5.iloc[8:shape5, :25]
dfstt1  = dfstt[dfstt["Unnamed: 0"].notna() & (dfstt["Unnamed: 3"].isin(locloai) == False) & (dfstt["Unnamed: 15"].isin(["Bán lẻ","Cắt lô","GS ôm kho/Duyệt giá"]) == True) & (dfstt["Unnamed: 24"].isin(loccn) ==False) & (dfstt["Unnamed: 16"].notna())]


with pd.ExcelWriter(tam) as writer:
    dfstt1.to_excel(writer, sheet_name='dtt')

wb = xw.Book(tam)
template = xw.Book(bc)


# Tạo một khoảng thời gian để lùi lại một ngày
one_day = timedelta(days=1)

# Lấy ngày hôm qua bằng cách trừ khoảng thời gian một ngày từ ngày hiện tại
yesterday = today - one_day

# Lấy phần ngày của ngày hôm qua
day_of_month_yesterday = yesterday.day
# thay ngay thang
row_num = 1
col_num = 1
thayngay = template.sheets["Tỉ lệ"]
thayngay.range(row_num, col_num).value = str(day_of_month_yesterday)

dfshapedtt = pd.read_excel(tam, sheet_name="dtt")

shapedtt  = dfshapedtt["Unnamed: 0"].count()

sheetdtt= template.sheets["data"]

sheetdtt1 = wb.sheets["dtt"]

#copy data

rngdtt= sheetdtt1.range(fr"B2:Z{shapedtt+1}").copy()
sheetdtt.range("A3").paste('values')


# save lai file
# save lai file

template.save(fr"D:\\.shortcut-targets-by-id\\17iVJ0hhkx3ZIc5d3mWBvN1A4R-QCnV4r\\06.2024\\04.BAOCAONGAY\\BC KEO HANG\\BC TỈ LỆ\\Tháng {t}\\BC TỈ LỆ_{str(today.day)}.{t}.xlsx")

sheetdtt1.range("A2:AC50000").clear_contents()
template.save()
wb.save()
wb.close()
