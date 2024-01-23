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

tonun = fr"D:\\.shortcut-targets-by-id\\17iVJ0hhkx3ZIc5d3mWBvN1A4R-QCnV4r\\06.2024\\04.BAOCAONGAY\\BC KEO HANG\\TON_UN.xlsx"
tonul = fr"D:\\.shortcut-targets-by-id\\17iVJ0hhkx3ZIc5d3mWBvN1A4R-QCnV4r\\06.2024\\04.BAOCAONGAY\\BC KEO HANG\\TON_UL.xlsx"
hangdd = fr"D:\\.shortcut-targets-by-id\\17iVJ0hhkx3ZIc5d3mWBvN1A4R-QCnV4r\\06.2024\\04.BAOCAONGAY\\BC KEO HANG\\HÀNG ĐI ĐƯỜNG.xlsx"
stt = fr"D:\\.shortcut-targets-by-id\\17iVJ0hhkx3ZIc5d3mWBvN1A4R-QCnV4r\\06.2024\\04.BAOCAONGAY\\BC KEO HANG\\bc.xlsx"
bc = fr"D:\\.shortcut-targets-by-id\\17iVJ0hhkx3ZIc5d3mWBvN1A4R-QCnV4r\\06.2024\\04.BAOCAONGAY\\BC KEO HANG\\SO KEO_TON KHO_SO BAN - CT.xlsx"
tam = fr"D:\\.shortcut-targets-by-id\\17iVJ0hhkx3ZIc5d3mWBvN1A4R-QCnV4r\\06.2024\\04.BAOCAONGAY\\BC KEO HANG\\tam.xlsx"
#lay data raw
loccn = ["222", "000"]
locloai = ["PALLET","VANCHUYEN"]

dfshape2 = pd.read_excel(tonun, sheet_name="BÁO CÁO TỒN KHO THEO KHO")
shape2 = dfshape2.shape[0]
dfshape3 = pd.read_excel(tonul, sheet_name="BÁO CÁO TỒN KHO THEO KHO")
shape3 = dfshape3.shape[0]
dfshape4 = pd.read_excel(hangdd, sheet_name="BẢNG KÊ THEO DÕI HÀNG ĐI ĐƯỜNG")
shape4 = dfshape4.shape[0]
dfshape5 = pd.read_excel(stt, sheet_name="BÁO CÁO DOANH THU THỰC")
shape5 = dfshape5.shape[0]
# loc cac dk

dftonun = dfshape2.iloc[5:shape2, :38]
dftonun1 = dftonun[dftonun["Unnamed: 0"].notna() & (dftonun["Unnamed: 2"].isin(["Bộ","CÁI","cái","Cái"]) == False)]

dftonul = dfshape3.iloc[5:shape3, :10]
dftonul1 = dftonul[dftonul["Unnamed: 0"].notna() & (dftonul["Unnamed: 2"].isin(["Bộ","cái","CÁI","Cái"]) == False)]

dfhangdd = dfshape4.iloc[6:shape4, :21]
dfhangdd1  = dfhangdd[(dfhangdd["Unnamed: 0"].isin(["Ô tô","Đường biển"]) == True) & (dfhangdd["Unnamed: 8"].isin(locloai) == False)]

dfstt = dfshape5.iloc[8:shape5, :25]
dfstt1  = dfstt[dfstt["Unnamed: 0"].notna() & (dfstt["Unnamed: 3"].isin(locloai) == False) & (dfstt["Unnamed: 15"].isin(["Bán lẻ","Cắt lô","GS ôm kho/Duyệt giá"]) == True) & (dfstt["Unnamed: 24"].isin(loccn) ==False) & (dfstt["Unnamed: 16"].notna())]


with pd.ExcelWriter(tam) as writer:
    dftonun1.to_excel(writer, sheet_name='ton un')
    dftonul1.to_excel(writer, sheet_name='ton ul')
    dfhangdd1.to_excel(writer, sheet_name='hang diduong')
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
thayngay = template.sheets["BC"]
thayngay.range(row_num, col_num).value = str(day_of_month_yesterday)


dfshapetonun = pd.read_excel(tam, sheet_name="ton un")
dfshapetonul  = pd.read_excel(tam, sheet_name="ton ul")
dfshapehangdiduong = pd.read_excel(tam, sheet_name="hang diduong")
dfshapedtt = pd.read_excel(tam, sheet_name="dtt")



shapetonun  = dfshapetonun["Unnamed: 0"].count()
shapetonul  = dfshapetonul["Unnamed: 0"].count()
shapehangdiduong  = dfshapehangdiduong["Unnamed: 1"].count()
shapedtt  = dfshapedtt["Unnamed: 0"].count()


sheettonun= template.sheets["TỒN UN"]
sheettonul = template.sheets["TỒN UL"]
sheethangdiduong= template.sheets["HANG DIDUONG"]
sheetdtt= template.sheets["DATA SALE"]


sheettonun1 = wb.sheets["ton un"]
sheettonul1 = wb.sheets["ton ul"]
sheethangdiduong1 = wb.sheets["hang diduong"]
sheetdtt1 = wb.sheets["dtt"]

#copy data


rngtonun = sheettonun1.range(fr"B2:AM{shapetonun+1}").copy()
sheettonun.range("A4").paste('values')


rngtonul= sheettonul1.range(fr"B2:K{shapetonul+1}").copy()
sheettonul.range("A4").paste('values')


rnghangdiduong= sheethangdiduong1.range(fr"B2:V{shapehangdiduong+1}").copy()
sheethangdiduong.range("A2").paste('values')

rngdtt= sheetdtt1.range(fr"B2:Z{shapedtt+1}").copy()
sheetdtt.range("A2").paste('values')


# save lai file

template.save(fr"D:\\.shortcut-targets-by-id\\17iVJ0hhkx3ZIc5d3mWBvN1A4R-QCnV4r\\06.2024\\04.BAOCAONGAY\\BC KEO HANG\\Tháng {t}\\KÉO HÀNG_{str(today.day)}.{t}.xlsx")
sheettonun1.range("A2:AO50000").clear_contents()
sheettonul1.range("A2:T50000").clear_contents()
sheethangdiduong1.range("A2:V50000").clear_contents()
sheetdtt1.range("A2:AC50000").clear_contents()
template.save()
wb.save()
wb.close()