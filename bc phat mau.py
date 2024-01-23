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

today = date.today()
t = str(today.month)

pathmau = fr"G:\\.shortcut-targets-by-id\\12I_9JRkJRh5iv-g-rvkTZu6vgx0hZc2Y\\03.2023\\04. BAO CAO NGAY\\BC PHAT MAU\\CP MAU THANG\\Báo cáo hàng xuất mẫu.xlsx"
pathdtt = fr"G:\\.shortcut-targets-by-id\\12I_9JRkJRh5iv-g-rvkTZu6vgx0hZc2Y\\03.2023\\04. BAO CAO NGAY\\BC PHAT MAU\\CP MAU THANG\\Báo cáo doanh thu thực tháng.xlsx"
bc = fr"G:\\.shortcut-targets-by-id\\12I_9JRkJRh5iv-g-rvkTZu6vgx0hZc2Y\\03.2023\\04. BAO CAO NGAY\\BC PHAT MAU\\CP MAU THANG\\CHI PHÍ PHÁT MẪU_CT.xlsx"
tam = fr"G:\\.shortcut-targets-by-id\\12I_9JRkJRh5iv-g-rvkTZu6vgx0hZc2Y\\03.2023\\04. BAO CAO NGAY\\BC PHAT MAU\\CP MAU THANG\\tam.xlsx"
#lay data raw
loccn = ["222", "000"]
locloai = ["PALLET","VANCHUYEN"]

dfshape1 = pd.read_excel(pathmau, sheet_name="BÁO CÁO HÀNG XUẤT MẪU")
shape1 = dfshape1.shape[0]
dfshape2 = pd.read_excel(pathdtt, sheet_name="BÁO CÁO DOANH THU THỰC")
shape2 = dfshape2.shape[0]

dfmau = dfshape1.iloc[8:shape1, :16]
dfmau1 = dfmau[dfmau["Unnamed: 0"].notna()]

dfdtt = dfshape2.iloc[7:shape2, :34]
dfdtt1 = dfdtt[dfdtt["Unnamed: 0"].notna() & (dfdtt["Unnamed: 5"].isin(["Viên","Hộp"]) == True) & (dfdtt["Unnamed: 28"].isin(loccn) == False)]



with pd.ExcelWriter(tam) as writer:
    dfmau1.to_excel(writer, sheet_name='xuatmau')
    dfdtt1.to_excel(writer, sheet_name='dtt')


wb = xw.Book(tam)
template = xw.Book(bc)


dfshapemau = pd.read_excel(tam, sheet_name="xuatmau")
dfshapedtt  = pd.read_excel(tam, sheet_name="dtt")

shapemau  = dfshapemau["Unnamed: 0"].count()
shapedtt  = dfshapedtt["Unnamed: 0"].count()

sheetmau= template.sheets["DATAMAU"]
sheetdtt = template.sheets["DATASALE"]


sheetmau1 = wb.sheets["xuatmau"]
sheetdtt1 = wb.sheets["dtt"]


#copy data


rngtonun = sheetmau1.range(fr"B2:P{shapemau+1}").copy()
sheetmau.range("A3").paste('values')

rngtonundan = sheetdtt1.range(fr"B2:AD{shapedtt+1}").copy()
sheetdtt.range("A4").paste('values')


# save lai file
template.save(fr"G:\\.shortcut-targets-by-id\\12I_9JRkJRh5iv-g-rvkTZu6vgx0hZc2Y\\03.2023\\04. BAO CAO NGAY\\BC PHAT MAU\\CP MAU THANG\\THANG {t}\\CHI PHÍ PHÁT MẪU_CT_{today - timedelta(days=1)}.xlsx")
sheetmau1.range("A2:P30000").clear_contents()
sheetdtt1.range("A2:AI30000").clear_contents()
wb.save()
wb.close()
