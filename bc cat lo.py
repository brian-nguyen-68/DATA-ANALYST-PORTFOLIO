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

pathmau = fr"G:\\.shortcut-targets-by-id\\12I_9JRkJRh5iv-g-rvkTZu6vgx0hZc2Y\\03.2023\\04. BAO CAO NGAY\\BC TON KHO + TOC DO RA HANG\\BC TON KHO\\THANG {t}.2023\\Báo cáo tồn kho theo kho {day_of_month_yesterday}.{t}.xlsx"
bc = fr"G:\\.shortcut-targets-by-id\\12I_9JRkJRh5iv-g-rvkTZu6vgx0hZc2Y\\03.2023\\04. BAO CAO NGAY\\BC CAT LO\\Báo cáo cắt lô gạch_sample.xlsx"


wb = xw.Book(pathmau)
template = xw.Book(bc)


dfshapetonun= pd.read_excel(pathmau, sheet_name="BC TỒN KHO THEO KHO UNIS")
dfshapetonul  = pd.read_excel(pathmau, sheet_name="BC TỒN KHO THEO KHO UNILUX")

shapetonun  = dfshapetonun["Unnamed: 0"].count()
shapetonul = dfshapetonul["Unnamed: 0"].count()

sheetun= template.sheets["tồn kho UN"]
sheetul = template.sheets["tồn kho UL"]
sheetht = template.sheets["HỆ THỐNG"]


sheettonun = wb.sheets["BC TỒN KHO THEO KHO UNIS"]
sheettonul = wb.sheets["BC TỒN KHO THEO KHO UNILUX"]


#copy data

rngg = sheetht.range("G4:G55").copy()
sheetht.range("AE4").paste('values')

rngl = sheetht.range("L4:L55").copy()
sheetht.range("AF4").paste('values')

rngq = sheetht.range("Q4:Q55").copy()
sheetht.range("AG4").paste('values')

rngv = sheetht.range("V4:V55").copy()
sheetht.range("AH4").paste('values')

rngaa= sheetht.range("AA4:AA55").copy()
sheetht.range("AI4").paste('values')

sheetun.range("A4:AU10000").clear_contents()
sheetul.range("A4:Y4000").clear_contents()

rngtonun = sheettonun.range(fr"A9:AU{shapetonun+7}").copy()
sheetun.range("A4").paste('values')

rngtonundan = sheettonul.range(fr"A8:Y{shapetonul+6}").copy()
sheetul.range("A4").paste('values')



# save lai file
template.save(fr"G:\\.shortcut-targets-by-id\\12I_9JRkJRh5iv-g-rvkTZu6vgx0hZc2Y\\03.2023\\04. BAO CAO NGAY\\BC CAT LO\\Thang {t}\\Báo cáo cắt lô gạch {day_of_month_yesterday}.{t}.xlsx")
wb.close()
