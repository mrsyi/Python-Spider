import xlrd
import xlwt

wb = xlwt.Workbook()
sheet6 = wb.add_sheet('6阴',cell_overwrite_ok=True)
sheet7 = wb.add_sheet('7阴',cell_overwrite_ok=True)
sheet8 = wb.add_sheet('8阴',cell_overwrite_ok=True)
sheet9 = wb.add_sheet('9阴'cell_overwrite_ok=True)
sheet10 = wb.add_sheet('10阴',cell_overwrite_ok=True)
sheet11 = wb.add_sheet('11阴',cell_overwrite_ok=True)
sheet12 = wb.add_sheet('12阴',cell_overwrite_ok=True)
sheet13 = wb.add_sheet('13阴',cell_overwrite_ok=True)
sheet14 = wb.add_sheet('14阴',cell_overwrite_ok=True)
sheet15 = wb.add_sheet('15阴',cell_overwrite_ok=True)
sheet16 = wb.add_sheet('16阴',cell_overwrite_ok=True)
sheet17 = wb.add_sheet('17阴',cell_overwrite_ok=True)
sheet18 = wb.add_sheet('18阴',cell_overwrite_ok=True)
sheet19 = wb.add_sheet('19阴',cell_overwrite_ok=True)
sheet21 = wb.add_sheet('20阴',cell_overwrite_ok=True)
sheet22 = wb.add_sheet('22阴',cell_overwrite_ok=True)
sheet23 = wb.add_sheet('23阴',cell_overwrite_ok=True)
sheet24 = wb.add_sheet('24阴',cell_overwrite_ok=True)
sheet25 = wb.add_sheet('25阴',cell_overwrite_ok=True)
sheet26 = wb.add_sheet('26阴',cell_overwrite_ok=True)
sheet27 = wb.add_sheet('27阴',cell_overwrite_ok=True)

for p in range(6,27):
    if p == 20 or p == 9:
        continue
    else:
        re = xlrd.open_workbook(str(p) + "#阴面.xlsx")
        print(re.sheet_names())
        sheet = re.sheet_by_name('Sheet')
        s = locals()['sheet' + str(p)]
        s.write(0,0,'房间号')
        s.write(1,0,sheet.row_values(1)[0])
        s.write(2,0,'日期')
        for i in range(2,13):
            s.write(0,i,'房间号')
        for i in range(1,13):
            s.write(2,i,'用量')

        for m in range(3,106):
            s.write(m,0,sheet.row_values(m)[0])
        for m in range(3,106):
            s.write(m,1,sheet.row_values(m)[1])
        for m in range(109,212):
            s.write(m-106,2,sheet.row_values(m)[0])
        for m in range(215,318):
            s.write(m-212,3,sheet.row_values(m)[0])
        for m in range(321,424):
            s.write(m-318,4,sheet.row_values(m)[0])
        for m in range(427,530):
            s.write(m-424,5,sheet.row_values(m)[0])
        for m in range(533,636):
            s.write(m-530,6,sheet.row_values(m)[0])
        for m in range(639,742):
            s.write(m-636,7,sheet.row_values(m)[0])
        for m in range(745,848):
            s.write(m-742,8,sheet.row_values(m)[0])
        for m in range(851,954):
            s.write(m-848,9,sheet.row_values(m)[0])
        for m in range(957,1060):
            s.write(m-954,10,sheet.row_values(m)[0])
        for m in range(1063,1166):
            s.write(m-1060,11,sheet.row_values(m)[0])
        for m in range(1169,1272):
            s.write(m-1166,12,sheet.row_values(m)[0])
        wb.save('阴面-9#不在.xls')
