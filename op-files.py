import xlrd
import xlwt

wb = xlwt.Workbook()
sheet6 = wb.add_sheet('6',cell_overwrite_ok=True)
sheet7 = wb.add_sheet('7',cell_overwrite_ok=True)
sheet8 = wb.add_sheet('8',cell_overwrite_ok=True)
sheet9 = wb.add_sheet('9',cell_overwrite_ok=True)
sheet10 = wb.add_sheet('10',cell_overwrite_ok=True)
sheet11 = wb.add_sheet('11',cell_overwrite_ok=True)
sheet12 = wb.add_sheet('12',cell_overwrite_ok=True)
sheet13 = wb.add_sheet('13',cell_overwrite_ok=True)
sheet14 = wb.add_sheet('14',cell_overwrite_ok=True)
sheet15 = wb.add_sheet('15',cell_overwrite_ok=True)
sheet16 = wb.add_sheet('16',cell_overwrite_ok=True)
sheet17 = wb.add_sheet('17',cell_overwrite_ok=True)
sheet18 = wb.add_sheet('18',cell_overwrite_ok=True)
sheet19 = wb.add_sheet('19',cell_overwrite_ok=True)
sheet21 = wb.add_sheet('21',cell_overwrite_ok=True)
sheet22 = wb.add_sheet('22',cell_overwrite_ok=True)
sheet23 = wb.add_sheet('23',cell_overwrite_ok=True)
sheet24 = wb.add_sheet('24',cell_overwrite_ok=True)
sheet25 = wb.add_sheet('25',cell_overwrite_ok=True)
sheet26 = wb.add_sheet('26',cell_overwrite_ok=True)
sheet27 = wb.add_sheet('27',cell_overwrite_ok=True)

for p in range(6,27):
    if p == 20 or p == 9:
        continue
    else:
        re = xlrd.open_workbook(str(p) + "#阳面.xlsx")
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
        s.write(1,2,sheet.row_values(107)[0])
        s.write(1,3,sheet.row_values(213)[0])
        s.write(1,4,sheet.row_values(319)[0])
        s.write(1,5,sheet.row_values(425)[0])
        s.write(1,6,sheet.row_values(531)[0])
        s.write(1,7,sheet.row_values(637)[0])
        s.write(1,8,sheet.row_values(743)[0])
        s.write(1,9,sheet.row_values(849)[0])
        s.write(1,10,sheet.row_values(955)[0])
        s.write(1,11,sheet.row_values(1061)[0])
        s.write(1,12,sheet.row_values(1167)[0])

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
        wb.save('阳面-9#不在.xls')
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

        s.write(1,2,sheet.row_values(107)[0])
        s.write(1,3,sheet.row_values(213)[0])
        s.write(1,4,sheet.row_values(319)[0])
        s.write(1,5,sheet.row_values(425)[0])
        s.write(1,6,sheet.row_values(531)[0])
        s.write(1,7,sheet.row_values(637)[0])
        s.write(1,8,sheet.row_values(743)[0])
        s.write(1,9,sheet.row_values(849)[0])
        s.write(1,10,sheet.row_values(955)[0])
        s.write(1,11,sheet.row_values(1061)[0])
        s.write(1,12,sheet.row_values(1167)[0])

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
