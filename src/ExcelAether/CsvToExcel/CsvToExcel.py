import sys
import xlwt
import chardet

if __name__ == '__main__':
    arge = sys.argv
    inputName = arge[1]
    outName = arge[2]

    f = open(inputName, 'rb')
    data = f.read()
    character = chardet.detect(data).get('encoding')

    f = open(inputName, 'r', encoding=character)
    xls = xlwt.Workbook()
    sheet = xls.add_sheet('sheet1', cell_overwrite_ok=True)
    x = 0
    while True:
        line = f.readline()
        if not line:
            break
        for i in range(len(line.split(','))):
            item = line.split(',')[i]
            sheet.write(x, i, item)
        x += 1
    f.close()
    xls.save(outName)
