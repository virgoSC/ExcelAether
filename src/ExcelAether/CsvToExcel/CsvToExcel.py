import sys
import xlwt
import chardet
import re

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
        line = line.strip()
        if not line:
            break
        for i in range(len(line.split(','))):
            item = line.split(',')[i]
            # 金额转换
            # if re.search('^(([1-9]{1}\\d*)|([0]{1}))(\\.(\\d){0,2})?$', item):
            if re.search('^([-+])?\\d+(\\.[0-9]{1,2})?$', item):
                item = float('%.2f' % float(item))

            sheet.write(x, i, item)
        x += 1
    f.close()
    xls.save(outName)
