import pandas as pd
import sys
import os
import chardet

if __name__ == '__main__':
    arge = sys.argv
    inputName = arge[1]
    outName = arge[2]

    # 获取文件类型
    fileType = os.path.splitext(inputName)[-1]
    if fileType == '.csv':
        # 获取文件编码
        f = open(inputName, 'rb')
        data = f.read()
        character = chardet.detect(data).get('encoding')
        print(character)
        f.close()
        data = pd.read_csv(inputName, encoding=character)
    else:
        data = pd.read_excel(inputName)
    data.to_csv(outName, encoding='UTF-8')
