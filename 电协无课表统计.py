import openpyxl
import numpy as np
import os

print("$$ 无课表处理\n\n")
DATAPATH = './Datas/'

files = next(os.walk(DATAPATH))[2]

TargetData = [[[],[],[],[],[],[],[]],
               [[],[],[],[],[],[],[]],
               [[],[],[],[],[],[],[]],
               [[],[],[],[],[],[],[]],
               [[],[],[],[],[],[],[]],
               [[],[],[],[],[],[],[]],]


def LoadData():
    totalFile = len(files)
    fileCounter = 0
    for fileName in files:
        fileCounter += 1
        wb = openpyxl.load_workbook(DATAPATH + fileName)  
        sheet = wb['Main']
    
        name = sheet.cell(row=11, column=3).value
        class_ = sheet.cell(row=12, column=3).value
    
        print(f">>> Analysis Process : {fileCounter / totalFile:>7.2%} - {name:5}  {class_}")
    
        for i in range(6):
            for j in range(7):
                tmpVal = sheet.cell(row=4 + i, column=3 + j).value
                if  not(tmpVal != None and tmpVal > 0):
                    TargetData[i][j].append(name)
    return fileCounter, fileName, totalFile

def SaveData():
    TargetFile = openpyxl.load_workbook("./TargetExcel_Template.xlsx")  
    TargetSheet = TargetFile['Main']
    
    for i in range(6):
            for j in range(7):
                print(f'>>> Saving Process : {(i*7+j)/41:>7.2%}')
                TargetSheet.cell(row=4 + i, column=3 + j).value = "\n".join(TargetData[i][j])
    return TargetFile

print("$$ 开始加载数据 ##########################\n")
LoadData()
print("\n$$ 加载数据完毕 ##########################\n")

print("$$ 开始保存数据 ##########################\n")
TargetFile = SaveData()
print("\n$$ 保存数据完毕 #########################\n")

TargetFile.save("./TargetExcel.xlsx")
print("\n$$ 汇总结果已保存到 './TargetExcel.xlsx'\n")