
import os
import xlrd
# 检查谁未交作业，初始值都设置为1，当数字大于1时就是未交人员
path = r'D:\离散作业\第十次作业'
dirs = os.listdir(path)
workbook = xlrd.open_workbook(r'mingdan.xlsx')
sheet_names= workbook.sheet_names()

for sheet_name in sheet_names:
    sheet2 = workbook.sheet_by_name(sheet_name)
    print (sheet_name)
    number = sheet2.col_values(0) # 获取第1列内容
    name = sheet2.col_values(1)
    tijiao = sheet2.col_values(2)


for j, file in enumerate(dirs):
    files = file.strip().split(".")[0]
    for i in range(0, len(number)):
        if tijiao[i] == 0:
                print("")
        else:
            ruslut = files in name[i]
            if ruslut:
                tijiao[i] = 0
            else:
                tijiao[i] += 1

for i in range(1, len(number)):
    if tijiao[i] != 0:
        print("{},{}".format(name[i], tijiao[i]))


