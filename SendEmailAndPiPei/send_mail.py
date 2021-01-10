#encoding=utf-8

import time

import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os
import xlrd
import shutil
os.chdir('C:\\Users\\wpx\\Desktop\\success')
suc = os.getcwd() # 建立发送成功名单，发送成功，自动将文件移动到success这个文件夹下面
path = r'C:\Users\wpx\Desktop\第九次作业'
dirs = os.listdir(path)
workbook = xlrd.open_workbook(r'C:\Users\wpx\Desktop\人员匹配\人员匹配\2020离散数学名单.xlsx') #空名单位置
sheet_names = workbook.sheet_names()



content = '这是一封来自离散数学班助的作业批改邮件。'

def send_mail(path,file,addr):
    # 因为可能会出现发送失败，所以163和QQ邮箱换着用
    # fromaddr = '163邮箱地址'
    # password = '授权码'

    fromaddr = 'QQ号'
    password = '授权码'
    toaddrs = addr

    content = '这是一封第9次离散作业批改邮件。'
    textApart = MIMEText(content)


    pdfFile = path+'\\'+file
    pdfApart = MIMEApplication(open(pdfFile, 'rb').read())
    pdfApart.add_header('Content-Disposition', 'attachment', filename=file)
    #
    m = MIMEMultipart()
    m.attach(textApart)
    m.attach(pdfApart)
    m['Subject'] = '第9次作业批改'

    try:
        server = smtplib.SMTP('smtp.qq.com')
        # server = smtplib.SMTP('smtp.163.com')
        server.login(fromaddr, password)
        server.sendmail(fromaddr, toaddrs, m.as_string())
        print('success')
        server.close()
        shutil.move(pdfFile, suc)
    except smtplib.SMTPException as e:
        print('error:', e)  # 打印错误

for sheet_name in sheet_names:
    sheet2 = workbook.sheet_by_name(sheet_name)
    print(sheet_name)
    number = sheet2.col_values(0) # 获取第1列内容
    name = sheet2.col_values(1)
    qq = sheet2.col_values(2)
    n = 0

for i in range(1, len(number)):
    for j, file in enumerate(dirs):
        files = file.strip().split(".")[0]
        ruslut = name[i] in file
        if ruslut:
            print("name:",name[i])
            addr = qq[i] + '@qq.com'
            print(addr)
            send_mail(path, file, addr)
            time.sleep(3)
            n += 1
print('已发送邮件数目：', n)



