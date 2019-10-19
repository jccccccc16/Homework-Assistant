# import re
# import os
# fileList = os.listdir("D:\\学习demo\\python\\文件批量改名\\texts")
#
# for fileName in fileList:
#
#     pat = ""
#     pattern = re.match(pat,fileName)
import jieba
import pymysql
import os
import re
import sys
import shutil

from email.mime.application import MIMEApplication
from smtplib import SMTP
from email.header import Header
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
import urllib
dir = ''    #作业所在目录
number = 0

#获取一个游标
def get_cur():
    conn = pymysql.connect(host='localhost',user='root',password='cjcisgood',database='class',port = 3306,charset='utf8')
    cur = conn.cursor()
    return cur,conn



def get_dir():
    global dir
    dir =input("作业所在的目录：")
    return dir


def cut_name(fileName):
    jieba.load_userdict('D:\\学习demo\\python\\文件批量改名\\error_name.txt')
    nameList = jieba.cut(fileName)
    nameTuple = []
    for word in nameList:
        nameTuple.append(word)
            
    nameTuple = tuple(nameTuple)
    print(nameTuple)
    return nameTuple


#查找学生的信息
def search_student(fileName):
    cur,conn = get_cur()
    

    nameTuple = cut_name(fileName)
    

    #执行数据库语句，找出该学生的信息
        
    sql_cmd ="SELECT * FROM STUDENT WHERE NAME in {nameTuple};".format(nameTuple=nameTuple)
        
    cur.execute(sql_cmd)
        
    student = cur.fetchone()

    #找到学生之后，标记为已交
    sql_commit = "UPDATE STUDENT SET IS_COMMIT = 1 WHERE NAME IN {nameTuple};".format(nameTuple=nameTuple)
    
    cur.execute(sql_commit)
    
    conn.commit()
    conn.close()
    return student


#修改文件名
def modify_name():
    print("="*20)
    
    global number
    number = int(input("这是实验几："))
    print("\n开始修改\n")
    global dir
    os.chdir(dir)
    fileList = os.listdir(dir)
    """
    匹配模式串
    """

    pat = '计科1182-.{2,3}-2018116212\d{2}-实验'+str(number)+'.\d{6}'
    
    #统计人数
    count = 0  

    for fileName in fileList:

        count = count+1

        print(fileName)
        pattern = re.match(pat,fileName)

        student = search_student(fileName) 

        if pattern:   #如果匹配

            print("%s格式正确跳过\n"%fileName)

            

            continue

        """
        修改文件名名字4
        """
        nameTuple = cut_name(fileName)
        
        file_format = nameTuple[-1]

        newFileName = '计科1182-' + student[1] +'-' + str(student[0]) + '-实验'+str(number) + '.'+file_format

        os.rename(fileName,newFileName)

        print("%s  修改成功!\n"%student[1])

    
    print("\n全部修改完毕！\n")
    print("共有 %d 人提交作业!\n" % count)
    

    




def how_many_students_commit():
    print("="*20)
    cur,conn = get_cur()
    
    sql_isnt_commit = "SELECT * FROM STUDENT WHERE IS_COMMIT = 0;"
    
    cur.execute(sql_isnt_commit)

    students = cur.fetchall()
    conn.commit()
    #遍历没有交作业的人

    if not students:
        print("\n\n作业交齐了!!\n\n")
    else:
        print("以下学生没有交作业:")
        count = 0
        for badguy in students:
            count = count +1
            print("%s\n"%badguy[1])
        print("\n共有 %d 人\n\n"%count)

    conn.close()
    

def zip_the_text():
    global number
    base_name = '计科1182-实验%d'%number
    format = 'zip'
    global dir
    root_dir = dir
    dir = shutil.make_archive(base_name=base_name,format=format,root_dir=root_dir)

    print("="*20)
    print("\n已将文件压缩在：%s\n"%dir)
    print('='*20)
    return dir


def send_the_text():
    
    # 创建一个带附件的邮件消息对象
    message = MIMEMultipart()
    
    # 创建文本内容
    text_content = MIMEText('计科1182-实验-%d'%number, 'plain', 'utf-8')
    message['Subject'] = Header('本月数据', 'utf-8')
    # 将文本内容添加到邮件消息对象中
    message.attach(text_content)
    """
    # 读取文件并将文件作为附件添加到邮件消息对象中
    with open('D:\\学习demo\\python\\文件批量改名\\texts\\计科1182-实验%d.zip'%number, 'rb') as f:
        xls = MIMEText(f.read(), 'base64', 'utf-8')
        xls['Content-Type'] = 'application/vnd.ms-excel'
        xls['Content-Disposition'] = 'attachment; filename=month-data.xlsx'          D:\学习demo\python\文件批量改名\texts
        message.attach(xls)
    """
    mp3part = MIMEApplication(open(dir+'\\计科1182-实验%d.zip'%number,'rb').read())
    mp3part.add_header('Content-Disposition', 'attachment', filename='实验.zip')
    message.attach(mp3part)

    # 创建SMTP对象
    smtper = SMTP('smtp.163.com')
    # 开启安全连接
    # smtper.starttls()
    sender = 'cjc617028197@163.com'
    receivers = '15360702785@163.com'

    message['from'] = sender
    message['to'] = receivers
    # 登录到SMTP服务器
    # 请注意此处不是使用密码而是邮件客户端授权码进行登录
    # 对此有疑问的读者可以联系自己使用的邮件服务器客服

    smtper.login(sender, 'cjc520')
    # 发送邮件
    smtper.sendmail(sender, receivers, message.as_string())
    # 与邮件服务器断开连接
    smtper.quit()
    print('发送完成!')

def notify_function(num):
    
    numbers = {

        1 : modify_name,
        2 : how_many_students_commit,
        3 : zip_the_text,
        4 : send_the_text,
        5 : restore,
    }

    method = numbers.get(num,'erro')
    if method:
        method()




def main():
    
    while(True):
        print('='*20)
       
        print('1.文件改名  2.查看未提交的学生  3.压缩文件  4.发送邮箱  5.将数据库重置并退出\n')
        num = int(input("你的选择:"))
        notify_function(num=num)
        if num == 5:
            print("\n程序结束")
            break


def restore():
    cur,conn = get_cur()
    sql_cmd = 'UPDATE STUDENT SET IS_COMMIT = 0;'
    cur.execute(sql_cmd)
    conn.commit()
    conn.close()
    

    



if __name__=="__main__":
    get_dir()
    main()




