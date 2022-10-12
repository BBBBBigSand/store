# -*- codeing = utf-8 -*-
# @Author : BBBBBigSand
# @File : sqlite应用.py
# @Software : PyCharm
import xlrd
import sqlite3

#连接excel文件
def ex(i):
    workbo = xlrd.open_workbook_xls('联想者.xls')
    wb = workbo.sheet_by_name(sheet_name='sheet1')
    row_d = wb.row_values(i)
    return row_d

#创建sqlite库
def create():
    conn = sqlite3.connect('data.db')
    c = conn.cursor()
    print("数据库打开成功")
    return c, conn


#为数据库创建表
def crea_table():
    c = []
    c = create()
    c[0].execute('''CREATE TABLE DATA
           (SEQ   primary key,
           NAME,       
           ID,             
           CONT,           
           TI,            
           AUCTIONSKU,     
           PROSIZE        );''')
    print("数据表创建成功")
    c[1].commit()


#插入excel中的数据
def ins():
    c = []
    c = create()
    for i in range(1, 601):
        row_d = ex(i)
        #row_d[0].lstrip("\t")
        #print(row_d)
        c[0].execute("INSERT INTO DATA (SEQ,NAME,ID,CONT,TI,AUCTIONSKU,PROSIZE) \
            VALUES (?,?,?,?,?,?,? ) ", (int(row_d[0]), row_d[1], int(row_d[2]), row_d[3], row_d[4], row_d[5], row_d[6]))
    c[1].commit()
    print("数据插入成功")
    cursor = c[0].execute("SELECT SEQ,NAME,ID,CONT,TI,AUCTIONSKU,PROSIZE from DATA ")
    for row in cursor:
        print(row)
    c[1].commit()
    print("数据操作成功")

#查询第40号数据
def select():
    c = []
    c = create()
    cursor = c[0].execute("SELECT SEQ,NAME,ID,CONT,TI,AUCTIONSKU,PROSIZE from DATA where SEQ =40")
    for row in cursor:
        print(row)
    c[1].commit()
    print("数据查询成功")


#更新第50行的数据
def update():
    c = []
    c = create()
    print("修改前：")
    cursor = c[0].execute("SELECT SEQ,NAME,ID,CONT,TI,AUCTIONSKU,PROSIZE from DATA where SEQ =50")
    for row in cursor:
        print(row)
    c[0].execute("UPDATE DATA set ID=11111111111 where SEQ=50")
    c[1].commit()
    print("修改后：")
    cursor = c[0].execute("SELECT SEQ,NAME,ID,CONT,TI,AUCTIONSKU,PROSIZE from DATA where SEQ =50")
    for row in cursor:
        print(row)
    print("数据操作成功")


#删除第10行的数据
def delete():
    c = []
    c = create()
    c[0].execute("DELETE from DATA where SEQ=10")
    c[1].commit()
    print("删除后：")
    cursor = c[0].execute("SELECT SEQ,NAME,ID,CONT,TI,AUCTIONSKU,PROSIZE from DATA where SEQ =10")
    for row in cursor:
        print(row)
    print("数据操作成功")


if __name__=="__main__":
    create()
    crea_table()
    ins()
    select()
    update()
    delete()
    c = create()
    c[1].close()
