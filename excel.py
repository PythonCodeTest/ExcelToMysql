# -*- coding: utf-8 -*-
import os
import re
import datetime
import pymysql
import xlrd

regg = 1
# 输入数据库localhost
dbip = input('input database IP:')
print(dbip)
dbip = '172.16.24.53'
# 输入数据库username  root
dbusername = input('input database username:')
print(dbusername)
dbusername = 'root'
# 输入数据库password  11111111
dbpwd = input('input database password:')
print(dbpwd)
dbpwd = '11111111'
# 输入数据库name  galaxy
dbname = input('input database name:')
print(dbname)
dbname = 'galaxy'
# 输入数据库tablename  df1
dbtablename = input('input database tablename:')
print(dbtablename)
dbtablename = 'df1'
# 输入数据库表字段项  avalue1,avalue2,avalue3,avalue4,avalue5
keys = input('input database table keys:')
print(keys)
keys = 'avalue1,avalue2,avalue3,avalue4,avalue5'

# 输入文件路径 C:\JITAutoTest\Jit_Test.xls
ExcelPath = input('input a Excel Path:')
print(ExcelPath)
ExcelPath = 'C:\JITAutoTest\Jit_Test.xls'
# 输入sheet名   Sheet2
ExcelSheetName = input('input a ExcelSheetName:')
print(ExcelSheetName)
ExcelSheetName = 'Sheet2'


def read_excel():
    # 打开文件
    workbook = xlrd.open_workbook(ExcelPath)
    # 根据sheet索引或者名称获取sheet内容
    sheetname = workbook.sheet_by_name(ExcelSheetName)
    # sheet的名称，行数，列数
    # print(sheetname.name, sheetname.nrows, sheetname.ncols)
    rr = sheetname.nrows

    for index in range(0, rr):
        # 获取第2行内容
        rows = sheetname.row_values(index)
        # print(rows)
        keyvalue1 = rows
        sql_excel(keyvalue1)


def sql_excel(keyvalue1):
    # 打开数据库连接
    # db = pymysql.connect(dbip, dbusername, dbpwd, dbname, charset='utf8')
    pin = dbtablename + "(" + keys + ")"
    # print(pin)
    lists = []
    nkeysnn = list(keys.split(",")).__len__()
    for index in range(0, nkeysnn):
        lists.append("'%s'")
    # print(lists)

    strlist = ','.join(lists)
    # print(strlist)
    pin2 = " VALUES (" + strlist + ")"
    # print(pin2)

    # 将list转化成元组
    keyvalue = tuple(keyvalue1)
    # print(keyvalue)

    # 使用cursor()方法获取操作游标
    cursor = db.cursor()

    # SQL 插入语句
    sql = "INSERT INTO " + pin + pin2 % (keyvalue)
    # print(sql)
    '''
    # 修改数据
    data1 = (8888, '13512345678')
    sql1 = "UPDATE trade SET saving = %.2f WHERE account = '%s' "% data1
    print(sql1)

    # 查询数据
    data2 = ('13512345678',)
    sql2 = "SELECT name,saving FROM trade WHERE account = '%s' "% data2

    for row in cursor.fetchall():
        print("Name:%s\tSaving:%.2f" % row)

    print(sql2)
    print('共查找出', cursor.rowcount, '条数据')

    # 删除数据
    data3 = ('13512345678', 1)
    sql3 = "DELETE FROM trade WHERE account = '%s' LIMIT %d" % data3
    print(sql3)
    print('成功删除', cursor.rowcount, '条数据')
    
    '''
    try:
        # 执行sql语句
        cursor.execute(sql)
        # 执行sql语句
        db.commit()

        # print("insert ok")
    except BaseException as e:
        # print(e)
        # 发生错误时回滚
        db.rollback()

    # 关闭数据库连接
    # db.close()


def cheack():
    # 打开文件
    workbook = xlrd.open_workbook(ExcelPath)
    # 根据sheet索引或者名称获取sheet内容
    sheetname = workbook.sheet_by_name(ExcelSheetName)
    # sheet的名称，行数，列数
    # print(sheetname.name, sheetname.nrows, sheetname.ncols)
    # 获取所有sheet名
    sheetsname = workbook.sheet_names()
    # print(workbook.sheet_names())
    nn = ExcelSheetName
    if nn in sheetsname:
        print('输入Sheet页名称存在')
    else:
        print('输入Sheet页名称不存在')
        regg = 0

    rr = sheetname.nrows
    cc = sheetname.ncols
    nkeys = list(keys.split(",")).__len__()
    # print(cc)
    # print(nkeys)

    if cc == nkeys:
        print('字段数和数据列数匹配')
    else:
        print('字段数和数据列数不匹配')
        regg = 0

    if os.path.exists(ExcelPath):
        print('excel文件存在')
    else:
        print('excel文件不存在')
        regg = 0
    # 判断表是否存在
    table_name = dbtablename
    exist_of_table(table_name)


def exist_of_table(table_name):
    try:
        conn = pymysql.connect(dbip, dbusername, dbpwd, dbname)
        cur = conn.cursor()
        sql_show_table = 'SHOW TABLES;'
        cur.execute(sql_show_table)
        tables = [cur.fetchall()]
        conn.close()
        # print(tables)
        table_list = re.findall('(\'.*?\')', str(tables))
        table_list = [re.sub("'", '', each) for each in table_list]
        if table_name in table_list:
            print('表存在')
            return 1
        else:
            print('表不存在')
            regg = 0
            return 0
    except BaseException:
        print('数据库连接失败,请检查配置项')


if __name__ == '__main__':
    db = pymysql.connect(dbip, dbusername, dbpwd, dbname, charset='utf8')
    cheack()
    if regg == 1:
        nowTime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(nowTime)
        read_excel()
    else:
        print('数据库连接失败,请检查配置项')

    lastTime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(lastTime)
    db.close()


        # keyvalue1 = ['Function', 'ExcelPath', 'SheetI', 'SheetName', 'ttt']
        # sql_excel(keyvalue1)
        # keyvalue2 = ['验证1', 'test1', 'a1', '结果1', '2017/01/01']
        # sql_excel(keyvalue2)
