# -*- coding: utf-8 -*-
import pymysql
import xlrd
import datetime

def read_excel():
    # 输入文件路径
    ExcelPath = input('input a Excel Path:')
    print(ExcelPath)
    # 输入sheet名
    ExcelSheetName = input('input a ExcelSheetName:')
    print(ExcelSheetName)
    # 打开文件C:\JITAutoTest\Jit_Test.xls
    workbook = xlrd.open_workbook(ExcelPath)
    # 获取所有sheet名
    # print(workbook.sheet_names())
    # 获取第二个sheet的名字
    # sheet2_name = workbook.sheet_names()[1]
    # print(sheet2_name)
    # 根据sheet索引或者名称获取sheet内容
    sheetname = workbook.sheet_by_name(ExcelSheetName)
    # sheet的名称，行数，列数
    print(sheetname.name, sheetname.nrows, sheetname.ncols)
    rr = sheetname.nrows
    for index in range(0, rr):
        # 获取第2行内容
        rows = sheetname.row_values(index)
        # 获取第三列内容
        # cols = sheetname.col_values(2)
        print(rows)

    # print(cols)
    # 获取单元格内容的三种方法第二行第一列
    # print(sheetname.cell(1, 0).value)
    # print(sheet2.cell(1, 0).value.encode('utf-8'))

    # 获取单元格内容的数据类型
    # print(sheetname.cell(1, 3).ctype)
    # print(sheetname.cell(1, 4).ctype)

    # ctype :  0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error

    if (sheetname.cell(1, 4).ctype == 3):
        date_value = xlrd.xldate_as_tuple(sheetname.cell_value(1, 4), workbook.datemode)
        date_tmp = datetime.date(*date_value[:3]).strftime('%Y/%m/%d')
        print(date_value)
        print(date_tmp)


def sql_excel():
    # 输入数据库localhost
    dbip = input('input database IP:')
    print(dbip)
    # 输入数据库username  root
    dbusername = input('input database username:')
    print(dbusername)
    # 输入数据库password  11111111
    dbpwd = input('input database password:')
    print(dbpwd)
    # 输入数据库name  galaxy
    dbname = input('input database name:')
    print(dbname)
    # 输入数据库tablename  df1
    dbtablename = input('input database tablename:')
    print(dbtablename)
    # 输入数据库表字段项  avalue1,avalue2,avalue3,avalue4,avalue5
    keys = input('input database table keys:')
    print(keys)


    # 打开数据库连接
    db = pymysql.connect(dbip, dbusername, dbpwd, dbname)
    pin = dbtablename + "(" + keys + ")"
    print(pin)

    keyvalue1 = ['Mac', 'Mohan', 'dd', 'M', 'test']
    # 将list转化成元组
    keyvalue=tuple(keyvalue1)

    print(keyvalue)
    # 使用cursor()方法获取操作游标
    cursor = db.cursor()
    # SQL 插入语句
    sql = "INSERT INTO "+pin+" VALUES ('%s', '%s', '%s', '%s', '%s' )" % (keyvalue)
    print(sql)
    try:
        # 执行sql语句
        cursor.execute(sql)
        # 执行sql语句
        db.commit()
        print("insert ok")
    except BaseException as e:
        # 发生错误时回滚
        print(e)
        db.rollback()
    # 关闭数据库连接
    db.close()


if __name__ == '__main__':
    sql_excel()

    # zhi = ('avalue1', 'avalue2', 'avalue3', 'avalue4', 'avalue5')
    # zhi1=",".join(tuple(zhi))