#!_*_coding:utf8_*_
import xlwt
import time



def xls_input(count,info):
    workbook = xlwt.Workbook(encoding='utf-8')
    booksheet = workbook.add_sheet('Sheet1')
    booksheet.write(0,0,'编号')
    booksheet.write(0,1,'流程编号')
    booksheet.write(0,2,'IDC编号')
    booksheet.write(0,3,'机房')
    booksheet.write(0,4,'用户')
    booksheet.write(0,5,'人数')
    booksheet.write(0,6,'入时间')
    booksheet.write(0,7,'出时间    ')
    booksheet.write(0,8,'时间')
    booksheet.write(0,9,'事由类型及说明')
    booksheet.write(0,10,'录入时间')
    booksheet.write(0,11,'其他')
    x = 1
    for i in count:
        booksheet.write(x, 0, info[i]['编号'])
        booksheet.write(x, 1, info[i]['流程号'])
        booksheet.write(x, 2, info[i]['IDC编号'])
        booksheet.write(x, 3, info[i]['机房'])
        booksheet.write(x, 4, info[i]['用户'])
        booksheet.write(x, 5, info[i]['人数'])
        booksheet.write(x, 6, info[i]['入时间'])
        booksheet.write(x, 7, info[i]['出时间'])
        booksheet.write(x, 8, str(info[i]['入时间']) + '--' + str(info[i]['出时间']))
        booksheet.write(x, 9, info[i]['事由类型及说明'])
        booksheet.write(x, 10, time.strftime('%Y-%m-%d', time.localtime(time.time())))
        booksheet.write(x, 11, info[i]['其他'])
        x += 1
    workbook.save('people.xls')
