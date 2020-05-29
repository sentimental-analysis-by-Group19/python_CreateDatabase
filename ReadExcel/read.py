#-*- codeing = utf-8 -*-
#@Time : 2020/4/28 17:51
#@Author : 于璐
#@File : readExcel.py
#@Software : PyCharm


import sqlite3
import xlrd


'''---------------------------------main函数-----------------------------------'''
def main():
    datalist = getDataxsl()
    dbpath = "data2.db"
    saveData2DB(datalist, dbpath)

'''-----------------------------------------------------------------------------'''



'''RankName__职位姓名
Title__留言标题
Tag1__留言标签1
Tag2__留言标签2
CDate__留言日期  不知道如何日期属于什么类型，就没有放入，而且感觉日期没啥用
Content__留言内容
RName__回复人
RContent__回复内容
RDate__回复日期
PleasedLev__满意程度
ServeScore__解决程度分数
MannerScore__办理态度分数
SpeedScore__办理速度分数
Assass__是否自动好评
AssassContent_评价内容
AssassDate_评价日期
'''
'''---------------------------------创建数据库-----------------------------------'''
def init_db(dbpath):
    sql = '''
        create table data2
        (
            RankName text,
            Title text,
            Tag1 text,
            Tag2 text,
            Content text,
            RName text,
            RContent text,
            PleasedLev text,
            ServeScore integer,
            MannerScore integer,
            SpeedScore integer,
            Assass text,
            AssassContent text  
        )

    '''
    # 创建数据库
    conn = sqlite3.connect(dbpath)
    cursor = conn.cursor()
    cursor.execute(sql)
    conn.commit()
    conn.close()
'''-----------------------------------------------------------------------------'''



'''---------------------------------解析excel,获取excel内容-----------------------------------'''
def getDataxsl():
    #打开文件
    workbook = xlrd.open_workbook(r'DATA.xls')
    #获取所有sheet
    sheet_name = workbook.sheet_names()[0]
    sheet = workbook.sheet_by_name(sheet_name)
    datalist = []


    #获取一行的内容
    for i in range(1,sheet.nrows):#sheet.nrows
        data = []
        for j in range(0,sheet.ncols):
            if j == 4 or j == 8 or j == 15: #没有将日期显示出来，因为不知道不会，感觉日期也没啥用
                continue
            data.append(str(sheet.cell(i,j).value))

        datalist.append(data)
    return datalist
'''-----------------------------------------------------------------------------'''


'''---------------------------------将列表中的信息放入数据库-----------------------------------'''
def saveDataDB(datalist, dbpath):
    init_db(dbpath)
    conn = sqlite3.connect(dbpath)
    cur = conn.cursor()

    #读取
    workbook = xlrd.open_workbook("DATA.xls")
    # 获取sheet
    sheet_name = workbook.sheet_names()[0]
    sheet = workbook.sheet_by_name(sheet_name)

    for data in datalist:
        for index in range(len(data)):# 行
            data[index] = '"'+data[index] +'"0'
            sql = '''
            insert into data2(
            RankName, Title, Tag1, Tag2, Content, RName, RContent,PleasedLev, ServeScore,MannerScore,SpeedScore,Assass,AssassContent)
            values(%s)'''%",".join(data)
    print(sql)


'''-----------------------------------------------------------------------------'''

def saveData2DB(datalist,dbpath):
    init_db(dbpath)
    conn = sqlite3.connect(dbpath)
    cur = conn.cursor()

    for data in datalist:
        for index in range(len(data)):
            # if index == 4 or index == 5:
            #     continue
            data[index] = "'"+data[index]+"'"
        sql = '''
                insert into data2 (
                 RankName, Title, Tag1, Tag2, Content, RName, RContent,PleasedLev, ServeScore,MannerScore,SpeedScore,Assass,AssassContent) 
                values(%s)'''%",".join(data)
        # print(sql)
        cur.execute(sql)
        conn.commit()
    cur.close()
    conn.close()
    print("导入数据结束")


if __name__ == '__main__':
    main()
