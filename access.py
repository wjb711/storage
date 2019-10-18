import pypyodbc
import pandas as pd
import os
import easygui as g
#定义conn
def mdb_conn(db_name, password = ""):
    """
    功能：创建数据库连接
    :param db_name: 数据库名称
    :param db_name: 数据库密码，默认为空
    :return: 返回数据库连接
    """
    str = 'Driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=" + db_name
    conn = pypyodbc.win_connect_mdb(str)

    return conn

#增加记录
def mdb_add(conn, cur, sql):
    """
    功能：向数据库插入数据
    :param conn: 数据库连接
    :param cur: 游标
    :param sql: sql语句
    :return: sql语句是否执行成功
    """
    try:
        cur.execute(sql)
        conn.commit()
        return True
    except:
        return False

#删除记录
def mdb_del(conn, cur, sql):
    """
    功能：向数据库删除数据
    :param conn: 数据库连接
    :param cur: 游标
    :param sql: sql语句
    :return: sql语句是否执行成功
    """
    try:
        cur.execute(sql)
        conn.commit()
        return True
    except:
        return False

#修改记录
def mdb_modi(conn, cur, sql):
    """
    功能：向数据库修改数据
    :param conn: 数据库连接
    :param cur: 游标
    :param sql: sql语句
    :return: sql语句是否执行成功
    """
    try:
        cur.execute(sql)
        conn.commit()
        return True
    except:
        return False

#查询记录
def mdb_sel(cur, sql):
    """
    功能：向数据库查询数据
    :param cur: 游标
    :param sql: sql语句
    :return: 查询结果集
    """
    try:
        cur.execute(sql)
        return cur.fetchall()
    except:
        return []

if __name__ == '__main__':
    file=g.fileopenbox('第一步：请选择access数据库')
    print(file)
    pathfile = file
    tablename=g.enterbox(msg="第二步，请输入要导出的精确表名",title="表名")
    lines=g.enterbox(msg="最后一步，请确认要导出的行数， 默认导出100行， 可以修改，越大越慢，可能会死机",title="行数", default=100)
    #tablename = 'IC卡部设备维修人员每日工作日志'

    conn = mdb_conn(pathfile)
    cur = conn.cursor()
    


##    #增
##    sql = "Insert Into " + tablename + " Values (33, 12, '天津', 0)"
##    if mdb_add(conn, cur, sql):
##       print("插入成功！")
##    else:
##       print("插入失败！")
##
##    #删
##    sql = "Delete * FROM " + tablename + " where id = 32"
##    if mdb_del(conn, cur, sql):
##       print("删除成功！")
##    else:
##       print("删除失败！")
##
##    #改
##    sql = "Update " + tablename + " Set IsFullName = 1 where ID = 33"
##    if mdb_modi(conn, cur, sql):
##       print("修改成功！")
##    else:
##       print("修改失败！")

    #查
    #sql = "SELECT * FROM " + tablename + " where id > 10"
    sql = "SELECT top "+lines+ " * FROM " + tablename +" order by ID desc"
    #sql="select * from IC卡部设备维修人员每日工作日志 order by top desc,id desc"
    #sql = "SELECT * FROM 岗位 where"
    #sql="select * from 岗位 where false"
    sel_data = mdb_sel(cur, sql)
    len0=len(sel_data[0])
    list0=[]
    for i in range(len0):
        list0.append('列表'+str(i+1))
    print(type(sel_data),sel_data,len(sel_data))
    
    #column=['ID','所属部门','工序','岗位']

    test=pd.DataFrame(columns=list0,data=sel_data)
##
    test.to_csv(tablename+'.csv',encoding="utf_8_sig")
    os.system(tablename+'.csv')

    cur.close()    #关闭游标
    conn.close()   #关闭数据库连接
