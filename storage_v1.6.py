import easygui as g
#xlrd读取是excel文件的
import xlrd,os,getpass
#xlutils是用来修改写入excel文件的
from xlutils.copy import copy
#datetime，取时间的
import datetime,time
from barcode.writer import ImageWriter
from barcode.codex import Code39
from PIL import Image,ImageDraw,ImageFont
from io import StringIO
import PIL,PIL.ImageOps
import pandas as pd

#定义now的格式
def now():
    return datetime.datetime.now().strftime('%Y%m%d %H:%M:%S')
def now1():
    return datetime.datetime.now().strftime('%Y%m%d%H%M%S')[2:]
#print (now()) 
#print (now())
def bar(barname,comments0,comments1,comments2,comments3):
    filename='image2'
    imagewriter = ImageWriter()
            
                #保存到图片中
                # add_checksum : Boolean   Add the checksum to code or not (default: True)
    ean = Code39(barname, writer=imagewriter, add_checksum=False)
                # 不需要写后缀，ImageWriter初始化方法中默认self.format = 'PNG'
            #    print ('保存到image2.png')
    ean.save(filename)
            
    img = Image.open(filename+'.png')
    img=PIL.ImageOps.invert(img)
    imgSize=img.size
    x,y=img.size
    print(img.size)
    box = (-50,120,x+50,(x+100)/9*4+120)
    img = img.crop(box)
    img=PIL.ImageOps.invert(img)
        #print '展示image2.png'
        #img.show()

    #img=img.resize((720,640))
    #box=(0,320,720,640)
    #img=img.crop(box)
    #imgSize=img.size
    print ('imgSize is:',imgSize)
    font = ImageFont.truetype('simhei.ttf', int((imgSize[0])*0.05))
    font1 = ImageFont.truetype('simhei.ttf', int((imgSize[0])*0.1))
    draw = ImageDraw.Draw(img)
    draw.text((imgSize[0]*0.05, (imgSize[1]*0.3)), comments0, (0,0,0),font=font)
    draw.text((imgSize[0]*0.15, (imgSize[1]*0.75)), comments1, (0,0,0),font=font)
    draw.text((imgSize[0]*0.65, (imgSize[1]*0.55)), comments2, (0,0,0),font=font)
    draw.text((imgSize[0]*0.15, (imgSize[1]*0.55)), comments3, (0,0,0),font=font)
    draw.text((imgSize[0]*0.45, (imgSize[1]*0.65)), barname, (0,0,0),font=font1)
    img.show()
    #time.sleep(1)

def switch(choice):
    if choice=='信息':
        pic()
    if choice=='入库':
        check_in()
        
    if choice=='出库':
        check_out()
    if choice=='新增':
        new_material()
        
    if choice=='导出':
        daochu()
    if choice=='搜索':
        search()
    #if choice=='库存':
    #    kucun() 
    if choice=='移库':
        move()
    if choice=='盘点':
        pandian()
    if choice=='预警':
        warning()
    if choice=='历史':
        history()
    if choice=='条码':
        generagteBarCode()
        
    return 'end'

def exist(x):
    #beijian = xlrd.open_workbook(workbook)
    with xlrd.open_workbook(workbook) as beijian:
        table0 = beijian.sheet_by_name('data')
        nrows = table0.nrows
        mode=0
        
             
        #历遍主数据，如果查到有一样的，mode取1，没查到取0，查到时记录行号i0
        for i in range(nrows ):
            #print (table.row_values(i)[0])
            
            if table0.row_values(i)[0]==x:
                #print ('same',rk)
                return True
                break
            else:
                continue
        if mode==0:
            g.msgbox('无此物编,或请新增 ', title='录入错误')
            menu()

def pic():
#print ('入库')
    #这里其实可以使用条码枪
    rk = g.enterbox("请输入物品编码：\n",title="查询单个物品信息")
    #print ('rk',rk)
    if rk==None:
        menu()
        
    with xlrd.open_workbook(workbook) as beijian:
        table0 = beijian.sheet_by_name('data')
        nrows = table0.nrows
        mode=0
        #历遍主数据，如果查到有一样的，mode取1，没查到取0，查到时记录行号i0
        for i in range(nrows ):
            #print (table.row_values(i)[0])
            
            if table0.row_values(i)[0]==rk:
                #print ('same',rk)
                mode=1
                i0=i
            else:
                pass
        if mode==0:
            g.msgbox('无此物编,或请新增 '+rk, title='录入错误')
            menu()
        else:
            information=table0.row_values(i0)[1]
            #print (information)
            ok_button='图片/数量/信息/库位'
            g.msgbox(rk+"\n"+'信息: '+table0.row_values(i0)[1]+"\n"+'当前库存: '+str(int(table0.row_values(i0)[4]))+"\n"+'库位:'+table0.row_values(i0)[2],ok_button,image='image/'+rk+'.jpg')

                
            menu()


def history():
    record=g.boolbox(msg='历史?', title=' ', choices=('单品历史', '整体历史限500条,更多请选择【导出】功能'), image=None, default_choice='Yes', cancel_choice='No')
    with xlrd.open_workbook(workbook) as beijian:
        table0 = beijian.sheet_by_name('history')
        nrows = table0.nrows
        list=[]
        if record==False:

            
            for i in range(0,500):
                j=nrows-1-i
                if j >=0:
                    #print (j)
                    list.append('\n'+table0.row_values(j)[0]+table0.row_values(j)[1]+table0.row_values(j)[2]+\
                                ' 数量：'+str(table0.row_values(j)[3])+table0.row_values(j)[4]+table0.row_values(j)[5]+' 时间：'\
                                +table0.row_values(j)[6]+' 库位:'+table0.row_values(j)[7]+' 录入：'+table0.row_values(nrows-1-i)[8])
            
        else:
            rk = g.enterbox("请输入物品编码：\n",title="历史")
            if exist(rk) ==True:
                for i in range(nrows):
                    if table0.row_values(nrows-1-i)[0]==rk:
                        list.append('\n'+table0.row_values(nrows-1-i)[0]+table0.row_values(nrows-1-i)[1]+table0.row_values(nrows-1-i)[2]
                                    +' 数量：'+str(int(table0.row_values(nrows-1-i)[3]))
                                    +table0.row_values(nrows-1-i)[4]
                                    +table0.row_values(nrows-1-i)[5]
                                    +' 时间：'+table0.row_values(nrows-1-i)[6]
                                    +' 库位：'+table0.row_values(nrows-1-i)[7]
                                    +' 录入：'+table0.row_values(nrows-1-i)[8]
                                    )

        S=''
        for s in list:
            S=S+s+'\n'
            #print ('S is:',S)
            #g.msgbox(S,title='搜索结果')
        g.msgbox(S)
        #g.msgbox(list)
                    
                    
                        
                    
            
        
        menu()

def generagteBarCode():
    rk = g.enterbox("请输入要生成的条码内容：\n",title="条码生成打印")
    barname=rk
    print ('rk is:',rk)
    if rk==None:
        menu()
    #beijian = xlrd.open_workbook(workbook)
    with xlrd.open_workbook(workbook) as beijian:
        table0 = beijian.sheet_by_name('data')
        nrows = table0.nrows
        for i in range(nrows ):
            if table0.row_values(i)[0]==rk:
                i0=i
                break
        try:
            print (i0)
            comments=table0.row_values(i0)[1]
            print (table0.row_values(i0)[1])
            filename='image2'
            imagewriter = ImageWriter()
            
                #保存到图片中
                # add_checksum : Boolean   Add the checksum to code or not (default: True)
            ean = Code39(barname, writer=imagewriter, add_checksum=False)
                # 不需要写后缀，ImageWriter初始化方法中默认self.format = 'PNG'
            #    print ('保存到image2.png')
            ean.save(filename)
            
            img = Image.open(filename+'.png')
            imgSize=img.size
            img=PIL.ImageOps.invert(img)
            imgSize=img.size
            x,y=img.size
            print(img.size)
            box = (-50,120,x+50,(x+100)/9*4+120)
            img = img.crop(box)
            img=PIL.ImageOps.invert(img)
            #box=(0,320,720,640)
            #img=img.crop(box)
            #imgSize=img.size
            print ('imgSize is:',imgSize)
            font = ImageFont.truetype('simhei.ttf', int((imgSize[0])*0.05))
            font1 = ImageFont.truetype('simhei.ttf', int((imgSize[0])*0.1))
            draw = ImageDraw.Draw(img)
            draw.text((imgSize[0]*0.05, (imgSize[1]*0.55)), comments, (0,0,0),font=font)
            draw.text((imgSize[0]*0.35, (imgSize[1]*0.75)), barname, (0,0,0),font=font1)
        except:
            filename='image2'
            imagewriter = ImageWriter()
            
                #保存到图片中
                # add_checksum : Boolean   Add the checksum to code or not (default: True)
            ean = Code39(barname, writer=imagewriter, add_checksum=False)
                # 不需要写后缀，ImageWriter初始化方法中默认self.format = 'PNG'
            #    print ('保存到image2.png')
            ean.save(filename)
            img = Image.open(filename+'.png')
            

            #设置要裁剪的区域

            #img=img.crop(box) 
            #imgSize=img.size
            img=PIL.ImageOps.invert(img)
            imgSize=img.size
            x,y=img.size
            print(img.size)
            box = (-50,120,x+50,(x+100)/9*4+120)
            img = img.crop(box)
            img=PIL.ImageOps.invert(img)
            print ('imgSize is:',imgSize)
            #font = ImageFont.truetype('simhei.ttf', int((imgSize[0])*0.04))
            #draw = ImageDraw.Draw(img)
            
            
        #print ('展示image2.png')
        img.show()
        time.sleep(1)
        #os.remove(filename+'.png')
        menu()

def move1():
    rk = g.enterbox("请输入物品编码：\n",title="移库")
    if exist(rk) ==True:
        with xlrd.open_workbook(workbook) as beijian:
            table0 = beijian.sheet_by_name('data')
            nrows = table0.nrows
            for i in range(nrows ):
                if table0.row_values(i)[0]==rk:
                    i0=i
                    break
            print (i0)
            newkuwei=g.enterbox(msg=table0.row_values(i)[0]+'\n'+'请输入新库位，默认值为原库位',title='移库',default=table0.row_values(i0)[2])
            beijian1 = copy(beijian)
            beijian1.get_sheet('data').write(i0,2,newkuwei)
            try:
                beijian1.save(workbook)
                g.msgbox("修改成功")
                menu()
            except:
                g.msgbox("保存失败，有人已打开excel表 \n 请关闭beijian.xls,再试")
    else:
        menu()

def move():
    button=g.buttonbox("库位操作",choices=("单品移库","新增库位","库位替换"))
    if button =="单品移库":
        rk = g.enterbox("请输入物品编码：\n",title="移库")
        if exist(rk) ==True:
            with xlrd.open_workbook(workbook) as beijian:
                table1 = beijian.sheet_by_name('locator')
                table1_nrows = table1.nrows
                table0 = beijian.sheet_by_name('data')
                nrows = table0.nrows
                list0=[]
                for i in range(table1_nrows ):
                    list0.append(table1.row_values(i)[0])
                
                for i in range(nrows ):
                    if table0.row_values(i)[0]==rk:
                        i0=i
                        break
                print (i0)
                newkuwei=g.enterbox(msg=table0.row_values(i)[0]+'\n'+'请输入新库位，默认值为原库位',title='移库',default=table0.row_values(i0)[2])
                if newkuwei in list0:
                    print ('yes, within kuwei')
                    beijian1 = copy(beijian)
                    beijian1.get_sheet('data').write(i0,2,newkuwei)
                    try:
                        beijian1.save(workbook)
                        g.msgbox("修改成功")
                        menu()
                    except:
                        g.msgbox("保存失败，有人已打开excel表 \n 请关闭beijian.xls,再试")
                else:
                    print ('no, out of kuwei')
                    g.msgbox("无此库位,请查正库位或新增库位")
                    menu()
    else:
        if button =="新增库位":
            st = g.enterbox("请输入新的库位代码：\n")
            with xlrd.open_workbook(workbook) as beijian:
                table1 = beijian.sheet_by_name('locator')
                table1_nrows = table1.nrows
                
                beijian1 = copy(beijian)
                
                beijian1.get_sheet('locator').write(table1_nrows,0,st)
                beijian1.save(workbook)
                g.msgbox(st+'已新增')
                menu()
        if button =="库位替换":
            st = g.enterbox("高级操作，请输入密码：\n")
            if st=='password':

                with xlrd.open_workbook(workbook) as beijian:
                    table1 = beijian.sheet_by_name('locator')
                    table1_nrows = table1.nrows
                    table0 = beijian.sheet_by_name('data')
                    nrows = table0.nrows
            
                    list0=[]
                    for i in range(table1_nrows ):
                        list0.append(table1.row_values(i)[0])
                    print ('list0 is:',list0)
                    #g.multchoicebox(msg="请选择你爱吃哪些水果?",title="",choices=("西瓜","香蕉","苹果","梨"))
                    old=g.choicebox(msg="请选择库位",title="",choices=list0)
                    new=g.choicebox(msg="请选择库位",title="",choices=list0)
                    print (old, new)
                    beijian1 = copy(beijian)
                    for i in range(nrows ):
                        if table0.row_values(i)[2]==old:
                            beijian1.get_sheet('data').write(i,2,new)
                    beijian1.save(workbook)
                    g.msgbox('已全部替换')
                    menu()
            else:
                g.msgbox('密码错误')
                menu()


        
def warning():
    yujing=g.boolbox(msg='预警?', title=' ', choices=('修改预警值', '预警报告'), image=None, default_choice='Yes', cancel_choice='No')
    if yujing==False:
        with xlrd.open_workbook(workbook) as beijian:
            table0 = beijian.sheet_by_name('data')
            nrows = table0.nrows
            mode=0
        #历遍主数据，如果查到有一样的，mode取1，没查到取0，查到时记录行号i0
            list0=[]
            for i in range(1,nrows ):
            #print (table.row_values(i)[0])
                if int(table0.row_values(i)[4])<int(table0.row_values(i)[5]):
            
            
                    list0.append([table0.row_values(i)[0],table0.row_values(i)[1],table0.row_values(i)[2],str(int(table0.row_values(i)[4])),str(int(table0.row_values(i)[5])),str(int(table0.row_values(i)[6])),str(int(table0.row_values(i)[6])-int(table0.row_values(i)[4]))])
            #S=''
            #for s in list0:
            #    S=S+s+'\n'
            #print ('S is:',S)
            #g.msgbox(S,title='搜索结果')
            print('list0:',list0)
            #print('S:',S)
            #list0=list0=[[1,2,3],[4,5,6],[7,8,9]]
            column=['物代ID','描述','库位','库存','最低预警','最高预警','待补足']

            test=pd.DataFrame(columns=column,data=list0)

            test.to_csv('./预警.csv',encoding="utf_8_sig")
            os.system('预警.csv')
            #g.msgbox(S)

    else:
        password=g.passwordbox(msg="请输入操作密码",title="捷德备件管理系统")
        if password=='password':
            wubian=g.enterbox("请输入精确物编：\n",title="盘点")
            with xlrd.open_workbook(workbook) as beijian:
                table0 = beijian.sheet_by_name('data')
                nrows = table0.nrows
                mode=0
            #历遍主数据，如果查到有一样的，mode取1，没查到取0，查到时记录行号i0
                list=[]
                for i in range(nrows ):
                #print (table.row_values(i)[0])
                    if wubian==table0.row_values(i)[0]:
                        print (wubian)
                        new_yujing=g.enterbox("请输入新最低预警值：\n",title=table0.row_values(i)[0],default=int(table0.row_values(i)[5]))
                        new_yujing_gao=g.enterbox("请输入新最高预警值：\n",title=table0.row_values(i)[0])
                        if new_yujing==None:
                            menu()
                        beijian1 = copy(beijian)

                        beijian1.get_sheet('data').write(i,5,int(new_yujing))
                        beijian1.get_sheet('data').write(i,6,int(new_yujing_gao))
                        try:
                            beijian1.save(workbook)
                            g.msgbox("修改成功")
                        except:
                            g.msgbox("保存失败，有人已打开excel表 \n 请关闭beijian.xls,再试")
    
    menu()
    
def pandian1():
    kuwei0=g.enterbox("请输入精确库位：\n",title="盘点")
    if kuwei0==None:
        menu()
        
    with xlrd.open_workbook(workbook) as beijian:
        table0 = beijian.sheet_by_name('data')
        nrows = table0.nrows
        list0=[]
        #历遍主数据，如果查到有一样的，mode取1，没查到取0，查到时记录行号i0
        for i in range(nrows ):
            #print (table.row_values(i)[0])
            
            if table0.row_values(i)[2]==kuwei0:
                #print ('same',rk)
                #list0.append(
                list0.append('\n'+table0.row_values(i)[0]+' '+table0.row_values(i)[1]+' 库位1: '+table0.row_values(i)[2]+' '+table0.row_values(i)[3]+'  库存:'+str(int(table0.row_values(i)[4]))+' 预警: '+str(int(table0.row_values(i)[5]))+' ')
            else:
                pass
        S=''
        for s in list0:
            S=S+s+'\n'
            #print ('S is:',S)
            #g.msgbox(S,title='搜索结果')
        g.msgbox(S)
        #g.msgbox(list0)
        menu()

def pandian():

        
    with xlrd.open_workbook(workbook) as beijian:
        table1 = beijian.sheet_by_name('locator')
        table1_nrows = table1.nrows
        table0 = beijian.sheet_by_name('data')
        nrows = table0.nrows
        
        list0=[]
        for i in range(table1_nrows ):
            list0.append(table1.row_values(i)[0])
        print ('list0 is:',list0)
        #g.multchoicebox(msg="请选择你爱吃哪些水果?",title="",choices=("西瓜","香蕉","苹果","梨"))
        kw=g.choicebox(msg="请选择库位",title="",choices=list0)
        list=[]
        for i in range(nrows ):
            if kw in table0.row_values(i)[2]:
                list.append('\n'+table0.row_values(i)[0]+' '+table0.row_values(i)[1]+' 库位: '+table0.row_values(i)[2]+' '+table0.row_values(i)[3]+'  库存:'+str(int(table0.row_values(i)[4]))+' 预警: '+str(int(table0.row_values(i)[5]))+' ')
        #print (list)
        S=''
        for s in list:
            S=S+s+'\n'
            #print ('S is:',S)
        g.msgbox(S,title='搜索结果')
        #历遍主数据，如果查到有一样的，mode取1，没查到取0，查到时记录行号i0
        #for i in range(nrows ):
            #print (table.row_values(i)[0])
            
        #    g.msgbox(i)
        menu()


def search():
    
    rk = g.enterbox("请输入关键词：\n",title="搜索")
    #print ('rk',rk)
    if rk==None:
        menu()
        
    with xlrd.open_workbook(workbook) as beijian:
        table0 = beijian.sheet_by_name('data')
        nrows = table0.nrows
        mode=0
        #历遍主数据，如果查到有一样的，mode取1，没查到取0，查到时记录行号i0
        list=[]
        
        for i in range(nrows ):
            #print (table.row_values(i)[0])
            
            if rk in table0.row_values(i)[1]:
                list.append('\n'+table0.row_values(i)[0]+' '+table0.row_values(i)[1]+' 库位: '+table0.row_values(i)[2]+' '+table0.row_values(i)[3]+'  库存:'+str(int(table0.row_values(i)[4]))+' 预警: '+str(int(table0.row_values(i)[5]))+' ')
            else:
                pass
        if list==[]:
            g.msgbox('没有结果，请尝试其它关键词',title='搜索结果')
        else:
            print ('list is :',list)
            S=''
            for s in list:
                S=S+s+'\n'
            #print ('S is:',S)
            g.msgbox(S,title='搜索结果')
        menu()

def kucun():
    
    
    with xlrd.open_workbook(workbook) as beijian:
        table0 = beijian.sheet_by_name('data')
        nrows = table0.nrows
        mode=0
        #历遍主数据，如果查到有一样的，mode取1，没查到取0，查到时记录行号i0
        list=[]
        for i in range(1,nrows ):
            #print (table.row_values(i)[0])
            
            
            list.append('\n'+table0.row_values(i)[0]+' '+table0.row_values(i)[1]+' 库位: '+table0.row_values(i)[2]+\
                        ' '+table0.row_values(i)[3]+'  库存:'+str(int(table0.row_values(i)[4]))+' 预警: '\
                        +str(int(table0.row_values(i)[5]))+' ')
        S=''
        for s in list:
            S=S+s+'\n'
            #print ('S is:',S)
            #g.msgbox(S,title='搜索结果')
        g.msgbox(S)

            
        menu()
#import image
#主界面，包含标题，信息，选项等内容
def daochu():
    #print ('hello')
    with xlrd.open_workbook(workbook) as beijian:
        beijian1 = copy(beijian)
        now1=datetime.datetime.now().strftime('%Y%m%d%H%M%S')
        beijian1.save(now1+'.xls')
        os.system(now1+'.xls')
        print ('end')
        menu()
    

def menu():
    title='捷德备件管理系统'
    msg='请问要执行哪类操作？'
    ok_button='确定？'
    #image='python.png'
    #g.msgbox(msg,title)
    choices=['搜索','信息','入库','出库','新增','移库','历史','导出','预警','盘点','条码','退出']
    #choice=g.choicebox(msg,title,choices)
    #g.msgbox(choice,title,ok_button,image)
    #g.ccbox(msg,choices=('y','n'))
    #用while做循环，并且调用自制的swich工具
    while True:
        choice=g.buttonbox(msg,title,choices)
    #g.indexbox(msg,title,choices)
        if switch(choice) == 'end':
            break
def check_in1():
#print ('入库')
    #这里其实可以使用条码枪
    rk = g.enterbox("请输入物品编码：\n",title="入库")
    #print ('rk',rk)
    if rk==None:
        menu()
        
    with xlrd.open_workbook(workbook) as beijian:
        table0 = beijian.sheet_by_name('data')
        nrows = table0.nrows
        mode=0
        #历遍主数据，如果查到有一样的，mode取1，没查到取0，查到时记录行号i0
        for i in range(nrows ):
            #print (table.row_values(i)[0])
            
            if table0.row_values(i)[0]==rk:
                #print ('same',rk)
                mode=1
                i0=i
            else:
                pass
        if mode==0:
            g.msgbox('无此物编,或请新增 '+rk, title='录入错误')
            menu()
        else:
            information=table0.row_values(i0)[1]
            #print (information)
            ok_button='确定？'
            #g.msgbox(rk+"\n"+'信息: '+table0.row_values(i0)[1]+"\n"+'当前库存: '+str(int(table0.row_values(i0)[4]))+"\n"+'库位:'+table0.row_values(i0)[2],ok_button,image='image/'+rk+'.jpg')
            title = '入库'
            fieldNames = ['数量(必填)*','来源','备注','库位']
            fieldValues = []
            #msg=rk+'\n'+'信息: '+table0.row_values(i0)[1]+"\n"+'当前库存: '+str(int(table0.row_values(i0)[4]))+"\n"+'预警库存: '+str(int(table0.row_values(i0)[5]))
            msg=rk+'\n'+'当前库存: '+str(int(table0.row_values(i0)[4]))
            fieldValues = g.multenterbox(msg,title,fieldNames,[1,'','',table0.row_values(i0)[2]])
            if fieldValues==None:
                menu()
            if fieldValues[0]=='':
                g.msgbox("数量不能为空，可以取消")
                menu()
            #print(fieldValues[1])
            #beijian = xlrd.open_workbook('beijian.xls')
            kuwei=fieldValues[3]
    ##        if kuwei==None:
    ##            g.msgbox("操作取消")
    ##            menu()

            table = beijian.sheet_by_name('history')
            nrows = table.nrows
            beijian1 = copy(beijian)
            beijian1.get_sheet('history').write(nrows,0,rk)
            beijian1.get_sheet('history').write(nrows,1,information)
            beijian1.get_sheet('history').write(nrows,2,fieldValues[1])
            beijian1.get_sheet('history').write(nrows,3,int(fieldValues[0]))
            beijian1.get_sheet('history').write(nrows,4,fieldValues[2])
            beijian1.get_sheet('history').write(nrows,5,'原库存'+str(int(table0.row_values(i0)[4]))+'+'+'变化量'+fieldValues[0]+'='+'最新库存'+str(int(table0.row_values(i0)[4])+int(fieldValues[0])))
            beijian1.get_sheet('history').write(nrows,6,now())
            beijian1.get_sheet('history').write(nrows,7,kuwei)
            beijian1.get_sheet('history').write(nrows,8,getpass.getuser())
            beijian1.get_sheet('data').write(i0,4,int(table0.row_values(i0)[4])+int(fieldValues[0]))
            beijian1.get_sheet('data').write(i0,2,kuwei)
            #int(table0.row_values(i0)[4])+int(fieldValues[1])
            try:
                beijian1.save(workbook)
                g.msgbox('ID: '+rk+"\n"+'信息: '+information+"\n"+'来源: '\
                         +fieldValues[1]+"\n"+'数量: '+fieldValues[0]+"\n"+'库位: '+kuwei+"\n"+'时间: '+now()+"\n"+'原库存'\
                         +str(int(table0.row_values(i0)[4]))+'+'+'变化量'+fieldValues[0]+'='\
                         +'最新库存'+str(int(table0.row_values(i0)[4])+int(fieldValues[0]))+"\n"+'录入人员:'+getpass.getuser(),"入库成功")
                check_in()
                #g.msgbox('原库存'+str(int(table0.row_values(i0)[4]))+'+'+'变化量'+fieldValues[1]+'='+'最新库存'+str(int(table0.row_values(i0)[4])+int(fieldValues[1])))
            except:
                g.msgbox("保存失败，有人已打开excel表 \n 请关闭beijian.xls,再试")
                
            menu()


def check_in():
#print ('入库')
    #这里其实可以使用条码枪
    rk = g.enterbox("请输入物品编码：\n",title="入库")
    #print ('rk',rk)
    if rk==None:
        menu()
        
    with xlrd.open_workbook(workbook) as beijian:
        table1 = beijian.sheet_by_name('locator')
        table1_nrows = table1.nrows

        list0=[]
        for i in range(table1_nrows ):
            list0.append(table1.row_values(i)[0])
        table0 = beijian.sheet_by_name('data')
        nrows = table0.nrows
        mode=0
        #历遍主数据，如果查到有一样的，mode取1，没查到取0，查到时记录行号i0
        for i in range(nrows ):
            #print (table.row_values(i)[0])
            
            if table0.row_values(i)[0]==rk:
                #print ('same',rk)
                mode=1
                i0=i
            else:
                pass
        if mode==0:
            g.msgbox('无此物编,或请新增 '+rk, title='录入错误')
            menu()
        else:
            information=table0.row_values(i0)[1]
            #print (information)
            ok_button='确定？'
            #g.msgbox(rk+"\n"+'信息: '+table0.row_values(i0)[1]+"\n"+'当前库存: '+str(int(table0.row_values(i0)[4]))+"\n"+'库位:'+table0.row_values(i0)[2],ok_button,image='image/'+rk+'.jpg')
            title = '入库'
            fieldNames = ['数量(必填)*','来源','批次','库位']
            fieldValues = []
            #msg=rk+'\n'+'信息: '+table0.row_values(i0)[1]+"\n"+'当前库存: '+str(int(table0.row_values(i0)[4]))+"\n"+'预警库存: '+str(int(table0.row_values(i0)[5]))
            msg=rk+'\n'+table0.row_values(i0)[1]+'\n'+'当前库存: '+str(int(table0.row_values(i0)[4]))
            fieldValues = g.multenterbox(msg,title,fieldNames,[1,'',now1(),table0.row_values(i0)[2]])
            if fieldValues==None:
                menu()
            if fieldValues[0]=='':
                g.msgbox("数量不能为空，可以取消")
                menu()
            #print(fieldValues[1])
            #beijian = xlrd.open_workbook('beijian.xls')
            kuwei=fieldValues[3]
            if kuwei in list0:
    ##        if kuwei==None:
    ##            g.msgbox("操作取消")
    ##            menu()

                table = beijian.sheet_by_name('history')
                nrows = table.nrows
                beijian1 = copy(beijian)
                beijian1.get_sheet('history').write(nrows,0,rk)
                beijian1.get_sheet('history').write(nrows,1,information)
                beijian1.get_sheet('history').write(nrows,2,fieldValues[1])
                beijian1.get_sheet('history').write(nrows,3,int(fieldValues[0]))
                beijian1.get_sheet('history').write(nrows,4,fieldValues[2])
                beijian1.get_sheet('history').write(nrows,5,'原库存'+str(int(table0.row_values(i0)[4]))+'+'+'变化量'+fieldValues[0]+'='+'最新库存'+str(int(table0.row_values(i0)[4])+int(fieldValues[0])))
                beijian1.get_sheet('history').write(nrows,6,now())
                beijian1.get_sheet('history').write(nrows,7,kuwei)
                beijian1.get_sheet('history').write(nrows,8,getpass.getuser())
                kucunliang=int(table0.row_values(i0)[4])+int(fieldValues[0])
                beijian1.get_sheet('data').write(i0,4,int(table0.row_values(i0)[4])+int(fieldValues[0]))
                beijian1.get_sheet('data').write(i0,2,kuwei)
                beijian1.get_sheet('data').write(i0,7,int(table0.row_values(i0)[6])-kucunliang)
                if kucunliang<table0.row_values(i0)[5]:
                    beijian1.get_sheet('data').write(i0,8,int(table0.row_values(i0)[6])-kucunliang)
                else:
                    beijian1.get_sheet('data').write(i0,8,0)
                
                #int(table0.row_values(i0)[4])+int(fieldValues[1])
                try:
                    beijian1.save(workbook)
                    #g.msgbox('ID: '+rk+"\n"+'信息: '+information+"\n"+'来源: '\
                    #         +fieldValues[1]+"\n"+'数量: '+fieldValues[0]+"\n"+'库位: '+kuwei+"\n"+'批次: '+fieldValues[2]+"\n"+'时间: '+now()+"\n"+'原库存'\
                    #         +str(int(table0.row_values(i0)[4]))+'+'+'变化量'+fieldValues[0]+'='\
                    #         +'最新库存'+str(int(table0.row_values(i0)[4])+int(fieldValues[0]))+"\n"+'录入人员:'+getpass.getuser(),"入库成功")
                    bar(rk,information,'数量:'+fieldValues[0],'库位:'+kuwei,'批次:'+fieldValues[2])
                    #check_in()
                    #g.msgbox('原库存'+str(int(table0.row_values(i0)[4]))+'+'+'变化量'+fieldValues[1]+'='+'最新库存'+str(int(table0.row_values(i0)[4])+int(fieldValues[1])))
                except:
                    g.msgbox("保存失败，有人已打开excel表 \n 请关闭beijian.xls,再试")
                    
                menu()
            else:
                g.msgbox("无此库位,请查正库位或新增库位")
                menu()
            

def check_out():
#print ('入库')
    #这里其实可以使用条码枪
    rk = g.enterbox("请输入物品编码：\n",title="出库")
    #print ('rk',rk)
    if rk==None:
        menu()
        
    with xlrd.open_workbook(workbook) as beijian:
        table0 = beijian.sheet_by_name('data')
        nrows = table0.nrows
        mode=0
        #历遍主数据，如果查到有一样的，mode取1，没查到取0，查到时记录行号i0
        for i in range(nrows ):
            #print (table.row_values(i)[0])
            
            if table0.row_values(i)[0]==rk:
                #print ('same',rk)
                mode=1
                i0=i
            else:
                pass
        if mode==0:
            g.msgbox('无此物编,或请新增 '+rk, title='录入错误')
            menu()
        else:
            information=table0.row_values(i0)[1]
            #print (information)
            ok_button='确定？'
            #g.msgbox(rk+"\n"+'信息: '+table0.row_values(i0)[1]+"\n"+'当前库存: '+str(int(table0.row_values(i0)[4]))
                     #+"\n"+' 库位: '+table0.row_values(i0)[2]+"\n"+' 预警库存: '+str(int(table0.row_values(i0)[5])),ok_button,image='image/'+rk+'.jpg')
            #g.msgbox(rk+"\n"+'信息: '+table0.row_values(i0)[1]+"\n"+'当前库存: '+str(int(table0.row_values(i0)[4]))+"\n"+'库位:'+table0.row_values(i0)[2]+"\n",ok_button,image='image/'+rk+'.jpg')
            title = '出库'
            fieldNames = ['数量(必填)*','去向','备注']
            fieldValues = []
            msg=rk+'\n'+'信息: '+table0.row_values(i0)[1]+"\n"+'当前库存: '+str(int(table0.row_values(i0)[4]))+"\n"+'预警库存: '+str(int(table0.row_values(i0)[5]))+"\n"+' 库位: '+table0.row_values(i0)[2]
            fieldValues = g.multenterbox(msg,title,fieldNames,[1,'',''])
            if fieldValues==None:
                menu()
            if fieldValues[0]=='':
                g.msgbox("数量不能为空，可以取消")
                menu()
            else:
                fieldValues[0]=str(0-int(fieldValues[0]))
            #print(fieldValues[1])
            #beijian = xlrd.open_workbook('beijian.xls')
            table = beijian.sheet_by_name('history')
            nrows = table.nrows
            beijian1 = copy(beijian)
            beijian1.get_sheet('history').write(nrows,0,rk)
            beijian1.get_sheet('history').write(nrows,1,information)
            beijian1.get_sheet('history').write(nrows,2,fieldValues[1])
            beijian1.get_sheet('history').write(nrows,3,int(fieldValues[0]))
            beijian1.get_sheet('history').write(nrows,4,fieldValues[2])
            beijian1.get_sheet('history').write(nrows,5,'原库存'+str(int(table0.row_values(i0)[4]))+'+'+'变化量'+fieldValues[0]+'='+'最新库存'+str(int(table0.row_values(i0)[4])+int(fieldValues[0])))
            beijian1.get_sheet('history').write(nrows,6,now())
            beijian1.get_sheet('history').write(nrows,8,getpass.getuser())
            kucunliang=int(table0.row_values(i0)[4])+int(fieldValues[0])
            beijian1.get_sheet('data').write(i0,4,kucunliang)
            beijian1.get_sheet('data').write(i0,7,int(table0.row_values(i0)[6])-kucunliang)
            if kucunliang<table0.row_values(i0)[5]:
                beijian1.get_sheet('data').write(i0,8,int(table0.row_values(i0)[6])-kucunliang)
            else:
                beijian1.get_sheet('data').write(i0,8,0)
            #int(table0.row_values(i0)[4])+int(fieldValues[1])
            try:
                beijian1.save(workbook)
                g.msgbox('ID: '+rk+"\n"+'信息: '+information+"\n"+'来源: '\
                         +fieldValues[1]+"\n"+'数量: '+fieldValues[1]+"\n"+'库位: '+fieldValues[2]+"\n"+'时间: '+now()+"\n"+'原库存'\
                         +str(int(table0.row_values(i0)[4]))+'+'+'变化量'+fieldValues[0]+'='+'最新库存'+str(int(table0.row_values(i0)[4])\
                        +int(fieldValues[0]))+"\n"+'录入人员:'+getpass.getuser(),"出库成功")
                #g.msgbox('原库存'+str(int(table0.row_values(i0)[4]))+'+'+'变化量'+fieldValues[1]+'='+'最新库存'+str(int(table0.row_values(i0)[4])+int(fieldValues[1])))
                #check_out()
            except:
                g.msgbox("保存失败，有人已打开excel表 \n 请关闭beijian.xls,再试")
                
            menu()


        
#新增物料模块    
def new_material():
    #password=g.passwordbox(msg="请输入操作密码",title="捷德备件管理系统")
    #if password=='password':
    with xlrd.open_workbook(workbook) as beijian:
        #主数据位于data表
        table = beijian.sheet_by_name('data')
        #nrows代表行数，意思是主数据现在排到了多少行。
        nrows = table.nrows
        ID0=table.row_values(nrows-1)[0]
        ID1=int(ID0[2:])
        ID2="NC"+str(ID1+1)
        print(ID2)
        #ncols = table.ncols
            
        #自定义流水编号，这里使用NC开头，并使用5000+当前行号，确保唯一性    
        ID='NC'+str(nrows+5010002700)
        ID=ID2
            
        #录入主数据内容描述   
        st = g.enterbox("请输入物品描述：\n")
        #点取消，退回主菜单
        if st==None:
            menu()
        #确定窗口，包含流水号和物品描述
        yujingzhi = g.enterbox("请输入物品最低预警值：\n",default=0)
        yujingzhi_gao = g.enterbox("请输入物品最高预警值：\n",default=0)
        if yujingzhi==None:
            menu()
        ID = g.enterbox("请确认物品唯一码（自动生成），如果产品自带条形码作为唯一码，请输入条码号：\n",default=ID)
        ccbox = g.ccbox(ID+": "+st,title="新增",choices=('确定','取消'))
        
        if ccbox==True:
                #print ('y')
            #beijian = xlrd.open_workbook('beijian.xls')文件写入，需要先复制一份，再保存覆盖
            #注意，不能在原
            beijian1 = copy(beijian)
            #在最下行，第一列，写入ID流水号码
            beijian1.get_sheet('data').write(nrows,0,ID)
            #在最下行，第一列，写入物品描述
            beijian1.get_sheet('data').write(nrows,1,st)
            beijian1.get_sheet('data').write(nrows,4,0)
            beijian1.get_sheet('data').write(nrows,5,int(yujingzhi))
            beijian1.get_sheet('data').write(nrows,6,int(yujingzhi_gao))
            beijian1.get_sheet('data').write(nrows,2,'未入库')
            g.msgbox(ID+": "+st+" 新增成功！请在image文件夹内新增物品图片"+ID+".jpg(640x480)")
            #保存，并覆盖原有excel表
            try:
                beijian1.save(workbook)
            except:
                g.msgbox("保存失败，有人已打开excel表 \n 请关闭beijian.xls,再试")
            menu()
        #点取消，退回主菜单
        else:
            g.msgbox("已被取消")
            menu()       
  
    
#主模块
if __name__=="__main__":
    #os.path.exists()
    workbook='database/beijian.xls'
    if os.path.exists(workbook):
    #menu()
    #先做一些定义，例如数据存储的excel表名称
        #password=g.passwordbox(msg="请输入操作密码",title="捷德备件管理系统")
        password='password'
        if password=='password':
            
        #调用主界面
            menu()
        else:
            g.msgbox('密码错误')
    else:
        g.msgbox('数据库database/beijian.xls不存在，程序无法继续，将关闭')
