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


def lifetime0(a0,delta):
    #保质期，第一个参数是生产日期，第二个是保质天数， 返回过期的日期
    
    #a0='20180511'
    #delta=400
    a1=datetime.datetime.strptime(a0,"%Y%m%d")
    a2=a1+datetime.timedelta(days=int(delta))
    #a2=datetime.datetime.strptime(str(a2),"%Y%m%d")
    a2=a2.strftime("%Y%m%d")
    return a2
    #print (datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))


#定义now的格式
def now():
    return datetime.datetime.now().strftime('%Y%m%d %H:%M:%S')
def now1():
    return datetime.datetime.now().strftime('%Y%m%d%H%M%S')
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
    print (barname,comments0,comments1,comments2,comments3)
    img.show()
    
    img.save('./barcode/'+comments3.replace(':','')+'.png')
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
        rk = g.enterbox("请输入批次号：\n",title="移库")
        if True:
            with xlrd.open_workbook(workbook) as beijian:
                table1 = beijian.sheet_by_name('locator')
                table1_nrows = table1.nrows
                table0 = beijian.sheet_by_name('batch')
                nrows = table0.nrows
                list0=[]
                for i in range(table1_nrows ):
                    list0.append(table1.row_values(i)[0])
                
                for i in range(nrows ):
                    if table0.row_values(i)[4]==rk:
                        i0=i
                        break
                print (i0)
                newkuwei=g.enterbox(msg=table0.row_values(i)[0]+'\n'+'请输入新库位，默认值为原库位',title='移库',default=table0.row_values(i0)[7])
                if newkuwei in list0:
                    print ('yes, within kuwei')
                    beijian1 = copy(beijian)
                    beijian1.get_sheet('batch').write(i0,7,newkuwei)
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
            list=[]
            for i in range(1,nrows ):
            #print (table.row_values(i)[0])
                if int(table0.row_values(i)[4])<int(table0.row_values(i)[5]):
            
            
                    list.append('\n'+table0.row_values(i)[0]+' '+table0.row_values(i)[1]+' 库位: '\
                                +table0.row_values(i)[2]+' '+table0.row_values(i)[3]+'  库存:'+str(int(table0.row_values(i)[4]))+\
                                ' 最低预警: '+str(int(table0.row_values(i)[5]))+' '+' 最高预警: '+str(int(table0.row_values(i)[6]))+' '\
                                +'待补足'+str(int(table0.row_values(i)[6])-int(table0.row_values(i)[4])))

            S=''
            for s in list:
                S=S+s+'\n'
            #print ('S is:',S)
            #g.msgbox(S,title='搜索结果')
            g.msgbox(S)

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
        table0 = beijian.sheet_by_name('batch')
        nrows = table0.nrows
        
        list0=[]
        for i in range(table1_nrows ):
            list0.append(table1.row_values(i)[0])
        print ('list0 is:',list0)
        #g.multchoicebox(msg="请选择你爱吃哪些水果?",title="",choices=("西瓜","香蕉","苹果","梨"))
        kw=g.choicebox(msg="请选择库位",title="",choices=list0)
        list=[]
        for i in range(nrows ):
            if kw in table0.row_values(i)[7]:
                list.append('\n'+table0.row_values(i)[0]+' '+table0.row_values(i)[1]+' 数量: '+str(int(table0.row_values(i)[3]))+'批次 '+table0.row_values(i)[4])
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
        beijian1.save('.\\export\\'+now1+'.xls')
        #os.system(now1+'.xls')
        os.system('explorer.exe .\\export\\'+now1+'.xls')
        print ('end')
        menu()
    

def menu():
    title='捷德库房管理系统'
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
            #g.msgbox(len(now1()))
            fieldValues = g.multenterbox(msg,title,fieldNames,[1,'',now1(),table0.row_values(i0)[2]])
            fieldValues[2]=fieldValues[2]+now1()[len(fieldValues[2]):]
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
                table_batch = beijian.sheet_by_name('batch')
                nrows = table.nrows
                nrows_batch = table_batch.nrows
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
                beijian1.get_sheet('batch').write(nrows_batch,0,rk)
                beijian1.get_sheet('batch').write(nrows_batch,1,information)
                beijian1.get_sheet('batch').write(nrows_batch,2,fieldValues[1])
                beijian1.get_sheet('batch').write(nrows_batch,3,int(fieldValues[0]))
                beijian1.get_sheet('batch').write(nrows_batch,4,fieldValues[2])
                lifetime_value=lifetime0(fieldValues[2][0:8],table0.row_values(i0)[9])
                beijian1.get_sheet('batch').write(nrows_batch,6,lifetime_value)
                beijian1.get_sheet('batch').write(nrows_batch,7,kuwei)
                #beijian1.get_sheet('batch').write(nrows,5,'原库存'+str(int(table0.row_values(i0)[4]))+'+'+'变化量'+fieldValues[0]+'='+'最新库存'+str(int(table0.row_values(i0)[4])+int(fieldValues[0])))
                #beijian1.get_sheet('batch').write(nrows,6,now())
                #beijian1.get_sheet('batch').write(nrows,7,kuwei)
                #beijian1.get_sheet('batch').write(nrows,8,getpass.getuser())
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
            msg=rk+'\n'+'信息: '+table0.row_values(i0)[1]+"\n"+'当前库存: '\
                 +str(int(table0.row_values(i0)[4]))+"\n"+'预警库存: '\
                 +str(int(table0.row_values(i0)[5]))+"\n"+' 库位: '+table0.row_values(i0)[2]
            #g.msgbox(msg)
            #checkout_num=int(g.enterbox(msg,title='出库数量'))
            title='出库'
            box=g.multenterbox(msg,title,["数量","去向"])
            checkout_num=int(box[0])
            dest=box[1]
            if checkout_num>int(table0.row_values(i0)[4]):
                g.msgbox('你的输入比库存量'+str(int(table0.row_values(i0)[4]))+'要大，错误,请核查数量后再录入')
                menu()
            table_batch = beijian.sheet_by_name('batch')
            nrows_batch = table_batch .nrows
            batch0={}
            num0=[]
            for x in range(nrows_batch):
                if table_batch.row_values(x)[0]==rk:
                    if table_batch.row_values(x)[4] !='':
                        num0.append(str(int(table_batch.row_values(x)[3])))
                        batch0[table_batch.row_values(x)[4]+'数量:'+str(int(table_batch.row_values(x)[3]))+':库位'+table_batch.row_values(x)[7]]=x
                        print (x,rk,table_batch.row_values(x)[4])
                    
            batch_list=sorted(list(batch0.keys()))
            batch_list0=[]
            batch_list_number=[]
            batch_line0=[]
            for item0 in batch_list:
                batch_list0.append(item0.split('数量')[0])
                batch_list_number.append(item0.split(':')[1])
                batch_line0.append(batch0[item0])
            print ('batch_list_number',batch_list_number)
            print ('batch_line0:',batch_line0)
            #g.msgbox(batch2)
            #batch_choice=g.multchoicebox(msg="请选择出库的批次(单选)",title="",choices=batch_list)
            #*********************************
            print (num0)
            msg=rk+'\n'+'信息: '+table0.row_values(i0)[1]+"\n"+'当前库存: '\
                 +str(int(table0.row_values(i0)[4]))+"\n"+'预警库存: '\
                 +str(int(table0.row_values(i0)[5]))+"\n"+' 最近一次库位: '+table0.row_values(i0)[2]
            title = '出库(先进先出原则)'
            fieldNames=batch_list
            list1here=[]
            sum=0
            i=0
            while sum<checkout_num:
                sum=sum+int(batch_list_number[i])
                list1here.append(batch_list_number[i])
                i=i+1
            print (sum,i,list1here[i-1])
            list1here[i-1]=checkout_num-(sum-int(list1here[i-1]))
            fieldValues=g.multenterbox(msg,title,fieldNames,list1here)
            sum=0
            for i in range(len(fieldValues)):
                if fieldValues[i] !='':
                    print (fieldValues[i],type(fieldValues[i]),batch_list_number[i],type(batch_list_number[i]))
                    if int(fieldValues[i])<=int(batch_list_number[i]):
                        sum=sum+int(fieldValues[i])
                    else:
                        print (fieldValues[i],type(fieldValues[i]),batch_list_number[i],type(batch_list_number[i]))
                        g.msgbox('错误！单批次输入数量大于批次剩余数量')
                        menu()
            print ('sum',sum)
            if sum!=checkout_num:
                g.msgbox('错误！总和与之前输入数量不符，请检查后重新输入')
                menu()
            else:
                #g.msgbox('前后数字一至')
                table = beijian.sheet_by_name('history')
                table_batch=beijian.sheet_by_name('batch')
                nrows = table.nrows
                beijian1 = copy(beijian)
                kucun=str(int(table0.row_values(i0)[4]))
                list_kuwei=[]
                list_batch=[]
                checkout_batch_num=[]
                for i in range(len(fieldValues)):
                    if fieldValues[i] !='':
                        list_batch.append(table_batch.row_values(batch_line0[i])[4])
                        list_kuwei.append(table_batch.row_values(batch_line0[i])[7])
                        checkout_batch_num.append(fieldValues[i])
                        #table_batch.row_values(i0)[7]
                        left_num=int(batch_list_number[i])-int(fieldValues[i])
                        beijian1.get_sheet('batch').write(batch_line0[i],3,left_num)
                        if left_num==0:
                            beijian1.get_sheet('batch').write(batch_line0[i],4,'')
                            beijian1.get_sheet('batch').write(batch_line0[i],6,'')
                            beijian1.get_sheet('batch').write(batch_line0[i],7,'')
                            print ('pici:',batch_line0[i])

                        #nrows=nrows+1
                        #batch_line0[i]
                #table = beijian.sheet_by_name('history')

                        beijian1.get_sheet('history').write(nrows,0,rk)
                        information=table0.row_values(i0)[1]
                        beijian1.get_sheet('history').write(nrows,1,information)
                        beijian1.get_sheet('history').write(nrows,2,dest)
                        #beijian1.get_sheet('history').write(nrows,3,'-'+fieldValues[i])
                        beijian1.get_sheet('history').write(nrows,3,-int(fieldValues[i]))
                        beijian1.get_sheet('history').write(nrows,4,batch_list0[i])
                        
                        beijian1.get_sheet('history').write(nrows,5,'原库存'+kucun+'-'+'变化量'+fieldValues[i]+'='+'最新库存'+str(int(kucun)-int(fieldValues[i])))
                        kucun=str(int(kucun)-int(fieldValues[i]))
                        table0.row_values(i0)[4]=str(int(table0.row_values(i0)[4])-int(fieldValues[i]))
                        beijian1.get_sheet('history').write(nrows,6,now())
                        beijian1.get_sheet('history').write(nrows,8,getpass.getuser())
                        nrows=nrows+1
                
                kucunliang=int(table0.row_values(i0)[4])
                beijian1.get_sheet('data').write(i0,4,kucun)
                beijian1.get_sheet('data').write(i0,7,int(table0.row_values(i0)[6])-kucunliang)
                if kucunliang<table0.row_values(i0)[5]:
                    beijian1.get_sheet('data').write(i0,8,int(table0.row_values(i0)[6])-kucunliang)
                else:
                    beijian1.get_sheet('data').write(i0,8,0)
                try:
                    beijian1.save(workbook)
                    #g.msgbox('ID: '+rk+"\n"+'信息: '+information+"\n"+'来源: '\
                    #         +fieldValues[1]+"\n"+'数量: '+fieldValues[1]+"\n"+'批次: '+fieldValues[2]+"\n"+'时间: '+now()+"\n"+'原库存'\
                    #         +str(int(table0.row_values(i0)[4]))+'+'+'变化量'+fieldValues[0]+'='+'最新库存'+str(int(table0.row_values(i0)[4])\
                    #        +int(fieldValues[0]))+"\n"+'录入人员:'+getpass.getuser(),"出库成功")
                    #g.msgbox('原库存'+str(int(table0.row_values(i0)[4]))+'+'+'变化量'+fieldValues[1]+'='+'最新库存'+str(int(table0.row_values(i0)[4])+int(fieldValues[1])))
                    #check_out()
                    print (now(),getpass.getuser(),dest,msg,checkout_num,list_batch,list_kuwei,checkout_batch_num)
                    list_for_print=[]
                    for i in range(len(list_batch)):
                        list_for_print.append('库位：'+list_kuwei[i]+' 批次：'+list_batch[i]+' 数量：'+checkout_batch_num[i]+'\n')
                    print (list_for_print)
                    g.msgbox('出库成功')
                    now1=datetime.datetime.now().strftime('%Y%m%d%H%M%S')
                    with open('./checkout/'+now1+'.txt',"a+") as f:
                        f.write("\n")
                        f.write("                                      出库单\n")
                        f.write("--------------------------------------------------------------------------------\n")
                        f.write("|日期:"+now1+"\t\t\t\t|操作员："+getpass.getuser()+'\n')
                        f.write("--------------------------------------------------------------------------------\n")
                        f.write("|ID:"+rk+"\t\t\t\t|接收人："+dest+'\n')
                        f.write("--------------------------------------------------------------------------------\n")
                        f.write("|出库数量:"+str(checkout_num)+"\t\t\t\t\t|剩余库存："+kucun+'\n')
                        f.write("--------------------------------------------------------------------------------\n")
                        f.write("|物品描述："+table0.row_values(i0)[1]+'\n')
                        f.write("--------------------------------------------------------------------------------\n")
                        list_for_print=[]
                        for i in range(len(list_batch)):
                            list_for_print.append('库位：'+list_kuwei[i]+' 批次：'+list_batch[i]+' 数量：'+checkout_batch_num[i]+'\n')
                            f.write('|库位：'+list_kuwei[i]+'\t\t|批次：'+list_batch[i]+'\t\t|数量：'+checkout_batch_num[i]+'\n')
                        f.write("\n")
                        f.write("---------------------------------结束-------------------------------------------\n")
                    
        
        
                    os.system('notepad.exe ./checkout/'+now1+'.txt')
                    print ('done')
                except:
                    g.msgbox("保存失败，有人已打开excel表 \n 请关闭beijian.xls,再试")
                
                menu()

            #*********************************
            print ('batch_choice:',batch_choice)
            batch_line=batch0[batch_choice]
            #g.msgbox(table_batch.row_values(batch_line)[0]+table_batch.row_values(batch_line)[1]+table_batch.row_values(batch_line)[4])
            default_checkout_num=int(table_batch.row_values(batch_line)[3])
            information=table0.row_values(i0)[1]
            #print (information)
            ok_button='确定？'
            #g.msgbox(rk+"\n"+'信息: '+table0.row_values(i0)[1]+"\n"+'当前库存: '+str(int(table0.row_values(i0)[4]))
                     #+"\n"+' 库位: '+table0.row_values(i0)[2]+"\n"+' 预警库存: '+str(int(table0.row_values(i0)[5])),ok_button,image='image/'+rk+'.jpg')
            #g.msgbox(rk+"\n"+'信息: '+table0.row_values(i0)[1]+"\n"+'当前库存: '+str(int(table0.row_values(i0)[4]))+"\n"+'库位:'+table0.row_values(i0)[2]+"\n",ok_button,image='image/'+rk+'.jpg')
            title = '出库'
            fieldNames = ['数量(必填)*','去向']
            fieldValues = []
            msg=rk+'\n'+'信息: '+table0.row_values(i0)[1]+"\n"+'当前库存: '+str(int(table0.row_values(i0)[4]))+"\n"+'预警库存: '+str(int(table0.row_values(i0)[5]))+"\n"+' 库位: '+table0.row_values(i0)[2]
            #fieldValues = g.multenterbox(msg,title,fieldNames,[default_checkout_num,'',''])
            fieldValues = g.multenterbox(msg,title,fieldNames,[default_checkout_num,''])
            fieldValues.append(batch_choice.split('数量')[0])
            if fieldValues==None:
                menu()
            elif fieldValues[0]=='':
                g.msgbox("数量不能为空，可以取消")
                menu()
            elif abs(int(fieldValues[0]))>default_checkout_num:
                g.msgbox('本批次'+batch_choice+'剩余'+str(default_checkout_num)+',单词出库数量不能大于批次数量，未执行，请重新核对数量')
                menu()
            else:
                fieldValues[0]=str(0-int(fieldValues[0]))
            #if abs(int(fieldValues[0]))==default_checkout_num:
            #    g.msgbox('same')
            #else:
                print (fieldValues[0],type(fieldValues[0]),default_checkout_num,type(default_checkout_num))
                #g.msgbox('diff',type(default_checkout_num),type(fieldValues[0]))
                left_num=default_checkout_num+int(fieldValues[0])
                print (left_num)
                #g.msgbox('left_num')
            #    a=g.enterbox('hello')
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
            #beijian1.get_sheet('batch').write(nrows,0,rk)
            #beijian1.get_sheet('batch').write(nrows,1,information)
            #beijian1.get_sheet('batch').write(nrows,2,fieldValues[1])
            beijian1.get_sheet('batch').write(batch_line,3,left_num)
            if left_num==0:
                line_number=batch_line
                beijian1.get_sheet('batch').write(line_number,4,'nihao')
                beijian1.get_sheet('batch').write(line_number,5,'diwuhang')
                beijian1.get_sheet('batch').write(line_number,6,'di6hang')
            #beijian1.get_sheet('batch').write(nrows,4,fieldValues[2])
            #beijian1.get_sheet('batch').write(nrows,5,'原库存'+str(int(table0.row_values(i0)[4]))+'+'+'变化量'+fieldValues[0]+'='+'最新库存'+str(int(table0.row_values(i0)[4])+int(fieldValues[0])))
            #beijian1.get_sheet('batch').write(nrows,6,now())
            #beijian1.get_sheet('batch').write(nrows,8,getpass.getuser())
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
                         +fieldValues[1]+"\n"+'数量: '+fieldValues[1]+"\n"+'批次: '+fieldValues[2]+"\n"+'时间: '+now()+"\n"+'原库存'\
                         +str(int(table0.row_values(i0)[4]))+'+'+'变化量'+fieldValues[0]+'='+'最新库存'+str(int(table0.row_values(i0)[4])\
                        +int(fieldValues[0]))+"\n"+'录入人员:'+getpass.getuser(),"出库成功")
                #g.msgbox('原库存'+str(int(table0.row_values(i0)[4]))+'+'+'变化量'+fieldValues[1]+'='+'最新库存'+str(int(table0.row_values(i0)[4])+int(fieldValues[1])))
                #check_out()
            except:
                g.msgbox("保存失败，有人已打开excel表 \n 请关闭beijian.xls,再试")
                
            menu()


        
#新增物料模块    
def new_material():
    password=g.passwordbox(msg="请输入操作密码",title="捷德备件管理系统")
    if password=='password':
        with xlrd.open_workbook(workbook) as beijian:
            #主数据位于data表
            table = beijian.sheet_by_name('data')
            #nrows代表行数，意思是主数据现在排到了多少行。
            nrows = table.nrows
            #ncols = table.ncols
                
            #自定义流水编号，这里使用NC开头，并使用5000+当前行号，确保唯一性    
            ID='J'+str(nrows+50000001)
            msg='新增物品'
            title='新增物品'
            fieldNames_new=['物品描述','最低预警','最高预警','保质期（天）','唯一ID']
            defaultvalues=['请输入物品的详细描述，如厂家，型号，颜色等','0','0','365',ID]
            
            box=g.multenterbox(msg,title,fieldNames_new,defaultvalues)
            if box==None:
                menu()
            else:
                
                #录入主数据内容描述   
                #st = g.enterbox("请输入物品描述：\n")
                st=box[0]
                #点取消，退回主菜单
                if st==None:
                    menu()
                #确定窗口，包含流水号和物品描述
                #yujingzhi = g.enterbox("请输入物品最低预警值：\n",default=0)
                yujingzhi=box[1]
                #yujingzhi_gao = g.enterbox("请输入物品最高预警值：\n",default=0)
                yujingzhi_gao=box[2]
                lifetime=box[3]
                ID=box[4]
                #if yujingzhi==None:
                #    menu()
                #ID = g.enterbox("请确认物品唯一码（自动生成），如果产品自带条形码作为唯一码，请输入条码号：\n",default=ID)
                #g.msgbox(box)
            

                beijian1 = copy(beijian)
                #在最下行，第一列，写入ID流水号码
                beijian1.get_sheet('data').write(nrows,0,ID)
                #在最下行，第一列，写入物品描述
                beijian1.get_sheet('data').write(nrows,1,st)
                beijian1.get_sheet('data').write(nrows,4,0)
                beijian1.get_sheet('data').write(nrows,5,int(yujingzhi))
                beijian1.get_sheet('data').write(nrows,6,int(yujingzhi_gao))
                beijian1.get_sheet('data').write(nrows,2,'未入库')
                beijian1.get_sheet('data').write(nrows,9,lifetime)
                
                #保存，并覆盖原有excel表
                try:
                    beijian1.save(workbook)
                    g.msgbox(ID+": "+st+" 新增成功！请在image文件夹内新增物品图片"+ID+".jpg(640x480)")
                except:
                    g.msgbox("保存失败，有人已打开excel表 \n 请关闭beijian.xls,再试")
                menu()

    
#主模块
if __name__=="__main__":
    #os.path.exists()
    workbook='database/beijian.xls'
    if os.path.exists(workbook):
    #menu()
    #先做一些定义，例如数据存储的excel表名称
        #password=g.passwordbox(msg="请输入操作密码",title="捷德卡体印刷库房管理系统")
        password='password'
        if password=='password':
            
        #调用主界面
            menu()
        else:
            g.msgbox('密码错误')
    else:
        g.msgbox('数据库database/beijian.xls不存在，程序无法继续，将关闭')
