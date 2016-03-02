# -*- coding: utf-8
##------------------------------------------------------------  
##   程序：汽车之家爬虫  
##   作者：王东升 Tel：13120330930 E-mail：13120330930@163.com
##   日期：2015-08-05 
##   语言：Python 2.7.10
##   需安装xlrd和xlrd库
##------------------------------------------------------------
import Queue
import cookielib
import threading
import urllib2
import time
import sys
import os 
import re
import xlwt
import xlrd
from bs4 import BeautifulSoup
from xlutils.copy import copy 
from xlrd import open_workbook 
from xlwt import easyxf
import json
import socket
import time
timeout = 120
socket.setdefaulttimeout(timeout)
sleep_download_time=10
ssleep_download_time=20
#'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'
ABClist=['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
##print len(ABClist)
allstr=u"品牌,车系,车型名称,厂商指导价(元),厂商,级别,发动机,变速箱,长*宽*高(mm),车身结构,最高车速(km/h),官方0-100km/h加速(s),实测0-100km/h加速(s),实测100-0km/h制动(m),实测油耗(L/100km),工信部综合油耗(L/100km),实测离地间隙(mm),整车质保,长度(mm),宽度(mm),高度(mm),轴距(mm),前轮距(mm),后轮距(mm),最小离地间隙(mm),整备质量(kg),车身结构,车门数(个),座位数(个),后排车门开启方式,油箱容积(L),货箱尺寸(mm),最大载重质量(kg),行李厢容积(L),发动机型号,排量(mL),排量(L),进气形式,气缸排列形式,气缸数(个),每缸气门数(个),压缩比,配气机构,缸径(mm),行程(mm),最大马力(Ps),最大功率(kW),最大功率转速(rpm),最大扭矩(N·m),最大扭矩转速(rpm),发动机特有技术,燃料形式,燃油标号,供油方式,缸盖材料,缸体材料,环保标准,电动机总功率(kW),电动机总扭矩(N·m),前电动机最大功率(kW),前电动机最大扭矩(N·m),后电动机最大功率(kW),后电动机最大扭矩(N·m),电池支持最高续航里程(km),电池容量(kWh),简称,挡位个数,变速箱类型,驱动方式,四驱形式,中央差速器结构,前悬架类型,后悬架类型,助力类型,车体结构,前制动器类型,后制动器类型,驻车制动类型,前轮胎规格,后轮胎规格,备胎规格,主/副驾驶座安全气囊,前/后排侧气囊,前/后排头部气囊(气帘),膝部气囊,胎压监测装置,零胎压继续行驶,安全带未系提示,ISOFIX儿童座椅接口,发动机电子防盗,车内中控锁,遥控钥匙,无钥匙启动系统,无钥匙进入系统,ABS防抱死,制动力分配(EBD/CBC等),刹车辅助(EBA/BAS/BA等),牵引力控制(ASR/TCS/TRC等),车身稳定控制(ESC/ESP/DSC等),上坡辅助,自动驻车,陡坡缓降,可变悬架,空气悬架,可变转向比,前桥限滑差速器/差速锁,中央差速器锁止功能,后桥限滑差速器/差速锁,电动天窗,全景天窗,运动外观套件,铝合金轮圈,电动吸合门,侧滑门,电动后备厢,感应后备厢,车顶行李架,真皮方向盘,方向盘调节,方向盘电动调节,多功能方向盘,方向盘换挡,方向盘加热,方向盘记忆,定速巡航,前/后驻车雷达,倒车视频影像,行车电脑显示屏,全液晶仪表盘,HUD抬头数字显示,座椅材质,运动风格座椅,座椅高低调节,腰部支撑调节,肩部支撑调节,主/副驾驶座电动调节,第二排靠背角度调节,第二排座椅移动,后排座椅电动调节,电动座椅记忆,前/后排座椅加热,前/后排座椅通风,前/后排座椅按摩,第三排座椅,后排座椅放倒方式,前/后中央扶手,后排杯架,GPS导航系统,定位互动服务,中控台彩色大屏,蓝牙/车载电话,车载电视,后排液晶屏,220V/230V电源,外接音源接口,CD支持MP3/WMA,多媒体系统,扬声器品牌,扬声器数量,近光灯,远光灯,日间行车灯,自适应远近光,自动头灯,转向辅助灯,转向头灯,前雾灯,大灯高度可调,大灯清洗装置,车内氛围灯,前/后电动车窗,车窗防夹手功能,防紫外线/隔热玻璃,后视镜电动调节,后视镜加热,内/外后视镜自动防眩目,后视镜电动折叠,后视镜记忆,后风挡遮阳帘,后排侧遮阳帘,后排侧隐私玻璃,遮阳板化妆镜,后雨刷,感应雨刷,空调控制方式,后排独立空调,后座出风口,温度分区控制,车内空气调节/花粉过滤,车载冰箱,自动泊车入位,发动机启停技术,并线辅助,车道偏离预警系统,主动刹车/主动安全系统,整体主动转向系统,夜视系统,中控液晶屏分屏显示,自适应巡航,全景摄像头"
#allstr=u"品牌,车系,车型名称,厂商指导价(元),厂商,级别,发动机,变速箱,长*宽*高(mm),车身结构,最高车速(km/h),官方0-100km/h加速(s),实测0-100km/h加速(s),实测100-0km/h制动(m),实测油耗(L/100km),工信部综合油耗(L/100km),实测离地间隙(mm),整车质保,长度(mm),宽度(mm),高度(mm),轴距(mm),前轮距(mm),后轮距(mm),最小离地间隙(mm),整备质量(kg),车身结构,车门数(个),座位数(个),油箱容积(L),行李厢容积(L),发动机型号,排量(mL),排量(L),进气形式,气缸排列形式,气缸数(个),每缸气门数(个),压缩比,配气机构,缸径(mm),行程(mm),最大马力(Ps),最大功率(kW),最大功率转速(rpm),最大扭矩(N•m),最大扭矩转速(rpm),发动机特有技术,燃料形式,燃油标号,供油方式,缸盖材料,缸体材料,环保标准,电动机总功率(kW),电动机总扭矩(N•m),前电动机最大功率(kW),前电动机最大扭矩(N•m),后电动机最大功率(kW),后电动机最大扭矩(N•m),电池支持最高续航里程(km),电池容量(kWh),简称,挡位个数,变速箱类型,驱动方式,四驱形式,中央差速器结构,前悬架类型,后悬架类型,助力类型,车体结构,前制动器类型,后制动器类型,驻车制动类型,前轮胎规格,后轮胎规格,备胎规格,主/副驾驶座安全气囊,前/后排侧气囊,前/后排头部气囊(气帘),膝部气囊,胎压监测装置,零胎压继续行驶,安全带未系提示,ISOFIX儿童座椅接口,发动机电子防盗,车内中控锁,遥控钥匙,无钥匙启动系统,无钥匙进入系统,ABS防抱死,制动力分配(EBD/CBC等),刹车辅助(EBA/BAS/BA等),牵引力控制(ASR/TCS/TRC等),车身稳定控制(ESC/ESP/DSC等),上坡辅助,自动驻车,陡坡缓降,可变悬架,空气悬架,可变转向比,前桥限滑差速器/差速锁,中央差速器锁止功能,后桥限滑差速器/差速锁,电动天窗,全景天窗,运动外观套件,铝合金轮圈,电动吸合门,侧滑门,电动后备厢,感应后备厢,车顶行李架,真皮方向盘,方向盘调节,方向盘电动调节,多功能方向盘,方向盘换挡,方向盘加热,方向盘记忆,定速巡航,前/后驻车雷达,倒车视频影像,行车电脑显示屏,全液晶仪表盘,HUD抬头数字显示,座椅材质,运动风格座椅,座椅高低调节,腰部支撑调节,肩部支撑调节,主/副驾驶座电动调节,第二排靠背角度调节,第二排座椅移动,后排座椅电动调节,电动座椅记忆,前/后排座椅加热,前/后排座椅通风,前/后排座椅按摩,第三排座椅,后排座椅放倒方式,前/后中央扶手,后排杯架,GPS导航系统,定位互动服务,中控台彩色大屏,蓝牙/车载电话,车载电视,后排液晶屏,220V/230V电源,外接音源接口,CD支持MP3/WMA,多媒体系统,扬声器品牌,扬声器数量,近光灯,远光灯,日间行车灯,自适应远近光,自动头灯,转向辅助灯,转向头灯,前雾灯,大灯高度可调,大灯清洗装置,车内氛围灯,前/后电动车窗,车窗防夹手功能,防紫外线/隔热玻璃,后视镜电动调节,后视镜加热,内/外后视镜自动防眩目,后视镜电动折叠,后视镜记忆,后风挡遮阳帘,后排侧遮阳帘,后排侧隐私玻璃,遮阳板化妆镜,后雨刷,感应雨刷,空调控制方式,后排独立空调,后座出风口,温度分区控制,车内空气调节/花粉过滤,车载冰箱,自动泊车入位,发动机启停技术,并线辅助,车道偏离预警系统,主动刹车/主动安全系统,整体主动转向系统,夜视系统,中控液晶屏分屏显示,自适应巡航,全景摄像头"
allstrlist=allstr.split(',')
filename=xlwt.Workbook()    #创建一个工作簿
sheet=filename.add_sheet("1")#创建一个表
row=0
for item1 in ABClist:
    try:
        homourl="http://www.autohome.com.cn/grade/carhtml/"+item1+".html"
        print homourl
        cookie_jar = cookielib.LWPCookieJar()
        cookie = urllib2.HTTPCookieProcessor(cookie_jar)
        opener = urllib2.build_opener(cookie)
        req = urllib2.Request(homourl)
        soures_home = opener.open(req).read()
    except:
        pass
    soupA= BeautifulSoup(soures_home,from_encoding="gbk")
    treeAdd=soupA.findAll("dl")
##    print len(treeAdd)
    for item in treeAdd:
        soupdd = BeautifulSoup(str(item))
        name_pinpai=soupdd.dt.div.text
        name_chexi=soupdd.findAll('h4')
##        print name_pinpai
        for item in name_chexi:
            soup= BeautifulSoup(str(item))
            print soup.text
            name_chexi1=soup.text
            name_nu=soup.a['href']
            find1=name_nu.find('cn/')
            find2=name_nu.find('/#',find1)
            name_nu1=name_nu[find1:find2][3:]
            name_nu_link="http://car.autohome.com.cn/config/series/"+name_nu1+".html"
            host=name_nu_link
            cookie_jar = cookielib.LWPCookieJar()
            cookie = urllib2.HTTPCookieProcessor(cookie_jar)
            opener = urllib2.build_opener(cookie)
            req = urllib2.Request(host)
            soures_home = opener.open(req).read()
            soup = BeautifulSoup(soures_home,from_encoding="gbk")
            aa=soup.findAll("script")#have:23 no:15
            if len(aa)>15:
                print name_nu_link
                a=aa
            else:
                name_nu_link1=name_nu
                print name_nu_link1
                host=name_nu_link1
                cookie_jar = cookielib.LWPCookieJar()
                cookie = urllib2.HTTPCookieProcessor(cookie_jar)
                opener = urllib2.build_opener(cookie)
                req = urllib2.Request(host)
                soures_home = opener.open(req).read()
##                print soures_home
                soup = BeautifulSoup(soures_home,from_encoding="gbk")
                if soup.find("div",attrs={"class":"modtab1"}):
                    ab=soup.find("div",attrs={"class":"modtab1"})#have:23 no:15("li",attrs={"class":w}
                    abc=ab.table.tr.td.a['href']
                    abc_name=ab.table.tr.td.a.text
                    name_nu_link2="http://car.autohome.com.cn/config/"+abc[:-1]+".html"
                    print name_nu_link2
                    print "#######################################################"
                    host=name_nu_link2
                    cookie_jar = cookielib.LWPCookieJar()
                    cookie = urllib2.HTTPCookieProcessor(cookie_jar)
                    opener = urllib2.build_opener(cookie)
                    req = urllib2.Request(host)
                    soures_home = opener.open(req).read()
                    soup = BeautifulSoup(soures_home,from_encoding="gbk")
                    ac=soup.findAll("script")#have:23 no:15
                    a=ac[1:]
                else:
                    sheet.write(row,0,name_pinpai)
                    sheet.write(row,1,name_chexi1)
                    sheet.write(row,2,"none")
                    row=row+1
                    a=['','','','','']
                    print "--------------------------------------------------------------none"
            if (str(a[4]))>20000:
                b=(str(a[4]))[22394:-14]
                c=b.split('var')
                if len(c)>2:##################判断车系详细信息是否存在####################################
                    d=[c[1][10:-11],c[2][10:-11],c[4][9:-11],c[5][7:-11]]
                    e=d[0].split('"valueitems":')
                    f=re.findall('"value":"',e[1])
                    row0=len(f)############################得到车系有多少车型
                    col0=0
                    while col0<row0:
                        sheet.write(row+col0,0,name_pinpai)
                        sheet.write(row+col0,1,name_chexi1)
                        col0=col0+1
                    s = json.loads(d[0])
                    ii=0#用来判断两个"车身型号"类型的
                    for item in s['result']['paramtypeitems']:#var1解析基本参数到车轮制动
##                        print "----------------------------------------"
##                        print item['name']
##                        print "1111111111111111111"
                        for item1 in item['paramitems']:
                            ii=ii+1
##                            print ii
                            x=item1['name']
##                            print x
                            colnu=allstrlist.index(x)
##                            print colnu
                            if colnu==9:
                                if ii>10:
                                    colnu=26
                                else:
                                    colnu=9
                            row1=0
                            for item2 in item1['valueitems']:
##                                print item2['value']
                                sheet.write(row+row1,colnu,item2['value'])
##                                sheet.write(row+row1,0,name_pinpai)
                                row1=row1+1
                    print "--------------------------------------------------------------OK1"
        #################################################################################################
                    s1 = json.loads(d[1])
                    for item in s1['result']['configtypeitems']:#var1解析车轮制动到高科技配置
##                        print "----------------------------------"
##                        print item['name']
##                        print "22222222222222222222"
                        for item1 in item['configitems']:
                            x=item1['name']
##                            print x
                            colnu=allstrlist.index(x)
##                            print colnu
                            row2=0
                            for item2 in item1['valueitems']:
##                                print item2['value']
                                sheet.write(row+row2,colnu,item2['value'])
##                                sheet.write(row+row2,0,item2['value'])
                                row2=row2+1
                    print "--------------------------------------------------------------OK2"
    ###########################################################################################################
                    try:
                        s2 = json.loads(d[2])
                        row3=0
                        for item in s2['result']['specitems']:#var1解析车外观颜色，前提是要存在外部颜色
                            strs=''
                            for item1 in item['coloritems']:
                                strs=strs+item1['name']+' '
                                #print item1['name']
    ##                        print strs                  
                            sheet.write(row+row3,205,strs)
                            row3=row3+1
                        print "--------------------------------------------------------------OK3"
                    except:
                        s2 = json.loads(d[2][4:])
                        row3=0
                        for item in s2['result']['specitems']:#var1解析车外观颜色，前提是要存在外部颜色
                            strs=''
                            for item1 in item['coloritems']:
                                strs=strs+item1['name']+' '
                                #print item1['name']
    ##                        print strs                  
                            sheet.write(row+row3,205,strs)
                            row3=row3+1
                        print "--------------------------------------------------------------OK3_2"
    #####################################################################################################
                    try:
                        s3 = json.loads(d[3][5:])
                        row4=0
                        for item in s3['result']['specitems']:#var1解析车内部颜色，前提是要存在内部颜色
                            strs=''
                            for item1 in item['coloritems']:
                                strs=strs+item1['name']+' '
                                #print item1['name']
    ##                        print strs
                            sheet.write(row+row4,206,strs)
                            row4=row4+1
##                        row=row+row0
                        print "--------------------------------------------------------------OK4"
                    except:
                        s3 = json.loads(d[3])
                        row4=0
                        try:
                            for item in s3['result']['specitems']:#var1解析车内部颜色，前提是要存在内部颜色
                                strs=''
                                for item1 in item['coloritems']:
                                    strs=strs+item1['name']+' '
                                    #print item1['name']
        ##                        print strs
                                sheet.write(row+row4,206,strs)
                                row4=row4+1
                        except: pass
                        print "--------------------------------------------------------------OK4_2"
                    row=row+row0
filename.save("./output/all.xls")












    
