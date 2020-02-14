# -*- coding: utf-8 -*-
import urllib.request
import urllib.error
import http.cookiejar
from PIL import ImageTk
from PIL import Image as Image
from tkinter import *
import PIL
import tkinter as tk
import io
import os
import time
import xlrd
import openpyxl
from urllib import parse
import random
import json

tobeCheckedCodeIMG = Image
finalResultRes = ""

success = 0
passexam = 0

headers = {
        "Content-type": "application/x-www-form-urlencoded",
        'Accept-Language': 'zh-CN,zh;q=0.8',
        'User-Agent': "Mozilla/5.0 (Windows NT 6.1; rv:32.0) Gecko/20100101 Firefox/32.0",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Connection": "close",
        "Cache-Control": "no-cache",
        "Referer": "http://cet.neea.edu.cn/cet/"
    }

# 验证码窗口相关代码
class GetCode(object):
    def __init__(self):
        global tobeCheckedCodeIMG
        global nowCode
        print("请在新窗口中输入验证码...")
        self.data={}  # 存放返回值
        self.root = tk.Tk()
        self.root.geometry('230x200')
        self.root.resizable(width=False,height=False)   # 固定长宽不可拉伸

        self.textLabel=tk.Label(self.root,text="输入验证码后按下回车(看不清直接回车)").pack()  # 标签
        self.textStr=StringVar()
        self.textEntry=tk.Entry(self.root,textvariable=self.textStr)  # 创建输入框
        self.textStr.set("")  # 清空输入框
        self.textEntry.pack()  # 输入框
        self.textEntry.bind('<Return>', self.return_code)  # 回车键按下自动递交
        self.textEntry.focus_set()  # 设置焦点

        im=PIL.Image.open(io.BytesIO(tobeCheckedCodeIMG))
        img=ImageTk.PhotoImage(im)
        tk.Label(self.root,image=img).pack() # 显示图片

        self.root.protocol('WM_DELETE_WINDOW', doNothing)  # 禁止关闭
        self.root.mainloop()

    def return_code(self, x):
        global nowCode
        # 返回输入框内容
        self.data["code"]=self.textStr.get()
        if(len(self.data["code"]) == 0):
            self.data["code"] = "NNNN"
        self.root.destroy()           # 关闭窗体
        nowCode = self.data["code"]


def doNothing():
        return
# 验证码窗口结束

# 爬取页面代码
def getPage(resultReq):
    finalRequlstRes = "ERROR"
    try:
        finalResultRes = urllib.request.urlopen(resultReq).read().decode("utf-8")  # 发送请求并保存页面
    except Exception as e:
        print("遇到HTTP错误:" + str(e))
        print("等待一秒")
        time.sleep(1)  # 等待1秒
        getPage(resultReq)
    else:
        # 开始检查返回页面正确性
        if (str(finalResultRes).find('核实后') != -1):
            # 信息不正确
            print("请重新核实信息!")
            finalResultRes = ""
        else:
            # 信息正确，开始检查验证码
            if (str(finalResultRes).find('验证码错误') == -1):
                # 验证码正确
                print("检查通过!")
            else:
                # 验证码错误
                print("验证码错误!")
                finalResultRes = "ERROR"
    return finalResultRes
# 爬取页面代码结束

# studentListExcel(studentList.xlsx)为学生信息表.该表结构如下：（第一行就应该开始放学生信息，不要有题头）
# |准考证号|姓名|
studentListExcel = 'studentList.xlsx'
# finalResult.xlsx为最终结果存储表，该表结构如下：(会自动生成题头)
# |准考证号|姓名|查询类型|大学名称|总分|听力|阅读|写作和翻译|口试准考证号|口试等级|
finalResultExcel = 'finalResult.xlsx'
# queryYear为查询成绩的年份。如要查询2019年上半年考试此变量应为19.不要去掉引号
queryYear = '19'
# queryTime为查询第几次考试。如要查询2019年上半年考试此变量应为1,下半年应为2.不要去掉引号
queryTime = '1'
# passExamMark为及格线。大于等于此成绩为及格.
passExamMark = 425
# 默认查询类型 CET4/CET6
cxlx = 'CET4'

if not os.path.exists(studentListExcel):
    print("学生信息表不存在!")
    exit()

# 读取xlsx表中信息并存入列表
sheet = xlrd.open_workbook(studentListExcel)
table = sheet.sheets()[0]  # 打开Sheet1
nrows = table.nrows  # 查询总行数
allStudents = nrows

# 第一名学生插入List中为第0个
print("正在读取第1个学生，准考证号:"+table.col_values(0)[0]+" 姓名:"+table.col_values(1)[0])
studentList = [[table.col_values(0)[0], table.col_values(1)[0]]]
# 从第二名(List中第1个)开始循环读取到最后一名学生插入List中
i = 1
while i < allStudents:
    print("正在读取第"+str(i+1)+"个学生，准考证号:"+table.col_values(0)[i]+" 姓名:"+table.col_values(1)[i])
    studentList.append([table.col_values(0)[i], table.col_values(1)[i]])
    i = i + 1

print("本次共导入"+str(allStudents)+"个学生")
print("正在打开存储数据文件...")
workbook = openpyxl.Workbook()  # 打开表
sheet = workbook.active  # 默认的第一张sheet
rawList = []
sheet.append(['准考证号','姓名','查询类型','大学名称','总分','听力','阅读','写作和翻译','口试准考证号','口试等级'])
workbook.save(finalResultExcel)  # 记得保存数据
workbook = openpyxl.load_workbook(finalResultExcel)  # 读取新建的文件
sheet = workbook.active
print("准备完成。")

queryYearTmp = input("本程序查询为"+queryYear+"年考试，是否正确？正确请直接按回车，不正确请输入年份后两位后回车。")
if queryYearTmp.isdecimal():
    if(len(queryYearTmp) == 2):
        queryYear = queryYearTmp
        print("输入了"+queryYear)
    else:
        print("输入只能为2位数!退出.")
        exit()
elif len(queryYearTmp) == 0 or queryYearTmp.isspace():
    print("无输入,继续")
else:
    print("输入非数字!退出.")
    exit()
queryTimeTmp = input("本程序查询为当年第"+queryTime+"次考试。上半年为1，下半年为2.正确请直接按回车，不正确请输入1或2后回车。")
if queryTimeTmp == '1' or queryTimeTmp == '2':
    queryTime = queryTimeTmp
    print("输入了"+queryTime)
elif queryTimeTmp == 1 or queryTimeTmp == 2:
    queryTime = str(queryTimeTmp)
    print("输入了" + queryTime)
elif len(queryTimeTmp) == 0 or queryTimeTmp.isspace():
    print("无输入,继续")
else:
    print("输入错误!只能输入1或2,退出.")
    exit()
cxlxTmp = input("本程序查询为"+cxlx+"考试。英语四级为CET4，英语六级为CET6.正确请直接回车,不正确请输入CET4或CET6后回车")
if cxlxTmp == 'CET4' or cxlxTmp == 'CET6':
    cxlx = cxlxTmp
    print("输入了"+cxlxTmp)
elif len(cxlxTmp) == 0 or cxlxTmp.isspace():
    print("无输入,继续")
else:
    print("输入错误!只能输入CET4或CET6,退出.")
    exit()

print("开始爬取。")
j = 0
while j < allStudents:
    # 开始准备...
    zkzh = str(studentList[j][0])
    xm = studentList[j][1]
    zkzh_parse = parse.quote(zkzh)
    print("进度: 第"+str(j+1)+"个，共"+str(allStudents)+"个。正在爬取: 姓名:" + xm + " 准考证号: " + zkzh + " 查询类型: "+cxlx)
    mainUrl = 'http://cache.neea.edu.cn/cet/query'
    captchaUrl = 'http://cache.neea.edu.cn/Imgs.do?c=CET&ik='+zkzh+'&t='+str(random.random())
    print("本次查询使用url: " + mainUrl)
    # 开始请求
    resultPage = "ERROR"
    while (resultPage == "ERROR"):
        # 开始爬取验证码
        reqCode = urllib.request.Request(url=captchaUrl, headers=headers)
        cjar = http.cookiejar.CookieJar()
        cookie = urllib.request.HTTPCookieProcessor(cjar)
        opener = urllib.request.build_opener(cookie)
        urllib.request.install_opener(opener)
        print("正在获取验证码地址...地址: " + captchaUrl)
        getCodeReady = False
        while (not getCodeReady):
            try:
                captchaNewUrl = urllib.request.urlopen(reqCode).read().decode("utf-8")[13:45]
                #captchaUrl = str(captchaUrl)[2:34]
                captchaNewUrl = "http://cet.neea.edu.cn/imgs/"+captchaNewUrl+".png"
                print("正在获取验证码...地址: " + captchaNewUrl)
                reqNewCode = urllib.request.Request(url=captchaNewUrl, headers=headers)
                tobeCheckedCodeIMG = urllib.request.urlopen(reqNewCode).read()
                GetCode()  # 弹出输入验证码窗口
            except Exception as e:
                print("遇到HTTP错误:" + str(e))
                print("暂停一秒后继续重试")
                time.sleep(1)
            else:
                getCodeReady = True
        print("输入了验证码: " + nowCode)
        dataUrl = mainUrl + "?data="+cxlx+"_"+queryYear+queryTime+"_DANGCI,"+zkzh+","+urllib.parse.quote(xm[:3])+"&v="+nowCode
        print("获取成绩Url: "+dataUrl)
        # 构建请求
        resultReq = urllib.request.Request(url=dataUrl, headers=headers)
        resultPage = str(getPage(resultReq))  # 返回结果
    # 判断是否信息不正确
    if (resultPage == ""):
        # 信息确实不正确，标记上
        rawList = [zkzh, xm, cxlx, "核实信息"]
        sheet.append(rawList)  # 追加一行
        print("正在保存...")
        workbook.save(finalResultExcel)  # 记得保存数据
    else:
        # 信息正确，洗数据吧。
        #print("正在提取成绩信息...信息原文: "+resultPage)
        resultPage = resultPage[16:]
        resultPage = resultPage[:-2]
        resultPage = resultPage.replace("'", '"')
        resultPage = resultPage.replace(",", ',"')
        resultPage = resultPage.replace(":", '":')
        resultPage = resultPage.replace("{", '{"')
        resultPage = resultPage.replace(":.00", ":0")
        print("正在提取信息: "+resultPage)
        try:
            resultJson = json.loads(resultPage)
        except Exception as e:
            print("解析结果发生错误: "+str(e))
            rawList = [zkzh, xm, cxlx, "接口错误: "+str(e)+" R:"+resultPage]
        else:
            success = success + 1
            rawList = [zkzh, xm, cxlx, resultJson['x'], resultJson['s'], resultJson['l'], resultJson['r'],
                       resultJson['w'], resultJson['kyz'], resultJson['kys']]
            if resultJson['s'] < passExamMark:
                print("该学生未通过!")
            else:
                passexam = passexam + 1
        print(rawList)  # 输出一下
        sheet.append(rawList)  # 追加一行
        print("正在保存...")
        workbook.save(finalResultExcel)  # 记得保存数据
    j = j + 1  # 别忘了自增开始下一名学生的处理。

passRate = passexam / success

rawList = ['', '总计', '成功获取:', success, '通过考试人数:', passexam, '通过率', str(passRate*100)+'%']
print("任务结束! 查询考试类型: "+cxlx+" "+queryYear+"年第"+queryTime+"次考试")
print("总计"+str(allStudents)+"个学生, 成功获取"+str(success)+"个学生, 有"+str(passexam)+"个学生通过考试. 通过率"+str(passRate*100)+"%")
sheet.append(rawList)  # 追加一行
print("正在保存...")
workbook.save(finalResultExcel)

