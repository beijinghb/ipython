#-*- coding: utf8 -*-
# Version 0.5
#配置smtp信息，发送工资条邮件

from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr
import smtplib
import xlrd,threading,queue
import os,sys,time,string,re
import configparser

def _format_addr(s): #格式化邮件信息
    name, addr = parseaddr(s)
    return formataddr((Header(name,'utf-8').encode(),addr))

class Sender(threading.Thread): #发送邮件--线程类对象
    def __init__(self):
        super(Sender,self).__init__()
    def run(self):
        global conf,q,errAccount
        server = smtplib.SMTP(conf['smtp'], 25) # SMTP协议默认端口是25
##        #server.set_debuglevel(1)
        
        try: # 检测登录是否OK?
            server.login(conf['from'], conf['pwd'])
        except Exception as e:
            lock.acquire()
            print(e)
            lock.release()
        else:
            while True:
                msg=q.get()
                confto=msg[0]
                
                try:# 检测邮件发送是否OK?
                    server.sendmail(conf['from'], msg[0], msg[1].as_string())
                except Exception as e:
                    lock.acquire()
                    errAccount.append(msg[0])
##                    print(e)
                    lock.release()
                else:
                    lock.acquire()
                    print('%-2s%-30s%-25s%s'%("√",msg[0],time.time(),self.name))
                    lock.release()
                q.task_done() #告诉队列取数后的操作已完毕。
                
            server.quit()
            print("%s is empty"%self.name)
            
        
        

def Ldump(*txt):
    lock.acquire()
    print(*txt)
    lock.release()



def Msg_encode(conf,th,i,td):
    # print("td=",td)
    conf['to']=td[0]
    content=html_head+"<table>"
    content+=th+td[2]
    content+="</table>"+html_end
    msg = MIMEText(content, 'html', 'utf-8')
    msg['From'] = _format_addr('财务 <%s>' % conf['from'])
    msg['To'] = _format_addr('%s <%s>' %(td[1],td[0]))
    msg['Subject'] = Header(conf['subject'], 'utf-8').encode()
    return conf['to'],msg
    



def getConf():
    pass

def htmlFile(th,d):
    print("htmlFile: th=",th)
    fname=r"d:\0701.html"
    if os.path.exists(fname):
        os.remove(fname) 
    with open(fname,'w') as f:
        content=html_head+"<table>"
        for k in d:
            for x in d[k]:
                content+=th[2]+x[2]+"</table>"+x[0]+"<table>"
        content+="</table>"+ html_end
        f.write(content)
    os.startfile(fname)
    


#@i_row   工资条标题栏起始行号
#@i_name  姓名的列号
#@i_mail  邮箱的列号
#@lab     标题栏数据--list（首行，次行）
#@d_lab   标题栏数据--dict（合并的列数，行数）
def th_encode(sh):
    cv=sh.col_values(0)
    i_row=cv.index('序号') #  工资条标题栏的行号
    rv=sh.row_values(i_row)
    i_name=rv.index("姓名")  #  工资条标题栏：姓名的列号
    i_mail=rv.index("邮箱") #  工资条标题栏：邮箱的列号

    # -- 标题栏首行数据 --
    d_lab={}
    lab=([],[]) 
    rv.reverse()
    t=0
    for x in rv:
        if x:
            d_lab[x]=([t,0])
            lab[0].append(x)
            t=0
            
        else:
            t+=1
    lab[0].reverse()
    lab[0].remove("邮箱")
    lab[0].remove("序号")

    # -- 标题栏次行数据 --
    # print(sh.cell_value(i_row+1,1))
    if sh.cell_value(i_row+1,0): #如果不为空，标题栏为2行
        pass
    else:
        rv2=sh.row_values(i_row+1)
        for i,x in enumerate(rv2):
            if x:
                lab[1].append(x)
            else: # -- 统计占两行的标题栏 --
                v=sh.cell_value(i_row,i)
                d_lab[v][1]=1
    th_html="<tr>"

    for x in lab[0]:
        th_html+="<th "
        if d_lab[x][0]:
            th_html+="colspan="+str(d_lab[x][0]+1)
        if d_lab[x][1]:
            th_html+="rowspan="+str(d_lab[x][1]+1)
        th_html+=">"+x+"</th>"
    th_html+="</tr>"
    if lab[1]:
        th_html+="<tr>"
        for x in lab[1]:
            th_html+="<th>"+x+"</th>"
        th_html+="</tr>"
    return i_mail,i_name,th_html

#@td  员工工资条数据html格式--list
#@prama sh   sheet对象
#@prama i    邮箱的列号
#@prama j    姓名的列号
#@prama d    td数据存储容器
def td_encode(sh,i,j,d):
    nrows=sh.nrows
    for n in range(1,nrows):
        rv=sh.row_values(n)
        mail=rv[i]
        data=[]
        if isinstance(mail,str) and re.match(r'^(\w+[\-\.]?\w+)@(\w+\-?\w+)(\.\w+)$',mail.strip()):
            td="<tr><td>"+rv[j]+"</td>"
            data.append(mail)
            rv.pop(i)
            rv.pop(0)
            data.append(rv[0])
            for y in rv[1:]:
                td+="<td>"+("%s"%y)+"</td>"
            td+="</tr>"
            data.append(td)
            d.append(data)
    

#@i_name  姓名的列号
#@i_mail  邮箱的列号
#@th_sign 标题栏分析状态（0：完成 1：未进行）
#@th_html 标题栏html的table格式
#@td_dhtml  员工工资条数据html格式--list
def readXLS(fname):
    th_sign=1
    bk=xlrd.open_workbook(fname)
    shname=bk.sheet_names()
    for s in shname:
        # -- 工资条标题栏生成 --
        if "部" in s: 
            sh=bk.sheet_by_name(s)
            if th_sign: #生成标题栏
                th.extend(th_encode(sh))
                # print("th=",th,"\n")
                th_sign=0
                print(("%-20s"%"....工资条标题栏"),"OK！")
            d[s]=[]
            td_encode(sh,th[0],th[1],d[s]) # 生成工资数据
    print(("%-22s"%"....工资数据"),"OK！")
##            break               
def setGlobal():
    global conf,html_head,html_end,q,lock,errAccount
    conf={}
    html_head='''<html>
            <head>
            <meta charset="GBK">
            <style type="text/css">
            #mainbox {margin:5 auto;}
            table {border-collapse:collapse;width:88%;margin:0 auto;}
            table,tr,th,td {
                border:1px solid #000;
                text-align:center;
                    }
            th{background-color:#eee}

            </style>
            </head>
            <body>
            <div id="mainbox">'''
    html_end='''    </div>
            </body>
            </html>'''
    q=queue.Queue()
    lock=threading.Lock()
    errAccount=[]

def SetConf():
    fname="payroll.cfg"
    print('        *--*--*--*--*--*--*--*--*--*--*--*--*')
    print('        |                                   |')
    print('    ----|       工资条发送程序                |----')
    print('        |                                   |')
    print('        *--*--*--*--*--*--*--*--*--*--*--*--*')
    conf={}
    cf = configparser.ConfigParser()
    if os.path.exists(fname):#确认配置文件存在，则读取配置
        cf.read(fname)
        conf['from']=cf.get('smtpset','from')
        conf['pwd']=cryptcode(cf.get('smtpset','pwd'),1)
        conf['smtp']=cf.get('smtpset','smtp')
    else:
        conf['from']=input("请设置发送邮箱：")
        conf['pwd']=input("请输入邮箱密码：")
        conf['smtp']=input("请设置SMTP：")    
        cf.add_section("smtpset")#增加section 
        cf.set("smtpset", "from", conf["from"])#增加option 
        cf.set("smtpset", "pwd", cryptcode(conf["pwd"])) 
        cf.set("smtpset", "smtp", conf["smtp"])  
        with open(fname, "w") as f: 
            cf.write(f)#写入配置文件文件中   
#  -- 设置邮件主题（某年某月工资条） --
    sub=time.strftime('%Y%m',time.localtime())
    month=int(sub[4:6])-1
    if month:
        conf['subject']='%s年%02d月工资条'%(sub[:4],month)
    else:
        conf['subject']='%s年%02d月工资条'%(int(sub[:4])-1,12)

    conf['fxls']=''
    while conf['fxls']=='':
        conf['fxls']=input('请输入工资表格文件：').lower()
        if not(re.match(r'.*\.xls(x?)$',conf['fxls'])):
            conf['fxls']=''
    print('    *--*--*--*--*--*--*--*--*--*--*--*--*')
    print('    |  配置信息如下：')
    print('    |  发送邮箱：%s'%conf['from'])
    print('    |  SMTP服务器：%s'%conf['smtp'])
    print('    |  工资条表格文件：%s'%conf['fxls'])
    print('    |  邮件标题：%s'%conf['subject'])
    print('    *--*--*--*--*--*--*--*--*--*--*--*--*')
    return conf
    

def iSelect():
    conf['cmd']=''
    inbox='浏览工资条(B)/'
    print("%-30s"%'    *--*--*--*--*--*--*--*--*--*--*--*--*')
    while not(conf['cmd']): 
        conf['cmd']=input(inbox+'发送邮件(S)/退出(Q):').lower()
        if conf['cmd']=='b':
            cmdBrow()
            inbox='重新浏览工资条(B)/'
            print("%-30s"%'    *--*--*--*--*--*--*--*--*--*--*--*--*')
            conf['cmd']=''
        elif conf['cmd']=='q':
            sys.exit()
        elif conf['cmd']=="s":
            cmdSend()
        else:
            conf['cmd']=''

# -- 浏览数据 --
def cmdBrow():
    htmlFile(th,d)

# -- 发送工资条 --            
def cmdSend():
    # print("Send mail is begining……")
    print("开始发送电子邮件……")
    # --工资条邮件生成 --
    for k in d:
        for v in d[k]:
                data=Msg_encode(conf,th[2],th[0],v)
                q.put(data)
    # -- 工资条邮件多线程发送 --
    Consumer=[Sender() for i in range(4)]
    start=time.time() #计时开始
    for c in Consumer:
        c.daemon = True
        c.start()
    q.join()

    # -- 发送完毕，输出结果 --
    # print("\nSend completed!  Total spending time:：%s"%(time.time()-start))
    # print("\nThe following is a list of error accounts：")
    print("\n发送完毕!  总计用时：%s"%(time.time()-start))
    print("\n以下邮件账号发送失败：")
    print("*"*25)
    for x in errAccount:
        print("%-2s%-25s"%("×",x))
    print("*"*25)

if __name__ == '__main__':
    setGlobal()
    if not(conf):
        setConf()
    pay_label="" #标题栏数据,默认为空
    i_height=1  #标题栏行高，默认为1
    fname=r"d:\code\js\123.xls"
    th=[]
    d={}  #@工资数据
    # print("Data analysis is begining……")
    print("数据分析开始……")
    readXLS(fname)
    # print("Data analysis is completed！")
    iSelect("数据分析完毕。")
    
       


    




 

