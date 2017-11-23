#!/usr/bin/python  
#encoding:utf-8 

#****************************************************************    
# Description: 主要函数文件  
#**************************************************************** 
import os,sys, urllib, httplib, profile, datetime, time ,re ,ast
from xml2dict import XML2Dict
import win32com.client  
import xml.etree.ElementTree as et
import sendEmail
  
#Excel表格中测试结果底色
OK_COLOR=0xffffff  
NG_COLOR=0xff  
#NT_COLOR=0xffff  
NT_COLOR=0xC0C0C0  
  
#Excel表格中测试结果汇总显示位置  
TESTTIME=[1, 14]  
TESTRESULT=[2, 14]  
  
#Excel模版设置  
#self.titleindex=6        #Excel中测试用例标题行索引  
#self.casebegin =7        #Excel中测试用例开始行索引  
#self.argbegin   =3       #Excel中参数开始列索引  
#self.argcount  =10        #Excel中支持的参数个数  

        
class create_excel:
    

    
    def read_data(self, iSheet, iRow, iCol):  
        try:
            sht = self.book.Worksheets(iSheet)   
            sValue=sht.Cells(iRow, iCol).value
            #为了兼容中文，这里做一下try处理。
            try:
                sValue=GetStr(sValue)
            except:
                sValue=str(sValue)
        except:  
            self.close()  
            print(str(iRow)+'行'+str(iCol)+'列读取数据失败')  
            exit()  
        #去除'.0'  
        if sValue[-2:]=='.0':  
            sValue = sValue[0:-2]  
        
        return sValue
  
    def write_data(self, iSheet, iRow, iCol, sData, color=OK_COLOR):  
        try:  
            sht = self.book.Worksheets(iSheet)
            sData=ToUnicode(sData)
            sht.Cells(iRow, iCol).Value = sData
            sht.Cells(iRow, iCol).Interior.Color=color  
            self.book.Save()  
        except:  
            self.close()  
            print(str(iRow)+'行'+str(iCol)+'列写入数据失败')  
            exit()
    
          
    def __init__(self, sFile,suiteid,dtitleindex=6, dcasebegin=7, dargbegin=3, dargcount=10): 
        #定义参数个数、请求发送方法、预期结果中检查的项
        global argsconut,reqmethod,reqHeaders            
        self.xlApp = win32com.client.Dispatch('Excel.Application')   #MS:Excel  WPS:et  
        try:  
            self.book = self.xlApp.Workbooks.Open(sFile)
        except:  
            print_error_info()  
            print "打开文件失败"  
            exit() 
        if suiteid == 'ALL':
            suiteid =  self.book.Worksheets(1).Name
        self.file=sFile  
        self.titleindex=dtitleindex  
        self.casebegin=dcasebegin  
        self.argbegin=dargbegin  
        self.argcount=dargcount  
        self.allresult=[]                     
        argsconut=self.read_data(suiteid, 2, 2)
        self.argscount=argsconut
        self.CheckKey=self.argbegin+self.argcount
        self.CheckVaule=self.CheckKey+1
        self.RealCol=self.CheckVaule+1  
        self.resultCol=self.RealCol+1
        reqmethod=self.read_data(suiteid, 4, 2) 
        reqHeaders=self.read_data(suiteid, 3, 2)
        if reqHeaders =="None":
            reqHeaders={}
        elif reqHeaders.startswith("'"):            
            reqHeaders=ast.literal_eval("{" + reqHeaders +'}')
        else:
            reqHeaders=ast.literal_eval("{'" + reqHeaders +'}')                                    
          
    def close(self):  
        #self.book.Close(SaveChanges=0)  
        self.book.Save()  
        self.book.Close()  
        #self.xlApp.Quit()  
        del self.xlApp  

    def get_all_sheetname(self):
        #获取所有sheet页名称
        list = []
        count = self.book.Worksheets.Count
        for i in range (int(count)):
            sheetname = self.book.Worksheets(i+1).Name
            list.append(str(sheetname))
        return list
    #获取用例个数      
    def get_ncase(self, iSheet):  
        try:  
            return self.get_nrows(iSheet)-self.casebegin+1  
        except:  
            self.close()  
            print('获取Case个数失败')  
            exit()  
      
    def get_nrows(self, iSheet):  
        try:  
            sht = self.book.Worksheets(iSheet)  
            return sht.UsedRange.Rows.Count  
        except:  
            self.close()  
            print('获取nrows失败')  
            exit()  
  
    def get_ncols(self, iSheet):  
        try:  
            sht = self.book.Worksheets(iSheet)  
            return sht.UsedRange.Columns.Count  
        except:  
            self.close()  
            print('获取ncols失败')  
            exit()  
      
    def del_testrecord(self, suiteid):  
        try:  
            #为提升性能特别从For循环提取出来  
            nrows=self.get_nrows(suiteid)+1  
#             ncols=16  
#             begincol=self.argbegin+self.argcount+2  
              
            #提升性能  
            sht = self.book.Worksheets(suiteid)  
  
            for row in range(self.casebegin, nrows):  
                #清除TestResul列中的测试结果，设置为NT  
                self.write_data(suiteid, row,  self.RealCol, ' ', OK_COLOR)  
                self.write_data(suiteid, row, self.resultCol, 'NT', NT_COLOR)  
        except:  
            self.close()  
            print('清除数据失败')  
            exit()  
#统一变为为unicode
def ToUnicode(text):
        """
        | ##@函数目的: 将字符串转化成unicode字符串
        | ##@参数说明：
        | ##@返回值：  text的字符串形式
        | ##@函数逻辑：先后以utf8、gbk、utf16的形式转化text。通常不考虑utf16be的情形
        """
        result = text
        if type(text) == str:
            try:
                result = text.decode("utf8")
                if result.encode("utf8") == text:
                    pass
                else:
                    raise Exception("not right conversion")
            except:
                try:
                    result = text.decode("gbk")
                    if result.encode("gbk") == text:
                        pass
                    else:
                        raise Exception("not right conversion")
                except:
                    try:
                        result = text.decode("utf16")
                        if result.encode("utf16") == text:
                            pass
                        else:
                            raise Exception("not right conversion")
                    except:
                        pass
    
        return result    

def GetUtf8Str(content):
    try:
        #如果是unicode字符，则进行utf-8编码
        value = content.encode("utf-8")
        return value
    except:
        #否则就是str类型
        #先进行gbk解码
        try:
            value = content.decode("gbk").encode("utf-8")
            return value
        except:
            #否则进行utf-8解码
            try:
                value = content.decode("utf-8").encode("utf-8")
                return value
            except:
                #如果都不是，返回空，暂时写到这
                return str(value)


def GetGBKStr(content):
    try:
        #如果是unicode字符，则进行gbk编码
        value = content.encode("gbk")
        return value
    except:
        #否则就是str类型
        #先进行gbk解码
        try:
            value = content.decode("gbk").encode("gbk")
            return value
        except:
            #否则进行utf-8解码
            try:
                value = content.decode("utf-8").encode("gbk")
                return value
            except:
                #如果都不是，返回空，暂时写到这
                return str(value)
def GetStr(text):
    if type(text) == str:
        return text
    else:
        try:
            return GetUtf8Str(text)
            
        except:
            return GetGBKStr(text)
#执行调用  
def HTTPInvoke(url,requestUri):
    proto,rest=urllib.splittype(url)
    host,rest =urllib.splithost(rest)
    conn = httplib.HTTPConnection(host)  
    if reqmethod.upper()=="GET":
        print url
        conn.request(reqmethod.upper(), url,headers=reqHeaders)
        rsps = conn.getresponse()
        if rsps.status==200:
            data = rsps.read()
            data = str(data)  
            conn.close()  
            return data
        elif rsps.status==301 or rsps.status==302:
            headerstr=rsps.getheaders()
            for i in headerstr:
                if i[0].lower()=='location':
                    url = i[1]
                    proto,rest=urllib.splittype(url)
                    host,rest =urllib.splithost(rest)
                    conn = httplib.HTTPConnection(host)  
                    conn.request('GET', url)
                    rsps = conn.getresponse()
                    if rsps.status==200:
                        data = rsps.read()
                        data = str(data)  
                        conn.close()  
                        return data
        else:
            data='[Response_Code_err]:'+str(rsps.status)
            data = str(data)  
            return data
    if reqmethod.upper()=="POST":
        print requestUri + '\t' + 'body=' + sArge
        conn.request(reqmethod.upper(), requestUri,body=sArge,headers=reqHeaders)  
        rsps = conn.getresponse()  
        if rsps.status==200:
            data = rsps.read()
            data = str(data)  
            conn.close()  
            return data
        elif rsps.status==301 or rsps.status==302:
            headerstr=rsps.getheaders()
            print headerstr
            for i in headerstr:
                if i[0].lower()=='location':
                    url = i[1]
                    proto,rest=urllib.splittype(url)
                    host,rest =urllib.splithost(rest)
                    conn = httplib.HTTPConnection(host)
                    conn.request('GET', url)
                    rsps = conn.getresponse()
                    if rsps.status==200:
                        data = rsps.read()
                        data = str(data)  
                        conn.close()  
                        return data
        else:
            data='[Response_Code_err]:'+str(rsps.status)
            return data 
def Correl(url,LB,RB):
    #关联函数
    #参数说明：必须参数：url，LB：左边界，RB：右边界
    #可选参数：headers：item类型，默认为{}，ORD:整型，如果匹配多个，则选择使用第几个，默认使用第一个。
    proto,rest=urllib.splittype(url)
    host,rest =urllib.splithost(rest)
    conn = httplib.HTTPConnection(host) 
    conn.request('GET', url,headers=reqHeaders)
    rsps = conn.getresponse()
    if rsps.status==200:
        data = rsps.read()
        data = str(data)  
        conn.close()  
        pattern  =  re.compile(str(LB)+'(.*?)'+str(RB))
        colvalue = pattern.findall(data)
        return colvalue

  
#获取用例基本信息[Interface,argcount,[ArgNameList]]  
def get_caseinfo(Data, SuiteID):  
    caseinfolist=[]  
    sInterface=Data.read_data(SuiteID, 1, 2)   
    argcount=int(Data.read_data(SuiteID, 2, 2))
      
    #获取参数名存入ArgNameList   
    ArgNameList=[]  
    for i in range(0, argcount):  
        ArgNameList.append(Data.read_data(SuiteID, Data.titleindex, Data.argbegin+i))    
      
    caseinfolist.append(sInterface)  
    caseinfolist.append(argcount)  
    caseinfolist.append(ArgNameList)  
    return caseinfolist  
  
#获取输入  
def get_input(Data, SuiteID, CaseID, caseinfolist):  
    global sArge
    sArge=''
    args = []  
    #对于get请求，将参数组合  
    if reqmethod.upper()=='GET':
        for j in range(0, caseinfolist[1]):  
            if Data.read_data(SuiteID, Data.casebegin+CaseID, Data.argbegin+j) != "None": 
                ArgValue =  Data.read_data(SuiteID, Data.casebegin+CaseID, Data.argbegin+j)
                if '$$' in ArgValue:#走关联分支
                    args = ArgValue.split('$$')
                    #print args
                    corvalue = Correl(args[0], args[1], args[2])
                    if corvalue == []:
                        sArge = 'correlerr'
                        #return sArge
                        #infolog="关联失败"
                        #ret1 = 'NG'
                        #Data.write_data(SuiteID, Data.casebegin+CaseID, 15,infolog,NG_COLOR)
                        #write_result(Date, SuiteID, Data.casebegin+CaseID, 16, ret1)
                    else:
                        sArge=sArge+caseinfolist[2][j]+'='+corvalue[0]+'&'                        
                else:
                    sArge=sArge+caseinfolist[2][j]+'='+ArgValue+'&'
                #print sArge
        #去掉结尾的&字符  
        if sArge[-1:]=='&':  
            sArge = sArge[0:-1]     
        #sInput=caseinfolist[0]+sArge    #为了post和get分开方便，不在这里组合接口名，在调用的地方组合接口名。
        return sArge 
    #对于post请求，因为不知道连接格式是=还是冒号，或者是其他的格式，所以不做拼接。直接取参数的第一个作为上传body。
    else:
        sArge=Data.read_data(SuiteID, Data.casebegin+CaseID, 3)
        if '$$' in sArge:#走关联分支
            args = sArge.split('$$')
            #print args
            corvalue = Correl(args[0], args[1], args[2])
            if corvalue == []:
                sArge = 'correlerr'
                return sArge
            else:
                return sArge
        
     

   
#结果判断   
def assert_result(sReal, sExpect):
    #预期结果和实际结果都是空的情况需要特别处理下
    if  sReal is None and sExpect == 'None':
        return 'OK'  
    sReal=ToUnicode(sReal)  
    sExpect=ToUnicode(sExpect)
    if sReal==sExpect:  
        return 'OK'  
    else:  
        return 'NG'  
  
#将测试结果写入文件  
def write_result(Data, SuiteId, CaseId, resultcol, *result):  
    if len(result)>1:  
        ret='OK'  
        for i in range(0, len(result)):  
            if result[i]=='NG':  
                ret='NG'  
                break  
        if ret=='NG':  
            Data.write_data(SuiteId, Data.casebegin+CaseId, resultcol,ret, NG_COLOR)  
        else:  
            Data.write_data(SuiteId, Data.casebegin+CaseId, resultcol,ret, OK_COLOR)  
        Data.allresult.append(ret)
        print ret  
    else:  
        if result[0]=='NG':  
            Data.write_data(SuiteId, Data.casebegin+CaseId, resultcol,result[0], NG_COLOR)  
        elif result[0]=='OK':  
            Data.write_data(SuiteId, Data.casebegin+CaseId, resultcol,result[0], OK_COLOR)  
        else:  #NT  
            Data.write_data(SuiteId, Data.casebegin+CaseId, resultcol,result[0], NT_COLOR)  
        Data.allresult.append(result[0])  
      
    #将当前结果立即打印  
    print 'case'+str(CaseId+1)+':', Data.allresult[-1]  
  
#打印测试结果  
def statisticresult(excelobj):  
    allresultlist=excelobj.allresult
    #print allresultlist
    count=[0, 0, 0]  
    for i in range(0, len(allresultlist)):  
        #print 'case'+str(i+1)+':', allresultlist[i]  
        count=countflag(allresultlist[i],count[0], count[1], count[2])  
    print 'Statistic result as follow:'  
    print 'OK:', count[0]  
    print 'NG:', count[1]  
    print 'NT:', count[2]
    return count
  
#解析XmlString返回Dict  
def get_xmlstring_dict(xml_string):  
    xml = XML2Dict()  
    return xml.fromstring(xml_string)  
      
#解析XmlFile返回Dict   
def get_xmlfile_dict(xml_file):  
    xml = XML2Dict()  
    return xml.parse(xml_file)  
  
#去除历史数据expect[real]  
def delcomment(excelobj, suiteid, iRow, iCol, str):  
    startpos = str.find('[')  
    if startpos>0:  
        str = str[0:startpos].strip()  
        excelobj.write_data(suiteid, iRow, iCol, str, OK_COLOR)  
    return str  
      
#检查每个item （非结构体）  
def check_result(excelobj, suiteid, caseid,real_value, checkcol):  
    ret='OK' 
    excelobj.write_data(suiteid, excelobj.casebegin+caseid, 16, 'OK', OK_COLOR) 
    #real=real_value[checklist[checkid]]['value']  
    real=real_value
    expect=excelobj.read_data(suiteid, excelobj.casebegin+caseid, checkcol)  
      
    #如果检查不一致测将实际结果写入expect字段，格式：expect[real]  
    #将return NG 
    result=assert_result(real, expect)
     
    if result=='NG':
        
        writestr=real
        if writestr==None:
            writestr="预期结果的value值为空！"
            excelobj.write_data(suiteid, excelobj.casebegin+caseid, 15, writestr, NG_COLOR)
            excelobj.write_data(suiteid, excelobj.casebegin+caseid, 16, 'NG', NG_COLOR)  
            ret='NG' 
        else:
            excelobj.write_data(suiteid, excelobj.casebegin+caseid, 15, writestr, NG_COLOR)
            excelobj.write_data(suiteid, excelobj.casebegin+caseid, 16, 'NG', NG_COLOR)  
            ret='NG'  

    return ret
   
  
#获取异常函数及行号  
def print_error_info():  
    """Return the frame object for the caller's stack frame."""  
    try:  
        raise Exception  
    except:  
        f = sys.exc_info()[2].tb_frame.f_back  
    print (f.f_code.co_name, f.f_lineno)    
  
#测试结果计数器，类似Switch语句实现  
def countflag(flag,ok, ng, nt):   
    calculation  = {'OK':lambda:[ok+1, ng, nt],    
                         'NG':lambda:[ok, ng+1, nt],                        
                         'NT':lambda:[ok, ng, nt+1]}       
    return calculation[flag]()



'''
def Sendmail(maillist , bodycontent,count,interfacename ):
    #发送邮件通知
    strHtml = ''
    interfacename = str(interfacename)
    strHtml += '<B><p style="font-size:16px">' + interfacename + '测试结果:</p></B>'
    strHtml += '<B><p style="font-size:14px">' + 'Statistic result as follow:</p></B>'
    strHtml += '<B><p style="font-size:14px">' + 'OK:\t' + str(count[0]) + '</p></B>'
    strHtml += '<B><p style="font-size:14px">' + 'NG:\t' + str(count[1]) + '</p></B>'
    strHtml += '<B><p style="font-size:14px">' + 'NT:\t' + str(count[2]) + '</p></B>'
    strHtml += '<table width="1000" border="1">'
    strHtml += '<tr>'
    strHtml += '<td>'+'接口url'+'</td>'
    strHtml += '<td>'+'预期key'+'</td>'
    strHtml += '<td>'+'预期value'+'</td>'
    strHtml += '<td>'+'实际结果'+'</td>'
    strHtml += '</tr>'
    
    for i in range(len(bodycontent)):
        content=''.join(bodycontent[i])
        content = GetStr(content)
        content = str(content)
        #print type(content)
        if i==0 :
            strHtml += '<tr>'
        strHtml += '<td>' + content + '</td>'
        if (i+1) % 4 == 0:
            strHtml += '</tr><tr>'
    strHtml = strHtml[:-4]    
    strHtml += '</table>'
     
        
    mailfr_name = 'caochengzhen'
    mailfr_addr = 'caochengzhen@sogou-inc.com'
    timenow = datetime.datetime.utcnow() + datetime.timedelta(hours=8)#东8区增加8小时
    title = '['+interfacename + '_TestResult]'+ timenow.strftime( '%Y-%m-%d %H:%M:%S' )
    body = strHtml
    mode = 'html'
    url = 'http://wiki.ie.sogou-inc.com/mailproxy?'
        
    params = urllib.urlencode({"uid": "caochengzhen@sogou-inc.com",
                               "fr_name": mailfr_name ,
                               "fr_addr": mailfr_addr ,
                               "title": title ,
                               "mode" : mode ,
                               "maillist" : maillist ,
                               "body" : body })

    if count[1] == 0:
        print 'All Case is OK!'
        pass
    else:
        urllib.urlopen( url , params )  
''' 
def Sendmail(maillist , bodycontent,count,interfacename ):  
    strHtml = ''
    interfacename = str(interfacename)
    strHtml += '<B><p style="font-size:16px">' + interfacename + '测试结果:</p></B>'
    strHtml += '<B><p style="font-size:14px">' + 'Statistic result as follow:</p></B>'
    strHtml += '<B><p style="font-size:14px">' + 'OK:\t' + str(count[0]) + '</p></B>'
    strHtml += '<B><p style="font-size:14px">' + 'NG:\t' + str(count[1]) + '</p></B>'
    strHtml += '<B><p style="font-size:14px">' + 'NT:\t' + str(count[2]) + '</p></B>'
    strHtml += '<table width="1000" border="1">'
    strHtml += '<tr>'
    strHtml += '<td>'+'接口url'+'</td>'
    strHtml += '<td>'+'预期key'+'</td>'
    strHtml += '<td>'+'预期value'+'</td>'
    strHtml += '<td>'+'实际结果'+'</td>'
    strHtml += '</tr>'
    
    for i in range(len(bodycontent)):
        content=''.join(bodycontent[i])
        content = GetStr(content)
        content = str(content)
        #print type(content)
        if i==0 :
            strHtml += '<tr>'
        strHtml += '<td>' + content + '</td>'
        if (i+1) % 4 == 0:
            strHtml += '</tr><tr>'
    strHtml = strHtml[:-4]    
    strHtml += '</table>'
     
        
    mail_from = 'venus@sogou-inc.com'
    mail_to = maillist
    timenow = datetime.datetime.utcnow() + datetime.timedelta(hours=8)#东8区增加8小时
    title = '【'+interfacename + '_接口测试结果】    '+ timenow.strftime( '%Y-%m-%d %H:%M:%S' )
    body = strHtml
    if count[1] == 0:
        print 'All Case is OK!'
        pass
    else:
        sendEmail.SendMail(mail_from, mail_to, title, body)
    