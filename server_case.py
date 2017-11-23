#!/usr/bin/python  
#encoding:utf-8 

#****************************************************************    
# Description: case处理区 
#****************************************************************  
  
from TestFrame import *  
import json  
 
  
def run(suiteid):  
    print '【'+suiteid+'】' + ' Test Begin,please waiting...'  
    global expectXmldir, realXmldir,reqmethod,checkitem,requesturi,mailbody

    #checkitem=excelobj.read_data(suiteid, 5, 2)
    requesturi=excelobj.read_data(suiteid, 1, 2)
      
      
    excelobj.del_testrecord(suiteid)  #清除历史测试数据  
    casecount=excelobj.get_ncase(suiteid) #获取case个数  
    caseinfolist=get_caseinfo(excelobj, suiteid) #获取Case基本信息  
    #print caseinfolist
    mailbody=[]      
    #遍历执行case  
    for caseid in range(0, casecount):
        #print caseid,casecount  

        #检查是否执行该Case  
        if excelobj.read_data(suiteid,excelobj.casebegin+caseid, 2)=='N':  
            excelobj.write_data(suiteid, excelobj.casebegin+caseid, 15, 'NT', NT_COLOR)
            ret1='NT'
            #写测试结果  
            write_result(excelobj, suiteid, caseid, excelobj.resultCol, ret1)  
            continue #当前Case结束，继续执行下一个Case  
        #获取预期结果的key字段
        checkitem=excelobj.read_data(suiteid, excelobj.casebegin+caseid, excelobj.CheckKey)
        #print checkitem   
        #拼接接口url
        if excelobj.read_data(suiteid,excelobj.casebegin+caseid, 2)=='E':
            sArge=excelobj.read_data(suiteid,excelobj.casebegin+caseid, 3) #如果是异常用例，不采用拼接参数，直接取第一个参数为url参数
            sArge = str(sArge)
        else:
            sArge=get_input(excelobj, suiteid, caseid, caseinfolist)#sInput得到的是除了接口以为的参数拼接串
            sArge = str(sArge)
        
        #拿到url的参数后，先做是否关联的判断
        sInput=requesturi+sArge 
        if sArge.startswith('correlerr'):
            real_value='关联失败'
            ret1='NG'
            print real_value
            ret1=check_result(excelobj, suiteid, caseid,real_value, excelobj.CheckVaule)
            if ret1 == 'NG':#对于不通过的测试结果拼接mailbody，便于发邮件
                expkey=excelobj.read_data(suiteid, excelobj.casebegin+caseid, excelobj.CheckKey)
                expvalue=excelobj.read_data(suiteid, excelobj.casebegin+caseid, excelobj.CheckVaule)
                mailbody.append(sInput)
                mailbody.append(expkey)
                mailbody.append(expvalue)
                mailbody.append(real_value) 
                    
            #写测试结果  
            write_result(excelobj, suiteid, caseid, excelobj.resultCol, ret1)
            continue            
            
           
        ResString=HTTPInvoke(sInput,requesturi)     #执行调用  
        print ResString
        #if ResString#处理jsonp的情况

        #获取返回码并比较 
        #print ResString,checkitem 
        #如果服务器返回码不是200，直接报错。
        if ResString.startswith('[Response_Code_err]'):
            real_value=ResString
            #print real_value
            ret1=check_result(excelobj, suiteid, caseid,real_value, excelobj.CheckVaule)
            if ret1 == 'NG':#对于不通过的测试结果拼接mailbody，便于发邮件
                expkey=excelobj.read_data(suiteid, excelobj.casebegin+caseid, excelobj.CheckKey)
                expvalue=excelobj.read_data(suiteid, excelobj.casebegin+caseid, excelobj.CheckVaule)
                mailbody.append(sInput)
                mailbody.append(expkey)
                mailbody.append(expvalue)
                mailbody.append(real_value) 
                    
            #写测试结果  
            write_result(excelobj, suiteid, caseid, excelobj.resultCol, ret1)
            continue
        #判断是使用xml还是json解析。
        if ResString.find('xml version=')>0:
            #判断并统一转换为utf8编码，不然会报“multi-byte encodings are not supported”错误 
            try:
                ResString=ResString.decode('gbk').encode('utf-8')
                ResString=ResString.replace('gbk', 'utf-8')
                ResString=ResString.replace('GBK', 'utf-8')
            except:
                pass
                
            try:
                ResString=ResString.decode('gb2312').encode('utf-8')
                ResString=ResString.replace('gb2312', 'utf-8')
                ResString=ResString.replace('GB2312', 'utf-8')
            except:
                pass
            try:
                real_value=et.fromstring(ResString).find(checkitem).text 
                ret1=check_result(excelobj, suiteid, caseid,real_value, excelobj.CheckVaule)
            except:
                print sInput+"返回不是标准的XML！"
                real_value= sInput+"返回不是标准的XML！"
                ret1=check_result(excelobj, suiteid, caseid,real_value, excelobj.CheckVaule)
        elif ResString.startswith('{'):
            try:
                ResString=ToUnicode(ResString)
                hjson = json.loads(ResString)
                expstr=excelobj.read_data(suiteid,excelobj.casebegin+caseid, 13)
                if str(checkitem).startswith('['):
                    #判断检查点是否需要检查多级json结构及是对应的值是否符合预期。示例：['data'][5]['imageId']会查找json结构中date下第4个list中imageId的value
#                     querystr = ''
#                     argslist = str(checkitem).split('>>')
#                     for i in range(len(argslist)):
#                         querystr += '["'+argslist[i]+'"]'
                    #print querystr
                    try:
                        checkitem = checkitem.replace("‘","'")
                        checkitem = checkitem.replace("’","'")
                        hjson = json.loads(ResString)
                        precmd='hjson'+checkitem                
                        real_value=eval(precmd)
                        ret1=check_result(excelobj, suiteid, caseid,real_value, excelobj.CheckVaule)
                    except:
                        real_value=checkitem+'解析失败，请检查层级及语法！'
                        ret1=check_result(excelobj, suiteid, caseid,real_value, excelobj.CheckVaule)                        
                else :
                    ResString=ToUnicode(ResString)
                    #print ResString
                    if ResString.find(checkitem)>0:
                        pattern  =  re.compile(r'%s(.*?),' % checkitem)
                        res = pattern.findall(ResString)
                        if res==[]:#如果要匹配的字段在最后，则没有逗号，只能匹配大括号
                            pattern  =  re.compile(r'%s(.*?)}' % checkitem)
                            res = pattern.findall(ResString)                            
                        #print res[0]
                    else:
                        print 'key:'+checkitem+'在response中不存在！'
                        real_value='key:'+checkitem+'在response中不存在！'
                        ret1=check_result(excelobj, suiteid, caseid,real_value, excelobj.CheckVaule)
                        if ret1 == 'NG':#对于不通过的测试结果拼接mailbody，便于发邮件
                            expkey=excelobj.read_data(suiteid, excelobj.casebegin+caseid, excelobj.CheckKey)
                            expvalue=excelobj.read_data(suiteid, excelobj.casebegin+caseid, excelobj.CheckVaule)
                            mailbody.append(sInput)
                            mailbody.append(expkey)
                            mailbody.append(expvalue)
                            mailbody.append(real_value) 
                        #写测试结果  
                        write_result(excelobj, suiteid, caseid, excelobj.resultCol, ret1)
                        continue
                    
                    real_value=res[0].replace('"','')
                    #print real_value
                    real_value=real_value[1:]
                    ret1=check_result(excelobj, suiteid, caseid,real_value, excelobj.CheckVaule)  
            except:
                print sInput+"返回不是标准的json！"
                real_value=sInput+"返回不是标准的json！"
                ret1=check_result(excelobj, suiteid, caseid,real_value, excelobj.CheckVaule) 
        else:#非json非xml文件
            expstr=excelobj.read_data(suiteid,excelobj.casebegin+caseid, 14)
            if ResString.find(expstr)>0:
                real_value=expstr
                ret1=check_result(excelobj, suiteid, caseid,real_value, excelobj.CheckVaule)
            else:
                print sInput+'中不存在'+expstr
                real_value=sInput+'中不存在'+expstr
                ret1=check_result(excelobj, suiteid, caseid,real_value, excelobj.CheckVaule)
        if ret1 == 'NG':#对于不通过的测试结果拼接mailbody，便于发邮件
            expkey=excelobj.read_data(suiteid, excelobj.casebegin+caseid, excelobj.CheckKey)
            expvalue=excelobj.read_data(suiteid, excelobj.casebegin+caseid, excelobj.CheckVaule)
            mailbody.append(sInput)
            mailbody.append(expkey)
            mailbody.append(expvalue)
            mailbody.append(real_value)                
        #写测试结果  
        write_result(excelobj, suiteid, caseid, excelobj.resultCol, ret1)
    print '【'+suiteid+'】' + ' Test End!' + '\n' +'**********************************************************************************************************'
    count = statisticresult(excelobj)
    return mailbody,count 