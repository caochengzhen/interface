#!/usr/bin/python  
#encoding:utf-8 
#****************************************************************  
# main.py
# Description: 入口文件，实际项目中只需要修改 ExcelPath和SheetName即可
# Author: caochengzhen@sogou-inc.com
#****************************************************************  

from server_case import *
import server_case
  
#设置读取文件和sheet页
ExcelPath=os.getcwd()+'/TestCaseDir/yuedu_Testcase.xlsx'
SheetName='ALL'#需要执行的sheet页名称，如果想执行所有sheet页，必须为大写的ALL
server_case.excelobj=create_excel(ExcelPath,SheetName)

#跑case，如果有错误发出邮件通知
if SheetName.upper() == 'ALL':
    #执行所有sheet页的用例
    Sheetnames = []
    Sheetnames = server_case.excelobj.get_all_sheetname()
    #print Sheetnames
    for SheetName in Sheetnames:
        mailbody,count = run(SheetName)  
        Sendmail('caochengzhen@sogou-inc.com', mailbody,count, SheetName)
else:    
    #执行指定sheet页的用例
    mailbody,count = run(SheetName)  
    Sendmail('caochengzhen@sogou-inc.com', mailbody,count, SheetName)  
    
#跑完case后，关闭excel
#server_case.excelobj.close()
