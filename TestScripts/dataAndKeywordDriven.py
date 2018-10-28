#encoding=utf-8
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from util.ObjectMap import *
from action.pageAction import *
from util.ParseExcel import *
import traceback
from TestScripts.keyWord import *
from TestScripts.dataDriven import *
from util.writeResult import *
from util.log import *
from config.config import *

pe=parseExcel(dataFilePath)
caseSheetName=u"测试用例"#测试用例sheet名
pe.get_sheet_by_name(caseSheetName)
logging.info(u"当前sheet：u'%s'" %pe.get_default_sheet())
caseRows=pe.get_all_rows()
for idx,row in enumerate(caseRows[1:]):
    frameWorkType=row[testCase_frameWorkName-1].value
    caseStepSheet=row[testCase_testStepSheetName-1].value
    dataDrivenSourceSheet=row[testCase_dataSourceSheetName-1].value
    ifExecute=row[testCase_isExecute-1].value
    if ifExecute.lower() == 'y':
        if frameWorkType ==u"关键字":
            logging.info(u"#####执行关键字驱动框架#####")
            result =keyWordFunction(pe,caseStepSheet)
            logging.info("result:u'%s'"%result)
            if result:#此条用例执行成功
                writeResult(pe,caseSheetName,"testCase",idx+2,u"成功")
            else:
                writeResult(pe,caseSheetName,"testCase", idx + 2, u"失败")

        elif frameWorkType == u"数据":
            logging.info(u"#####执行数据驱动框架#####")
            result=dataDrivenFunction(pe,caseStepSheet,dataDrivenSourceSheet)
            if result:#说明需要添加的联系人和添加成功的联系人数量一样，那这条用例就成功了
                writeResult(pe, caseSheetName,"dataSheet", idx + 2, u"成功")
            else:#有联系人添加失败了，用例执行失败，
                writeResult(pe, caseSheetName,"dataSheet", idx + 2, u"失败")

    else:#不是y
        writeResult(pe,caseSheetName, "testCase", idx + 2, u"忽略")
























