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

pe=parseExcel(dataFilePath)
caseSheetName=u"测试用例"#测试用例sheet名
pe.get_sheet_by_name(caseSheetName)
print u"当前sheet：",pe.get_default_sheet()
caseRows=pe.get_all_rows()
#print "caseRows:",caseRows
for idx,row in enumerate(caseRows[1:]):
    frameWorkType=row[testCase_frameWorkName-1].value
    caseStepSheet=row[testCase_testStepSheetName-1].value
    dataDrivenSourceSheet=row[testCase_dataSourceSheetName-1].value
    ifExecute=row[testCase_isExecute-1].value
    #print "ifExecute",ifExecute
    if ifExecute.lower() == 'y':
        # print u"用例名称：", row[0].value
        # print u"用例描述：", row[1].value
        # print u"调用框架类型：", frameWorkType
        # print u"用例步骤sheet名：", caseStepSheet
        # print u"数据驱动的数据源sheet名：", dataDrivenSourceSheet
        # print u"是否执行：", ifExecute
        if frameWorkType ==u"关键字":
            print "#####执行关键字驱动框架#####"
            result =keyWordFunction(pe,caseStepSheet)
            print "result:",result
            if result:#此条用例执行成功
                writeResult(pe,caseSheetName,"testCase",idx+2,u"成功")
                # global caseSheetName
                # global caseStepSheet
                # global dataDrivenSourceSheet
                #writeResult(excelObj, sheetType, rowNo, result, errorInfo=None, captureScreenPath=None):
                # pe.get_sheet_by_name(caseSheetName)  # 把默认sheet获取到测试用例sheet上
                # pe.write_cell_content(idx + 2, testCase_testResult, u"成功",color="green")  # 在结果列中写入忽略
                # pe.write_cell_current_time(idx+2,testCase_runTime)#写入执行时间
            else:
                writeResult(pe,caseSheetName,"testCase", idx + 2, u"失败")
                # pe.get_sheet_by_name(caseSheetName)  # 把默认sheet获取到测试用例sheet上
                # pe.write_cell_content(idx + 2, testCase_testResult, u"失败",color="red")  # 在结果列中写入忽略
                # pe.write_cell_current_time(idx + 2, testCase_runTime)  # 写入执行时间

        elif frameWorkType == u"数据":
            print "#####执行数据驱动框架#####"
            result=dataDrivenFunction(pe,caseStepSheet,dataDrivenSourceSheet)
            if result:#说明需要添加的联系人和添加成功的联系人数量一样，那这条用例就成功了
                writeResult(pe, caseSheetName,"dataSheet", idx + 2, u"成功")
                # writeResult(excelObj,sheetName,sheetType,rowNo,result,errorInfo=None,captureScreenPath=None):
                # pe.get_sheet_by_name(caseSheetName)
                # pe.write_cell_content(idx+2,testCase_testResult,u"成功",color="green")
                # pe.write_cell_current_time(idx+2,testCase_runTime)

            else:#有联系人添加失败了，用例执行失败，
                writeResult(pe, caseSheetName,"dataSheet", idx + 2, u"失败")
                # pe.get_sheet_by_name(caseSheetName)
                # pe.write_cell_content(idx + 2, testCase_testResult, u"失败",color="red")
                # pe.write_cell_current_time(idx + 2, testCase_runTime)

    else:#不是y
        writeResult(pe,caseSheetName, "testCase", idx + 2, u"忽略")
        # pe.get_sheet_by_name(caseSheetName)#把默认sheet获取到测试用例sheet上
        # pe.write_cell_content(idx+2,testCase_testResult,u"忽略")#在结果列中写入忽略
        # pe.write_cell_content(idx+2,testCase_runTime,"")#清空时间列的值
























