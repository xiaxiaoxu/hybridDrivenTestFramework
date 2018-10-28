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
from util.writeResult import *

def keyWordFunction(excelObj,caseStepSheet):
    excelObj.get_sheet_by_name(caseStepSheet)
    stepRows = excelObj.get_all_rows()
    totalStepNum = len(stepRows) - 1
    #print "totalStepNum:", totalStepNum
    successStepNum = 0
    for idx1, row1 in enumerate(stepRows[1:]):
        caseStepDescripion = row1[testStep_testStepDescribe-1].value
        keyWord = row1[testStep_keyWords-1].value
        locatorType = row1[testStep_locationType-1].value
        locatorExpression = row1[testStep_locatorExpression-1].value
        operateValue = str(row1[testStep_operateValue-1].value) if isinstance(row1[testStep_operateValue-1].value, long) else row1[testStep_operateValue-1].value
        if (locatorType and locatorExpression):
            command = "%s('%s','%s',u'%s')" % (keyWord, locatorType, locatorExpression.replace("'", "\""),
                                               operateValue) if operateValue else "%s('%s','%s')" % (
            keyWord, locatorType, locatorExpression.replace("'", "\""))
        elif operateValue:
            command = "%s(u'%s')" % (keyWord, operateValue)
        else:
            command = "%s()" % keyWord
        print caseStepDescripion
        #print "command:", command
        try:
            eval(command)
        except Exception, e:
            errorInfo=traceback.format_exc()
            captureScreenPath=captureScreen()
            writeResult(excelObj, caseStepSheet,"caseStep", idx1 + 2, u"fail",errorInfo=errorInfo,captureScreenPath=captureScreenPath)
            # excelObj.write_cell_content(idx1 + 2, testStep_testResult, u"fail", color='red')
            # excelObj.write_cell_current_time(idx1 + 2, testStep_runTime)
            # excelObj.write_cell_content(idx1 + 2, testStep_errorInfo, traceback.format_exc())
            # excelObj.write_cell_content(idx1 + 2, testStep_errorPic, captureScreen())
            # 写入错误截图需要封装函数，返回图片地址

        else:  # 没有报错
            successStepNum += 1
            writeResult(excelObj, caseStepSheet,"caseStep", idx1 + 2, u"pass")
            #print "successStepNum:", successStepNum
            # excelObj.write_cell_content(idx1 + 2, testStep_testResult, u"pass", color="green")
            # excelObj.write_cell_current_time(idx1 + 2, testStep_runTime)
            # excelObj.write_cell_content(idx1 + 2, testStep_errorInfo, "")
            # excelObj.write_cell_content(idx1 + 2, testStep_errorPic, "")
    if totalStepNum == successStepNum:  # 此条用例执行成功
        return 1#代表成功
    else:
        return 0#代表失败


