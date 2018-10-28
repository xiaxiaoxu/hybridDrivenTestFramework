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
from config.config import *
from util.log import *

def dataDrivenFunction(excelObj,caseStepSheet,dataDrivenSourceSheet):
    logging.info(u"步骤sheet：u'%s'"%caseStepSheet)
    logging.info(u"数据源sheet: u'%s'"%dataDrivenSourceSheet)
    excelObj.get_sheet_by_name(dataDrivenSourceSheet)
    requiredContactNum = 0
    successfullyAddedContactNum = 0
    isExecuteCol = excelObj.get_single_col(5)
    emailCol = excelObj.get_single_col(1)
    excelObj.get_sheet_by_name(caseStepSheet)
    stepRows = excelObj.get_all_rows()
    totalStepNum = len(stepRows) - 1
    #logging.info("totalStepNum: %s"%totalStepNum)
    for idx1, cell in enumerate(isExecuteCol[1:]):
        if cell.value == 'y':
            successStepNum = 0
            requiredContactNum += 1  # 记录一次需要添加的联系人
            for row in stepRows[1:]:
                caseStepDescripion = row[testStep_testStepDescribe-1].value
                keyWord = row[testStep_keyWords-1].value
                locatorType = row[testStep_locationType-1].value
                locatorExpression = row[testStep_locatorExpression-1].value
                operateValue = str(row[testStep_operateValue-1].value) if isinstance(row[testStep_operateValue-1].value, long) else row[testStep_operateValue-1].value
                if operateValue and operateValue.isalpha():
                    excelObj.get_sheet_by_name(dataDrivenSourceSheet)
                    operateValue = excelObj.get_cell_content_by_coordinate(operateValue + str(idx1 + 2))
                    logging.info("operateValue:u'%s'"%operateValue)
                if (locatorType and locatorExpression):
                    command = "%s('%s','%s',u'%s')" % (
                        keyWord, locatorType, locatorExpression.replace("'", "\""),
                        operateValue) if operateValue else "%s('%s','%s')" % (
                        keyWord, locatorType, locatorExpression.replace("'", "\""))
                elif operateValue:
                    command = "%s(u'%s')" % (keyWord, operateValue)
                logging.info(u'%s'%caseStepDescripion)
                try:
                    if operateValue != u"否":
                        eval(command)

                except Exception, e:  # 某个步骤执行失败
                    logging.info(u"执行步骤-u'%s'-失败" % caseStepDescripion)
                else:  # 执行步骤成功
                    logging.info(u"执行步骤-u'%s'-成功" % caseStepDescripion)
                    successStepNum += 1
            if totalStepNum == successStepNum:#说明一条联系人添加成功
                #写结果
                logging.info(u"添加联系人-u'%s'-成功"%emailCol[idx1+2].value)
                successfullyAddedContactNum +=1#记录一次添加一条联系人成功
                writeResult(excelObj,dataDrivenSourceSheet,"dataSheet",idx1+2,u"pass",)
            else:#说明添加联系人没成功，写fail
                logging.info(u"添加联系人-u'%s'-失败" % emailCol[idx1 + 2].value)
                writeResult(excelObj, dataDrivenSourceSheet, "dataSheet", idx1 + 2, u"fail", )

        else:  # 忽略
            writeResult(excelObj, dataDrivenSourceSheet, "dataSheet", idx1 + 2, u"忽略", )

    if requiredContactNum == successfullyAddedContactNum:#说明需要添加的联系人和添加成功的联系人数量一样，那这条用例就成功了
        return 1
    else:#有联系人添加失败了，用例执行失败，
        return 0

