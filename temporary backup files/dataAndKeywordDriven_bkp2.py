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

pe=parseExcel(dataFilePath)
pe.get_sheet_by_name(u"测试用例")
print u"当前sheet：",pe.get_default_sheet()
caseRows=pe.get_all_rows()
for idx,row in enumerate(caseRows[1:]):
    frameWorkType=row[2].value
    caseStepSheet=row[3].value
    dataDrivenSourceSheet=row[4].value
    ifExecute=row[5].value
    if ifExecute.lower() == 'y':
        # print u"用例名称：", row[0].value
        # print u"用例描述：", row[1].value
        # print u"调用框架类型：", frameWorkType
        # print u"用例步骤sheet名：", caseStepSheet
        # print u"数据驱动的数据源sheet名：", dataDrivenSourceSheet
        # print u"是否执行：", ifExecute
        if frameWorkType ==u"关键字":
            print "-----执行关键字驱动框架-----"
            pe.get_sheet_by_name(caseStepSheet)
            stepRows=pe.get_all_rows()
            totalStepNum=len(stepRows)-1
            print "totalStepNum:",totalStepNum
            successStepNum=0
            for idx1,row1 in enumerate(stepRows[1:]):
                caseStepDescripion=row1[0].value
                keyWord=row1[1].value
                locatorType=row1[2].value
                locatorExpression=row1[3].value
                operateValue=str(row1[4].value) if isinstance(row1[4].value,long) else row1[4].value
                # print u"测试步骤描述：",caseStepDescripion
                # print u"关键字：",keyWord
                # print u"操作元素的定位方式：",locatorType
                # print u"操作元素的定位表达式：",locatorExpression
                # print u"操作值：",operateValue
                if (locatorType  and locatorExpression):
                    command="%s('%s','%s',u'%s')"%(keyWord,locatorType,locatorExpression.replace("'","\""),operateValue) if operateValue else "%s('%s','%s')"%(keyWord,locatorType,locatorExpression.replace("'","\""))
                elif operateValue :
                    command ="%s(u'%s')"%(keyWord,operateValue)
                else:
                    command="%s()"%keyWord
                print caseStepDescripion
                print "command:",command
                try:
                    eval(command)
                except Exception,e:
                    pe.write_cell_content(idx1 + 2, 7, u"fail",color='red')
                    pe.write_cell_current_time(idx1 + 2, 6)
                    pe.write_cell_content(idx1+2,8,traceback.format_exc())
                    pe.write_cell_content(idx1+2,9,captureScreen())
                    #写入错误截图需要封装函数，返回图片地址

                else:#没有报错
                    successStepNum+=1
                    print "successStepNum:",successStepNum
                    pe.write_cell_content(idx1+2,7,u"pass",color="green")
                    pe.write_cell_current_time(idx1+2,6)
                    pe.write_cell_content(idx1 + 2, 8, "")
                    pe.write_cell_content(idx1 + 2, 9, "")
            if totalStepNum == successStepNum:#此条用例执行成功
                pe.get_sheet_by_name(u"测试用例")  # 把默认sheet获取到测试用例sheet上
                pe.write_cell_content(idx + 2, 8, u"成功",color="green")  # 在结果列中写入忽略
                pe.write_cell_current_time(idx+2,7)#写入执行时间
            else:
                pe.get_sheet_by_name(u"测试用例")  # 把默认sheet获取到测试用例sheet上
                pe.write_cell_content(idx + 2, 8, u"失败",color="red")  # 在结果列中写入忽略
                pe.write_cell_current_time(idx + 2, 7)  # 写入执行时间


        elif frameWorkType == u"数据":
            print "-----执行数据驱动框架-----"
            print u"步骤sheet：",caseStepSheet
            print u"数据源sheet",dataDrivenSourceSheet
            pe.get_sheet_by_name(dataDrivenSourceSheet)
            requiredContactNum = 0
            successfullyAddedContactNum=0
            isExecuteCol=pe.get_single_col(5)
            emailCol=pe.get_single_col(1)
            print "isExecuteCol:",isExecuteCol
            print "emailCol:",emailCol
            pe.get_sheet_by_name(caseStepSheet)
            stepRows = pe.get_all_rows()
            totalStepNum = len(stepRows) - 1

            print "totalStepNum:", totalStepNum
            for idx1,cell in enumerate(isExecuteCol[1:]):
                if cell.value =='y':
                    successStepNum = 0
                    requiredContactNum+=1#记录一次需要添加的联系人
                    for row in stepRows[1:]:
                        caseStepDescripion = row[0].value
                        keyWord = row[1].value
                        locatorType = row[2].value
                        locatorExpression = row[3].value
                        operateValue = str(row[4].value) if isinstance(row[4].value, long) else row[4].value
                        # print "caseStepDescripion:",caseStepDescripion
                        # print "keyWord:",keyWord
                        # print "locatorType:",locatorType
                        # print "locatorExpression:",locatorExpression
                        # print "operateValue:",operateValue
                        if operateValue and operateValue.isalpha():
                            pe.get_sheet_by_name(dataDrivenSourceSheet)
                            operateValue=pe.get_cell_content_by_coordinate(operateValue+str(idx1+2))
                            #print "operateValue:",operateValue
                        if (locatorType and locatorExpression):
                            command = "%s('%s','%s',u'%s')" % (
                            keyWord, locatorType, locatorExpression.replace("'", "\""),
                            operateValue) if operateValue else "%s('%s','%s')" % (
                            keyWord, locatorType, locatorExpression.replace("'", "\""))
                        elif operateValue:
                            command = "%s(u'%s')" % (keyWord, operateValue)
                        print caseStepDescripion
                        print "command:",command
                        try:
                            if operateValue != u"否":
                                eval(command)

                        except Exception,e:#某个步骤执行失败
                            print u"执行步骤-%s-失败"%caseStepDescripion
                        else:#执行步骤成功
                            print u"执行步骤-%s-成功" % caseStepDescripion
                            successStepNum +=1
                            print "successStepNum:", successStepNum

                    #print "successStepNum:",successStepNum
                    if totalStepNum == successStepNum:#说明一条联系人添加成功
                        #写结果
                        print u"添加联系人-%s-成功"%emailCol[idx1+2].value
                        successfullyAddedContactNum +=1#记录一次添加一条联系人成功
                        pe.get_sheet_by_name(dataDrivenSourceSheet)
                        pe.write_cell_current_time(idx1+2,7)

                        pe.write_cell_content(idx1+2,8,u"pass",color='green')
                    else:#说明添加联系人没成功，写fail
                        print u"添加联系人-%s-失败" % emailCol[idx1 + 2].value
                        pe.get_sheet_by_name(dataDrivenSourceSheet)
                        pe.write_cell_current_time(idx1 + 2, 7)
                        pe.write_cell_content(idx1 + 2, 8, u"fail",color="red")


                else: #忽略
                    pe.get_sheet_by_name(dataDrivenSourceSheet)
                    pe.write_cell_content(idx1+2,7,'')
                    pe.write_cell_content(idx1+2,8,u"忽略")

            if requiredContactNum == successfullyAddedContactNum:#说明需要添加的联系人和添加成功的联系人数量一样，那这条用例就成功了
                pe.get_sheet_by_name(u"测试用例")
                pe.write_cell_content(idx+2,8,u"成功",color="green")
                pe.write_cell_current_time(idx+2,7)

            else:#有联系人添加失败了，用例执行失败，
                pe.get_sheet_by_name(u"测试用例")
                pe.write_cell_content(idx + 2, 8, u"失败",color="red")
                pe.write_cell_current_time(idx + 2, 7)

    else:#不是y
        pe.get_sheet_by_name(u"测试用例")#把默认sheet获取到测试用例sheet上
        pe.write_cell_content(idx+2,8,u"忽略")#在结果列中写入忽略
        pe.write_cell_content(idx+2,7,"")#清空时间列的值























