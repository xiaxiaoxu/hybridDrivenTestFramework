#encoding=utf-8
from util.ParseExcel import *
from config.config import *
import traceback
from util.ParseExcel import *
from config.config import *


def writeResult(excelObj,sheetName,sheetType,rowNo,result,errorInfo=None,captureScreenPath=None):
    colorOfResult = {"pass": "green", "fail": "red", u"成功": "green",u"失败": "red", u"忽略": None, "": None}
    print "当前所在sheet:",excelObj.get_default_sheet()
    print excelObj, sheetName, sheetType, rowNo, result, errorInfo,captureScreenPath
    sheetTypeDict={
        "testCase":[testCase_runTime, testCase_testResult],#测试用例sheet名，时间列和结果列的列号
        "caseStep":[testStep_runTime, testStep_testResult,testStep_errorInfo,testStep_errorPic],#关键字驱动测试步骤sheet名，时间列和结果列号
        "dataSheet":[dataSource_runTime, dataSource_result]}#数据驱动的数据源sheet名，时间、结果列号
    try:
        excelObj.get_sheet_by_name(sheetName)
        if sheetType != "caseStep":#sheetType不是关键字步骤sheet，说明是用例sheet或者数据源sheet
            print "sheetType1:",sheetType
            print "sheetTypeDict[sheetType][1]:",sheetTypeDict[sheetType][1]
            #直接把结果写上去
            excelObj.write_cell_content(rowNo, sheetTypeDict[sheetType][1], result, color=colorOfResult[result])
            if result != "" and result != u"忽略":
                print "result-1:",result
                excelObj.write_cell_current_time(rowNo,sheetTypeDict[sheetType][0])
            else:#联系人被忽略，时间列清空
                print "result-2:", result
                excelObj.write_cell_content(rowNo,sheetTypeDict[sheetType][0],"",color=colorOfResult[result])
        else :#sheetType是步骤sheet
            print "sheetType2:", sheetType
            print "result:",result
            #先把结果列和时间列给写上
            excelObj.write_cell_content(rowNo, sheetTypeDict[sheetType][1], result, color=colorOfResult[result])
            excelObj.write_cell_current_time(rowNo, sheetTypeDict[sheetType][0])
            if errorInfo and captureScreenPath:#有错误信息和截图信息，把错误信息和截图信息写进去
                print "sheetTypeDict[sheetType][2]:",sheetTypeDict[sheetType][2]
                print "sheetTypeDict[sheetType][3]:",sheetTypeDict[sheetType][3]
                print "errorInfo:",errorInfo
                print "captureScreenPath:",captureScreenPath
                excelObj.write_cell_content(rowNo, sheetTypeDict[sheetType][2], errorInfo)
                excelObj.write_cell_content(rowNo, sheetTypeDict[sheetType][3], captureScreenPath)
            else:#没有截图信息和错误信息，把截图信息和错误信息清空
                print "sheetTypeDict[sheetType][3]",sheetTypeDict[sheetType][3]
                print "sheetTypeDict[sheetType][2]",sheetTypeDict[sheetType][2]
                excelObj.write_cell_content(rowNo, sheetTypeDict[sheetType][3],"")#是成功的，没有错误信息
                excelObj.write_cell_content(rowNo, sheetTypeDict[sheetType][2], "")#是成功的，没有截图信息

    except Exception,e:
        raise e
