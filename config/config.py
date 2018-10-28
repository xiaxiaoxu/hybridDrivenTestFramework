#encoding=utf-8
import os
firefoxDriverPath='c:\\geckodriver'
chromeDriverPath='c:\\chromedriver'
ieDriverPath='d:\\IEDriverServer'

#当前工程所在目录的绝对路径
projectPath=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

#测试数据文件的绝对路径
dataFilePath=projectPath + u"\\testData\\126邮箱创建联系人并发邮件.xlsx"

# 异常图片存放目录
screenPicturesDir = projectPath + "\\errorScreenShots\\"




# 测试数据文件中，测试用例表中部分列对应的数字序号
#都是第一个sheet测试用例的内容
testCase_testCaseName = 1
testCase_frameWorkName = 3
testCase_testStepSheetName = 4
testCase_dataSourceSheetName = 5
testCase_isExecute = 6
testCase_runTime = 7
testCase_testResult = 8

# 用例步骤表中，部分列对应的数字序号
#测试步骤对应的文件中后边的关键字的列号
testStep_testStepDescribe = 1
testStep_keyWords = 2
testStep_locationType = 3
testStep_locatorExpression = 4
testStep_operateValue = 5
testStep_runTime = 6
testStep_testResult = 7
testStep_errorInfo = 8
testStep_errorPic = 9

# 数据源表中，是否执行列对应的数字编号
#联系人的sheet
dataSource_isExecute = 6
dataSource_email = 2
dataSource_runTime = 7
dataSource_result = 8




if __name__ == '__main__':
    print projectPath
    print dataFilePath
    print screenPicturesDir