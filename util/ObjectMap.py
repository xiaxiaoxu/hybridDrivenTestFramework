#encoding=utf-8
from selenium.webdriver.support.ui import WebDriverWait


#获取单个元素对象
def getElement(driver,locatorType,locatorExpression):
    wait = WebDriverWait(driver, 20)
    try:
        element=wait.until(lambda x:x.find_element(by=locatorType,value=locatorExpression))
        return element
    except Exception,e:
        raise e

#获取多个相同页面元素对象，以list返回
def getElements(driver,locatorType,locatorExpression):
    wait=WebDriverWait(driver,20)
    try:
        elements=wait.until(lambda x:x.find_elements(by=locatorType,value=locatorExpression))
        return elements
    except Exception,e:
        raise e

if __name__=='__main__':
    #测试代码
    from selenium import webdriver
    driver=webdriver.Firefox(executable_path='c:\\geckodriver')
    driver.get("http://www.baidu.com")
    searchBox=getElement(driver,'id','kw')
    print searchBox.tag_name
    aList=getElements(driver,'tag name','a')
    print len(aList)
    driver.quit()
