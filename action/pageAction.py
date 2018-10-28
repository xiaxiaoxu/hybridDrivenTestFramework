#encoding=utf-8
from util.ObjectMap import *
from util.clipboard import *
from util.keyboard import *
from config.config import *
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from config.config import *
from selenium import webdriver
from datetime import datetime
from selenium.webdriver import ActionChains
import os
#定义全局driver变量
driver=None
wait=None
locatorTypeDict = {
        "xpath": By.XPATH,
        "id": By.ID,
        "name": By.NAME,
        "css_selector": By.CSS_SELECTOR,
        "class_name": By.CLASS_NAME,
        "tag_name": By.TAG_NAME,
        "link_text": By.LINK_TEXT,
        "partial_link_text": By.PARTIAL_LINK_TEXT}
#之所以括号中后边写*arg，是因为拼接数据的时候，可能会拼成有参数的，这时候做个容错
#传了参数也不会报错。
def open_browser(browserName,*arg):
    global driver,wait
    try:
        if browserName.lower() == 'firefox':
            driver=webdriver.Firefox(executable_path=firefoxDriverPath)#稍后路径放到config文件中
        elif browserName.lower() == 'chrome':
            #创建chrome浏览器的一个options实例对象
            chrome_options=Options()
            #添加屏蔽ignore-certificate-errors提示信息的设置参数项
            chrome_options.add_experimental_option("excludeSwitches",["ignore-certificate-errors"])
            driver=webdriver.Chrome(executable_path=chromeDriverPath, chrome_options=chrome_options)# 稍后把路径放到config文件中
        else:#剩余的就是ie了
            driver = webdriver.Ie(executable_path=ieDriverPath) # 稍后把路径放到config文件中

        #driver对象确定之后，就可以确定wait了
        wait=WebDriverWait(driver,20)
    except Exception,e:
        raise e

def find_element(locatorType,locatorExpression,*arg):
    global driver
    try:
        element=getElement(driver,locatorType,locatorExpression)
    except Exception,e:
        raise e

def actionChainsClick(element,*arg):
    global driver
    try:
        # 开始模拟鼠标双击操作
        action_chains = ActionChains(driver)
        action_chains.click(element).perform()
        action_chains.click()
    except Exception,e:
        raise e

def visit_url(url,*arg):
    #访问某个网址
    global driver
    try:
        driver.get(url)
    except Exception,e:
        raise e

def close_browser(*arg):
    #关闭浏览器
    global driver
    try:
        driver.quit()
    except Exception,e:
        raise e

def sleep(seconds,*arg):
    try:
        time.sleep(int(seconds))
    except Exception,e:
        raise e

def clear(locatorType,locatorExpression,*arg):
    #找到输入框元素对象，然后清楚输入框内容
    global driver
    try:
        getElement(driver,locatorType,locatorExpression).clear()
    except Exception,e:
        raise e

def input_string(locatorType,locatorExpression,content):
    global driver
    try:
        getElement(driver,locatorType,locatorExpression).send_keys(content)
    except Exception,e:
        raise e

def click(locatorType,locatorExpression,*arg):
    global driver
    try:
        getElement(driver,locatorType,locatorExpression).click()
    except Exception,e:
        raise e

def assert_string_in_pagesource(assertString,*arg):
    #断言页面源码中是否包含要断言的字符串
    global driver
    try:
        assert assertString in driver.page_source,u"%s not found in page source!"%assertString
    except AssertionError,e:
        #raise AssertionError(e)
        raise e
    except Exception,e:
        raise e

def assert_title(titleStr,*arg):
    #断言页面标题是否是给定的字符串
    global driver
    try:
        assert titleStr in driver.title,u"%s not found in title!" %titleStr
    except AssertionError,e:
        raise e
    except Exception,e:
        raise e

def getTitle(*arg):
    #获取页面标题
    global driver
    try:
        return driver.title
    except Exception,e:
        raise e

def getPageSource(*arg):
    #获取页面源码，这个没有用到
    global driver
    try:
        return driver.page_source
    except Exception,e:
        raise e

def switch_to_frame(locatorType,frameLocatorExpression,*arg):
    #查找到frame，并切进frame
    global driver
    try:
        driver.switch_to.frame(getElement(driver,locatorType,frameLocatorExpression))
    except Exception,e:
        print "switch to frame error"
        raise e

def switch_to_default_content(*arg):
    #切换页面从frame到默认窗口中
    global driver
    try:
        driver.switch_to.default_content()
    except Exception,e:
        raise e

def paste_string(pasteContent,*arg):
    #模拟ctrl + v操作
    try:
        Clipboard.setText(pasteContent)#Clipboard类的静态函数
        #等待2秒，防止剪贴板没设置好内容就粘贴，会失败
        time.sleep(2)
        KeyboardKeys.twoKeys('ctrl','v')#Keyboardkeys类的静态函数
    except Exception,e:
        raise e

def press_tab_key(*arg):#没用到
    #模拟tab按键
    try:
        KeyboardKeys.oneKey('tab')
    except Exception,e:
        raise e

def press_enter_key(*arg):
    #模拟enter按键
    try:
        KeyboardKeys.oneKey('enter')
    except Exception,e:
        raise e

def maximize_browser():
    #窗口最大化
    global driver
    try:
        pass
        #driver.maximize_window()
    except Exception,e:
        raise e

def waitFrameToBeAvailableAndSwitchToIt(locatorType,locatorExpression,*arg):
    #检查frmae是否存在，存在则切进frame中
    global wait
    global locatorTypeDict
    try:
        element = wait.until(
            EC.frame_to_be_available_and_switch_to_it((locatorTypeDict[locatorType.lower()], locatorExpression)))
        return element
    except Exception, e:
        raise e


def waitVisibilityOfElementLocated(locatorType,locatorExpression,*arg):
    #等待页面元素出现在DOM中并且可见，存在则返回该元素
    global wait
    global locatorTypeDict
    try:
        element=wait.until(EC.visibility_of_element_located((locatorTypeDict[locatorType.lower()],locatorExpression)))
        return element
    except Exception,e:
        raise e


def captureScreen(*arg):
    global driver
    currentDate=time.strftime("%Y-%m-%d")
    currentTime=datetime.now().strftime('%H-%M-%S-%f')#时分秒毫秒
    dirName=os.path.join(screenPicturesDir,currentDate)
    if not os.path.exists(dirName):
        os.makedirs(dirName)
    #print "dirName:",dirName
    screenShotNameAndPath="%s\\%s.png"%(dirName,currentTime)
    #print "screenPicturesDir:",screenPicturesDir
    try:
        driver.get_screenshot_as_file(screenShotNameAndPath.replace('\\',r'\\'))
    except Exception,e:
        raise e
    else:
        return screenShotNameAndPath

if __name__ == '__main__':
    from selenium import webdriver
    open_browser('firefox')
    visit_url('http:\\126.com')
    switch_to_frame("id", "x-URS-iframe")
    input_string('xpath', '//input[@name="email"]', u'xiaxiaoxu1987')
    input_string('xpath', '//input[@name="password"]', u'gloryroad')
    click('id', 'dologin')
    sleep(u'5')
    switch_to_default_content()
    click('xpath','//div[text()="通讯录"]')
    waitVisibilityOfElementLocated('xpath', '//span[text()="新建联系人"]')
    click('xpath', '//span[text()="新建联系人"]')
    waitVisibilityOfElementLocated('xpath', '//a[@title="编辑详细姓名"]/preceding-sibling::div/input')



