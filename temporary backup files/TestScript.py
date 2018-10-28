#encoding=utf-8
#encoding=utf-8
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from util.ObjectMap import *
from action.pageAction import *

print "start browser..."
open_browser('firefox')
maximize_browser()
print "start browser done..."
print "access 126 mail login page..."
visit_url("http://mail.126.com")
sleep(5)
assert_string_in_pagesource(u"126网易免费邮--你的专业电子邮局")
print "access 126 mail login page done"
switch_to_frame('id',"x-URS-iframe")
#wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID,"x-URS-iframe")))
input_string('xpath',"//input[@name='email']",'xiaxiaoxu1987')
input_string('xpath',"//input[@name='password']",'gloryroad')
press_enter_key()
#pwd.send_keys(Keys.RETURN)
print "user login..."
sleep(5)
switch_to_default_content()
assert_title(u"网易邮箱")
print "login done"
# click('xpath',"//div[text()='通讯录']")
# click('xpath',"//span[text()='新建联系人']")
# input_string('xpath',"//a[@title='编辑详细姓名']/preceding-sibling::div/input",u"徐凤钗")
# input_string('xpath',"//*[@id='iaddress_MAIL_wrap']//input","593152023@qq.com")
# click('xpath',"//span[text()='设为星标联系人']/preceding-sibling::span/b")
# input_string('xpath',"//*[@id='iaddress_TEL_wrap']//dd//input",'18141134488')
# input_string('xpath',"//textarea",'my wife')
# click('xpath',"//span[.='确 定']")
print u"进入首页。。。"
sleep(3)
click('xpath',"//div[.="
              "'首页']")
assert_string_in_pagesource(u"已发送")

print u"进入首页成功"
print "write message..."
click('xpath',"//span[text()='写 信']")
input_string('xpath',"//div[contains(@id,'_mail_emailinput')]/input","367224698@qq.com")
input_string('xpath',"//div[@aria-label='邮件主题输入框，请输入邮件主题']/input",u"测试邮件")
#sleep(30)
#click('xpath',"//div[@title='点击添加附件']/input[@size='1' and @type='file']")

#attachButton = getElement(driver,'xpath',"//div[@title='点击添加附件']/input[@size='1' and @type='file']")
#attachButton=waitVisibilityOfElementLocated('xpath',"//div[@title='点击添加附件']/input[@size='1' and @type='file']")

# 导入支持双击操作的模块
#sleep(10)
#'xpath',"//div[@title='点击添加附件']/input[@size='1' and @type='file']"
#attachButton=getElement(driver,'xpath',"//div[@title='点击添加附件']/input[@size='1' and @type='file']")

attachButton=find_element('xpath',"//div[@title='点击添加附件']/input[@size='1' and @type='file']")
attachButton.send_keys("d:\\test.txt")
#actionChainsClick(attachButton)
# from selenium.webdriver import ActionChains
#

waitVisibilityOfElementLocated('xpath','//span[text()="上传完成"]')
#switch_to_frame('xpath',"//iframe[@tabindex=1]")
waitFrameToBeAvailableAndSwitchToIt('xpath',"//iframe[@tabindex=1]")
input_string('xpath','/html/body',u"发给夏晓旭的一封信")
switch_to_default_content()
#print u"写信完成"
print "write message done"
#点击发送按钮
click('xpath',"//header//span[text()='发送']")
print "start to send email.."
time.sleep(3)
assert_string_in_pagesource(u"发送成功")
print "send emial done"

close_browser()
