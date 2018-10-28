#encoding=utf-8
#author-夏晓旭
import logging
import logging.config
from config.config import *

#读取日志的配置文件
logging.config.fileConfig(projectPath+"\\config\\Logger.conf")

#选择一个日志格式
logger=logging.getLogger("example02")

def error(message):
    #打印debug级别的信息
    logger.error(message)

def info(message):
    #打印info级别的信息
    logger.info(message)

def warning(message):
    #打印warning级别的信息
    logger.warning(message)

if __name__=="__main__":
    info("hi")
    print "config file path:",projectPath+"\\config\\Logger.conf"
    error("world!")
    warning("gloryroad!")