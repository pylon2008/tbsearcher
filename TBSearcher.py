# coding=GBK
import logging, datetime, socket, traceback, os
from win32com.shell import shell, shellcon
from xpylon.xethernet.IEProxy import *
from xpylon.xethernet.IEExplorer import *
from xpylon.xethernet.NetManager import *
from xpylon.xutil.Activation import *

def initLogging():
    LOG_FILENAME="TBSearcher.log"
    logging.basicConfig(filename=LOG_FILENAME, level=logging.DEBUG, format='%(asctime)s - %(levelname)s: %(message)s')
    curTime = datetime.datetime.now
    strTime = str(curTime)
    logging.debug("===============================================Begin Log===============================================")
    logging.debug( socket.gethostname() )
 
def test_del_cookie():
    clearIECookie()

def search_baobei():
    logging.debug('do nothing')
    test_del_cookie()

def tbsearch_2897106_dowork():
    initLogging()
    netManger = None
    
    # init net manager
    try:
        netManger = NetManager()
        config = ConfigIni("Config.ini")
        netType = config.getKeyValue(u"网络连接类型")
        ethernet = config.getKeyValue(u"网络连接名称")
        user = config.getKeyValue(u"用户名")
        password = config.getKeyValue(u"密码")
        netManger.setEthernetInfo(netType, ethernet, user, password)
    except:
        logging.error("初始化失败，请检查配置文件：Config.ini")
        traceStr = traceback.format_exc()
        logging.error(traceStr)
        
    hasUnvisit = True
    batIdx = 0
    while hasUnvisit:
        logging.debug("\r\n\r\n")
        logging.debug("batIdx: %d", batIdx)

        if isOutOfData()==True:
            logging.error("isOutOfData" + str(datetime.datetime.now()))
            time.sleep(24*60*60)

        #init ev
        try:
            os.startfile("C:\\Program Files\\Internet Explorer\\iexplore.exe")
        except:
            logging.error("空白页打开异常")
            traceStr = traceback.format_exc()
            logging.error(traceStr)

        # view baobei
        try:
            hasUnvisit = search_baobei()
            closeAllRunningIE()
        except:
            traceStr = traceback.format_exc()
            logging.error(traceStr)
            closeAllRunningIE()
        break

        # change IP
        try:
            netManger.changeIP()
        except:
            traceStr = traceback.format_exc()
            logging.error(traceStr)

        # next view
        batIdx += 1






class PyResorceRelease(object):
    def __init__(self):
        a = 0

    def __del__(self):
        try:
            logging.error("release all py exe resorce")
            closeAllRunningIE()
        except:
            traceStr = traceback.format_exc()
            logging.error(traceStr)
            
def tbsearch_2897106():
    releaser = PyResorceRelease()
    tbsearch_2897106_dowork()


if __name__=='__main__':
    #tbsearch_2897106()
    print isActive(u"tbsearch")
