# coding=GBK
import logging, datetime, socket, traceback, os
from win32com.shell import shell, shellcon
from xpylon.xethernet.IEProxy import *
from xpylon.xethernet.IEExplorer import *
from xpylon.xethernet.NetManager import *
from xpylon.xutil.Activation import *

###################################################################################

class TaobaoSearcher(object):
    def __init__(self):
        self.mainIE = None 
        self.pageIdx = 0
        self.baobeiSet = []
        
        self.readUrlConfig()

        numUnvisit = self.numAllUnvisit()
        if 1 > numUnvisit:
            self.numBaobei = 0

        if self.numBaobei<=0:
            return None

        logging.debug("numUnvisit: %d, self.numBaobei: %d", numUnvisit, self.numBaobei)
        # 随机抽取访问对象
        unvisitIdx = self.allUnvisitIdx()
        self.randomVisit = []
        for i in range(self.numBaobei):
            while True:
                r = random.randint(0, len(unvisitIdx)-1)
                rr = unvisitIdx[r]
                if rr not in self.randomVisit:
                    if self.baobeiSet[rr][0] > 0:
                        self.randomVisit.append(rr)
                        break
        randomStr = "self.randomVisit: " + str(self.randomVisit)
        logging.debug(randomStr)
        
    def numVisit(self):
        return len(self.randomVisit)

    def numAllUnvisit(self):
        numUnvisit = 0
        for baobei in self.baobeiSet:
            if baobei[0]>0:
                numUnvisit += 1
        return numUnvisit

    def allUnvisitIdx(self):
        unvisit = []
        for i in range(len(self.baobeiSet)):
            baobei = self.baobeiSet[i]
            if baobei[0]>0:
                unvisit.append(i)
        return unvisit

    def hasUnvisit(self):
        numUnvisit = self.numAllUnvisit()
        return numUnvisit > 0

    def createBaobei(self, visitIdx):
        realIdx = self.randomVisit[visitIdx]
        self.baobeiSet[realIdx][0] -= 1
        url = self.baobeiSet[realIdx][1]
        ieExplorer = IEExplorer()
        ieExplorer.openURL(url)
        ieExplorer.setVisible(1)
        store = TaobaoBaobei(ieExplorer)
        self.mainIE.append(store)
        debugInfo = "createBaobei visitIdx(" + str(visitIdx) + "): " + url
        logging.debug(debugInfo)

    def getBaobei(self, visitIdx):
        return self.mainIE[visitIdx]

    def readUrlConfig(self):
        # 读取所有宝贝
        try:
            filePath = "UrlConfig.xls"
            wb = xlrd.open_workbook(filePath)
            sheet = wb.sheet_by_index(0)
            for row_index in range(sheet.nrows):
                numVisit = sheet.cell(row_index,0).value
                numVisit = (int)(numVisit)
                url = sheet.cell(row_index,1).value
                self.baobeiSet.append( [numVisit,url] )
        except:
            logging.error("UrlConfig.xls read error!")
            traceStr = traceback.format_exc()
            logging.error(traceStr)
           
    def writeUrlConfig(self):
        try:
            wb = xlwt.Workbook()
            sheet = wb.add_sheet('sheet 1')
            for row_index in range(len(self.baobeiSet)):
                baobei = self.baobeiSet[row_index]
                num = baobei[0]-1
                url = baobei[1]
                sheet.write(row_index,0,num)
                sheet.write(row_index,1,url)
            sheet.col(1).width = 3333*8
            filePath = "UrlConfig_backup.xls"
            wb.save(filePath)
            win32file.DeleteFile(u"UrlConfig.xls")
            win32file.CopyFile(u"UrlConfig_backup.xls", u"UrlConfig.xls", False)
        except:
            logging.error("UrlConfig_backup.xls write error!")
            traceStr = traceback.format_exc()
            logging.error(traceStr)
       
    def closeAllIE(self):
        numVisitBaobei = self.numVisit()
        for mainIdx in range(numVisitBaobei):
            baobei = self.getBaobei(mainIdx)
            debugInfo = "mainIdx: "+ str(mainIdx) + ", type(baobei): " + str(type(baobei)) + ", baobei.getNumSubIE(): " + str(baobei.getNumSubIE())
            logging.debug(debugInfo)
            for subIdx in range(baobei.getNumSubIE()):
                subIE = baobei.getNewSubIE(subIdx)
                debugInfo = "subIdx: "+str(subIdx)+ ", type(subIE): "+ str(type(subIE))
                logging.debug(debugInfo)
                while subIE.waitBusy(IE_TIME_OUT_NEW_PAGE)==True:
                    subIE.stop()
                    time.sleep(0.1)
                subIE.setForeground()
                time.sleep(IE_INTERVAL_TIME_CLOSE)
                subIE.quit()
            while baobei.getMainIE().waitBusy(IE_TIME_OUT_NEW_PAGE)==True:
                baobei.getMainIE().stop()
                time.sleep(0.1)
            baobei.getMainIE().setForeground()
            time.sleep(IE_INTERVAL_TIME_CLOSE)
            baobei.getMainIE().quit()

###################################################################################

def initLogging():
    LOG_FILENAME="TBSearcher.log"
    logging.basicConfig(filename=LOG_FILENAME, level=logging.DEBUG, format='%(asctime)s - %(levelname)s: %(message)s')
    curTime = datetime.datetime.now
    strTime = str(curTime)
    logging.debug("===============================================Begin Log===============================================")
    logging.debug( socket.gethostname() )
 
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

##        if isOutOfData()==True:
##            logging.error("isOutOfData" + str(datetime.datetime.now()))
##            time.sleep(24*60*60)
##        if isActive(u"tbsearch")==False:
##            logging.error("isActive" + str(datetime.datetime.now()))
##            time.sleep(24*60*60)

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

        # clear IE cookie
        try:
            clearIECookie()
        except:
            traceStr = traceback.format_exc()
            logging.error(traceStr)
            
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
    tbsearch_2897106()
##    while True:       
##        print isActive(u"tbsearch")
    #doActiveFile(u"active.txt")
