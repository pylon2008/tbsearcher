# coding=GBK
import logging, datetime, socket, traceback, os
import xlwt, xlrd
from win32com.shell import shell, shellcon
from xpylon.xethernet.IEProxy import *
from xpylon.xethernet.IEExplorer import *
from xpylon.xethernet.NetManager import *
from xpylon.xutil.Activation import *
from xpylon.xutil.xstring import *
import urllib
from bs4 import BeautifulSoup

#IE_TIME_OUT_NEW_PAGE = 20
###################################################################################
class BaobeiSearher(object):
    def __init__(self, searchKey, targetUrl):
        self.searchKey = searchKey
        self.targetUrl = targetUrl
        self.targetTitle = None
        self.searchPageIE = None
        self.curPageIdx = 0
        self.allSearchPages = []

    def doSearch(self):
        # open taobao
        url = u"http://www.taobao.com/"
        self.searchPageIE = IEExplorer()
        self.searchPageIE.openURL(url)
        self.searchPageIE.setVisible(1)
        while self.searchPageIE.waitBusy(IE_TIME_OUT_NEW_PAGE)==True:
            self.searchPageIE.stop()
            time.sleep(0.1)

        #input search key
        nodeSearchInput = self.getSearchUnputNode()
        nodeSearchInput.click()
        nodeSearchInput.focus()
        enumHumanInput(nodeSearchInput, self.searchKey)

        #search
        nodeSearchButton = self.getSearchButtonNode()
        nodeSearchButton.click()
        nodeSearchButton.focus()
        self.getTargetUrlInfo()
        while self.searchPageIE.waitBusy(IE_TIME_OUT_NEW_PAGE)==True:
            self.searchPageIE.stop()
            time.sleep(0.1)

        #find the target
        self.doFindTargetBaobei()
        
#http://s.taobao.com/search?q=%C5%ED%D3%A6%C1%C1&app=detail
#http://s.taobao.com/search?q=男士短袖衬衫&suggest=0_5&wq=男士&suggest_query=男士&source=suggest&initiative_id=tbindexz_20130706&spm=1.1000386.5803581.d4908513&sourceId=tb.index&search_type=item&commend=all

    def doFindTargetBaobei(self):
        while True:
            self.getCurPageSearchItem()
            break
        time.sleep(10)

    def getCurPageSearchItem(self):
        nextPageNode = self.getNextPageNode()
        nextPageNode.scrollIntoView(True)
        while self.searchPageIE.waitBusy(IE_TIME_OUT_NEW_PAGE)==True:
            self.searchPageIE.stop()
            time.sleep(0.1)
        
        strDbgInfo = u"Page:" + str(self.curPageIdx) + u", locationURL: " + self.searchPageIE.locationURL()
        logging.debug(strDbgInfo)
        body = self.searchPageIE.getBody()
        nodesDiv = getSubNodesByTag(body, u"div")
        nodesItem = []
        for node in nodesDiv:
            if node.className == u"item-box":
                nodesItem.append(node)
        strDbgInfo = u"cur page item count: " + str(len(nodesItem))
        logging.debug(strDbgInfo)
        for item in nodesItem:
            print "item.nodeItems.length: ", item.childNodes.length

    def getNextPageNode(self):
        body = self.searchPageIE.getBody()
        nodesA = getSubNodesByTag(body, u"a")
        nodesNextPage = []
        for node in nodesA:
            if node.className==u"page-next":
                nodesNextPage.append(node)
        if len(nodesNextPage) != 1:
            strDbg = u"num of next page button: " + str( len(nodesNextPage) )
            logging.error(strDbg)
            raise ValueError, strDbg
        return nodesNextPage[0]
        
    def getTargetUrlInfo(self):
        content = urllib.urlopen(self.targetUrl).read()
        soup = BeautifulSoup(content)
        self.targetTitle = soup.find(u'title')
        
    def getSearchUnputNode(self):
        body = self.searchPageIE.getBody()
        nodesInput = getSubNodesByTag(body, u"input")
        nodeSearchInput = getNodeByAttr(nodesInput, u"id", u"q")
        if nodeSearchInput == None:
            raise ValueError, u"Can't find the input edit"
        return nodeSearchInput

    def getSearchButtonNode(self):
        body = self.searchPageIE.getBody()
        nodesInput = getSubNodesByTag(body, u"button")
        nodeSearchButton = None
        for node in nodesInput:
            value = u"submit"
            try: 
                if node.getAttribute(u"type")==u"submit":
                    value = None
                    value = node.getAttribute(u"tabIndex")
            except:
                a = 0
            if value==0:
                nodeSearchButton = node
                break

        if nodeSearchButton == None:
            raise ValueError, u"Can't find the submit button"
        return nodeSearchButton



###################################################################################

class TaobaoSearcher(object):
    def __init__(self):
        self.numBaobei = 1
        self.baobeiSet = []
        self.searcher = None
        
        self.readUrlConfig()
        self.baobeiIndex = self.getRandomBaobeiIndex()
        if self.baobeiIndex==None:
            return
        self.randomKey = self.getRandomSearchKey()

    def getRandomSearchKey(self):
        baobei = self.baobeiSet[self.baobeiIndex]
        keystr = baobei[1]
        keystr = str2unicode(keystr)
        keys = keystr.split(u" ")
        numKey = len(keys)
        keyIdx = random.randint(0, numKey-1)
        return keys[keyIdx]

    def getRandomBaobeiIndex(self):
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
        if len(self.randomVisit) != 1:
            logging.error("random cal error")

        return self.randomVisit[0]
        
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

    def doSearch(self):
        self.searcher = BaobeiSearher(self.randomKey, self.baobeiSet[self.baobeiIndex][2])
        self.searcher.doSearch()

    def getBaobei(self, visitIdx):
        return self.mainIE[visitIdx]

    def readUrlConfig(self):
        # 读取所有宝贝
        try:
            filePath = "SearchConfig.xls"
            wb = xlrd.open_workbook(filePath)
            sheet = wb.sheet_by_index(0)
            for row_index in range(sheet.nrows):
                numVisit = sheet.cell(row_index,0).value
                numVisit = (int)(numVisit)
                key = sheet.cell(row_index,1).value
                url = sheet.cell(row_index,2).value
                self.baobeiSet.append( [numVisit, key, url] )
        except:
            logging.error("SearchConfig.xls read error!")
            traceStr = traceback.format_exc()
            logging.error(traceStr)
           
    def writeUrlConfig(self):
        try:
            wb = xlwt.Workbook()
            sheet = wb.add_sheet('sheet 1')
            for row_index in range(len(self.baobeiSet)):
                baobei = self.baobeiSet[row_index]
                num = baobei[0]-1
                key = baobei[1]
                url = baobei[2]
                sheet.write(row_index,0,num)
                sheet.write(row_index,1,key)
                sheet.write(row_index,2,url)
            sheet.col(1).width = 3333*6
            sheet.col(2).width = 3333*8
            filePath = "SearchConfig_backup.xls"
            wb.save(filePath)
            win32file.DeleteFile(u"SearchConfig.xls")
            win32file.CopyFile(u"SearchConfig_backup.xls", u"SearchConfig.xls", False)
        except:
            logging.error("SearchConfig_backup.xls write error!")
            traceStr = traceback.format_exc()
            logging.error(traceStr)
       
##    def closeAllIE(self):
##        numVisitBaobei = self.numVisit()
##        for mainIdx in range(numVisitBaobei):
##            baobei = self.getBaobei(mainIdx)
##            debugInfo = "mainIdx: "+ str(mainIdx) + ", type(baobei): " + str(type(baobei)) + ", baobei.getNumSubIE(): " + str(baobei.getNumSubIE())
##            logging.debug(debugInfo)
##            for subIdx in range(baobei.getNumSubIE()):
##                subIE = baobei.getNewSubIE(subIdx)
##                debugInfo = "subIdx: "+str(subIdx)+ ", type(subIE): "+ str(type(subIE))
##                logging.debug(debugInfo)
##                while subIE.waitBusy(IE_TIME_OUT_NEW_PAGE)==True:
##                    subIE.stop()
##                    time.sleep(0.1)
##                subIE.setForeground()
##                time.sleep(IE_INTERVAL_TIME_CLOSE)
##                subIE.quit()
##            while baobei.getMainIE().waitBusy(IE_TIME_OUT_NEW_PAGE)==True:
##                baobei.getMainIE().stop()
##                time.sleep(0.1)
##            baobei.getMainIE().setForeground()
##            time.sleep(IE_INTERVAL_TIME_CLOSE)
##            baobei.getMainIE().quit()

###################################################################################

def initLogging():
    LOG_FILENAME="TBSearcher.log"
    logging.basicConfig(filename=LOG_FILENAME, level=logging.DEBUG, format='%(asctime)s - %(levelname)s: %(message)s')
    curTime = datetime.datetime.now
    strTime = str(curTime)
    logging.debug("===============================================Begin Log===============================================")
    logging.debug( socket.gethostname() )

    # define a Handler which writes INFO messages or higher to the sys.stderr
    console = logging.StreamHandler();
    console.setLevel(logging.DEBUG);    
    # set a format which is simpler for console use
    #formatter = logging.Formatter('LINE %(lineno)-4d : %(levelname)-8s %(message)s')
    formatter = logging.Formatter('%(message)s')
    # tell the handler to use this format
    console.setFormatter(formatter);
    logging.getLogger('').addHandler(console); 
 
def search_baobei():
    searcher = TaobaoSearcher()
    searcher.doSearch()
    searcher.writeUrlConfig()
    
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
##        try:
##            os.startfile("C:\\Program Files\\Internet Explorer\\iexplore.exe")
##        except:
##            logging.error("空白页打开异常")
##            traceStr = traceback.format_exc()
##            logging.error(traceStr)

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
