# coding=GBK
import logging, datetime, socket, traceback, os
import xlwt, xlrd
from xpylon.xethernet.IEProxy import *
from xpylon.xethernet.IEExplorer import *
from xpylon.xethernet.NetManager import *
from xpylon.xutil.Activation import *
from xpylon.xutil.xstring import *
import urllib
from bs4 import BeautifulSoup

#IE_TIME_OUT_NEW_PAGE = 20
NUM_SUB_PAGE_MIN = 2
NUM_SUB_PAGE_MAX = 4
TIME_BAOBEI_VIEW_MIN = 280
TIME_BAOBEI_VIEW_MAX = 400

###################################################################################
class TaobaoBaobeiViewer(object):
    def __init__(self, mainIE):
        self.mainIE = mainIE
        self.subIESet = []
        self.imgHrefNodes = []
        self.subNodes = []
        self.timeBegOp = None                   # 开始操作宝贝的起点时间，从开始滚动开始

    def getRandomSubIE(self):
        numSubIE = random.randint(NUM_SUB_PAGE_MIN, NUM_SUB_PAGE_MAX)
        numImgHrefNodes = len(self.imgHrefNodes)
        if numSubIE > numImgHrefNodes:
            numSubIE = numImgHrefNodes

        allIdxs = []
        for i in range(numSubIE):
            xx = random.randint(0, numImgHrefNodes-1)
            while True:
                if xx in allIdxs:
                    xx = random.randint(0, numImgHrefNodes-1)
                else:
                    allIdxs.append(xx)
                    break
                
        for i in range(len(allIdxs)):
            self.subNodes.append( self.imgHrefNodes[allIdxs[i]] )

    def getImgHrefNodes(self):
        body = self.mainIE.getBody()
        nodesImg = getSubNodesByTag(body, "img")
        for node in nodesImg:
            nodeParent = node.parentElement
            if nodeParent!=None:
                if nodeParent.tagName==u"a" or nodeParent.tagName==u"A":
                    href = nodeParent.getAttribute("href")
                    if type(href)==unicode and href!=u"":
                        if u"detail" in href:
                            self.imgHrefNodes.append(nodeParent)

    def getNumSubIE(self):
        return len(self.subNodes)

    def getTimeBegOp(self):
        return self.timeBegOp

    def createNewSubIE(self, subIdx):
        subNode = self.subNodes[subIdx]
        url = subNode.getAttribute("href")
        ie = IEExplorer()
        ie.openURL(url)
        ie.setVisible(1)
        self.subIESet.append(ie)

    def getMainIE(self):
        return self.mainIE
    
    def getNewSubIE(self, subIdx):
        return self.subIESet[subIdx]

    def baobeiSrcollBeg(self):
        self.timeBegOp = datetime.datetime.now()
        mainIE = self.getMainIE()
        while mainIE.waitBusy(IE_TIME_OUT_NEW_PAGE)==True:
            mainIE.stop()
            time.sleep(0.1)
        mainIE.waitReadyState(IE_TIME_OUT_NEW_PAGE)
        mainIE.setForeground()
        while mainIE.waitBusy(IE_TIME_OUT_NEW_PAGE)==True:
            mainIE.stop()
            time.sleep(0.1)
        isReady = mainIE.waitReadyState(IE_TIME_OUT_NEW_PAGE)
        time.sleep(1)
        timeOut = random.randint(3, 5)
        mainIE.stayInSubPage(timeOut)

    def openCurBaobei(self):
        logging.debug("openCurBaobei")
        self.getImgHrefNodes()
        self.getRandomSubIE()
        
        numSubIE = self.getNumSubIE()
        dbgInfo = u"numSubIE: " + str2unicode(str(numSubIE))
        logging.debug(dbgInfo)
        for subIdx in range(numSubIE):
            debugInfo = "subIdx: " + str(subIdx) + ", url: " + self.subNodes[subIdx].getAttribute("href")
            logging.debug(debugInfo)
            debugInfo = "come back to mainIE waitBusy before setForeground: "
            logging.debug(debugInfo)
            # 回宝贝界面
            while self.mainIE.waitBusy(IE_TIME_OUT_NEW_PAGE)==True:
                self.mainIE.stop()
                time.sleep(0.1)
            self.mainIE.waitReadyState(IE_TIME_OUT_NEW_PAGE)
            self.mainIE.setForeground()
            time.sleep(1)
            self.mainIE.resizeMax()
            time.sleep(1)
            debugInfo = "come back to mainIE waitBusy after setForeground: "
            logging.debug(debugInfo)
            while self.mainIE.waitBusy(IE_TIME_OUT_NEW_PAGE)==True:
                self.mainIE.stop()
                time.sleep(0.1)
            self.mainIE.waitReadyState(IE_TIME_OUT_NEW_PAGE)

            # 打开子页面
            logging.debug("self.mainIE.scrollToNode")
            subNode = self.subNodes[subIdx]
            self.mainIE.scrollToNode(subNode)
            subNode.focus()
            logging.debug("self.createNewSubIE")
            self.createNewSubIE(subIdx)

            # 滚动子页面
            logging.debug("subIE.stayInSubPage")
            subIE = self.getNewSubIE(subIdx)
            while subIE.waitBusy(IE_TIME_OUT_NEW_PAGE)==True:
                subIE.stop()
                time.sleep(0.1)
            isReady = subIE.waitReadyState(IE_TIME_OUT_NEW_PAGE)
            timeOut = random.randint(3, 5)
            subIE.stayInSubPage(timeOut)

            if subIdx != numSubIE-1:
                time.sleep(2)           


###################################################################################
def getTBIDFromUrl(url):
    allIdName = [u"id=", u"ad_id=", u"cm_id=", u"pm_id="]
    allNum = [u"0", u"1", u"2", u"3", u"4", u"5", u"6", u"7", u"8", u"9"]
    idElement = u"id="
    idCount = url.count(idElement)
    begPos = 0
    idValue = None
    for idIdx in range(idCount):
        begPos = url.find(idElement, begPos, len(url)-1)
        idPre = url[begPos-3:begPos]
        idName = idPre + idElement
        if idName not in allIdName:
            idName = idElement
        if idName == idElement:
            begIdx = begPos+len(idElement)
            endIdx = len(url)
            for i in range(begIdx, endIdx):
                if url[i] not in allNum:
                    endIdx = i
                    break
            idValue = url[begIdx:endIdx]
            idValue = str2unicode(idValue)
            break
        begPos += len(idElement)
    return idValue

def test_getTBIDFromUrl():
    url0 = u"http://detail.tmall.com/item.htm?id=9556473603&spm=a230r.1.14.3.5RD65G&ad_id=&am_id=&cm_id=140105335569ed55e27b&pm_id="
    url1 = u"http://item.taobao.com/item.htm?spm=a230r.1.14.52.5RD65G&id=24075560410"
    url2 = u"http://item.taobao.com/item.htm?spm=a230r.1.14.124.5RD65G&id=18352266988"
    url3 = u"http://item.taobao.com/item.htm?spm=a230r.1.14.11.5RD65G&id=17351284242&ad_id=&am_id=&cm_id=140105335569ed55e27b&pm_id="
    url4 = u"http://item.taobao.com/item.htm?spm=a230r.1.14.70.5RD65G&id=19083043608"
    url5 = u"http://item.taobao.com/item.htm?spm=a230r.1.14.73.5RD65G&id=19083043608"
    urls = [url0, url1, url2, url3, url4, url5]
    for url in urls:
        print getTBIDFromUrl(url)

class SearchRecord(object):
    def __init__(self, node):
        self.rcdNode = node
        self.summaryStr = None
        self.summaryNode = None
        self.extractSummary()

    def extractSummary(self):
        self.summaryStr = u"Not extract the summary"
        try:
            self.summaryNode = self.extractSummaryNode()
            self.summaryStr = self.summaryNode.title
            self.summaryStr = str2unicode(self.summaryStr)
        except:
            traceStr = traceback.format_exc()
            logging.error(traceStr)

    def extractSummaryNode(self):
        nodeH3 = getSubNodesByTag(self.rcdNode, u"h3")
        if len(nodeH3) != 1 or nodeH3[0].className!=u"summary":
            raise ValueError, u"find the summary node error"
        return nodeH3[0].childNodes[0]

    def getSummaryNode(self):
        return self.summaryNode
    
    def getSummaryID(self):
        url = self.summaryNode.getAttribute(u"href")
        return getTBIDFromUrl(url)
        
    def getSummaryStr(self):
        return self.summaryStr


###################################################################################
class BaobeiSearher(object):
    def __init__(self, searchKey, targetUrl):
        self.searchKey = searchKey
        self.targetUrl = targetUrl
        self.targetID = getTBIDFromUrl(self.targetUrl)
        self.targetTitle = None
        self.targetViewer = None
        self.searchPageIE = None
        self.curPageIdx = 0
        self.curPageInnerIdx = -1
        self.allSearchPages = []
        self.randomBaobei = []

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
        logging.debug(self.searchKey)
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
        
    def doFindTargetBaobei(self):
        # page next loop
        while True:
            self.getCurPageSearchItem()
            self.procViewRandomBaobei()
            if self.isTargetInThisPage()==True:
                break
            else:
                self.curPageIdx += 1
                nextPageNode = self.getNextPageNode()
                self.searchPageIE.scrollToNode(nextPageNode)
                nextPageNode.click()
                nextPageNode.focus()

        dbgInfo = u"randomBaobei: " + str2unicode( str(self.randomBaobei) )
        logging.debug(dbgInfo)
        # go to the targe baobei page
        targetNode = self.allSearchPages[self.curPageIdx][self.curPageInnerIdx]
        rcd = SearchRecord(targetNode)
        summaryNode = rcd.getSummaryNode()
        self.searchPageIE.scrollToNode(summaryNode)
        summaryNode.focus()
        url = summaryNode.getAttribute(u"href")
        baobeiIE = IEExplorer()
        baobeiIE.openURL(url)
        baobeiIE.setVisible(1)
        self.targetViewer = TaobaoBaobeiViewer(baobeiIE)
        self.targetViewer.baobeiSrcollBeg()
        self.targetViewer.openCurBaobei()

        # stay in baobei page
        timeTotal = random.randint(TIME_BAOBEI_VIEW_MIN, TIME_BAOBEI_VIEW_MAX)
        timeBeg = self.targetViewer.getTimeBegOp()
        timeNow = datetime.datetime.now()
        timePass = (timeNow-timeBeg).seconds
        timeSleep = timeTotal - timePass
        if timeSleep <= 0:
            timeSleep = 10
        dbgInfo = u"stay in baobei page time: " + str2unicode(str(timeSleep))
        logging.debug(dbgInfo)
        time.sleep(timeSleep)

    def viewRandomBaobei(self, randomIdx):
        randomNode = self.allSearchPages[self.curPageIdx][randomIdx]
        rcd = SearchRecord(randomNode)
        summaryNode = rcd.getSummaryNode()
        self.searchPageIE.scrollToNode(summaryNode)
        summaryNode.focus()
        url = summaryNode.getAttribute(u"href")
        baobeiIE = IEExplorer()
        baobeiIE.openURL(url)
        baobeiIE.setVisible(1)
        isReady = baobeiIE.waitReadyState(IE_TIME_OUT_NEW_PAGE)
        timeOut = random.randint(3, 5)
        baobeiIE.stayInSubPage(timeOut)

        # set the search page into top
        while self.searchPageIE.waitBusy(IE_TIME_OUT_NEW_PAGE)==True:
            self.targetViewer.getMainIE().stop()
            time.sleep(0.1)
        self.searchPageIE.waitReadyState(IE_TIME_OUT_NEW_PAGE)
        self.searchPageIE.setForeground()
        time.sleep(1)
        #self.searchPageIE.resizeMax()
        time.sleep(1)
   
    def procViewRandomBaobei(self):
        if self.isTargetInThisPage()==True:
            targetBaobei = (self.curPageIdx, self.curPageInnerIdx)
            self.randomBaobei.append(targetBaobei)
        if len(self.allSearchPages[-1]) < 3:
            return
        
        if self.curPageIdx == 0:
            # page 0, we have a random baobei view
            randomIdx = None
            while True:
                innerIdx = random.randint(0, len(self.allSearchPages[-1])-1)
                toupleIdx = (self.curPageIdx, innerIdx)
                if toupleIdx not in self.randomBaobei:
                    randomIdx = innerIdx
                    self.randomBaobei.append(toupleIdx)
                    break
            self.viewRandomBaobei(randomIdx)

        if self.isTargetInThisPage()==True:
            # page 0, we have a random baobei view
            randomIdx = None
            while True:
                innerIdx = random.randint(0, len(self.allSearchPages[-1])-1)
                toupleIdx = (self.curPageIdx, innerIdx)
                if toupleIdx not in self.randomBaobei:
                    randomIdx = innerIdx
                    self.randomBaobei.append(toupleIdx)
                    break
            self.viewRandomBaobei(randomIdx)
                
        
    def isTargetInThisPage(self):
        return self.curPageInnerIdx != -1

    def getCurPageSearchItem(self):
        nodesItem = []
        self.allSearchPages.append(nodesItem)
        self.refreshOutAllItem()
        while self.searchPageIE.waitBusy(IE_TIME_OUT_NEW_PAGE)==True:
            self.searchPageIE.stop()
            time.sleep(0.1)

        # get all item node
        strDbgInfo = u"Page:" + str(self.curPageIdx) + u", locationURL: " + self.searchPageIE.locationURL()
        logging.debug(strDbgInfo)
        body = self.searchPageIE.getBody()
        nodesDiv = getSubNodesByTag(body, u"div")
        for node in nodesDiv:
            if node.className == u"item-box" and node.childNodes.length>=6:
                nodesItem.append(node)
        strDbgInfo = u"cur page item count: " + str(len(nodesItem))
        logging.debug(strDbgInfo)

        # get the target node
        self.allSearchPages[self.curPageIdx] = nodesItem
        for i in range(len(nodesItem)):
            node = nodesItem[i]
            try:
                rcd = SearchRecord(node)
                dbgInfo = str2unicode(str(i)) + u":" + rcd.getSummaryStr() + rcd.getSummaryID()
                logging.debug( dbgInfo )
                if self.isRecordTarget(rcd)==True:
                    self.curPageInnerIdx = i
            except:
                traceStr = traceback.format_exc()
                logging.error(traceStr)

    def isRecordTarget(self, rcd):
        if rcd.getSummaryStr() in self.targetTitle:
            titleBeg = u"<title>"
            titleEnd = u"-淘宝网</title>"
            beg = titleBeg+rcd.getSummaryStr()
            if beg in self.targetTitle:
                return rcd.getSummaryID() == self.targetID
        return False
    
    def refreshOutAllItem(self):
        self.searchPageIE.stayInSubPage(10)

    def getNextPageNode(self):
        body = self.searchPageIE.getBody()
        nodesA = getSubNodesByTag(body, u"a")
        nodesNextPage = []
        for node in nodesA:
            if node.className==u"page-next" and \
               node.getAttribute(u"title")==u"下一页":
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
        self.targetTitle = str2unicode(str(self.targetTitle))
        
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


###################################################################################

def initLogging():
    LOG_FILENAME="TBSearcher.log"
    logging.basicConfig(filename=LOG_FILENAME, level=logging.DEBUG, format='%(asctime)s - %(levelname)s: %(message)s')
    curTime = datetime.datetime.now
    strTime = str(curTime)
    logging.debug("===============================================Begin Log===============================================")
    logging.debug( socket.gethostname() )

##    # define a Handler which writes INFO messages or higher to the sys.stderr
##    console = logging.StreamHandler();
##    console.setLevel(logging.DEBUG);    
##    # set a format which is simpler for console use
##    #formatter = logging.Formatter('LINE %(lineno)-4d : %(levelname)-8s %(message)s')
##    formatter = logging.Formatter('%(message)s')
##    # tell the handler to use this format
##    console.setFormatter(formatter);
##    logging.getLogger('').addHandler(console); 
 
def search_baobei():
    searcher = TaobaoSearcher()
    searcher.doSearch()
    searcher.writeUrlConfig()
    
    # write URL config
    searcher.writeUrlConfig()
    return searcher.hasUnvisit()
    
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

        if isActive(u"tbsearch")==False:
            logging.error("isActive" + str(datetime.datetime.now()))
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
    #test_getTBIDFromUrl()
    tbsearch_2897106()
##        print isActive(u"tbsearch")
    #doActiveFile(u"active.txt")
