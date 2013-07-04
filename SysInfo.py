#coding=GBK

import wmi 
import os 
import sys 
import platform 
import time 
import chardet

def str2unicode(string):
    uniStr = u""
    if isinstance(string, unicode):
        uniStr = string
    else:
        result = chardet.detect(string)
        print "result:", result
        uniStr = string.decode(result["encoding"])
    return uniStr
    
#sys.getdefaultencoding()
def sys_version():  
    c = wmi.WMI() 
    #获取操作系统版本 
    for mysys in c.Win32_OperatingSystem():
        print "Version:%s" % mysys.Caption.encode("GBK"),"Vernum:%s" % mysys.BuildNumber
        print mysys.keys
        #print  mysys.OSArchitecture
        print mysys.NumberOfProcesses #当前系统运行的进程总数
 
def cpu_mem(): 
    c = wmi.WMI()        
    #CPU类型和内存 
    for processor in c.Win32_Processor(): 
        print "Processor ID: %s" % processor.DeviceID
        print "Processor Manufacturer: %s" % processor.Manufacturer
        print "Processor PNPDeviceID: %s" % processor.PNPDeviceID 
        print "processor.InstallDate: ", processor.InstallDate
        print "Process Name: %s" % processor.Name.strip() 
    for Memory in c.Win32_PhysicalMemory(): 
        print "Memory Capacity: %.fMB" %(int(Memory.Capacity)/1048576) 
 
def cpu_use(): 
    #5s取一次CPU的使用率 
    c = wmi.WMI() 
    while True: 
        for cpu in c.Win32_Processor(): 
            timestamp = time.strftime('%a, %d %b %Y %H:%M:%S', time.localtime()) 
            print '%s | Utilization: %s: %d %%' % (timestamp, cpu.DeviceID, cpu.LoadPercentage) 
            time.sleep(5)    
              
def disk(): 
    c = wmi.WMI()    
    #获取硬盘分区 
    for physical_disk in c.Win32_DiskDrive (): 
        for partition in physical_disk.associators ("Win32_DiskDriveToDiskPartition"): 
            for logical_disk in partition.associators ("Win32_LogicalDiskToPartition"): 
                print physical_disk.Caption.encode("GBK"), partition.Caption.encode("GBK"), logical_disk.Caption 
    
    #获取硬盘使用百分情况 
    for disk in c.Win32_LogicalDisk (DriveType=3): 
        print disk.Caption, "%0.2f%% free" % (100.0 * long (disk.FreeSpace) / long (disk.Size)) 
 
def network(): 
    c = wmi.WMI()    
    #获取MAC和IP地址 
    for interface in c.Win32_NetworkAdapterConfiguration (IPEnabled=1): 
        print "MAC: %s" % interface.MACAddress 
    for ip_address in interface.IPAddress: 
        print "ip_add: %s" % ip_address 
    print 
    
    #获取自启动程序的位置 
    for s in c.Win32_StartupCommand (): 
        print "[%s] %s <%s>" % (s.Location.encode("GBK"), s.Caption.encode("GBK"), s.Command.encode("UTF8"))  
    
    
    #获取当前运行的进程 
    for process in c.Win32_Process (): 
        print process.ProcessId, process.Name 

def getPlatform():
    platformStr = u""
    try:
        platformStr = platform.platform()
        platformStr = str2unicode(platformStr)
    except:
        platformStr = u""
    return platformStr

def getEnableMacs():
    Macs = u""
    try:
        c = wmi.WMI()    
        #获取MAC和IP地址 
        for interface in c.Win32_NetworkAdapterConfiguration (IPEnabled=1): 
            mac = interface.MACAddress
            print "interface.Caption: ", interface.Caption
            print "interface.Description: ", interface.Description
            print "interface.DNSDomain: ", interface.DNSDomain
            print "interface.DNSHostName: ", interface.DNSHostName
            print "interface.DomainDNSRegistrationEnabled: ", interface.DomainDNSRegistrationEnabled
            print "interface.Index: ", interface.Index
            print "interface.IPEnabled: ", interface.IPEnabled
            print "interface.IPXAddress: ", interface.IPXAddress
            print "interface.IPXFrameType: ", interface.IPXFrameType
            print "interface.IPXMediaType: ", interface.IPXMediaType
            print "interface.ServiceName: ", interface.ServiceName
            print "interface.WINSHostLookupFile: ", interface.WINSHostLookupFile
             
             
            
            
            mac = str2unicode(processor.Name)
            Macs = Macs + mac + u"-"
        Macs = Macs[0:-1]
    except:
        Macs = u""
    return Macs

def getCpuInfo():
    cpu = u""
    try:
        c = wmi.WMI()        
        #CPU类型和内存 
        for processor in c.Win32_Processor(): 
            mac = str2unicode(processor.Name)
            cpu = cpu + mac + u"-"
        cpu = cpu[0:-1]
    except:
        cpu = u""
    return cpu

def getSelfInfo():
    element = u"@@@"
    platformStr = getPlatform()
    macs = getEnableMacs()
    cpu = getCpuInfo()
    strSelfInfo = element + platformStr + element + macs + element + cpu + element
    print "type(strSelfInfo): ", type(strSelfInfo)
    print "strSelfInfo: ", strSelfInfo

def main(): 
    #sys_version() 
    #cpu_mem() 
    #disk() 
    #network() 
    #cpu_use()
    getSelfInfo()

    aa = "处理器"
    bb = str2unicode(aa)
 
if __name__ == '__main__': 
    main() 
    print "platform.system(): ", platform.system() 
    print "platform.release(): ", platform.release() 
    print "platform.version(): ", platform.version() 
    print "platform.platform(): ", platform.platform() 
    print "platform.machine(): ", platform.machine()
