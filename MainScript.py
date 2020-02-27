from splinter import Browser
from selenium import webdriver
import openpyxl,time,threading,datetime

def GetLgHeNet(LstIP,browser,DctResult):
    for Ip in LstIP:
        StrResult  = 'None'
        print('GetLgHeNet : %s .Inprogress : %d/%d'%(Ip,LstIP.index(Ip)+1,len(LstIP)))
        try:
            time.sleep(3)
            browser.visit('https://lg.he.net/')
            time.sleep(3)
            browser.click_link_by_id('command_ping')
            browser.find_by_id('ip').fill(Ip.strip())
            browser.find_by_id('raw').click()
            browser.find_by_value('Probe').click()
            StrResult = browser.find_by_id('lg_return').text
        except Exception as error : print('GetLgHeNet : %s\n%s'%(Ip,error))
        
        if Ip in DctResult: DctResult[Ip]['LgHeNet'] = StrResult
        else: DctResult[Ip] = {'LgHeNet':StrResult}



    
def GetCenturyLink(LstIP,browser,DctResult):
    for Ip in LstIP:
        StrResult  = 'None'
        print('GetCenturyLink : %s .Inprogress : %d/%d'%(Ip,LstIP.index(Ip)+1,len(LstIP)))
        try:
            time.sleep(3)
            browser.visit('https://lookingglass.centurylink.com/')
            time.sleep(3)
            browser.find_by_value('Singapore').click()
            browser.find_by_xpath('//div[@class="col-12 col-sm-9 col-md-6 col-lg-6 col-xl-6"]/input').first.fill(Ip)
            browser.find_by_xpath('//div[@class="col-7 col-sm-7 col-md-6 col-lg-6 col-xl-6"]/input').first.fill('32')
            browser.find_by_xpath('//div[@class="col-7 col-sm-7 col-md-6 col-lg-6 col-xl-6"]/select').last.select('5')
            browser.find_by_xpath('//button[@class="btn  btn-primary btn-sm"]').first.click()
            StrResult = browser.find_by_xpath('//div[@class="container-fluid"]/div[3]/div[2]').text
        except Exception as error : print('GetCenturyLink : %s\n%s'%(Ip,error))

        if Ip in DctResult: DctResult[Ip]['CenturyLink'] = StrResult
        else: DctResult[Ip] = {'CenturyLink':StrResult}

def GetPCCW(LstIP,browser,DctResult):
    for Ip in LstIP:
        StrResult  = 'None'
        print('GetPCCW : %s .Inprogress : %d/%d'%(Ip,LstIP.index(Ip)+1,len(LstIP)))
        try:
            time.sleep(3)
            browser.visit('https://lookingglass.pccwglobal.com/')
            time.sleep(3)
            browser.find_by_xpath('//div[@id="srcContainer"]/select/option[@value="sin01"]').click()
            browser.find_by_xpath('//div[@id="rProtocolContainer"]/select/option[@value="ipv4"]').click()
            browser.find_by_xpath('//div[@id="rProfileContainer"]/select/option[@value="standard"]').click()
            browser.find_by_id('offNet').first.click()
            browser.find_by_id('newOffNet').first.fill(Ip)
            browser.find_by_id('addOff').first.click()
            #browser.find_by_xpath('//div[@id="serviceContainer"]/select/option[@value="ping"]').click()
            #browser.find_by_xpath('//div[@id="packetCountContainer"]/select/option[@value="100"]').click()
            browser.find_by_id('submit').first.click()
            while True:
                if browser.is_text_present('Query Complete'): break
                else: time.sleep(3)
            StrResult = browser.find_by_xpath('//div[@id="rsDiv"]').text
        except Exception as error : print('GetPCCW : %s\n%s'%(Ip,error))

        if Ip in DctResult: DctResult[Ip]['PCCW'] = StrResult
        else: DctResult[Ip] = {'PCCW':StrResult}


if __name__ == "__main__":
    while True:
        Stop = True
        StrTime = input('Nhập vào list các giờ cần chạy, mỗi giờ cách nhau dấu phẩy or dấu cách : ').strip()
        if ',' in StrTime: LstTmp = StrTime.split(',')
        else: LstTmp = StrTime.split(' ')
        for Time in LstTmp: 
            if Time.isdigit() == False: Stop = False
        if Stop : break
        print('\n')
    LstTime = []
    for Time in LstTmp: LstTime.append(int(Time))
    LstTimeCheck = []
    while True: 
        TimeCurrent = datetime.datetime.now()
        DateTimeCurrent = str(datetime.datetime.now().strftime('%d:%m:%Y %H:%M:%S'))
        StrTimeCheck = '%s_%s'%(str(TimeCurrent.day),str(TimeCurrent.hour))
        if TimeCurrent.hour in LstTime and StrTimeCheck not in LstTimeCheck:
            try:
                StrTimeTmp = str(datetime.datetime.now().strftime('%d_%m_%Y_%H_%M_%S'))
                NameFileResult = './/DataInfo/Report_%s.xlsx'%StrTimeTmp
                Wb = openpyxl.load_workbook(r'./DataInfo/LstIP.xlsx')
                SheetName = Wb.sheetnames
                Ws = Wb[SheetName[0]]
                LstIP,threads,DctResult = [],[],{}
                browser1 = Browser('chrome')
                browser2 = Browser('chrome')
                browser3 = Browser('chrome')

                for IndexRow in range(2,Ws.max_row+1): LstIP.append(Ws.cell(row = IndexRow,column = 1).value)
                x = threading.Thread(target=GetLgHeNet, args=(LstIP,browser1,DctResult))
                threads.append(x)
                x.start()
                x = threading.Thread(target=GetCenturyLink, args=(LstIP,browser2,DctResult))
                threads.append(x)
                x.start()
                x = threading.Thread(target=GetPCCW, args=(LstIP,browser3,DctResult))
                threads.append(x)
                x.start()
                for index, thread in enumerate(threads):
                    thread.join()

                browser1.quit()
                browser2.quit()
                browser3.quit()

                for IndexRow in range(2,Ws.max_row+1): 
                    Ip = Ws.cell(row = IndexRow,column = 1).value
                    LgHeNet, CenturyLink, PCCWglobal = ' ', ' ', ' '
                    if Ip in DctResult:
                        DctTmp = DctResult[Ip]
                        if 'LgHeNet' in DctTmp: LgHeNet = DctTmp['LgHeNet']
                        if 'CenturyLink' in DctTmp: CenturyLink = DctTmp['CenturyLink']
                        if 'PCCW' in DctTmp: PCCWglobal = DctTmp['PCCW']
                    
                    Ws.cell(row = IndexRow,column = 2).value = LgHeNet
                    Ws.cell(row = IndexRow,column = 3).value = CenturyLink
                    Ws.cell(row = IndexRow,column = 4).value = PCCWglobal
                    Ws.cell(row = IndexRow,column = 5).value = ' '

                Wb.save(NameFileResult)
                LstTimeCheck.append(StrTimeCheck)
            except Exception as error : print('Main : %s'%error)
        else:
            print('List time collect mỗi ngày : %s'%(','.join(LstTmp)))
            print('Time current : %s'%DateTimeCurrent)
            time.sleep(900)
