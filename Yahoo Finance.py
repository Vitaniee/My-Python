#-*- coding = utf-8 -*-
#@Time: 2020/7/14 13:45
#Author: Vitanie
#@File: Yahoo Finance.py
#@Software: PyCharm



from bs4 import BeautifulSoup
import re
import urllib.request,urllib.error
import xlwt
import random

def main():
    sector_list=["_basic_materials", "_communication_services", "_consumer_cyclical", "_consumer_defensive","_energy","_financial_services", "_healthcare_","industrials", "_real_estate", "_technology","_utilities"]
    for sector in sector_list:
            getSectorInformation(sector)
            print(sector,"is finished")

#get formation of each sector
def getSectorInformation(sector):
    book = xlwt.Workbook()# creat a workbook for the sector
    url1="https://finance.yahoo.com/screener/predefined/ms"+sector+"?count=100&offset=0"
    symbol_list=getCompanySymbol(url1)
    print(symbol_list)
    count=0
    for symbol in symbol_list:
        url2="https://finance.yahoo.com/quote/"+symbol+"/history?period1=1577836800&period2=1594684800&interval=1mo&filter=history&frequency=1mo"
        historicalStatistics=getHistoricalStatistics(url2)
        saveData(symbol, historicalStatistics,sector,book)
        # print(historicalStatistics)
        count += 1
        print(count)# count how many comanies' statistics have been succesfully saved

# ask URL for HTML
def askURL(url):
    #creat a user-agent list and randomly change the user-agent, which can tell the url that WE ARE NOT SPIDERS(hhhhhh
    user_agent=["Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/535.1 (KHTML, like Gecko) Chrome/14.0.835.163 Safari/535.1","Mozilla/5.0 (Windows NT 6.1; WOW64; rv:6.0) Gecko/20100101 Firefox/6.0","Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/534.50 (KHTML, like Gecko) Version/5.1 Safari/534.50","Opera/9.80 (Windows NT 6.1; U; zh-cn) Presto/2.9.168 Version/11.50","Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0; .NET CLR 2.0.50727; SLCC2; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.3; .NET4.0C; Tablet PC 2.0; .NET4.0E)","Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; InfoPath.3)","Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 5.1; Trident/4.0; GTB7.0)","Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1)","Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1)","Mozilla/5.0 (Windows; U; Windows NT 6.1; ) AppleWebKit/534.12 (KHTML, like Gecko) Maxthon/3.0 Safari/534.12","Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.3; .NET4.0C; .NET4.0E)","Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.3; .NET4.0C; .NET4.0E; SE 2.X MetaSr 1.0)","Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/534.3 (KHTML, like Gecko) Chrome/6.0.472.33 Safari/534.3 SE 2.X MetaSr 1.0","Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.3; .NET4.0C; .NET4.0E)","Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.1 (KHTML, like Gecko) Chrome/13.0.782.41 Safari/535.1 QQBrowser/6.9.11079.201","Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.3; .NET4.0C; .NET4.0E) QQBrowser/6.9.11079.201","Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)"]
    head={"user-agent":user_agent[random.randint(0,len(user_agent)-1)]}
    request=urllib.request.Request(url,headers=head)# disguise
    html=""
    try:
        response=urllib.request.urlopen(request)
        html=response.read().decode("utf-8")
        # print(html)
    except urllib.error.URLError as e:
        if hasattr(e,"code"):
            print(e.code)
        if hasattr(e,"reason"):
            print(e.reason)
    return html

#get the fist 100 symbols of companys in a sector
def getCompanySymbol(url1):
    symbol_list=[] #creat a list to save symbols
    html=askURL(url1)
    soup=BeautifulSoup(html, "html.parser")# analysis by using bs4
    # print(html)
    # search for the symbol we want
    for item in soup.find_all("tr"):
        # print(item)
        aclass=item.find_all("a")
        aclass=str(aclass)
        # print(aclass)
        findSymbol=re.compile(r'href="/quote.*?p=(.*?)"')
        symbol=re.findall(findSymbol,aclass)
        # remove the blank list
        if len(symbol)==0:
            continue
        else:
            symbol=symbol[0]
            symbol_list.append(symbol)
    # print(symbol_list)
    # print(len(symbol_list))
    # if failed, try again
    if len(symbol_list)==0:
        getCompanySymbol(url1)
    else:
        return symbol_list

#get the historical statistics of a company
def getHistoricalStatistics(url2):
    historicalStatistics=[]# create a list to save the statistics of a company
    html=askURL(url2)# call askURL() to get HTML
    soup=BeautifulSoup(html, "html.parser")# analysis
    # print(html)
    # search for the statisitcs we want
    for item in soup.find_all("tr",class_="BdT Bdc($seperatorColor) Ta(end) Fz(s) Whs(nw)"):
        span=item.find_all("span")
        # remove the blank list for "dividend" and save the real list we want
        if len(span)==7:
            span=str(span)
            findNumber=re.compile(r'>(.+?)<',re.S)
            number=re.findall(findNumber,span)
            time=number[0]
            close=number[8]# close value of the stock
            volume=number[12]
            info=[time,close,volume]
            # print(info)
            historicalStatistics.append(info)
        else:
            continue
    # if failed, try again
    if len(historicalStatistics)==0 or historicalStatistics==None:
        getHistoricalStatistics(url2)
    else:
        historicalStatistics.reverse()
        return historicalStatistics

#save the data of each company in a sheet,save the data of the whole sector in a book
def saveData(symbol,historicalStatistics,sector,book):
    file_name=sector+".xls" # set the file name
    print("saving...")
    sheet=book.add_sheet(symbol) # create a sheet for each company
    col=("Time","Close price","Volume") # write the column
    for i in range(3):
        sheet.write(0,i+1,col[i])
    # write the statistics of a campany
    for i in range(len(historicalStatistics)):
        sheet.write(i+1,0,i+1)
        for j in range(3):
            sheet.write(i+1,j+1,historicalStatistics[i][j])
    book.save(file_name)#save the file

main()#call main function



