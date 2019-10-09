import requests
from lxml import html
import xlwt
import xlrd

def saveFile(pathName, fileName, data, sheetName):
    workbook = xlwt.Workbook(encoding='utf-8')
    sheet = workbook.add_sheet(sheetName)
    i = 0
    for key in data.keys():
        j = 0
        sheet.write(j, i, key)
        for val in data[key]:
            sheet.write(j+1, i, data[key][j])
            j = j + 1
        i = i + 1
    workbook.save(pathName+ '/'+ fileName+'.xls')

def parseROEData(stockid):
    Titles = ['ROE']
    url = ('http://vip.stock.finance.sina.com.cn/corp/view/vFD_FinancialGuideLineHistory.php?stockid='
               +stockid+ '&typecode=financialratios62')         # ROE
    data = {}
    timeReg = '//div[@id="con02-2"]//table//tbody//tr//td[1]/text()'
    itemReg = '//div[@id="con02-2"]//table//tbody//tr//td[2]/text()'
    page = requests.Session().get(url)
    htmlPage = html.fromstring(page.text)
    timeRegVal = htmlPage.xpath(timeReg)
    ROEVal = htmlPage.xpath(itemReg)
    yearTimeVal = []
    ROEYearVal = []
    for i in range(0,len(timeRegVal)):
        if timeRegVal[i][5:] == '12-31':
            yearTimeVal.append(timeRegVal[i])
            ROEYearVal.append(ROEVal[i])
    data['时间'] = yearTimeVal
    data['ROE'] = ROEYearVal
    return data

def checkROE(data,years,minVal):
    roeData = data['ROE']
    bigYears = 0
    if len(roeData) < years:
        years = len(roeData)
    for i in range(years):
        if '\xa0'==roeData[i]:
            bigYears = bigYears + 1
            continue
        if float(roeData[i])>minVal:
            bigYears = bigYears + 1
    if bigYears >= years:
        return True
    else:
        return False

def iterStockList(stockMap):
    for stockid in stockMap.keys():
        data = parseROEData(stockid)
        if checkROE(data,6,20):
            print(stockMap[stockid])
            saveFile('ROE',stockMap[stockid]+stockid, data, 'ROE')

def readFile(fileName):
    book = xlrd.open_workbook(fileName)
    stockMap = {}
    stockMap.values()
    for i in range(0,2):
        sheet = book.sheet_by_index(i)
        for i in range(1,sheet.nrows):
            stockMap[str(sheet.cell(i, 2).value)[0:6]] = sheet.cell(i,3).value
    return stockMap

# 市值，市盈率（TTM），ROE，营业收入，净利润，
stockMap = readFile('股票列表.xlsx')
iterStockList(stockMap)





