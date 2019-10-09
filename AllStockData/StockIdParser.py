import requests
from lxml import html
import xlwt

def saveSheet(workbook, data, sheetName):
    sheet = workbook.add_sheet(sheetName)
    i = 0
    for key in data.keys():
        j = 0
        sheet.write(j, i, key)
        for val in data[key]:
            sheet.write(j+1, i, data[key][j])
            j = j + 1
        i = i + 1

def parseDividendData(stockid):
    url = 'http://vip.stock.finance.sina.com.cn/corp/go.php/vISSUE_ShareBonus/stockid/' + stockid + '.phtml'
    page = requests.Session().get(url)
    htmlPage = html.fromstring(page.text)
    dividendTitle = ['公告日', '送股', '增股', '每十股分红']
    data = {}
    for i in range(0, 4):
        dividendItem = '//div[@id="con02-0"]//table//tbody//tr//td[' + str(i + 1) + ']/text()'
        dividendItemVal = htmlPage.xpath(dividendItem)
        if i == 1:
            if len(dividendItemVal) != len(data[dividendTitle[i-1]]):
                data[dividendTitle[i - 1]] = data[dividendTitle[i-1]][0:len(dividendItemVal)]
        data[dividendTitle[i]] = dividendItemVal
    return data