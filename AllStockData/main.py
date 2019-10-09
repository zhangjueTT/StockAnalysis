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

def getUrl(stockid):
    Titles = ['时间','经营活动产生的现金流量净额', '投资活动产生的现金流量净额', '筹资活动产生的现金流量净额', '现金及现金等价物净增加额',
              '期末现金及现金等价物余额','每股收益','主营业务增长率','净利润增长率','净利润','营业总收入','流动资产',
              '非流动资产','总资产','流动负债','非流动负债','总负债']
    url = []
    url.append('http://money.finance.sina.com.cn/corp/view/vFD_FinanceSummaryHistory.php?stockid='
               +stockid+ '&typecode=MANANETR&cate=xjll0')       # 经营活动产生的现金流量净额
    url.append('http://money.finance.sina.com.cn/corp/view/vFD_FinanceSummaryHistory.php?stockid='
               +stockid+ '&typecode=INVNETCASHFLOW&cate=xjll0') # 投资活动产生的现金流量净额
    url.append('http://money.finance.sina.com.cn/corp/view/vFD_FinanceSummaryHistory.php?stockid='
               +stockid+ '&typecode=FINNETCFLOW&cate=xjll00')   # 筹资活动产生的现金流量净额
    url.append('http://money.finance.sina.com.cn/corp/view/vFD_FinanceSummaryHistory.php?stockid='
               +stockid+ '&typecode=CASHNETR&cate=xjll0')       # 现金及现金等价物净增加额
    url.append('http://money.finance.sina.com.cn/corp/view/vFD_FinanceSummaryHistory.php?stockid='
               +stockid+ '&typecode=FINALCASHBALA&cate=xjll0') # 期末现金及现金等价物余额
    url.append('http://money.finance.sina.com.cn/corp/view/vFD_FinancialGuideLineHistory.php?stockid='
               +stockid+ '&typecode=financialratios61')         # 每股收益
    url.append('http://money.finance.sina.com.cn/corp/view/vFD_FinancialGuideLineHistory.php?stockid='
               +stockid+ '&typecode=financialratios43')         # 主营业务增长率
    url.append('http://money.finance.sina.com.cn/corp/view/vFD_FinancialGuideLineHistory.php?stockid='
               +stockid+ '&typecode=financialratios44')         # 净利润增长率
    url.append('http://money.finance.sina.com.cn/corp/view/vFD_FinanceSummaryHistory.php?stockid='
               +stockid+ '&type=NETPROFIT&cate=liru0')          # 净利润
    url.append('http://money.finance.sina.com.cn/corp/view/vFD_FinanceSummaryHistory.php?stockid='
               +stockid+ '&type=BIZTOTINCO&cate=liru0')         # 营业总收入
    url.append('http://money.finance.sina.com.cn/corp/view/vFD_FinanceSummaryHistory.php?stockid='
               +stockid+ '&type=TOTCURRASSET&cate=zcfz0')       # 流动资产
    url.append('http://money.finance.sina.com.cn/corp/view/vFD_FinanceSummaryHistory.php?stockid='
               +stockid+ '&type=TOTALNONCASSETS&cate=zcfz0')    # 非流动资产
    url.append('http://money.finance.sina.com.cn/corp/view/vFD_FinanceSummaryHistory.php?stockid='
               +stockid+ '&type=TOTASSET&cate=zcfz0')           # 总资产
    url.append('http://money.finance.sina.com.cn/corp/view/vFD_FinanceSummaryHistory.php?stockid='
               +stockid+ '&type=TOTALCURRLIAB&cate=zcfz0')      # 流动负债
    url.append('http://money.finance.sina.com.cn/corp/view/vFD_FinanceSummaryHistory.php?stockid='
               +stockid+ '&type=TOTALNONCLIAB&cate=zcfz0')     #非流动负债
    url.append('http://money.finance.sina.com.cn/corp/view/vFD_FinanceSummaryHistory.php?stockid='
               +stockid+ '&type=TOTALNONCLIAB&cate=zcfz0')     #总负债
    return url,Titles

def parseTotalData(stockid):
    url,Titles = getUrl(stockid)
    data = {}
    timeReg = '//div[@id="con02-2"]//table//tbody//tr//td[1]/text()'
    itemReg = '//div[@id="con02-2"]//table//tbody//tr//td[2]/text()'
    page = requests.Session().get(url[0])
    htmlPage = html.fromstring(page.text)
    timeRegVal = htmlPage.xpath(timeReg)
    data[Titles[0]] = timeRegVal
    for i in range(0,len(url)):
        page = requests.Session().get(url[i])
        htmlPage = html.fromstring(page.text)
        titleNameVal = htmlPage.xpath(itemReg)
        data[Titles[i+1]] = titleNameVal
    return data

def categoryDividendData(categoryName, stockidArray, stockNameArray):
    workbook = xlwt.Workbook(encoding='utf-8')
    for i in range(0,len(stockidArray)):
        data = parseDividendData(stockidArray[i])
        saveSheet(workbook, data, stockNameArray[i])
    workbook.save(categoryName + '.xls')

def categoryTotalData(categoryName, stockidArray, stockNameArray):
    workbook = xlwt.Workbook(encoding='utf-8')
    for i in range(0,len(stockidArray)):
        data = parseTotalData(stockidArray[i])
        saveSheet(workbook, data, stockNameArray[i])
    workbook.save(categoryName + '.xls')

# stockidArray = ['002304','600519','000858','000596','603369','000568']
# stockNameArray = ['洋河','茅台','五粮液','古井贡酒','今世缘','泸州老窖']
# categoryName1 = '白酒财务数据'
# categoryName2 = '白酒分红数据'

stockidArray = ['600036','002415','002236','000333','000651','002027','600900','601336','601318','002304']
stockNameArray = ['招商银行','海康威视','大华股份','美的集团','格力电器','分众传媒','长江电力','新华保险','中国平安','洋河']
categoryName1 = '持股数据'
categoryName2 = '持股分红数据'
categoryTotalData(categoryName1, stockidArray, stockNameArray)
categoryDividendData(categoryName2, stockidArray, stockNameArray)




