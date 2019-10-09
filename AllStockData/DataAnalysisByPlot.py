import matplotlib.pyplot as plt
import xlrd

def readData(fileName,sheetIndexArr,colArr):
    file = xlrd.open_workbook(fileName)
    totalDataList = []
    for i in sheetIndexArr:
        ItemDataList = []
        sheet = file.sheet_by_index(i)
        for j in colArr:
            result = sheet.col_values(j)
            ItemDataList.append(result)
        totalDataList.append(ItemDataList)
    return totalDataList

def drawTrending(totalDataList,item,timeRange,companyList,legendList):
    plt.rcParams['font.sans-serif'] = ['SimHei'] # 用来正常显示中文
    plt.rcParams['axes.unicode_minus'] = False   # 用来正常显示负号
    fig, ax = plt.subplots()
    timeRange = timeRange + 1               # 考虑标题的影响
    title = totalDataList[0][item][0]
    for i in companyList:
        newTimeRange = timeRange
        if len(totalDataList[i][item])<timeRange:
            newTimeRange = len(totalDataList[i][item])
        drawData = totalDataList[i][item][1:newTimeRange]
        for j in range(0,len(drawData)):
            if '\xa0'==drawData[j]:     # 排除空值的情况
                drawData[j] = 0
                continue
            drawData[j] = drawData[j].replace(',','')
            drawData[j] = float(drawData[j])
        timeData = totalDataList[0][0][1:newTimeRange]
        timeData.reverse()
        drawData.reverse()                  # 保证新数据绘制在最后
        ax.plot(range(1,newTimeRange), drawData)
    plt.xticks(range(1,newTimeRange), timeData)     # 修改x坐标显示值
    ax.set_title(title)
    ax.set_xlabel('time (Quarter)')
    plt.grid(True, linestyle='-.')
    plt.legend(legendList)
    plt.gcf().autofmt_xdate()           # 自动旋转日期标记
    plt.show()


# '0: 时间','1: 现金及现金等价物净增加额','2: 期末现金及现金等价物余额'
# '3: 每股收益', '4: 主营业务增长率', '5: 净利润增长率', '6: 净利润', '7: 营业总收入'
# '8: 总资产','9: 总负债'
# categoryName1 = '白酒财务数据.xls'
# legendList = ['洋河','茅台','五粮液','古井贡酒','今世缘','泸州老窖']

categoryName1 = '持股数据.xls'
legendList = ['招商银行','海康威视','大华股份','美的集团','格力电器','长江电力','新华保险','中国平安','洋河']
totalDataList = readData(categoryName1,range(0,len(legendList)),[0,4,5,6,7,8,9,10,13,16])
drawTrending(totalDataList,6,30,[1,3,5,6,7,8],['海康威视','美的集团','长江电力','新华保险','中国平安','洋河'])