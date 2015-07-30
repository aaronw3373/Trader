from index import *

# Input variables
vars_model = xlrd.open_workbook("inputVars.xlsx")
varsSheet = vars_model.sheets()[0]
varInputPrice1 = varsSheet.row_values(4)[5]
varInputPercent1 = varsSheet.row_values(5)[5]
varInputPercent2 = varsSheet.row_values(6)[5]
varInputPercent3 = varsSheet.row_values(7)[5]
varInputPercent5 = varsSheet.row_values(8)[5]
# varInputPercentHigh = varsSheet.row_values(10)[5]

# Input prices
data_model = xlrd.open_workbook("input.xlsx")
sheet = data_model.sheets()[0]

# Parsing data
print("staring parsing data")

# TODO:
# IMPORTANT!!! make this a function add add where to read from the file to work on all input
colStart = 7
colEnd = 3424
dateRead = 2
closeRead = 3
openRead = 4
highRead = 5
lowRead = 6

# parse that data to get a list of workable values
dataDate = []
dataClose = []
dataOpen = []
dataHigh = []
dataLow = []
sheetParser(sheet,colStart,colEnd,dateRead, setColWrite(), "Date",dataDate, style1)
sheetParser(sheet,colStart,colEnd,closeRead, setColWrite(), "Close",dataClose,style0)
sheetParser(sheet,colStart,colEnd,openRead, setColWrite(), "Open",dataOpen,style0)
sheetParser(sheet,colStart,colEnd,highRead, setColWrite(), "High",dataHigh,style0)
sheetParser(sheet,colStart,colEnd,lowRead, setColWrite(), "Low",dataLow,style0)

# Moving Average Calculator
print('finding averages...')
data200Avg = []
data100Avg = []
data50Avg = []
data30Avg = []
data10Avg = []
numDayAvg(dataClose, 200, setColWrite(), "200 Day Avg", style0, data200Avg)
numDayAvg(dataClose, 100, setColWrite(), "100 Day Avg", style0, data100Avg)
numDayAvg(dataClose, 50, setColWrite(), "50 Day Avg", style0, data50Avg)
numDayAvg(dataClose, 30, setColWrite(), "30 Day Avg", style0, data30Avg)
numDayAvg(dataClose, 10, setColWrite(), "10 Day Avg", style0, data10Avg)

# number of days returns
print("finding returns...")
data2Rtn = []
data3Rtn = []
data5Rtn = []
data1Rtn = []
numDayRtn(dataClose, 2, setColWrite(), "2 Day Return", data2Rtn)
numDayRtn(dataClose, 3, setColWrite(), "3 Day Return", data3Rtn)
numDayRtn(dataClose, 5, setColWrite(), "5 Day Return", data5Rtn)
numDayRtn(dataClose, 1, setColWrite(), "1 Day Return", data1Rtn)

# Overnight Return
dataNightRtn = []
overnightRtn(dataOpen,dataClose,setColWrite(), "Overnight Return", dataNightRtn)

# daytime Return
dataDayRtn = []
dayRtn(dataOpen, dataClose, setColWrite(), "Daytime Return", dataDayRtn)



#SIGNALS
setColWrite()
print("testing for signals...")


topLine(data10Avg, [data30Avg, data50Avg, data100Avg, data200Avg, dataClose], setColWrite(), "10 Day Top")
topLine(data30Avg, [data10Avg, data50Avg, data100Avg, data200Avg, dataClose], setColWrite(), "30 Day Top")
topLine(data50Avg, [data10Avg, data30Avg, data100Avg, data200Avg, dataClose], setColWrite(), "50 Day Top")
topLine(data100Avg, [data10Avg, data30Avg, data50Avg, data200Avg, dataClose], setColWrite(), "100 Day Top")
topLine(data200Avg, [data10Avg, data30Avg, data50Avg, data100Avg, dataClose], setColWrite(), "200 Day Top")
topLine(dataClose, [data10Avg, data30Avg, data50Avg, data100Avg, data200Avg], setColWrite(), "Close Top")


bottomLine(data10Avg, [data30Avg, data50Avg, data100Avg, data200Avg, dataClose], setColWrite(), "10 Day Bottom")
bottomLine(data30Avg, [data10Avg, data50Avg, data100Avg, data200Avg, dataClose], setColWrite(), "30 Day Bottom")
bottomLine(data50Avg, [data10Avg, data30Avg, data100Avg, data200Avg, dataClose], setColWrite(), "50 Day Bottom")
bottomLine(data100Avg, [data10Avg, data30Avg, data50Avg, data200Avg, dataClose], setColWrite(), "100 Day Bottom")
bottomLine(data200Avg, [data10Avg, data30Avg, data50Avg, data100Avg, dataClose], setColWrite(), "200 Day Bottom")
bottomLine(dataClose, [data10Avg, data30Avg, data50Avg, data100Avg, data200Avg], setColWrite(), "Close Bottom")


priceAbove(dataClose, data10Avg, setColWrite(), "price above 10")
priceAbove(dataClose, data30Avg, setColWrite(), "price above 30")
priceAbove(dataClose, data50Avg, setColWrite(), "price above 50")
priceAbove(dataClose, data100Avg, setColWrite(), "price above 100")
priceAbove(dataClose, data200Avg, setColWrite(), "price above 200")


crossAbove(data10Avg, data30Avg, setColWrite(), "10 above 30")
crossAbove(data10Avg, data50Avg, setColWrite(), "10 above 50")
crossAbove(data10Avg, data100Avg, setColWrite(), "10 above 100")
crossAbove(data10Avg, data200Avg, setColWrite(), "10 above 200")

crossAbove(data30Avg, data10Avg, setColWrite(), "30 above 10")
crossAbove(data30Avg, data50Avg, setColWrite(), "30 above 50")
crossAbove(data30Avg, data100Avg, setColWrite(), "30 above 100")
crossAbove(data30Avg, data200Avg, setColWrite(), "30 above 200")

crossAbove(data50Avg, data10Avg, setColWrite(), "50 above 10")
crossAbove(data50Avg, data30Avg, setColWrite(), "50 above 30")
crossAbove(data50Avg, data100Avg, setColWrite(), "50 above 100")
crossAbove(data50Avg, data200Avg, setColWrite(), "50 above 200")

crossAbove(data100Avg, data10Avg, setColWrite(), "100 above 10")
crossAbove(data100Avg, data30Avg, setColWrite(), "100 above 30")
crossAbove(data100Avg, data50Avg, setColWrite(), "100 above 50")
crossAbove(data100Avg, data200Avg, setColWrite(), "100 above 200")

crossAbove(data200Avg, data10Avg, setColWrite(), "200 above 10")
crossAbove(data200Avg, data30Avg, setColWrite(), "200 above 30")
crossAbove(data200Avg, data50Avg, setColWrite(), "200 above 50")
crossAbove(data200Avg, data100Avg, setColWrite(), "200 above 100")

crossBelow(data10Avg, data30Avg, setColWrite(), "10 below 30")
crossBelow(data10Avg, data50Avg, setColWrite(), "10 below 50")
crossBelow(data10Avg, data100Avg, setColWrite(), "10 below 100")
crossBelow(data10Avg, data200Avg, setColWrite(), "10 below 200")

crossBelow(data30Avg, data10Avg, setColWrite(), "30 below 10")
crossBelow(data30Avg, data50Avg, setColWrite(), "30 below 50")
crossBelow(data30Avg, data100Avg, setColWrite(), "30 below 100")
crossBelow(data30Avg, data200Avg, setColWrite(), "30 below 200")

crossBelow(data50Avg, data10Avg, setColWrite(), "50 below 10")
crossBelow(data50Avg, data30Avg, setColWrite(), "50 below 30")
crossBelow(data50Avg, data100Avg, setColWrite(), "50 below 100")
crossBelow(data50Avg, data200Avg, setColWrite(), "50 below 200")

crossBelow(data100Avg, data10Avg, setColWrite(), "100 below 10")
crossBelow(data100Avg, data30Avg, setColWrite(), "100 below 30")
crossBelow(data100Avg, data50Avg, setColWrite(), "100 below 50")
crossBelow(data100Avg, data200Avg, setColWrite(), "100 below 200")

crossBelow(data200Avg, data10Avg, setColWrite(), "200 below 10")
crossBelow(data200Avg, data30Avg, setColWrite(), "200 below 30")
crossBelow(data200Avg, data50Avg, setColWrite(), "200 below 50")
crossBelow(data200Avg, data100Avg, setColWrite(), "200 below 100")

# variable signals
crossVarPrice(dataClose, varInputPrice1, setColWrite(), "close cross " + str(varInputPrice1))
crossVarPercent(data1Rtn, varInputPercent1,setColWrite(), "1D X " + str(varInputPercent1))
crossVarPercent(data2Rtn, varInputPercent2,setColWrite(), "2D X " + str(varInputPercent2))
crossVarPercent(data3Rtn, varInputPercent3,setColWrite(), "3D X " + str(varInputPercent3))
crossVarPercent(data5Rtn, varInputPercent5,setColWrite(), "5D X " + str(varInputPercent5))


# TODO: Finish variable signals below


# crossVarPercent(dataDayRtn, float(VarInput0),setColWrite(), "Day X " + VarInput0)
# crossVarPercent(dataNightRtn, float(VarInput),setColWrite(), "Nt X " + VarInput)

# NEED ONE DAY RETURN LIMIT

# varInput10 = raw_input("Please enter days for max close: ")
# print "you entered", varInput10
# highBtwDays(dataClose, int(varInput10), setColWrite(), varInput10 + " high close")

# varInput11 = raw_input("Please enter days for max close: ")
# print "you entered", varInput11
# highBtwDays(dataHigh, int(varInput11), setColWrite(), varInput11 + " high close")

# varInput12 = raw_input("Please enter days for min close: ")
# print "you entered", varInput12
# lowBtwDays(dataClose, int(varInput12), setColWrite(), varInput12 + " low close")

# varInput13 = raw_input("Please enter days for min close: ")
# print "you entered", varInput13
# lowBtwDays(dataLow, int(varInput13), setColWrite(), varInput13 + " low close")

# varInput14 = raw_input("Please enter days for return: ")
# print "you entered", varInput14
# varDayRtn(data1Rtn, int(varInput14), setColWrite(), varInput14 + " day Rtn")



# TODO: SELECT TESTS AND VALIDATE TRADING SIGNAL



wb.save('end_model.xls')
