from index import *


# Input variables
vars_model = xlrd.open_workbook("inputVars.xlsx")
varsSheet = vars_model.sheets()[0]
varInputPrice1 = varsSheet.row_values(4)[5]
varInputPercent1 = varsSheet.row_values(5)[5]
varInputPercent2 = varsSheet.row_values(6)[5]
varInputPercent3 = varsSheet.row_values(7)[5]
varInputPercent5 = varsSheet.row_values(8)[5]
varInputPercent1Limit = varsSheet.row_values(9)[5]

varInputDayIH = varsSheet.row_values(10)[4]
varInputDayEH = varsSheet.row_values(11)[4]
varInputDayIL = varsSheet.row_values(12)[4]
varInputDayEL = varsSheet.row_values(13)[4]
varInputPercentIH = varsSheet.row_values(10)[5]
varInputPercentEH = varsSheet.row_values(11)[5]
varInputPercentIL = varsSheet.row_values(12)[5]
varInputPercentEL = varsSheet.row_values(13)[5]

varInputPercentDay = varsSheet.row_values(14)[5]
varInputPercentNt = varsSheet.row_values(15)[5]

varInputNumDays = varsSheet.row_values(16)[5]
varInputPercentDays = varsSheet.row_values(16)[5]
# varInputPercentHigh = varsSheet.row_values(10)[5]

# Input prices
data_model = xlrd.open_workbook("input.xlsx")
sheet = data_model.sheets()[0]

# Parsing data
print("staring parsing data")

# TODO improve the readFile function to be better at finding the right values and error handling
stockInfo = []
def readFile():
  for i in range(0, len(sheet.row_values(5))):
    if sheet.row_values(5)[i] != "":
      def findEnd():
        for j in range(7, len(sheet.col_values(i))):
          if sheet.col_values(i+1)[j] == "":
            return j
      starterOBJ = {
        "stockName": sheet.row_values(5)[i],
        "rowStart": 7,
        "rowEnd": findEnd(),
        "dateRead": i,
        "closeRead": i+1,
        "openRead": i+2,
        "highRead": i+3,
        "lowRead": i+4
      }
      stockInfo.append(starterOBJ)
      return

readFile()

for stock in stockInfo:
  # resetColWrite()
  setColWrite()

  # parse that data to get a list of workable values
  dataDate = []
  dataClose = []
  dataOpen = []
  dataHigh = []
  dataLow = []

  sheetParser(sheet,stock["rowStart"],stock["rowEnd"],stock["dateRead"], setColWrite(), "Date",dataDate, style1)
  sheetParser(sheet,stock["rowStart"],stock["rowEnd"],stock["closeRead"], setColWrite(), "Close",dataClose,style0)
  sheetParser(sheet,stock["rowStart"],stock["rowEnd"],stock["openRead"], setColWrite(), "Open",dataOpen,style0)
  sheetParser(sheet,stock["rowStart"],stock["rowEnd"],stock["highRead"], setColWrite(), "High",dataHigh,style0)
  sheetParser(sheet,stock["rowStart"],stock["rowEnd"],stock["lowRead"], setColWrite(), "Low",dataLow,style0)

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

  # TODO: Finish variable signals

  crossVarPrice(dataClose, varInputPrice1, setColWrite(), "close cross " + str(varInputPrice1))
  crossVarPercent(data1Rtn, varInputPercent1,setColWrite(), "1D X " + str(varInputPercent1))
  crossVarPercent(data2Rtn, varInputPercent2,setColWrite(), "2D X " + str(varInputPercent2))
  crossVarPercent(data3Rtn, varInputPercent3,setColWrite(), "3D X " + str(varInputPercent3))
  crossVarPercent(data5Rtn, varInputPercent5,setColWrite(), "5D X " + str(varInputPercent5))

  # NEED ONE DAY RETURN LIMIT

  highBtwDays(dataClose, int(varInputDayEH), setColWrite(), str(varInputDayEH) + " high close")
  highBtwDays(dataHigh, int(varInputDayIH), setColWrite(), str(varInputDayIH) + " high close")
  lowBtwDays(dataClose, int(varInputDayEL), setColWrite(), str(varInputDayEL) + " low close")
  lowBtwDays(dataLow, int(varInputDayIL), setColWrite(), str(varInputDayIL) + " low close")

  crossVarPercent(dataDayRtn, varInputPercentDay,setColWrite(), "Day X " + str(varInputPercentDay))
  crossVarPercent(dataNightRtn, varInputPercentNt,setColWrite(), "Nt X " + str(varInputPercentNt))

  # varDayRtn(data1Rtn, int(varInput14), setColWrite(), varInput14 + " day Rtn")



  # TODO: SELECT TESTS AND VALIDATE TRADING SIGNAL

  wb.save(str(stock["stockName"]) + ".xls")


wb.save("end_model.xls")
