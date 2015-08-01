import time
start_time = time.time()
from index import *
import xlrd
import sys

# Input variables
vars_model = xlrd.open_workbook(sys.argv[2])
varsSheet = vars_model.sheets()[0]
varInputPrice1 = varsSheet.row_values(4)[5]
varInputPercent1 = varsSheet.row_values(5)[5]
varInputPercent2 = varsSheet.row_values(6)[5]
varInputPercent3 = varsSheet.row_values(7)[5]
varInputPercent5 = varsSheet.row_values(8)[5]
varInputPercent1Limit = varsSheet.row_values(9)[5]
varInputDayIH = int(varsSheet.row_values(10)[4])
varInputDayEH = int(varsSheet.row_values(11)[4])
varInputDayIL = int(varsSheet.row_values(12)[4])
varInputDayEL = int(varsSheet.row_values(13)[4])
varInputPercentIH = varsSheet.row_values(10)[5]
varInputPercentEH = varsSheet.row_values(11)[5]
varInputPercentIL = varsSheet.row_values(12)[5]
varInputPercentEL = varsSheet.row_values(13)[5]
varInputPercentDay = varsSheet.row_values(14)[5]
varInputPercentNt = varsSheet.row_values(15)[5]
# varInputNumDays = int(varsSheet.row_values(16)[4])
varInputPercentDays = varsSheet.row_values(16)[5]

# GET INPUT FILE
inputDF = pd.read_excel(sys.argv[1])

# read the input dataframe and create objects of stocks
# TODO: improve error handling and rowstart awareness
def findEnd(i, inputDF):
  for j in range(5, len(inputDF)):
    if pd.isnull(inputDF.iloc[j,i+1]):
      return j

def readFile():
  for i in range(0, len(inputDF.columns)):
    if pd.notnull(inputDF.iloc[3,i]):
      stockOBJ = {
        "stockName": inputDF.iloc[3,i],
        "rowStart": 5,
        "rowEnd": findEnd(i, inputDF),
        "dateRead": i,
        "closeRead": i+1,
        "openRead": i+2,
        "highRead": i+3,
        "lowRead": i+4
      }
      stockInfo.append(stockOBJ)
      # take out this return to do the whole set of stocks not just the first
      return

def save_xls(list_dfs, xls_path):
    writer = pd.ExcelWriter(xls_path)
    for n, df in enumerate(list_dfs):
        df.to_excel(writer,'sheet%s' % n, engine="openpyxl")
    writer.save()


print("reading file into objects...")
stockInfo = []
readFile()
print("Read in time was %g seconds" % (time.time() - start_time))

print("starting calculations...")
for stock in stockInfo:
  start_time2 = time.time()
  resetColName()

  # Parse the data into a new DataFrame
  # 0-4 date, close, open, high, low
  df = sheetParser(inputDF,stock)

  # Get moving averages over x numver of days
  # 5-9
  df =  pd.concat([df, numDayAvg(df[1], 200)],axis=1)
  df =  pd.concat([df, numDayAvg(df[1], 100)],axis=1)
  df =  pd.concat([df, numDayAvg(df[1], 50)],axis=1)
  df =  pd.concat([df, numDayAvg(df[1], 30)],axis=1)
  df =  pd.concat([df, numDayAvg(df[1], 10)],axis=1)

  # Get percent return over number of days
  # 10 -15
  df = pd.concat([df, numDayRtn(df[1], 2)],axis=1)
  df = pd.concat([df, numDayRtn(df[1], 3)],axis=1)
  df = pd.concat([df, numDayRtn(df[1], 5)],axis=1)
  df = pd.concat([df, nightRtn(df[1], df[2])],axis=1)
  df = pd.concat([df, dayRtn(df[2], df[1])],axis=1)
  df = pd.concat([df, numDayRtn(df[1], 1)],axis=1)

  # Start signals
  # top line
  # 16-21
  df = pd.concat([df,topLine(df[1], [df[5],df[6],df[7],df[8],df[9]])], axis=1)
  df = pd.concat([df,topLine(df[5], [df[1],df[6],df[7],df[8],df[9]])], axis=1)
  df = pd.concat([df,topLine(df[6], [df[1],df[5],df[7],df[8],df[9]])], axis=1)
  df = pd.concat([df,topLine(df[7], [df[1],df[5],df[6],df[8],df[9]])], axis=1)
  df = pd.concat([df,topLine(df[8], [df[1],df[5],df[6],df[7],df[9]])], axis=1)
  df = pd.concat([df,topLine(df[9], [df[1],df[5],df[6],df[7],df[8]])], axis=1)

  # bottom line
  # 22-27
  df = pd.concat([df,bottomLine(df[1], [df[5],df[6],df[7],df[8],df[9]])], axis=1)
  df = pd.concat([df,bottomLine(df[5], [df[1],df[6],df[7],df[8],df[9]])], axis=1)
  df = pd.concat([df,bottomLine(df[6], [df[1],df[5],df[7],df[8],df[9]])], axis=1)
  df = pd.concat([df,bottomLine(df[7], [df[1],df[5],df[6],df[8],df[9]])], axis=1)
  df = pd.concat([df,bottomLine(df[8], [df[1],df[5],df[6],df[7],df[9]])], axis=1)
  df = pd.concat([df,bottomLine(df[9], [df[1],df[5],df[6],df[7],df[8]])], axis=1)

  # price above
  # 28-32
  df = pd.concat([df, priceAbove(df[1], df[5])],axis=1)
  df = pd.concat([df, priceAbove(df[1], df[6])],axis=1)
  df = pd.concat([df, priceAbove(df[1], df[7])],axis=1)
  df = pd.concat([df, priceAbove(df[1], df[8])],axis=1)
  df = pd.concat([df, priceAbove(df[1], df[9])],axis=1)

  # cross above
  # 33-36
  df = pd.concat([df, crossAbove(df[9], df[8])],axis=1)
  df = pd.concat([df, crossAbove(df[9], df[7])],axis=1)
  df = pd.concat([df, crossAbove(df[9], df[6])],axis=1)
  df = pd.concat([df, crossAbove(df[9], df[5])],axis=1)
  # 37-40
  df = pd.concat([df, crossAbove(df[8], df[9])],axis=1)
  df = pd.concat([df, crossAbove(df[8], df[7])],axis=1)
  df = pd.concat([df, crossAbove(df[8], df[6])],axis=1)
  df = pd.concat([df, crossAbove(df[8], df[5])],axis=1)
  # 41-44
  df = pd.concat([df, crossAbove(df[7], df[9])],axis=1)
  df = pd.concat([df, crossAbove(df[7], df[8])],axis=1)
  df = pd.concat([df, crossAbove(df[7], df[6])],axis=1)
  df = pd.concat([df, crossAbove(df[7], df[5])],axis=1)
  # 45-48
  df = pd.concat([df, crossAbove(df[6], df[9])],axis=1)
  df = pd.concat([df, crossAbove(df[6], df[8])],axis=1)
  df = pd.concat([df, crossAbove(df[6], df[7])],axis=1)
  df = pd.concat([df, crossAbove(df[6], df[5])],axis=1)
  # 49-52
  df = pd.concat([df, crossAbove(df[5], df[9])],axis=1)
  df = pd.concat([df, crossAbove(df[5], df[8])],axis=1)
  df = pd.concat([df, crossAbove(df[5], df[7])],axis=1)
  df = pd.concat([df, crossAbove(df[5], df[6])],axis=1)

  # cross below
  # 53-56
  df = pd.concat([df, crossBelow(df[9], df[8])],axis=1)
  df = pd.concat([df, crossBelow(df[9], df[7])],axis=1)
  df = pd.concat([df, crossBelow(df[9], df[6])],axis=1)
  df = pd.concat([df, crossBelow(df[9], df[5])],axis=1)
  # 57-60
  df = pd.concat([df, crossBelow(df[8], df[9])],axis=1)
  df = pd.concat([df, crossBelow(df[8], df[7])],axis=1)
  df = pd.concat([df, crossBelow(df[8], df[6])],axis=1)
  df = pd.concat([df, crossBelow(df[8], df[5])],axis=1)
  # 61-64
  df = pd.concat([df, crossBelow(df[7], df[9])],axis=1)
  df = pd.concat([df, crossBelow(df[7], df[8])],axis=1)
  df = pd.concat([df, crossBelow(df[7], df[6])],axis=1)
  df = pd.concat([df, crossBelow(df[7], df[5])],axis=1)
  # 65-68
  df = pd.concat([df, crossBelow(df[6], df[9])],axis=1)
  df = pd.concat([df, crossBelow(df[6], df[8])],axis=1)
  df = pd.concat([df, crossBelow(df[6], df[7])],axis=1)
  df = pd.concat([df, crossBelow(df[6], df[5])],axis=1)
  # 69-72
  df = pd.concat([df, crossBelow(df[5], df[9])],axis=1)
  df = pd.concat([df, crossBelow(df[5], df[8])],axis=1)
  df = pd.concat([df, crossBelow(df[5], df[7])],axis=1)
  df = pd.concat([df, crossBelow(df[5], df[6])],axis=1)

  # variable signals
  # 73-77
  df = pd.concat([df, crossVarPrice(df[1], varInputPrice1)],axis=1)
  df = pd.concat([df, crossVarPercent(df[10], varInputPercent2)],axis=1)
  df = pd.concat([df, crossVarPercent(df[11], varInputPercent3)],axis=1)
  df = pd.concat([df, crossVarPercent(df[12], varInputPercent5)],axis=1)
  # NEED ONE DAY RETURN LIMIT
  df = pd.concat([df, crossVarPercent(df[15], varInputPercent1)],axis=1)
  # 78-81
  df = pd.concat([df, highBtwDays(df[1], varInputDayEH, varInputPercentEH)],axis=1)
  df = pd.concat([df, highBtwDays(df[3], varInputDayIH, varInputPercentIH)],axis=1)
  df = pd.concat([df, lowBtwDays(df[1], varInputDayEL, varInputPercentEL)],axis=1)
  df = pd.concat([df, lowBtwDays(df[4], varInputDayIL, varInputPercentIL)],axis=1)
  # 82-83
  df = pd.concat([df, crossVarPercent(df[14], varInputPercentDay)],axis=1)
  df = pd.concat([df, crossVarPercent(df[13], varInputPercentNt)],axis=1)

  # varDayRtn(data1Rtn, int(varInput14))



  # print df.head(35)
  print(str(stock["stockName"]) + " %g seconds" % (time.time() - start_time2))

  # save_xls([df],"end_model.xlsx")

print("Total Elapsed time was %g seconds" % (time.time() - start_time))
# save sheet
