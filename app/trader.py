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
varInputNumDays = int(varsSheet.row_values(16)[5])
varInputPercentDays = varsSheet.row_values(16)[5]

# GET INPUT FILE
inputDF = pd.read_excel(sys.argv[1])

# read the input dataframe and create objects of stocks
# TODO: improve error handling and rowstart awareness
def readFile():
  for i in range(0, len(inputDF.columns)):
    if pd.notnull(inputDF.iloc[3,i]):
      def findEnd():
        for j in range(5, len(inputDF)):
          if pd.isnull(inputDF.iloc[j,i+1]):
            return j
      stockOBJ = {
        "stockName": inputDF.iloc[3,i],
        "rowStart": 5,
        "rowEnd": findEnd(),
        "dateRead": i,
        "closeRead": i+1,
        "openRead": i+2,
        "highRead": i+3,
        "lowRead": i+4
      }
      stockInfo.append(stockOBJ)
      return

stockInfo = []
readFile()

for stock in stockInfo:
  # Parse the data into a new DataFrame
  df = sheetParser(inputDF,stock)

  # Get moving averages over x numver of days
  df =  pd.concat([df, numDayAvg(df[1], 200)],axis=1)
  df =  pd.concat([df, numDayAvg(df[1], 100)],axis=1)
  df =  pd.concat([df, numDayAvg(df[1], 50)],axis=1)
  df =  pd.concat([df, numDayAvg(df[1], 30)],axis=1)
  df =  pd.concat([df, numDayAvg(df[1], 10)],axis=1)


  # Get percent return over number of days
  df = pd.concat([df, numDayRtn(df[1], 2)],axis=1)
  df = pd.concat([df, numDayRtn(df[1], 3)],axis=1)
  df = pd.concat([df, numDayRtn(df[1], 5)],axis=1)
  df = pd.concat([df, nightRtn(df[1], df[2])],axis=1)
  df = pd.concat([df, dayRtn(df[2], df[1])],axis=1)
  df = pd.concat([df, numDayRtn(df[1], 1)],axis=1)


  # Start signals

  print df.head(15)
