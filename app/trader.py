print("starting up...")
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
varInputNumDays = int(varsSheet.row_values(16)[4])
varInputPercentDays = varsSheet.row_values(16)[5]

# final test and validation input
# col0 - col9
# opperation (none, and, or), signal 1, signal 2, signal 3
col0opp = varsSheet.row_values(27)[2]
col1opp = varsSheet.row_values(28)[2]
col2opp = varsSheet.row_values(29)[2]
col3opp = varsSheet.row_values(30)[2]
col4opp = varsSheet.row_values(31)[2]
col5opp = varsSheet.row_values(32)[2]
col6opp = varsSheet.row_values(33)[2]
col7opp = varsSheet.row_values(34)[2]
col8opp = varsSheet.row_values(35)[2]
col9opp = varsSheet.row_values(36)[2]

col0sig1 = varsSheet.row_values(27)[3]
col1sig1 = varsSheet.row_values(28)[3]
col2sig1 = varsSheet.row_values(29)[3]
col3sig1 = varsSheet.row_values(30)[3]
col4sig1 = varsSheet.row_values(31)[3]
col5sig1 = varsSheet.row_values(32)[3]
col6sig1 = varsSheet.row_values(33)[3]
col7sig1 = varsSheet.row_values(34)[3]
col8sig1 = varsSheet.row_values(35)[3]
col9sig1 = varsSheet.row_values(36)[3]

col0sig2 = varsSheet.row_values(27)[4]
col1sig2 = varsSheet.row_values(28)[4]
col2sig2 = varsSheet.row_values(29)[4]
col3sig2 = varsSheet.row_values(30)[4]
col4sig2 = varsSheet.row_values(31)[4]
col5sig2 = varsSheet.row_values(32)[4]
col6sig2 = varsSheet.row_values(33)[4]
col7sig2 = varsSheet.row_values(34)[4]
col8sig2 = varsSheet.row_values(35)[4]
col9sig2 = varsSheet.row_values(36)[4]

col0sig3 = varsSheet.row_values(27)[5]
col1sig3 = varsSheet.row_values(28)[5]
col2sig3 = varsSheet.row_values(29)[5]
col3sig3 = varsSheet.row_values(30)[5]
col4sig3 = varsSheet.row_values(31)[5]
col5sig3 = varsSheet.row_values(32)[5]
col6sig3 = varsSheet.row_values(33)[5]
col7sig3 = varsSheet.row_values(34)[5]
col8sig3 = varsSheet.row_values(35)[5]
col9sig3 = varsSheet.row_values(36)[5]


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
      # return

def save_xls(list_dfs, xls_path):
  writer = pd.ExcelWriter(xls_path)
  for n, df in enumerate(list_dfs):
    df.to_excel(writer,'sheet%s' % n, engine="openpyxl")
  writer.save()

def strParserEval(input):
  words = input.split()
  res = "signals"
  for i in range(0, len(words)):
    res += '["' + words[i] + '"]'
  try:
    count = 0
    maximum = 2
    while type(res) == unicode or type(res) == str:
      res = eval(res)
      count +=1
      if count == maximum:
        break
  except:
    print "Unexpected error: trader.py lines 341-346"
    return "Error"
  else:
    return res

def makeCol(opp, sig1, sig2, sig3):
  array = []
  numTrue = 0
  numSigs = 0
  if sig1:
    numSigs +=1
    sig1 = strParserEval(sig1)
    if type(sig1) == str:
      return None
  if sig2:
    numSigs +=1
    sig2 = strParserEval(sig2)
    if type(sig2) == str:
      return None
  if sig3:
    numSigs +=1
    sig3 = strParserEval(sig3)
    if type(sig3) == str:
      return None

  if opp == "none":
    if numSigs == 1:
      for i in range(0, len(sig1)):
        if sig1[i] == 1:
          array.append(1)
          numTrue += 1
        else:
          array.append(0)
    else:
      print("invalid sigs: none", numSigs)
      return None
  elif opp == "and":
    if numSigs == 2:
      for i in range(0, len(sig1)):
        if sig1[i] + sig2[i] == 2:
          array.append(1)
          numTrue += 1
        else:
          array.append(0)
    elif numSigs == 3:
      for i in range(0, len(sig1)):
        if sig1[i] + sig2[i] + sig3[i]== 3:
          array.append(1)
          numTrue += 1
        else:
          array.append(0)
    else:
      print("invalid sigs: and", numSigs)
      return None
  elif opp == "or":
    if numSigs == 2:
      for i in range(0, len(sig1)):
        if sig1[i] + sig2[i] >= 1:
          array.append(1)
          numTrue += 1
        else:
          array.append(0)
    elif numSigs == 3:
      for i in range(0, len(sig1)):
        if sig1[i] + sig2[i] + sig3[i] >= 1:
          array.append(1)
          numTrue += 1
        else:
          array.append(0)
    else:
      print("invalid sigs: or", numSigs)
      return None
  else:
    print("invalid opperation")
    return None
  result = pd.Series(array, name=setColName())
  # print numTrue
  return result

def findNumCol():
  numCol = 0
  for i in range(27, 37):
    if varsSheet.row_values(i)[3]:
      numCol += 1
  return numCol

def canMakeCol(colNum):
  if varsSheet.row_values(27 + colNum)[3]:
    return eval(makeColsArr[colNum])
  else:
    return 0


print("reading file...")
# read_time = time.time()
stockInfo = []
readFile()
# print("Read in time was %g seconds" % (time.time() - read_time))

print("starting calculations...")
for stock in stockInfo:
  start_time2 = time.time()
  resetColName()
#
  # Parse the data into a new DataFrame
  # 0-4 date, close, open, high, low
  df = sheetParser(inputDF,stock)

  # Get moving averages over x numver of days
  # 5-9
  df = pd.concat([df, numDayAvg(df[1], 200)],axis=1)
  df = pd.concat([df, numDayAvg(df[1], 100)],axis=1)
  df = pd.concat([df, numDayAvg(df[1], 50)],axis=1)
  df = pd.concat([df, numDayAvg(df[1], 30)],axis=1)
  df = pd.concat([df, numDayAvg(df[1], 10)],axis=1)

  # Get percent return over number of days
  # 10 -15
  df = pd.concat([df, numDayRtn(df[1], 2)],axis=1)
  df = pd.concat([df, numDayRtn(df[1], 3)],axis=1)
  df = pd.concat([df, numDayRtn(df[1], 5)],axis=1)
  df = pd.concat([df, nightRtn(df[1], df[2])],axis=1)
  df = pd.concat([df, dayRtn(df[2], df[1])],axis=1)
  df = pd.concat([df, numDayRtn(df[1], 1)],axis=1)

  # stringify all the signal definitions to be called when needed
  signals = {
    "topLine": {
      "close": "topLine(df[1], [df[5],df[6],df[7],df[8],df[9]])",
      "200": "topLine(df[5], [df[1],df[6],df[7],df[8],df[9]])",
      "100": "topLine(df[6], [df[1],df[5],df[7],df[8],df[9]])",
      "50": "topLine(df[7], [df[1],df[5],df[6],df[8],df[9]])",
      "30": "topLine(df[8], [df[1],df[5],df[6],df[7],df[9]])",
      "10": "topLine(df[9], [df[1],df[5],df[6],df[7],df[8]])"
    },
    "bottomLine":{
      "close": "bottomLine(df[1], [df[5],df[6],df[7],df[8],df[9]])",
      "200": "bottomLine(df[5], [df[1],df[6],df[7],df[8],df[9]])",
      "100": "bottomLine(df[6], [df[1],df[5],df[7],df[8],df[9]])",
      "50": "bottomLine(df[7], [df[1],df[5],df[6],df[8],df[9]])",
      "30": "bottomLine(df[8], [df[1],df[5],df[6],df[7],df[9]])",
      "10": "bottomLine(df[9], [df[1],df[5],df[6],df[7],df[8]])"
    },
    "priceAbove":{
      "close": {
        "200": "priceAbove(df[1], df[5])",
        "100": "priceAbove(df[1], df[6])",
        "50": "priceAbove(df[1], df[7])",
        "30": "priceAbove(df[1], df[8])",
        "10": "priceAbove(df[1], df[9])"
      },
      "10": {
        "close": "priceAbove(df[9], df[1])",
        "30": "priceAbove(df[9], df[8])",
        "50": "priceAbove(df[9], df[7])",
        "100": "priceAbove(df[9], df[6])",
        "200": "priceAbove(df[9], df[5])"
      },
      "30": {
        "close": "priceAbove(df[8], df[1])",
        "10": "priceAbove(df[8], df[9])",
        "50": "priceAbove(df[8], df[7])",
        "100": "priceAbove(df[8], df[6])",
        "200": "priceAbove(df[8], df[5])"
      },
      "50": {
        "close": "priceAbove(df[7], df[1])",
        "10": "priceAbove(df[7], df[9])",
        "30": "priceAbove(df[7], df[8])",
        "100": "priceAbove(df[7], df[6])",
        "200": "priceAbove(df[7], df[5])"
      },
      "100": {
        "close": "priceAbove(df[6], df[1])",
        "10": "priceAbove(df[6], df[9])",
        "30": "priceAbove(df[6], df[8])",
        "50": "priceAbove(df[6], df[7])",
        "200": "priceAbove(df[6], df[5])"
      },
      "200": {
        "close": "priceAbove(df[5], df[1])",
        "10": "priceAbove(df[5], df[9])",
        "30": "priceAbove(df[5], df[8])",
        "50": "priceAbove(df[5], df[7])",
        "100": "priceAbove(df[5], df[6])"
      }
    },
    "priceBelow":{
      "close": {
        "200": "priceBelow(df[1], df[5])",
        "100": "priceBelow(df[1], df[6])",
        "50": "priceBelow(df[1], df[7])",
        "30": "priceBelow(df[1], df[8])",
        "10": "priceBelow(df[1], df[9])"
      },
      "10": {
        "close": "priceBelow(df[9], df[1])",
        "30": "priceBelow(df[9], df[8])",
        "50": "priceBelow(df[9], df[7])",
        "100": "priceBelow(df[9], df[6])",
        "200": "priceBelow(df[9], df[5])"
      },
      "30": {
        "close": "priceBelow(df[8], df[1])",
        "10": "priceBelow(df[8], df[9])",
        "50": "priceBelow(df[8], df[7])",
        "100": "priceBelow(df[8], df[6])",
        "200": "priceBelow(df[8], df[5])"
      },
      "50": {
        "close": "priceBelow(df[7], df[1])",
        "10": "priceBelow(df[7], df[9])",
        "30": "priceBelow(df[7], df[8])",
        "100": "priceBelow(df[7], df[6])",
        "200": "priceBelow(df[7], df[5])"
      },
      "100": {
        "close": "priceBelow(df[6], df[1])",
        "10": "priceBelow(df[6], df[9])",
        "30": "priceBelow(df[6], df[8])",
        "50": "priceBelow(df[6], df[7])",
        "200": "priceBelow(df[6], df[5])"
      },
      "200": {
        "close": "priceBelow(df[5], df[1])",
        "10": "priceBelow(df[5], df[9])",
        "30": "priceBelow(df[5], df[8])",
        "50": "priceBelow(df[5], df[7])",
        "100": "priceBelow(df[5], df[6])"
      }
    },
    "crossAbove":{
      "10": {
        "30": "crossAbove(df[9], df[8])",
        "50": "crossAbove(df[9], df[7])",
        "100": "crossAbove(df[9], df[6])",
        "200": "crossAbove(df[9], df[5])"
      },
      "30": {
        "10": "crossAbove(df[8], df[9])",
        "50": "crossAbove(df[8], df[7])",
        "100": "crossAbove(df[8], df[6])",
        "200": "crossAbove(df[8], df[5])"
      },
      "50": {
        "10": "crossAbove(df[7], df[9])",
        "30": "crossAbove(df[7], df[8])",
        "100": "crossAbove(df[7], df[6])",
        "200": "crossAbove(df[7], df[5])"
      },
      "100": {
        "10": "crossAbove(df[6], df[9])",
        "30": "crossAbove(df[6], df[8])",
        "50": "crossAbove(df[6], df[7])",
        "200": "crossAbove(df[6], df[5])"
      },
      "200": {
        "10": "crossAbove(df[5], df[9])",
        "30": "crossAbove(df[5], df[8])",
        "50": "crossAbove(df[5], df[7])",
        "100": "crossAbove(df[5], df[6])"
      }
    },
    "crossBelow":{
      "10": {
        "30": "crossBelow(df[9], df[8])",
        "50": "crossBelow(df[9], df[7])",
        "100": "crossBelow(df[9], df[6])",
        "200": "crossBelow(df[9], df[5])"
      },
      "30": {
        "10": "crossBelow(df[8], df[9])",
        "50": "crossBelow(df[8], df[7])",
        "100": "crossBelow(df[8], df[6])",
        "200": "crossBelow(df[8], df[5])"
      },
      "50": {
        "10": "crossBelow(df[7], df[9])",
        "30": "crossBelow(df[7], df[8])",
        "100": "crossBelow(df[7], df[6])",
        "200": "crossBelow(df[7], df[5])"
      },
      "100": {
        "10": "crossBelow(df[6], df[9])",
        "30": "crossBelow(df[6], df[8])",
        "50": "crossBelow(df[6], df[7])",
        "200": "crossBelow(df[6], df[5])"
      },
      "200": {
        "10": "crossBelow(df[5], df[9])",
        "30": "crossBelow(df[5], df[8])",
        "50": "crossBelow(df[5], df[7])",
        "100": "crossBelow(df[5], df[6])"
      }
    },
    "variable":{
      "crossPrice": "crossVarPrice(df[1], varInputPrice1)",
      "crossPercent": {
        "2": "crossVarPercent(df[10], varInputPercent2)",
        "3": "crossVarPercent(df[11], varInputPercent3)",
        "5": "crossVarPercent(df[12], varInputPercent5)",
        "1": "crossVarPercent(df[15], varInputPercent1)",
        "day": "crossVarPercent(df[14], varInputPercentDay)",
        "night": "crossVarPercent(df[13], varInputPercentNt)"
      },
      "high": {
        "interday": "highBtwIDays(df[1], df[3], varInputDayIH, varInputPercentIH)",
        "endDay": "highBtwEDays(df[4], df[3], varInputDayEH, varInputPercentEH)"
      },
      "low": {
        "interday": "lowBtwIDays(df[1], df[3], varInputDayIL, varInputPercentIL)",
        "endDay": "lowBtwEDays(df[4], df[3], varInputDayEL, varInputPercentEL)"
      },
      "returnLimit": "varRtnLimit(df[3], df[4],varInputPercent1Limit)",
      "dayReturn": "varDayRtn(df[1], varInputNumDays,varInputPercentDays)"
    }
  }

  # Section 3
  # Asign Columns
  makeColsArr = ["makeCol(col0opp,col0sig1,col0sig2,col0sig3)",
  "makeCol(col1opp,col1sig1,col1sig2,col1sig3)",
  "makeCol(col2opp,col2sig1,col2sig2,col2sig3)",
  "makeCol(col3opp,col3sig1,col3sig2,col3sig3)",
  "makeCol(col4opp,col4sig1,col4sig2,col4sig3)",
  "makeCol(col5opp,col5sig1,col5sig2,col5sig3)",
  "makeCol(col6opp,col6sig1,col6sig2,col6sig3)",
  "makeCol(col7opp,col7sig1,col7sig2,col7sig3)",
  "makeCol(col8opp,col8sig1,col8sig2,col8sig3)",
  "makeCol(col9opp,col9sig1,col9sig2,col9sig3)"]

  col0 = canMakeCol(0)
  col1 = canMakeCol(1)
  col2 = canMakeCol(2)
  col3 = canMakeCol(3)
  col4 = canMakeCol(4)
  col5 = canMakeCol(5)
  col6 = canMakeCol(6)
  col7 = canMakeCol(6)
  col8 = canMakeCol(7)
  col9 = canMakeCol(8)

  numCol = findNumCol()

  # print col0






  # run final tests


  print(str(stock["stockName"]) + " %g seconds" % (time.time() - start_time2))

  # save_xls([df],"end_model.xlsx")

print("Total Elapsed time was %g seconds" % (time.time() - start_time))
# save sheet
# save_xls([], results.xlsx)
