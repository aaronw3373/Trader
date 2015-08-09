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

# final columns input
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

col0sig4 = varsSheet.row_values(27)[6]
col1sig4 = varsSheet.row_values(28)[6]
col2sig4 = varsSheet.row_values(29)[6]
col3sig4 = varsSheet.row_values(30)[6]
col4sig4 = varsSheet.row_values(31)[6]
col5sig4 = varsSheet.row_values(32)[6]
col6sig4 = varsSheet.row_values(33)[6]
col7sig4 = varsSheet.row_values(34)[6]
col8sig4 = varsSheet.row_values(35)[6]
col9sig4 = varsSheet.row_values(36)[6]

col0sig5 = varsSheet.row_values(27)[7]
col1sig5 = varsSheet.row_values(28)[7]
col2sig5 = varsSheet.row_values(29)[7]
col3sig5 = varsSheet.row_values(30)[7]
col4sig5 = varsSheet.row_values(31)[7]
col5sig5 = varsSheet.row_values(32)[7]
col6sig5 = varsSheet.row_values(33)[7]
col7sig5 = varsSheet.row_values(34)[7]
col8sig5 = varsSheet.row_values(35)[7]
col9sig5 = varsSheet.row_values(36)[7]


# final test input
#  opp, if sum(number),  col+days prior* 10,
# final test 1
test1part1opp = varsSheet.row_values(44)[2]
test1part1sum = varsSheet.row_values(44)[3]
test1part1col0 = varsSheet.row_values(44)[4]
test1part1col1 = varsSheet.row_values(44)[5]
test1part1col2 = varsSheet.row_values(44)[6]
test1part1col3 = varsSheet.row_values(44)[7]
test1part1col4 = varsSheet.row_values(44)[8]
test1part1col5 = varsSheet.row_values(44)[9]
test1part1col6 = varsSheet.row_values(44)[10]
test1part1col7 = varsSheet.row_values(44)[11]
test1part1col8 = varsSheet.row_values(44)[12]
test1part1col9 = varsSheet.row_values(44)[13]
test1part1skip = varsSheet.row_values(44)[14]

test1part2opp = varsSheet.row_values(45)[2]
test1part2sum = varsSheet.row_values(45)[3]
test1part2col0 = varsSheet.row_values(45)[4]
test1part2col1 = varsSheet.row_values(45)[5]
test1part2col2 = varsSheet.row_values(45)[6]
test1part2col3 = varsSheet.row_values(45)[7]
test1part2col4 = varsSheet.row_values(45)[8]
test1part2col5 = varsSheet.row_values(45)[9]
test1part2col6 = varsSheet.row_values(45)[10]
test1part2col7 = varsSheet.row_values(45)[11]
test1part2col8 = varsSheet.row_values(45)[12]
test1part2col9 = varsSheet.row_values(45)[13]
test1part2skip = varsSheet.row_values(45)[14]

test1part3opp = varsSheet.row_values(46)[2]
test1part3sum = varsSheet.row_values(46)[3]
test1part3col0 = varsSheet.row_values(46)[4]
test1part3col1 = varsSheet.row_values(46)[5]
test1part3col2 = varsSheet.row_values(46)[6]
test1part3col3 = varsSheet.row_values(46)[7]
test1part3col4 = varsSheet.row_values(46)[8]
test1part3col5 = varsSheet.row_values(46)[9]
test1part3col6 = varsSheet.row_values(46)[10]
test1part3col7 = varsSheet.row_values(46)[11]
test1part3col8 = varsSheet.row_values(46)[12]
test1part3col9 = varsSheet.row_values(46)[13]
test1part3skip = varsSheet.row_values(46)[14]

test1part4opp = varsSheet.row_values(47)[2]
test1part4sum = varsSheet.row_values(47)[3]
test1part4col0 = varsSheet.row_values(47)[4]
test1part4col1 = varsSheet.row_values(47)[5]
test1part4col2 = varsSheet.row_values(47)[6]
test1part4col3 = varsSheet.row_values(47)[7]
test1part4col4 = varsSheet.row_values(47)[8]
test1part4col5 = varsSheet.row_values(47)[9]
test1part4col6 = varsSheet.row_values(47)[10]
test1part4col7 = varsSheet.row_values(47)[11]
test1part4col8 = varsSheet.row_values(47)[12]
test1part4col9 = varsSheet.row_values(47)[13]
test1part4skip = varsSheet.row_values(47)[14]

test1part5opp = varsSheet.row_values(48)[2]
test1part5sum = varsSheet.row_values(48)[3]
test1part5col0 = varsSheet.row_values(48)[4]
test1part5col1 = varsSheet.row_values(48)[5]
test1part5col2 = varsSheet.row_values(48)[6]
test1part5col3 = varsSheet.row_values(48)[7]
test1part5col4 = varsSheet.row_values(48)[8]
test1part5col5 = varsSheet.row_values(48)[9]
test1part5col6 = varsSheet.row_values(48)[10]
test1part5col7 = varsSheet.row_values(48)[11]
test1part5col8 = varsSheet.row_values(48)[12]
test1part5col9 = varsSheet.row_values(48)[13]
test1part5skip = varsSheet.row_values(48)[14]

test1part6opp = varsSheet.row_values(49)[2]
test1part6sum = varsSheet.row_values(49)[3]
test1part6col0 = varsSheet.row_values(49)[4]
test1part6col1 = varsSheet.row_values(49)[5]
test1part6col2 = varsSheet.row_values(49)[6]
test1part6col3 = varsSheet.row_values(49)[7]
test1part6col4 = varsSheet.row_values(49)[8]
test1part6col5 = varsSheet.row_values(49)[9]
test1part6col6 = varsSheet.row_values(49)[10]
test1part6col7 = varsSheet.row_values(49)[11]
test1part6col8 = varsSheet.row_values(49)[12]
test1part6col9 = varsSheet.row_values(49)[13]
test1part6skip = varsSheet.row_values(49)[14]

test1part7opp = varsSheet.row_values(50)[2]
test1part7sum = varsSheet.row_values(50)[3]
test1part7col0 = varsSheet.row_values(50)[4]
test1part7col1 = varsSheet.row_values(50)[5]
test1part7col2 = varsSheet.row_values(50)[6]
test1part7col3 = varsSheet.row_values(50)[7]
test1part7col4 = varsSheet.row_values(50)[8]
test1part7col5 = varsSheet.row_values(50)[9]
test1part7col6 = varsSheet.row_values(50)[10]
test1part7col7 = varsSheet.row_values(50)[11]
test1part7col8 = varsSheet.row_values(50)[12]
test1part7col9 = varsSheet.row_values(50)[13]
test1part7skip = varsSheet.row_values(50)[14]

# final test 2
test2part1opp = varsSheet.row_values(53)[2]
test2part1sum = varsSheet.row_values(53)[3]
test2part1col0 = varsSheet.row_values(53)[4]
test2part1col1 = varsSheet.row_values(53)[5]
test2part1col2 = varsSheet.row_values(53)[6]
test2part1col3 = varsSheet.row_values(53)[7]
test2part1col4 = varsSheet.row_values(53)[8]
test2part1col5 = varsSheet.row_values(53)[9]
test2part1col6 = varsSheet.row_values(53)[10]
test2part1col7 = varsSheet.row_values(53)[11]
test2part1col8 = varsSheet.row_values(53)[12]
test2part1col9 = varsSheet.row_values(53)[13]
test2part1skip = varsSheet.row_values(53)[14]

test2part2opp = varsSheet.row_values(54)[2]
test2part2sum = varsSheet.row_values(54)[3]
test2part2col0 = varsSheet.row_values(54)[4]
test2part2col1 = varsSheet.row_values(54)[5]
test2part2col2 = varsSheet.row_values(54)[6]
test2part2col3 = varsSheet.row_values(54)[7]
test2part2col4 = varsSheet.row_values(54)[8]
test2part2col5 = varsSheet.row_values(54)[9]
test2part2col6 = varsSheet.row_values(54)[10]
test2part2col7 = varsSheet.row_values(54)[11]
test2part2col8 = varsSheet.row_values(54)[12]
test2part2col9 = varsSheet.row_values(54)[13]
test2part2skip = varsSheet.row_values(54)[14]

test2part3opp = varsSheet.row_values(55)[2]
test2part3sum = varsSheet.row_values(55)[3]
test2part3col0 = varsSheet.row_values(55)[4]
test2part3col1 = varsSheet.row_values(55)[5]
test2part3col2 = varsSheet.row_values(55)[6]
test2part3col3 = varsSheet.row_values(55)[7]
test2part3col4 = varsSheet.row_values(55)[8]
test2part3col5 = varsSheet.row_values(55)[9]
test2part3col6 = varsSheet.row_values(55)[10]
test2part3col7 = varsSheet.row_values(55)[11]
test2part3col8 = varsSheet.row_values(55)[12]
test2part3col9 = varsSheet.row_values(55)[13]
test2part3skip = varsSheet.row_values(55)[14]

test2part4opp = varsSheet.row_values(56)[2]
test2part4sum = varsSheet.row_values(56)[3]
test2part4col0 = varsSheet.row_values(56)[4]
test2part4col1 = varsSheet.row_values(56)[5]
test2part4col2 = varsSheet.row_values(56)[6]
test2part4col3 = varsSheet.row_values(56)[7]
test2part4col4 = varsSheet.row_values(56)[8]
test2part4col5 = varsSheet.row_values(56)[9]
test2part4col6 = varsSheet.row_values(56)[10]
test2part4col7 = varsSheet.row_values(56)[11]
test2part4col8 = varsSheet.row_values(56)[12]
test2part4col9 = varsSheet.row_values(56)[13]
test2part4skip = varsSheet.row_values(56)[14]

test2part5opp = varsSheet.row_values(57)[2]
test2part5sum = varsSheet.row_values(57)[3]
test2part5col0 = varsSheet.row_values(57)[4]
test2part5col1 = varsSheet.row_values(57)[5]
test2part5col2 = varsSheet.row_values(57)[6]
test2part5col3 = varsSheet.row_values(57)[7]
test2part5col4 = varsSheet.row_values(57)[8]
test2part5col5 = varsSheet.row_values(57)[9]
test2part5col6 = varsSheet.row_values(57)[10]
test2part5col7 = varsSheet.row_values(57)[11]
test2part5col8 = varsSheet.row_values(57)[12]
test2part5col9 = varsSheet.row_values(57)[13]
test2part5skip = varsSheet.row_values(57)[14]

finalReturnDays = int(varsSheet.row_values(59)[3])

# GET Stock Data File
inputDF = pd.read_excel(sys.argv[1])

# List of functions for reading and maniplulating files
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

def makeCol(opp, sig1, sig2, sig3, sig4, sig5):
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
  if sig4:
    numSigs +=1
    sig4 = strParserEval(sig4)
    if type(sig4) == str:
      return None
  if sig5:
    numSigs +=1
    sig5 = strParserEval(sig5)
    if type(sig5) == str:
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
    elif numSigs == 4:
      for i in range(0, len(sig1)):
        if sig1[i] + sig2[i] + sig3[i] + sig4[i] == 4:
          array.append(1)
          numTrue += 1
        else:
          array.append(0)
    elif numSigs == 5:
      for i in range(0, len(sig1)):
        if sig1[i] + sig2[i] + sig3[i] + sig4[i] + sig5[i]== 5:
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
    elif numSigs == 4:
      for i in range(0, len(sig1)):
        if sig1[i] + sig2[i] + sig3[i] + sig4[i] >= 1:
          array.append(1)
          numTrue += 1
        else:
          array.append(0)
    elif numSigs == 5:
      for i in range(0, len(sig1)):
        if sig1[i] + sig2[i] + sig3[i] + sig4[i] + sig5[i] >= 1:
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

def parseTestCols(testNum, partNum):
    array = []
    col = "test" + str(testNum) + "part" + str(partNum) + "col"
    for i in range(0, 10):
      test = col + str(i)
      res =  eval(test)
      if res:
        array.append(res)
    return array

def finalTestParams(testNum, partNum):
    testCols = parseTestCols(testNum,partNum)
    numCols = len(testCols)
    opp = eval("test" + str(testNum) + "part" + str(partNum) + "opp")
    skip = eval("test" + str(testNum) + "part" + str(partNum) + "skip").lower()
    sumNum = 0
    if skip == "skip" or skip == "true":
      skip = True
    else:
      skip = False
    if opp == "sum":
      sumNum = int(eval("test" + str(testNum) + "part" + str(partNum) + "sum"))
    elif opp == "and":
      sumNum = len(testCols)
    elif opp == "or":
      sumNum = 1
    else:
      sumNum = 1
    colArr = []
    for i in range(0, numCols):
      colArr.append({
        "data": eval(testCols[i].split()[0].lower()),
        "daysPrior": eval(testCols[i].split()[1])
        })
    return colArr, sumNum, numCols, skip

def finalTestPart(testNum, partNum):
    params =  finalTestParams(testNum, partNum)
    cols = params[0]
    sumNum = params[1]
    numCols = params[2]
    skip =  params[3]
    array = []
    for i in range(0, len(cols[0]["data"])):
      testSum = 0
      for j in range(0, numCols):
        days = cols[j]["daysPrior"]
        if i - j > 0:
          if cols[j]["data"][i-days] == 1:
            testSum += 1
      if testSum >= sumNum:
        array.append(1)
      else:
        array.append(0)
    num = 0
    for k in range(0, len(array)):
      if array[k] == 1:
        num += 1
    return array, skip

def finalTest1():
    numTests = 0
    for n in range(0, 7):
      if varsSheet.row_values(44 + n)[2]:
        numTests += 1
    resArray = []
    finalArr = []
    tests = []
    skips = []
    for j in range(1, numTests+1):
      test = finalTestPart(1, j)
      tests.append(test[0])
      skips.append(test[1])
    for i in range(0, len(tests[0])):
      res = 0
      finalRes = 0
      for k in range(0, numTests):
        if skips[k]:
          if tests[k][i] == 1:
            finalRes = 1
        else:
          if tests[k][i] == 1:
            res = 1
      resArray.append(res)
      finalArr.append(finalRes)
    result = pd.Series(resArray, name=setColName())
    return resArray, result, finalArr

def finalTest2(dependent, done):
    numTests = 0
    for n in range(0, 5):
      if varsSheet.row_values(53 + n)[2]:
        numTests += 1
    resArray = []
    tests = []
    skips = []
    for j in range(1, numTests+1):
      test = finalTestPart(2, j)
      tests.append(test[0])
      skips.append(test[1])
    for i in range(0, len(tests[0])):
      res = 0
      for k in range(0, numTests):
        if skips[k]:
          if tests[k][i] == 1:
            res = 2
        else:
          if tests[k][i] == 1:
            res = 1
      if res == 1 and dependent[i] == 1:
        resArray.append(1)
      elif res == 2:
        resArray.append(1)
      else:
        resArray.append(0)
    final = []
    for l in range(0, len(done)):
      if done[l] == 1 or resArray[l] == 1:
        final.append(1)
      else:
        final.append(0)
    result = pd.Series(final, name=setColName())
    return result

def calcRtns(final2, rtn1, numDays):
    array = []
    for i in range(0, len(final2)):
      if final2[i] == 1:
        rtn = 0
        for j in range(1, numDays+1):
          if (i + j) < len(rtn1):
            rtn += rtn1[i + j]
        array.append(rtn)
      else:
        array.append(None)

    result = pd.Series(array, name=stock["stockName"], index=df[0])
    return result

def rtnStats(rtn):
    totRtn = 0
    rtns = []
    hit = 0
    mx = -10
    mn = 10
    winHit = 0
    winRtn = 0
    lossHit = 0
    winStr = 0
    lossStr = 0
    winStrRtn = 0
    lossStrRtn = 0
    mxWinStr = 0
    mxWinStrRtn = 0
    mxLossStr = 0
    mxLossStrRtn = 0
    for i in range(0, len(rtn)):
      if pd.notnull(rtn[i]):
        totRtn += rtn[i]
        hit += 1
        rtns.append(rtn[i])
        if rtn[i] > mx:
          mx = rtn[i]
        if rtn[i] < mn:
          mn = rtn[i]
        if rtn[i] > 0:
          winHit += 1
          winRtn += rtn[i]
          winStr += 1
          winStrRtn += rtn[i]
          if lossStr > mxLossStr:
            mxLossStr = lossStr
            mxLossStrRtn = lossStrRtn
          lossStr = 0
          lossStrRtn = 0
        if rtn[i] < 0:
          lossHit += 1
          lossStr += 1
          lossStrRtn += rtn[i]
          if winStr > mxWinStr:
            mxWinStr = winStr
            mxWinStrRtn = winStrRtn
          winStr = 0
          winStrRtn = 0
    drawDown = mn - mx
    avgWin = winRtn / winHit
    winPer = winHit / hit
    avgRtn = totRtn / hit
    data = [hit,
      mx,
      mn,
      avgRtn,
      totRtn,
      winHit,
      winRtn,
      lossHit,
      mxWinStr,
      mxWinStrRtn,
      mxLossStr,
      mxLossStrRtn,
      drawDown,
      rtns.sort()]
    index = [
      "Hit Count",
      "Max",
      "Min",
      "Average Return",
      "Total Return",
      "Win #",
      "Win %",
      "Loss #",
      "Count Win Streak",
      "Win Streak %",
      "Count Loss Streak",
      "Loss Streak %",
      "Max Drawdown",
      "List of Returns"]
    result = pd.Series(data, name=stock["stockName"], index=index)
    return result

makeColsArr = ["makeCol(col0opp,col0sig1,col0sig2,col0sig3,col0sig4,col1sig5)",
  "makeCol(col1opp,col1sig1,col1sig2,col1sig3,col1sig4,col1sig5)",
  "makeCol(col2opp,col2sig1,col2sig2,col2sig3,col2sig4,col2sig5)",
  "makeCol(col3opp,col3sig1,col3sig2,col3sig3,col3sig4,col3sig5)",
  "makeCol(col4opp,col4sig1,col4sig2,col4sig3,col4sig4,col4sig5)",
  "makeCol(col5opp,col5sig1,col5sig2,col5sig3,col5sig4,col5sig5)",
  "makeCol(col6opp,col6sig1,col6sig2,col6sig3,col6sig4,col6sig5)",
  "makeCol(col7opp,col7sig1,col7sig2,col7sig3,col7sig4,col7sig5)",
  "makeCol(col8opp,col8sig1,col8sig2,col8sig3,col8sig4,col8sig5)",
  "makeCol(col9opp,col9sig1,col9sig2,col9sig3,col9sig4,col9sig5)"]
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

#
# Parse through the stocks
#
print("reading file...")
stockInfo = []
readFile()

resultsFrame = pd.DataFrame()
returnsFrame = pd.DataFrame()

#
# FOR EACH STOCK
#
print("starting calculations...")
for stock in stockInfo:
  start_time2 = time.time()
  resetColName()

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


  # Section 3
  # Assign Columns
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


  # run final tests

  final1 = finalTest1()
  dependent = final1[0]
  done = final1[2]

  final2 = finalTest2(dependent, done)

  # calculate returns returns.
  resReturns = calcRtns(final2, df[15], finalReturnDays)
  returnsFrame = pd.concat([returnsFrame,resReturns], axis = 1)

  stats = rtnStats(resReturns)
  resultsFrame = pd.concat([resultsFrame, stats], axis = 1)

  # save in temp stock df
  dfRes = pd.concat([stats, resReturns])
  df = pd.concat([df, dfRes],axis=1)


  print(str(stock["stockName"]) + " %g seconds" % (time.time() - start_time2))

# save and join the tables.
resultsFrame = pd.concat([resultsFrame, returnsFrame])
save_xls([resultsFrame], "results.xlsx")

print("Total Elapsed time was %g seconds" % (time.time() - start_time))
