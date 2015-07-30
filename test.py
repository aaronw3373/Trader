
print("starting up...")
import xlrd
import xlwt
from datetime import datetime

print("reading models...")

data_model = xlrd.open_workbook("input.xlsx")
sheet = data_model.sheets()[0]

wb = xlwt.Workbook()
ws = wb.add_sheet("model")

xlwt.add_palette_colour("custom_green", 0x21)
wb.set_colour_RGB(0x21,200, 255, 200)

xlwt.add_palette_colour("custom_red", 0x22)
wb.set_colour_RGB(0x22,255, 200, 200)

style1 = xlwt.easyxf(num_format_str='M/D/YYYY')

style0 = xlwt.easyxf('font: color-index black',
    num_format_str='#,###0.000')

style2 = xlwt.easyxf('pattern:  pattern solid, fore_color custom_green; font: color-index green',
    num_format_str='#,##0.00%')
style3 = xlwt.easyxf('pattern:  pattern solid, fore_color custom_red; font: color-index red',
    num_format_str='#,##0.00%')
style4 = xlwt.easyxf('pattern:  pattern solid, fore_color white; font: color-index black',
    num_format_str='#,##0.00%')

style5 = xlwt.easyxf('pattern:  pattern solid, fore_color custom_green; font: color-index green',
    num_format_str='#0')
style6 = xlwt.easyxf('pattern:  pattern solid, fore_color white; font: color-index black',
    num_format_str='#0')
style7 = xlwt.easyxf('pattern:  pattern solid, fore_color custom_red; font: color-index red',
    num_format_str='#0')

cwn = 0
def setColWrite():
  global cwn
  cwn +=1
  print cwn
  return cwn

print("model imported, parsing data...");

# limits on the data store

# IMPORTANT!!! make this a function add add where to read from the file to work on all input
lowerRange = 7
upperRange = 3424


# Input
# parse that data to get a list of workable values
dataDate = []
dataClose = []
dataOpen = []
dataHigh = []
dataLow = []

def sheetParser(input,low, high, colRead, colWrite, head, output, style):
  ws.write(1,colWrite,head)
  num = 2
  for i in range(low, high):
    output.append(input.row_values(i)[colRead])
    ws.write(num,colWrite,input.row_values(i)[colRead], style)
    num += 1

sheetParser(sheet,lowerRange,upperRange,2, setColWrite(), "Date",dataDate, style1)
sheetParser(sheet,lowerRange,upperRange,3, setColWrite(), "Close",dataClose,style0)
sheetParser(sheet,lowerRange,upperRange,4, setColWrite(), "Open",dataOpen,style0)
sheetParser(sheet,lowerRange,upperRange,5, setColWrite(), "High",dataHigh,style0)
sheetParser(sheet,lowerRange,upperRange,6, setColWrite(), "Low",dataLow,style0)

print('finding averages...')

# Moving Average Calculator
data200Avg = []
data100Avg = []
data50Avg = []
data30Avg = []
data10Avg = []
def numDayAvg(input, numDays, colWrite, head, style, output):
  ws.write(1,colWrite,head)
  length = len(input)
  num = 2
  for i in range(0, length):
    total = 0
    count = 0;
    if i - numDays >= -1:
      for j in range(0, numDays):
        if (i - j) >= 0:
          count += 1
          total += input[i - j]
      avg = total / count
      output.append(avg)
      ws.write(num,colWrite,avg, style)
    else:
      output.append(0)
      ws.write(num,colWrite,0, style)
    num += 1

numDayAvg(dataClose, 200, setColWrite(), "200 Day Avg", style0, data200Avg)
numDayAvg(dataClose, 100, setColWrite(), "100 Day Avg", style0, data100Avg)
numDayAvg(dataClose, 50, setColWrite(), "50 Day Avg", style0, data50Avg)
numDayAvg(dataClose, 30, setColWrite(), "30 Day Avg", style0, data30Avg)
numDayAvg(dataClose, 10, setColWrite(), "10 Day Avg", style0, data10Avg)

print("finding returns...")
# number of days returns
data2Rtn = []
data3Rtn = []
data5Rtn = []
data1Rtn = []
def numDayRtn(input, numDays, colWrite, head, output):
  ws.write(1,colWrite,head)
  length = len(input)
  num = 2
  for i in range(0, length):
    if (i - numDays) >= 0:
      diff = (input[i] - input[i-numDays])/input[i-numDays]
      style = style2
      if diff < 0:
        style = style3
      if diff == 0:
        style = style4
      output.append(diff)
      ws.write(num,colWrite,diff, style)
    else:
      output.append(0)
      ws.write(num,colWrite,0, style4)
    num += 1

numDayRtn(dataClose, 2, setColWrite(), "2 Day Return", data2Rtn)
numDayRtn(dataClose, 3, setColWrite(), "3 Day Return", data3Rtn)
numDayRtn(dataClose, 5, setColWrite(), "5 Day Return", data5Rtn)
numDayRtn(dataClose, 1, setColWrite(), "1 Day Return", data1Rtn)

# Overnight Return
dataNightRtn = []
def overnightRtn(open, close, colWrite, head, output):
  ws.write(1,colWrite,head)
  length = len(close)
  num = 2
  for i in range(0, length):
    if (i > 0):
      diff = (open[i] - close[i-1])/close[i-1]
      style = style2
      if diff < 0:
        style = style3
      if diff == 0:
        style = style4
      output.append(diff)
      ws.write(num,colWrite,diff, style)
    else:
      output.append(0)
      ws.write(num,colWrite,0, style4)
    num += 1

overnightRtn(dataOpen,dataClose,setColWrite(), "Overnight Return", dataNightRtn)

# daytime Return
dataDayRtn = []
def dayRtn(open, close, colWrite, head, output):
  ws.write(1,colWrite,head)
  length = len(close)
  num = 2
  for i in range(0, length):
    diff = (close[i] - open[i])/open[i]
    style = style2
    if diff < 0:
      style = style3
    if diff == 0:
      style = style4
    output.append(diff)
    ws.write(num,colWrite,diff, style)
    num += 1

dayRtn(dataOpen, dataClose, setColWrite(), "Daytime Return", dataDayRtn)

print("testing for signals...")

#SIGNALS


def topLine(test, others, colWrite, head):
  ws.write(1,colWrite,head)
  length = len(test)
  num = 2
  for i in range(0, length):
    res = 1
    style = style5
    for other in others:
      if test[i] < other[i]:
        res = 0
        style = style6
    ws.write(num,colWrite,res,style)
    num += 1

topLine(data10Avg, [data30Avg, data50Avg, data100Avg, data200Avg, dataClose], setColWrite(), "10 Day Top")
topLine(data30Avg, [data10Avg, data50Avg, data100Avg, data200Avg, dataClose], setColWrite(), "30 Day Top")
topLine(data50Avg, [data10Avg, data30Avg, data100Avg, data200Avg, dataClose], setColWrite(), "50 Day Top")
topLine(data100Avg, [data10Avg, data30Avg, data50Avg, data200Avg, dataClose], setColWrite(), "100 Day Top")
topLine(data200Avg, [data10Avg, data30Avg, data50Avg, data100Avg, dataClose], setColWrite(), "200 Day Top")
topLine(dataClose, [data10Avg, data30Avg, data50Avg, data100Avg, data200Avg], setColWrite(), "Close Top")

def bottomLine(test, others, colWrite, head):
  ws.write(1,colWrite,head)
  length = len(test)
  num = 2
  for i in range(0, length):
    if test[i] != 0:
      res = 1
      style = style7
      for other in others:
        if other[i] != 0:
          if test[i] > other[i]:
            res = 0
            style = style6
      ws.write(num,colWrite,res,style)
    else:
      res = 0
      style = style6
      ws.write(num,colWrite,res,style)
    num += 1

bottomLine(data10Avg, [data30Avg, data50Avg, data100Avg, data200Avg, dataClose], setColWrite(), "10 Day Bottom")
bottomLine(data30Avg, [data10Avg, data50Avg, data100Avg, data200Avg, dataClose], setColWrite(), "30 Day Bottom")
bottomLine(data50Avg, [data10Avg, data30Avg, data100Avg, data200Avg, dataClose], setColWrite(), "50 Day Bottom")
bottomLine(data100Avg, [data10Avg, data30Avg, data50Avg, data200Avg, dataClose], setColWrite(), "100 Day Bottom")
bottomLine(data200Avg, [data10Avg, data30Avg, data50Avg, data100Avg, dataClose], setColWrite(), "200 Day Bottom")
bottomLine(dataClose, [data10Avg, data30Avg, data50Avg, data100Avg, data200Avg], setColWrite(), "Close Bottom")

def priceAbove(test, other, colWrite, head):
  ws.write(1,colWrite,head)
  length = len(test)
  num = 2
  for i in range(0, length):
    if test[i] != 0 and other[i] != 0:
      res = 0
      style = style6
      if test[i] > other[i]:
          res = 1
          style = style5
      ws.write(num,colWrite,res,style)
    else:
      res = 0
      style = style6
      ws.write(num,colWrite,res,style)
    num += 1

priceAbove(dataClose, data10Avg, setColWrite(), "price above 10")
priceAbove(dataClose, data30Avg, setColWrite(), "price above 30")
priceAbove(dataClose, data50Avg, setColWrite(), "price above 50")
priceAbove(dataClose, data100Avg, setColWrite(), "price above 100")
priceAbove(dataClose, data200Avg, setColWrite(), "price above 200")

def crossAbove(test, other, colWrite, head):
  ws.write(1,colWrite,head)
  length = len(test)
  num = 2
  for i in range(0, length):
    if test[i] != 0 and other[i] != 0:
      res = 0
      style = style6
      if test[i-1] < other[i-1] and test[i] > other[i]:
          res = 1
          style = style5
      ws.write(num,colWrite,res,style)
    else:
      res = 0
      style = style6
      ws.write(num,colWrite,res,style)
    num += 1

def crossBelow(test, other, colWrite, head):
  ws.write(1,colWrite,head)
  length = len(test)
  num = 2
  for i in range(0, length):
    if test[i] != 0 and other[i] != 0:
      res = 0
      style = style6
      if test[i-1] > other[i-1] and test[i] < other[i]:
          res = 1
          style = style7
      ws.write(num,colWrite,res,style)
    else:
      res = 0
      style = style6
      ws.write(num,colWrite,res,style)
    num += 1

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

def crossVarPrice(test, var, colWrite, head):
  ws.write(1,colWrite,head)
  length = len(test)
  num = 2
  for i in range(0, length):
    if test[i] != 0:
      res = 0
      style = style6
      if test[i] > var and test[i-1] < var:
          res = 1
          style = style5
      elif test[i] < var and test[i-1] > var:
          res = 1
          style = style7
      ws.write(num,colWrite,res,style)
    else:
      res = 0
      style = style6
      ws.write(num,colWrite,res,style)
    num += 1

# varInputPrice = raw_input("Please enter a price to cross: ")
# print "you entered", varInputPrice
# crossVarPrice(dataClose, float(varInputPrice), setColWrite(), "close cross " + varInputPrice)


def crossVarPercent(test, var, colWrite, head):
  ws.write(1,colWrite,head)
  length = len(test)
  num = 2
  if var >= 0:
    for i in range(0, length):
      if test[i] != 0:
        res = 0
        style = style6
        if test[i] > var:
            res = 1
            style = style5
        ws.write(num,colWrite,res,style)
      else:
        res = 0
        style = style6
        ws.write(num,colWrite,res,style)
      num += 1
  elif var < 0:
    for i in range(0, length):
      if test[i] != 0:
        res = 0
        style = style6
        if test[i] < var:
            res = 1
            style = style7
        ws.write(num,colWrite,res,style)
      else:
        res = 0
        style = style6
        ws.write(num,colWrite,res,style)
      num += 1

# VarInput2 = raw_input("Please enter a decimal percent to cross 2 Day Return: ")
# print "you entered", VarInput2
# crossVarPercent(data2Rtn, float(VarInput2),setColWrite(), "2D X " + VarInput2)

# VarInput3 = raw_input("Please enter a decimal percent to cross 3 Day Return: ")
# print "you entered", VarInput3
# crossVarPercent(data3Rtn, float(VarInput3),setColWrite(), "3D X " + VarInput3)

# VarInput5 = raw_input("Please enter a decimal percent to cross 5 Day Return: ")
# print "you entered", VarInput5
# crossVarPercent(data5Rtn, float(VarInput5),setColWrite(), "5D X " + VarInput5)

# VarInput1 = raw_input("Please enter a decimal percent to cross 1 Day Return: ")
# print "you entered", VarInput1
# crossVarPercent(data1Rtn, float(VarInput1),setColWrite(), "1D X " + VarInput1)

# VarInput0 = raw_input("Please enter a decimal percent to cross Daytime Return: ")
# print "you entered", VarInput0
# crossVarPercent(dataDayRtn, float(VarInput0),setColWrite(), "Day X " + VarInput0)

# VarInput = raw_input("Please enter a decimal percent to cross Nighttime Return: ")
# print "you entered", VarInput
# crossVarPercent(dataNightRtn, float(VarInput),setColWrite(), "Nt X " + VarInput)


def highBtwDays(test, numDays, colWrite, head):
  ws.write(1,colWrite,head)
  length = len(test)
  num = 2
  for i in range(0, length):
    if i - numDays >= -1:
      res = 1
      style = style5
      for j in range(0,numDays):
        if test[i] < test[i-j]:
          res = 0
          style = style6
      ws.write(num,colWrite,res,style)
    else:
      res = 0
      style = style6
      ws.write(num,colWrite,res,style)
    num += 1

# varInput10 = raw_input("Please enter days for max close: ")
# print "you entered", varInput10
# highBtwDays(dataClose, int(varInput10), setColWrite(), varInput10 + " high close")

# varInput11 = raw_input("Please enter days for max close: ")
# print "you entered", varInput11
# highBtwDays(dataHigh, int(varInput11), setColWrite(), varInput11 + " high close")

def lowBtwDays(test, numDays, colWrite, head):
  ws.write(1,colWrite,head)
  length = len(test)
  num = 2
  for i in range(0, length):
    if i - numDays >= -1:
      res = 1
      style = style7
      for j in range(0,numDays):
        if test[i] > test[i-j]:
          res = 0
          style = style6
      ws.write(num,colWrite,res,style)
    else:
      res = 0
      style = style6
      ws.write(num,colWrite,res,style)
    num += 1

# varInput12 = raw_input("Please enter days for min close: ")
# print "you entered", varInput12
# lowBtwDays(dataClose, int(varInput12), setColWrite(), varInput12 + " low close")

# varInput13 = raw_input("Please enter days for min close: ")
# print "you entered", varInput13
# lowBtwDays(dataLow, int(varInput13), setColWrite(), varInput13 + " low close")


def varDayRtn(test, numDays, colWrite, head):
  ws.write(1,colWrite,head)
  length = len(test)
  num = 2
  for i in range(0, length):
    if i - numDays >= -1:
      rtn = 0
      for j in range(0, numDays):
        rtn += test[i-j]
      if rtn >= 0:
        style = style2
      else:
        style = style3
      ws.write(num,colWrite,rtn,style)
    else:
      rtn = 0
      style = style4
      ws.write(num,colWrite,rtn,style)
    num += 1

# varInput14 = raw_input("Please enter days for return: ")
# print "you entered", varInput14
# varDayRtn(data1Rtn, int(varInput14), setColWrite(), varInput14 + " day Rtn")




# saving spreadsheet
wb.save('end_model.xls')
