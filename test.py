
print("starting up...")
import xlrd
import xlwt
from datetime import datetime

print("reading models...")

data_model = xlrd.open_workbook("data_model.xlsx")
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



print("model imported, parsing data...");

# limits on the data store
lowerRange = 7
upperRange = 3384


# Input
# parse that data to get a list of workable values
dataClose = []
dataOpen = []
dataDate = []
# data4 = []
# data5 = []

def sheetParser(input,low, high, colRead, colWrite, head, output, style):
  ws.write(1,colWrite,head)
  num = 2
  for i in range(low, high):
    output.append(input.row_values(i)[colRead])
    ws.write(num,colWrite,input.row_values(i)[colRead], style)
    num += 1

sheetParser(sheet,lowerRange,upperRange,2, 0, "Date",dataDate, style1)
sheetParser(sheet,lowerRange,upperRange,102, 2, "Price Open",dataOpen,style0)
sheetParser(sheet,lowerRange,upperRange,103, 1, "Price Close",dataClose,style0)

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

numDayAvg(dataClose, 200, 4, "200 Day Avg", style0, data200Avg)
numDayAvg(dataClose, 100, 5, "100 Day Avg", style0, data100Avg)
numDayAvg(dataClose, 50, 6, "50 Day Avg", style0, data50Avg)
numDayAvg(dataClose, 30, 7, "30 Day Avg", style0, data30Avg)
numDayAvg(dataClose, 10, 8, "10 Day Avg", style0, data10Avg)

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

numDayRtn(dataClose, 2, 9, "2 Day Return", data2Rtn)
numDayRtn(dataClose, 3, 10, "3 Day Return", data3Rtn)
numDayRtn(dataClose, 5, 11, "5 Day Return", data5Rtn)
numDayRtn(dataClose, 1, 12, "1 Day Return", data1Rtn)

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

overnightRtn(dataOpen,dataClose,13, "Overnight Return", dataNightRtn)

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

dayRtn(dataOpen, dataClose, 14, "Daytime Return", dataDayRtn)

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

topLine(data10Avg, [data30Avg, data50Avg, data100Avg, data200Avg, dataClose], 16, "10 Day Top")
topLine(data30Avg, [data10Avg, data50Avg, data100Avg, data200Avg, dataClose], 17, "30 Day Top")
topLine(data50Avg, [data10Avg, data30Avg, data100Avg, data200Avg, dataClose], 18, "50 Day Top")
topLine(data100Avg, [data10Avg, data30Avg, data50Avg, data200Avg, dataClose], 19, "100 Day Top")
topLine(data200Avg, [data10Avg, data30Avg, data50Avg, data100Avg, dataClose], 20, "200 Day Top")
topLine(dataClose, [data10Avg, data30Avg, data50Avg, data100Avg, data200Avg], 21, "Close Top")

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

bottomLine(data10Avg, [data30Avg, data50Avg, data100Avg, data200Avg, dataClose], 22, "10 Day Bottom")
bottomLine(data30Avg, [data10Avg, data50Avg, data100Avg, data200Avg, dataClose], 23, "30 Day Bottom")
bottomLine(data50Avg, [data10Avg, data30Avg, data100Avg, data200Avg, dataClose], 24, "50 Day Bottom")
bottomLine(data100Avg, [data10Avg, data30Avg, data50Avg, data200Avg, dataClose], 25, "100 Day Bottom")
bottomLine(data200Avg, [data10Avg, data30Avg, data50Avg, data100Avg, dataClose], 26, "200 Day Bottom")
bottomLine(dataClose, [data10Avg, data30Avg, data50Avg, data100Avg, data200Avg], 27, "Close Bottom")

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

priceAbove(dataClose, data10Avg, 28, "Price above 10")
priceAbove(dataClose, data30Avg, 29, "Price above 30")
priceAbove(dataClose, data50Avg, 30, "Price above 50")
priceAbove(dataClose, data100Avg, 31, "Price above 100")
priceAbove(dataClose, data200Avg, 32, "Price above 200")

priceAbove(data10Avg, dataClose, 33, "10 above Price")
priceAbove(data10Avg, data30Avg, 34, "10 above 30")
priceAbove(data10Avg, data50Avg, 35, "10 above 50")
priceAbove(data10Avg, data100Avg, 36, "10 above 100")
priceAbove(data10Avg, data200Avg, 37, "10 above 200")

priceAbove(data30Avg, dataClose, 38, "30 above Price")
priceAbove(data30Avg, data10Avg, 39, "30 above 10")
priceAbove(data30Avg, data50Avg, 40, "30 above 50")
priceAbove(data30Avg, data100Avg, 41, "30 above 100")
priceAbove(data30Avg, data200Avg, 42, "30 above 200")

priceAbove(data50Avg, dataClose, 43, "50 above Price")
priceAbove(data50Avg, data10Avg, 44, "50 above 10")
priceAbove(data50Avg, data30Avg, 45, "50 above 30")
priceAbove(data50Avg, data100Avg, 46, "50 above 100")
priceAbove(data50Avg, data200Avg, 47, "50 above 200")

priceAbove(data100Avg, dataClose, 48, "100 above Price")
priceAbove(data100Avg, data10Avg, 49, "100 above 10")
priceAbove(data100Avg, data30Avg, 50, "100 above 30")
priceAbove(data100Avg, data50Avg, 51, "100 above 50")
priceAbove(data100Avg, data200Avg, 52, "100 above 200")

priceAbove(data200Avg, dataClose, 53, "200 above Price")
priceAbove(data200Avg, data10Avg, 54, "200 above 10")
priceAbove(data200Avg, data30Avg, 55, "200 above 30")
priceAbove(data200Avg, data50Avg, 56, "200 above 50")
priceAbove(data200Avg, data100Avg, 57, "200 above 100")


def priceBelow(test, other, colWrite, head):
  ws.write(1,colWrite,head)
  length = len(test)
  num = 2
  for i in range(0, length):
    if test[i] != 0 and other[i] != 0:
      res = 0
      style = style6
      if test[i] < other[i]:
          res = 1
          style = style7
      ws.write(num,colWrite,res,style)
    else:
      res = 0
      style = style6
      ws.write(num,colWrite,res,style)
    num += 1

priceBelow(dataClose, data10Avg, 58, "Price below 10")
priceBelow(dataClose, data30Avg, 59, "Price below 30")
priceBelow(dataClose, data50Avg, 60, "Price below 50")
priceBelow(dataClose, data100Avg, 61, "Price below 100")
priceBelow(dataClose, data200Avg, 62, "Price below 200")

priceBelow(data10Avg, dataClose, 63, "10 below Price")
priceBelow(data10Avg, data30Avg, 64, "10 below 30")
priceBelow(data10Avg, data50Avg, 65, "10 below 50")
priceBelow(data10Avg, data100Avg, 66, "10 below 100")
priceBelow(data10Avg, data200Avg, 67, "10 below 200")

priceBelow(data30Avg, dataClose, 68, "30 below Price")
priceBelow(data30Avg, data10Avg, 69, "30 below 10")
priceBelow(data30Avg, data50Avg, 70, "30 below 50")
priceBelow(data30Avg, data100Avg, 71, "30 below 100")
priceBelow(data30Avg, data200Avg, 72, "30 below 200")

priceBelow(data50Avg, dataClose, 73, "50 below Price")
priceBelow(data50Avg, data10Avg, 74, "50 below 10")
priceBelow(data50Avg, data30Avg, 75, "50 below 30")
priceBelow(data50Avg, data100Avg, 76, "50 below 100")
priceBelow(data50Avg, data200Avg, 77, "50 below 200")

priceBelow(data100Avg, dataClose, 78, "100 below Price")
priceBelow(data100Avg, data10Avg, 79, "100 below 10")
priceBelow(data100Avg, data30Avg, 80, "100 below 30")
priceBelow(data100Avg, data50Avg, 81, "100 below 50")
priceBelow(data100Avg, data200Avg, 82, "100 below 200")

priceBelow(data200Avg, dataClose, 83, "200 below Price")
priceBelow(data200Avg, data10Avg, 84, "200 below 10")
priceBelow(data200Avg, data30Avg, 85, "200 below 30")
priceBelow(data200Avg, data50Avg, 86, "200 below 50")
priceBelow(data200Avg, data100Avg, 87, "200 below 100")

# variable signals



wb.save('end_model.xls')
