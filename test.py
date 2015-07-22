
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

style1 = xlwt.easyxf(num_format_str='D/M/YYYY')

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
    for j in range(0, numDays):
      if (i - j) >= 0:
        count += 1
        total += input[i - j]
    avg = total / count
    output.append(avg)
    ws.write(num,colWrite,avg, style)
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
    if (i -numDays) >= 0:
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
    res = 1
    style = style7
    for other in others:
      if test[i] > other[i]:
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


wb.save('end_model.xls')
