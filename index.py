import xlrd
import xlwt
from datetime import datetime

# New styles and colors
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

# Function definitions
cwn = 0
def setColWrite():
  global cwn
  cwn +=1
  return cwn

def sheetParser(input,low, high, colRead, colWrite, head, output, style):
  ws.write(1,colWrite,head)
  num = 2
  for i in range(low, high):
    output.append(input.row_values(i)[colRead])
    ws.write(num,colWrite,input.row_values(i)[colRead], style)
    num += 1

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
