
print("starting up...")
import xlrd
import xlwt
from datetime import datetime

data_model = xlrd.open_workbook("data_model.xlsx")
sheet = data_model.sheets()[0]

wb = xlwt.Workbook()
ws = wb.add_sheet("model")

print("model imported, parsing data...");

# limits on the data store
lowerRange = 7
upperRange = 3384

# parse that data to get a list of workable values
data1 = []
data2 = []
data3 = []
data4 = []
def sheetParser(input,low, high, colRead, colWrite, head, output):
  ws.write(1,colWrite,head)
  num = 2
  for i in range(low, high):
    output.append(input.row_values(i)[colRead])
    ws.write(num,colWrite,input.row_values(i)[colRead])
    num += 1

sheetParser(sheet,lowerRange,upperRange,4, 1, "Price Close",data1)
sheetParser(sheet,lowerRange,upperRange,5, 2, "Price Open",data2)
sheetParser(sheet,lowerRange,upperRange,6, 3, "Price High",data3)
sheetParser(sheet,lowerRange,upperRange,7, 4, "Price Low",data4)

print('finding averages...')

# Moving Average Calculator
def numDayAvg(input, numDays, colWrite, head):
  ws.write(1,colWrite,head)
  length = len(input)
  res = [];
  num = 2
  for i in range(0, length):
    avg = 0
    count = 0;
    for j in range(0, numDays):
      if (i - j) >= 0:
        count += 1
        avg += input[i - j]
    res.append(avg / count)
    ws.write(num,colWrite,(avg / count))
    num += 1

numDayAvg(data1, 200, 5, "200 Day Avg")
numDayAvg(data1, 100, 6, "100 Day Avg")
numDayAvg(data1, 50, 7, "50 Day Avg")
numDayAvg(data1, 30, 8, "30 Day Avg")
numDayAvg(data1, 10, 9, "10 Day Avg")


wb.save('end_model.xls')
