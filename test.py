import xlrd
data_model = xlrd.open_workbook("data_model.xlsx")
sheet = data_model.sheets()[0]
print("model imported");
lowRange = 7
upRange = 3376
data = []
for i in range(lowRange, upRange):
  data.append(sheet.row_values(i)[4])

l = len(data);
res = [];
for i in range(3300, l):
  avg = 0
  for j in range(0, 9):
    avg += data[i - j]
  res.append(avg / 10)
print res

# print len(data)
# print sum(data) / len(data)
