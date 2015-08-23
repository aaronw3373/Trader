from fs.opener import fsopendir
data_fs = fsopendir('lib/data')
for item in data_fs:
  print item


array = []
  for i in range(0, len(input)):
    AVG = 0
    if numDays > i:
      AVG = None
    else:
      for j in range(0, numDays):
        AVG += input[i-j]
    array.append(AVG)
  avg = pd.Series(array, name=setColName())
  return avg
