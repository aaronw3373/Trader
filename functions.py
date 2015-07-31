import pandas as pd
import numpy as np

# column writter counter
colName = -1
def setColName():
  global colName
  colName +=1
  return colName

# move data from input dataframe to organized stock dataframe
def sheetParser(input, stock):
  df = pd.DataFrame({
    setColName(): input.iloc[stock["rowStart"]:stock["rowEnd"],stock["dateRead"]].values,
    setColName(): input.iloc[stock["rowStart"]:stock["rowEnd"],stock["closeRead"]].values,
    setColName(): input.iloc[stock["rowStart"]:stock["rowEnd"],stock["openRead"]].values,
    setColName(): input.iloc[stock["rowStart"]:stock["rowEnd"],stock["highRead"]].values,
    setColName():input.iloc[stock["rowStart"]:stock["rowEnd"],stock["lowRead"]].values
    })
  return df

# find moving averages
def numDayAvg(input, numDays):
  temp = pd.rolling_mean(input,numDays)
  avg = pd.Series(temp, name=setColName())
  return avg

# find returns
def numDayRtn(input, numDays):
  temp = input.pct_change(periods=numDays)
  rtn = pd.Series(temp, name=setColName())
  return rtn

def nightRtn(closeYesterday, openToday):
  array = []
  array.append(np.nan)
  for i in range(1, len(openToday)):
    change = (openToday[i]-closeYesterday[i-1])/closeYesterday[i-1]
    array.append(change)
  rtn = pd.Series(array, name=setColName())
  return rtn

def dayRtn(openToday, closeToday):
  array = []
  print closeToday[0]
  for i in range(0, len(openToday)):
    change = (closeToday[i]-openToday[i])/openToday[i]
    array.append(change)
  rtn = pd.Series(array, name=setColName())
  return rtn
