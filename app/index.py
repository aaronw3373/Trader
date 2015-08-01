import openpyxl
import pandas as pd
import numpy as np

# column writter counter
colName = -1
def setColName():
  global colName
  colName +=1
  return colName
def resetColName():
  global colName
  colName = -1

# move data from input dataframe to organized stock dataframe
def sheetParser(input, stock):
  newFrame = pd.DataFrame({
    setColName(): input.iloc[stock["rowStart"]:stock["rowEnd"],stock["dateRead"]].values,
    setColName(): input.iloc[stock["rowStart"]:stock["rowEnd"],stock["closeRead"]].values,
    setColName(): input.iloc[stock["rowStart"]:stock["rowEnd"],stock["openRead"]].values,
    setColName(): input.iloc[stock["rowStart"]:stock["rowEnd"],stock["highRead"]].values,
    setColName():input.iloc[stock["rowStart"]:stock["rowEnd"],stock["lowRead"]].values
    })
  return newFrame

# find moving averages
def numDayAvg(input, numDays):
  temp = pd.rolling_mean(input, numDays)
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
  for i in range(0, len(openToday)):
    change = (closeToday[i]-openToday[i])/openToday[i]
    array.append(change)
  rtn = pd.Series(array, name=setColName())
  return rtn

# Tests
# top line
def topLine(test, others):
  array = []
  for i in range(0, len(test)):
    if pd.notnull(test[i]):
      res = 1
      for other in others:
        if test[i] < other[i]:
          res = 0
      array.append(res)
    else:
      array.append(0)
  results = pd.Series(array, name=setColName())
  return results


# bottom line
def bottomLine(test, others):
  array = []
  for i in range(0, len(test)):
    if pd.notnull(test[i]):
      res = 1
      for other in others:
        if test[i] > other[i]:
          res = 0
      array.append(res)
    else:
      array.append(0)
  results = pd.Series(array, name=setColName())
  return results

# price above
def priceAbove(test, other):
  array = []
  for i in range(0, len(test)):
    if pd.notnull(other[i]):
      res = 0
      if test[i] > other[i]:
        res = 1
      array.append(res)
    else:
      array.append(0)
  results = pd.Series(array, name=setColName())
  return results

def crossAbove(test, other):
  array = []
  for i in range(0, len(test)):
    if pd.notnull(test[i]) and pd.notnull(other[i]):
      res = 0
      if test[i-1] < other[i-1] and test[i] > other[i]:
        res = 1
      array.append(res)
    else:
      array.append(0)
  results = pd.Series(array, name=setColName())
  return results

def crossBelow(test, other):
  array = []
  for i in range(0, len(test)):
    if pd.notnull(test[i]) and pd.notnull(other[i]):
      res = 0
      if test[i-1] > other[i-1] and test[i] < other[i]:
        res = 1
      array.append(res)
    else:
      array.append(0)
  results = pd.Series(array, name=setColName())
  return results

def crossVarPrice(test, var):
  array = []
  array.append(0)
  for i in range(1, len(test)):
    if pd.notnull(test[i]):
      res = 0
      if test[i] > var and test[i-1] < var:
          res = 1
      elif test[i] < var and test[i-1] > var:
          res = 1
      array.append(res)
    else:
      array.append(0)
  results = pd.Series(array, name=setColName())
  return results

def crossVarPercent(test, var):
  array = []
  if var >= 0:
    for i in range(0, len(test)):
      if pd.notnull(test[i]):
        res = 0
        if test[i] > var:
            res = 1
        array.append(res)
      else:
        array.append(0)
  elif var < 0:
    for i in range(0, len(test)):
      if pd.notnull(test[i]):
        res = 0
        if test[i] < var:
            res = 1
        array.append(res)
      else:
        array.append(0)
  results = pd.Series(array, name=setColName())
  return results

def highBtwDays(test, numDays, percent):
  array = []
  for i in range(0, len(test)):
    if i - numDays >= -1:
      res = 1
      for j in range(0,numDays):
        if test[i] < test[i-j]:
          res = 0
      array.append(res)
    else:
      array.append(0)
  results = pd.Series(array, name=setColName())
  return results


def lowBtwDays(test, numDays, percent):
  array = []
  for i in range(0, len(test)):
    if i - numDays >= -1:
      res = 1
      for j in range(0,numDays):
        if test[i] > test[i-j]:
          res = 0
      array.append(res)
    else:
      array.append(0)
  results = pd.Series(array, name=setColName())
  return results
