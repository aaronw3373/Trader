import xlrd
import pandas as pd
import numpy as np

# Input variables
vars_model = xlrd.open_workbook("inputVars.xlsx")
varsSheet = vars_model.sheets()[0]
varInputPrice1 = varsSheet.row_values(4)[5]
varInputPercent1 = varsSheet.row_values(5)[5]
varInputPercent2 = varsSheet.row_values(6)[5]
varInputPercent3 = varsSheet.row_values(7)[5]
varInputPercent5 = varsSheet.row_values(8)[5]
varInputPercent1Limit = varsSheet.row_values(9)[5]
varInputDayIH = varsSheet.row_values(10)[4]
varInputDayEH = varsSheet.row_values(11)[4]
varInputDayIL = varsSheet.row_values(12)[4]
varInputDayEL = varsSheet.row_values(13)[4]
varInputPercentIH = varsSheet.row_values(10)[5]
varInputPercentEH = varsSheet.row_values(11)[5]
varInputPercentIL = varsSheet.row_values(12)[5]
varInputPercentEL = varsSheet.row_values(13)[5]
varInputPercentDay = varsSheet.row_values(14)[5]
varInputPercentNt = varsSheet.row_values(15)[5]
varInputNumDays = varsSheet.row_values(16)[5]
varInputPercentDays = varsSheet.row_values(16)[5]

inputDF = pd.read_excel("input.xlsx")
print inputDF.iloc[5,3]
