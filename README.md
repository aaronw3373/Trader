#Trader
<!-- By: Aaron Weinberg -->

#About
<!-- Trader is used to find stock trading signals to determine when to buy and sell stock.
Trader takes as input a portfolio of stock price data.
Then runs a series of tests to determain if trading signals are true.
And returns a portfolio of each stock and the returns over time for the input signals. -->

#How to use
This is a command line application
To use go into the root of the application and in the command line type:
python trader.py <--stock price file--> <--var input file-->
stock price being a spreadsheet of all the stocks price data
and var input file being a spreadsheet of the input variables

for testing run:
python trader.py lib/input.xlsx lib/inputVars.xlsx

#Dependencies
python,
pandas,
xlrd

#Installation
Fork this repository in any directory on your computer.
install the following dependencies if you don't already have them:

python,
pandas,
xlrd

For installing Python:
I'm using version 2.7.10 as it is more stable:
https://www.python.org/downloads/

For installing pandas:
http://pandas.pydata.org/pandas-docs/stable/install.html




<!-- Go into that directory and type: python test.py -->
<!-- Once complete if there are no errors then you are set to use trader -->

