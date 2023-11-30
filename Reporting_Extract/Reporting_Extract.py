import datetime
from datetime import date 
from datetime import datetime as dt
import timeit
import os
os.getcwd()
# go to the folder where you saved the package
os.chdir('P:\\Product Specialists\\Tools\\Python Tools\\Reporting_Extract')
import ReportingDataExtract.Extract as Extract

print("""
If you would like to generate database for new fund, input \"Generate\" or \"G\"
If you would like to update existing database, input \"Update\" or \"U\"\n
""")


# Put the list of all the portfolios below in "tickers"
tickers = ["PF92166","2SICFAL", "C2525R", "D2525","I40","M22285","PF56699",
            "PF71258","X2947","GIFGL","X2975","PF88017","PF49055",
            "X20400","PF103995","MULTI","UNOB","PF55645","PF58925",
            "PF58175","PF82728","PF96625","PF92671","S2525","L837",
            "PF56415","X1144","PF71798","PF47190","M29448","PF90548", 
            "E2525", "D1945L", "D1945C", "D318M", "X68085","PF40733",
            "PF60091","PHMDE","PF81742","PF87396","PF56896","PF37442",
            "PF82494","PF56190","PF45040","PF55902","5742995","PF37614",
           "PF69405","M27828","PF88902","S105195","PF61685","PF48158",
           "PF103973", "PF82700"]

##### Function input #####
function = input("Input the function: ")

##### Date input #####
temp = int(input('\nInput first business day of the following month(YYYYMMDD):'))
print("\n")
date_input = dt.strptime(str(temp), '%Y%m%d').date()

##### Ticker input #####
ticker = input('Ticker:')
print("\n")

##### File path #####
path = "P:\\Product Specialists\\Tools\\Python Tools\\db_trial\\"


##### Start extracting data #####
start_all = timeit.default_timer() # calculate processing time for all the funds
if function == "Generate" or function == "G" or function == "g": # generate database from scratch, load all history
    if ticker == "ALL" or ticker == "All": # generate all portolio data in the list 'tickers'
        for ticker in tickers:
            start_fund = timeit.default_timer()# calculate processing time for each fund
            Extract.GetReporting(ticker, date_input, "Generate", path)
            print(ticker + " Process Time: " + str(timeit.default_timer() - start_fund))# print processing time
    else: # generate data for one fund
        Extract.GetReporting(ticker, date_input, "Generate", path)
        
elif function == "Update" or function == "U" or function == "u": # update database, load only newest data for time series
    if ticker == "ALL" or ticker == "All":  # update all portolio data in the list 'tickers'
        for ticker in tickers:
            start_fund = timeit.default_timer() # calculate processing time for each fund
            Extract.GetReporting(ticker, date_input, "Update", path)
            print(ticker + " Process Time: " + str(timeit.default_timer() - start_fund)) # print processing time
    else: # generate data for one fund
        Extract.GetReporting(ticker, date_input, "Update", path)
        
else: # when the function in incorrectly entered
    print("Incorrect input")
        
        
print("\n\nOverall Process Time: " + str(timeit.default_timer() - start_all))    # print overall processing time