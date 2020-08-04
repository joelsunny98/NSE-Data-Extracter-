# Importing the urllib and pandas library
import pandas as pd 
import urllib
# Importing the the variations of Datetime modules to manipulate the date
from datetime import datetime
from datetime import date
from datetime import timedelta

#URL for PWOI is initialised
url = "https://www1.nseindia.com/content/nsccl/fao_participant_oi_08072020.csv"

#find the current date
now = datetime.now()
#convert date to a format ready to put into URL
today = now.strftime("%d%m%Y")

#The url is changed according to the day
new_url = url[0:59] + today + url[67: ]

#CSV file for that day is exctracted
data = pd.DataFrame(pd.read_csv(new_url))
#the data already on the excel sheet is extracted
sheet = pd.DataFrame(pd.read_excel('data.xlsx'))

#The new data is appended to the data on the Excel Sheet
updated_sheet = sheet.append(data)
