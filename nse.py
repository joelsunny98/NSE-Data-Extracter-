# Importing the urllib and pandas library
import pandas as pd 
import urllib
# Importing the the variations of Datetime modules to manipulate the date
from datetime import datetime
from datetime import date
from datetime import timedelta

def pre_date():
    '''
    DESCRIPTION: This function will return the previous day's date
    INPUT: Null
    OUTPUT: Previous day's date in format ready to be used in URL
    '''
    # Get today's name in number
    myday = datetime.weekday(date.today())

    # Get today's date
    today = date.today()

    #Yesterday's date
    yesterday = str(today - timedelta(days = 1))

    # Extractting day month and year individually
    day = yesterday[8: ]
    month = yesterday[5:7]
    year = yesterday[0:4]

    # Finding out if the previous day was a weekend for which there is no data
    if myday == 0 or myday == 6:
        print("The market was closed Yesterday")
    else:
        # The market are open on weekdays
        # The date is return in format that can be put inbetween the URL
        return day + month + year

def extract(day):
    '''
    DESCRIPTION: Extracts data for the Participant wise Open Interest(PWOI)
    data from the NSE Website and updates it onto an excel file.
    INPUT: The CSV file from which you want the data
    OUTPUT: Null, it does not return any result but stores the extracted data
    on to the Excel Sheet.
    '''
    #URL For PWOI is initialised
    url = "https://www1.nseindia.com/content/nsccl/fao_participant_oi_08072020.csv"
    # The URL is changed accordingly to the required day
    new_url = url[0:59] + day + url[67: ]

    #CSV file for that day is extracted
    data = pd.DataFrame(pd.read_csv(new_url))
    # the data already on the excel sheet is extracted
    sheet = pd.DataFrame(pd.read_excel('data.xlsx'))
    
    #The new data is appended to the data on the excel sheet
    updated_sheet = sheet.append(data)

    # The updated data is stored on the excel sheet
    updated_sheet.to_excel('data.xlsx')
    print("...DATA EXTRACTION COMPLETE")
    print("Please check the file: data.xlsx")

    #Find the current date
    now = datetime.now()
    # Convert data into a format ready to be used in the url
    today = now.strftime("%d%m%Y")

    try: 
        extract(today)
    except:
        print("There is an error")



