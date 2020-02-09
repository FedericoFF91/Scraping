import urllib.request
import pickle
from pynput import mouse
from pynput import keyboard
from ctypes import windll
from time import sleep
import pandas as pd
import pyautogui
import os
from bs4 import BeautifulSoup
import unicodedata
import datetime
import re
import matplotlib.pyplot as plt
import requests
from textblob import TextBlob
from itertools import repeat
from selenium import webdriver
import win32clipboard
import argparse
#from MongoClass import MongoDataHandling
import pyautogui
import datetime as dt
from sklearn.linear_model import LinearRegression
import statsmodels.tsa.stattools as ts
from functools import reduce
import requests
import pytz
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime
from pandas_datareader import DataReader
import fix_yahoo_finance as yf
from pandas_datareader import data as pdr
import calendar

###############################################
####### Define the class to scrape data #######
###############################################

class ScrapeHistoricData():

    def __init__(self):
        pass

    @classmethod
    def GetClipboardData(self):
        """
        Extracts the data currently saved in the clipboard and returns it in string format.

        Arguments:

        Returns:
        clipboard_data -- The data contained in the clipboard in string format
        """
        win32clipboard.OpenClipboard()
        clipboard_data = win32clipboard.GetClipboardData()
        win32clipboard.CloseClipboard()
        return(clipboard_data)
    
    @classmethod
    def DownloadFullHistoryDataSeries(self,Index,FinancialIndex=False,Commodity=False,Currency=False,Bond=False,Etf=False):

        # define functions to click and pause
        def pause(wait = 2):
            sleep(wait)

        def click():
            m.press(Button.left)
            pause()
            m.release(Button.left)

        def rclick():
            m.press(Button.right)
            pause()
            m.release(Button.right)
        
        m = mouse.Controller()
        Button = mouse.Button
        k = keyboard.Controller()
        Key = keyboard.Key

        # open browser and get the page, click on space in order to move down and 
        # visualize the object
        driver = webdriver.Firefox()
        if FinancialIndex == True:
            PageUrl = r'https://www.investing.com/indices/'+str(Index)+r'-historical-data'
        if Commodity == True:
            PageUrl = r'https://www.investing.com/currencies/'+str(Index)+r'-historical-data'
        if Currency == True:
            PageUrl = r'https://www.investing.com/quotes/'+str(Index)+r'-historical-data'
        if Bond == True:
            PageUrl = r'https://www.investing.com/rates-bonds/'+str(Index)+r'-historical-data'
        if Etf == True:
            PageUrl = r'https://www.investing.com/etfs/'+str(Index)+r'-historical-data'
        driver.get(PageUrl)
        pause(2)
        # move down 
        k.press(Key.space)
        k.release(Key.space)
        pause(2)
        # click on the DatePicker table in order
        # to open the tables to select the dates
        # ClickPosition = (824, 267)
        if FinancialIndex == True:
            ClickPosition = (825, 317) 
        if Commodity == True:
            ClickPosition = (824, 286) 
        if Currency == True:
            ClickPosition = (824, 269)             
        if Bond == True:
            ClickPosition = (825, 268)
        if Etf == True:
            ClickPosition = (824, 317)                          
        m.position = ClickPosition
        pause(1)
        click()
        # write the date into the start date DatePicker
        # do not need to write all the date, only the year that we want
        # in this case 2000 because all the time series will start from there
        pause(2)
        k.type('2001')
        pause()
        k.press(Key.enter)
        # click on Apply in order to made the changes
        #ApplyPosition = (806,470)
        if FinancialIndex == True:
            ApplyPosition = (802,526)
        if Commodity == True:
            ApplyPosition = (806,495)
        if Currency == True:
            ApplyPosition = (806,479)
        if Bond == True:
            ApplyPosition = (806,481)
        if Etf == True:
            ApplyPosition = (806,527)
        m.position = ApplyPosition
        pause(2)
        click()
        # Select and Copy to Clipboard
        # Used this way instead of use Beautiful Soup
        # because it did not work properly (it requires too much time to write that correctly)
        # so I decided to simply select everything and right click and copy to clipboard
        # after I will get only the historical table copied to clipboard
        pause(10)
        k.press(Key.ctrl)
        k.press("a")
        k.release(Key.ctrl)
        k.release("a")
        pause()
        CenterPage = (562,518)
        m.position = CenterPage
        rclick()
        pause()
        CopyPosition = (621,536)
        m.position = CopyPosition
        click()
        pause()
        ClipboardData = self.GetClipboardData()
        # Convert the Data copied into the clipboard to TextBlob
        FXTimeSeriesText = TextBlob(ClipboardData)
        # find the position of the first word in the table
        FirstWord = FXTimeSeriesText.find("Change %") +10
        # find the position of the last word
        LastWord = FXTimeSeriesText.find("Highest") - 2
        # select the table
        FXTable = FXTimeSeriesText[FirstWord:LastWord]
        # convert to a list
        ListFXTable = []
        for i in range(0,len(FXTable.words),8):
            ListFXTable.append(list(FXTable.words[i:(i+8)]))
        # convert to a PandaDataframe
        if Currency == True:
            FXTable = re.sub('\t- ','\tNA',str(FXTable))
            FXTable = TextBlob(FXTable)
            ListFXTable = []
            for i in range(0,len(FXTable.words),9):
                ListFXTable.append(list(FXTable.words[i:(i+9)]))
        if Etf == True:
            FXTable = re.sub('\t- ','\tNA',str(FXTable))
            FXTable = TextBlob(FXTable)
            ListFXTable = []
            for i in range(0,len(FXTable.words),9):
                ListFXTable.append(list(FXTable.words[i:(i+9)]))

        FXTab = pd.DataFrame(ListFXTable)
        # close the browser
        driver.close()
        # format the dataset
        # Copy the Dataset for Currency  (it's not beautiful but it's doing its work)
        FXTab2 = FXTab.copy()
        if Currency == True:
            FXTab2.columns = ['Date1','Date2','Date3','Price','Open','Max','Min','Volume','Var']
            FXTab2['Date'] = FXTab2['Date1'] + FXTab2['Date2'] + FXTab2['Date3']
            FXTab2['Date'] = pd.to_datetime(FXTab2['Date'], format="%b%d%Y", errors='coerce')
            FXTab2 = FXTab2[['Date','Price','Open','Max','Min']]
            FXTab2['Price'] = pd.to_numeric(FXTab2['Price'].str.replace(',',''), errors='coerce')
            FXTab2['Open'] = pd.to_numeric(FXTab2['Open'].str.replace(',',''), errors='coerce')
            FXTab2['Max'] = pd.to_numeric(FXTab2['Max'].str.replace(',',''), errors='coerce')
            FXTab2['Min'] = pd.to_numeric(FXTab2['Min'].str.replace(',',''), errors='coerce')
            return FXTab2
        if Etf == True:
            FXTab2.columns = ['Date1','Date2','Date3','Price','Open','Max','Min','Volume','Var']
            FXTab2['Date'] = FXTab2['Date1'] + FXTab2['Date2'] + FXTab2['Date3']
            FXTab2['Date'] = pd.to_datetime(FXTab2['Date'], format="%b%d%Y", errors='coerce')
            FXTab2 = FXTab2[['Date','Price','Open','Max','Volume','Min']]
            FXTab2['Price'] = pd.to_numeric(FXTab2['Price'].str.replace(',',''), errors='coerce')
            FXTab2['Open'] = pd.to_numeric(FXTab2['Open'].str.replace(',',''), errors='coerce')
            FXTab2['Max'] = pd.to_numeric(FXTab2['Max'].str.replace(',',''), errors='coerce')
            FXTab2['Min'] = pd.to_numeric(FXTab2['Min'].str.replace(',',''), errors='coerce')
            return FXTab2        
        FXTab.columns = ['Date1','Date2','Date3','Price','Open','Max','Min','Var']
        FXTab['Date'] = FXTab['Date1'] + FXTab['Date2'] + FXTab['Date3']
        FXTab['Date'] = pd.to_datetime(FXTab['Date'], format="%b%d%Y", errors='coerce')
        FXTab = FXTab[['Date','Price','Open','Max','Min']]
        FXTab['Price'] = pd.to_numeric(FXTab['Price'].str.replace(',',''), errors='coerce')
        FXTab['Open'] = pd.to_numeric(FXTab['Open'].str.replace(',',''), errors='coerce')
        FXTab['Max'] = pd.to_numeric(FXTab['Max'].str.replace(',',''), errors='coerce')
        FXTab['Min'] = pd.to_numeric(FXTab['Min'].str.replace(',',''), errors='coerce')
        # Save into the Database ('Same collection')
        # if Update.lower() in ('true','yes','t','y'):
        #     Tab = self.UpdateTimeSeries(FXTab,'MarketData',str(BaseCurrency).upper() +str(SecondCurrency).upper())
        # elif Update.lower() in ('false','no','f','n'):
        #     Tab = self.UploadTimeSeries(FXTab,'MarketData',str(BaseCurrency).upper() +str(SecondCurrency).upper())
        # else:
        #     raise argparse.ArgumentTypeError('Boolean Value Expected with '' ')            
        return FXTab


if __name__ == '__main__':

    # Initialize the scraper
    Scrap = ScrapeHistoricData()

    os.chdir(r'C:\Users\Federico\Documents\Trading\GoldAnalysis')
    # Download time series
    Gold = Scrap.DownloadFullHistoryDataSeries('xau-usd',Commodity=True)
    Gold.to_csv('GoldHistoricPrice.csv')
    CRBCommodityIndex = Scrap.DownloadFullHistoryDataSeries('thomson-reuters---jefferies-crb',FinancialIndex=True)
    CRBCommodityIndex.to_csv('CRBCommodityIndexHistoricPrice.csv')
    DXY = Scrap.DownloadFullHistoryDataSeries('us-dollar-index',Currency=True)
    DXY.to_csv('DXYHistoricPrice.csv')
    US10YearBondYield = Scrap.DownloadFullHistoryDataSeries('u.s.-10-year-bond-yield',Bond=True)
    US10YearBondYield.to_csv('US10YearBondYieldHistoricPrice.csv')
    GLDEtf = Scrap.DownloadFullHistoryDataSeries('spdr-gold-trust',Etf=True)
    GLDEtf.to_csv('GLDEtfHistoricPrice.csv')
    IAUEtf = Scrap.DownloadFullHistoryDataSeries('ishares-comex-gold-trust',Etf=True)
    IAUEtf.to_csv('IAUEtfHistoricPrice.csv')
    GDXEtf = Scrap.DownloadFullHistoryDataSeries('market-vectors-gold-miners',Etf=True)
    GDXEtf.to_csv('GDXEtfHistoricPrice.csv')
    DbUSDIndexEtf = Scrap.DownloadFullHistoryDataSeries('powershares-db-usd-index-bullish',Etf=True)
    DbUSDIndexEtf.to_csv('DbUSDIndexHistoricPrice.csv')
    CRBIndexEtf = Scrap.DownloadFullHistoryDataSeries('lyxor-commodities-crb-tr',Etf=True)
    CRBIndexEtf.to_csv('CRBIndexETFHistoricPrice.csv')
    TenYearNoteEtf = Scrap.DownloadFullHistoryDataSeries('ishares-7-10-year-treasury-bond',Etf=True)
    TenYearNoteEtf.to_csv('TenYearNoteETFHistoricPrice.csv')
    TenTwentyYearNoteEtf = Scrap.DownloadFullHistoryDataSeries('ishares-lehman-10-20-y-tr.-bond',Etf=True)
    TenTwentyYearNoteEtf.to_csv('TenTwentyYearNoteETFHistoricPrice.csv')
    USTIPSEtf = Scrap.DownloadFullHistoryDataSeries('pimco-1-5-year-us-tips-index-fund',Etf=True)
    USTIPSEtf.to_csv('USTIPSETFHistoricPrice.csv')
    USDJPYEtf = Scrap.DownloadFullHistoryDataSeries('rydex-currencyshares-japanese-yen',Etf=True)
    USDJPYEtf.to_csv('USDJPYETFHistoricPrice.csv')
    TwentyYearTreasuryEtf = Scrap.DownloadFullHistoryDataSeries('ishares-lehman-20-year-treas',Etf=True)
    TwentyYearTreasuryEtf.to_csv('TwentyYearTreasuryETFHistoricPrice.csv')


    


