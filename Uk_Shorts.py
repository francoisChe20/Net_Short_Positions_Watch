import pandas as pd

import copy
#-------------------
import requests
from bs4 import BeautifulSoup
import io
#import tia.bbg.datamgr as dm
#-------------------
from datetime import timedelta
from datetime import time
from datetime import date
from datetime import datetime
import time as t
import locale
from datetime import date
#-------------------
import os
import glob
from openpyxl import load_workbook
#-------------------
import re

import PyPDF2
from fuzzywuzzy import process
import unidecode
import warnings
#-------------------
import matplotlib.pyplot as plt
from matplotlib.pyplot import figure

from selenium import webdriver # pour pouvoir surfer sur l'internet
#from webdriver_manager.chrome import ChromeDriverManager # pour utiliser Google Chrome
from selenium.webdriver.common.keys import Keys # pour écrire du texte sur une page, utile pour se connecter sur le site
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By


warnings.filterwarnings("ignore")


def get_uk_database() :

    print("Launching...")

    files = glob.glob(r"C:\Users\33659\Desktop\Net shorts\Data Bases\UK\\*")
    for f in files:
        os.remove(f) # On enlève tous les fichiers précédents 
    
    lien = f"https://www.fca.org.uk/markets/short-selling/notification-and-disclosure-net-short-positions"
    
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory' : r"C:\Users\33659\Desktop\Net shorts\Data Bases\UK"}
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument("--disable-search-engine-choice-screen")

    chrome = webdriver.Chrome(options=chrome_options)
    chrome.get(lien)
    
    t.sleep(1)
    chrome.find_element(By.XPATH, '//*[@id="cookiepopup"]/div/div[2]/button[1]').click()
    t.sleep(1)
    chrome.find_element(By.XPATH,'//*[@id="section-end-of-the-transition-period"]/div/p[7]/a').click()
    t.sleep(2)


    chrome.quit()

def color_arrows(dataframe):
# Cette fonction sera utilisée ensuite pour le style du dataframe
# de reporting. Le style de l'objet renvoyé (et notamment la
# syntaxe "c: red", qui rappelle celle d'un dictionnaire
# sans y coller parfaitement) peut sembler étonnant, mais il
# est important de le conserver, cf docu sur pandas Styler
    list_of_styles = []
    for elem in dataframe:
        if elem == '↗':
            list_of_styles.append("color: red")
        elif elem == '↘':
            list_of_styles.append("color: #1DAD1D")
        else:
            try:
                var = float(elem)
            except:
                list_of_styles.append("background-color: yellow")
            else:
                if var > 0:
                    list_of_styles.append("color: red")
                elif var < 0:
                    list_of_styles.append("color: #1DAD1D")
                else:
                    list_of_styles.append("color: #9A9A9A")
    return list_of_styles


def align_center(df):
    list_of_styles = []
    for elem in df:
        list_of_styles.append('text-align: center')

    return list_of_styles

def filter(x):
    x=re.sub(",","",x)
    x=re.sub("  "," ",x)
    x=re.sub("/.","",x)
    return x


def find_last_pos(stock, fund, df_histo):
    target_df = df_histo[(df_histo['Position Holder'] == fund) & (df_histo['Name of Share Issuer'] == stock)]
    if target_df.shape[0] == 0:
        last_pos = 0
        last_date = 'NEW'
    else:
        last_date = df_histo[(df_histo['Position Holder'] == fund) & (df_histo['Name of Share Issuer'] == stock)]['Position Date'].max()
        last_pos = df_histo[(df_histo['Position Holder'] == fund) & (df_histo['Name of Share Issuer'] == stock) & (df_histo['Position Date'] == last_date)]['Net Short Position (%)'].values[0]
        if last_pos < 0.5 and (datetime.today().date() - last_date.date()).days > 90:
            last_pos = 0
            last_date = 'NEW'

    return [last_pos, last_date]




def report_short_uk():

    isin = []
    target_stock = []
    holder = []
    last_pos_date = []
    last_pos = []
    new_pos = []
    side = []
    var = []
    total_short_pos = []


    day = datetime.today().weekday()
    if day == 0:
        yesterday_date = datetime.today().date() - timedelta(days=3)
    elif day == 6:
        yesterday_date = datetime.today().date() - timedelta(days=2)
    else:
        yesterday_date = datetime.today().date() - timedelta(days=1)



    df = pd.read_excel(r"C:\Users\33659\Desktop\Net shorts\Data Bases\UK\short-positions-daily-update.xlsx", sheet_name = 'Current Disclosures ' + datetime.strftime(yesterday_date, "%d.%m.%Y"))
    df_histo = pd.read_excel(r"C:\Users\33659\Desktop\Net shorts\Data Bases\UK\short-positions-daily-update.xlsx", sheet_name = 'Historic Disclosures ' + datetime.strftime(yesterday_date, "%d.%m.%Y"))

    max_date = df["Position Date"].max()
    df_current = df[df["Position Date"] == max_date].reset_index(drop=True)


    for i in df_current.index:
        stock = df_current['Name of Share Issuer'][i]
        fund = df_current['Position Holder'][i]
        isin.append(df_current['ISIN'][i])
        target_stock.append(stock)
        holder.append(fund)
        previous_pos, previous_pos_date = find_last_pos(stock, fund, df_histo)
        last_pos_date.append(previous_pos_date)
        last_pos.append(previous_pos)
        new_pos.append(df_current['Net Short Position (%)'][i])
        
        var.append((df_current['Net Short Position (%)'][i] - previous_pos))
        if (df_current['Net Short Position (%)'][i] - previous_pos) > 0:
            if previous_pos_date == 'NEW':
                side.append('NEW')
            else:
                side.append('↗')
        else:
            side.append('↘')

        total_short_pos.append(df[df['Name of Share Issuer']==df_current['Name of Share Issuer'][i]]['Net Short Position (%)'].sum())

    df_result = pd.DataFrame({})
    df_result['ISIN'] = isin
    df_result['Company Name']= target_stock
    df_result['Company Name'] = df_result['Company Name'].str.upper()
    df_result['Position Holder']= holder
    df_result['Position Holder'] = df_result['Position Holder'].str.upper()
    df_result['Previous Position Date'] = last_pos_date
    df_result['Previous Position'] = last_pos
    df_result['Net Short Position'] = new_pos
    df_result['Side'] = side
    df_result['Change'] = var
    df_result['Aggregate Net Short Positions'] = total_short_pos

    df_style = df_result.style.apply(color_arrows, subset=pd.IndexSlice[:, ['Side', 'Change']])

    # Apply align_center styling to specific columns (remove duplicate 'Previous Position Date')
    df_style = df_style.apply(align_center, subset=pd.IndexSlice[:, ['Previous Position Date', 'Net Short Position', 'Previous Position', 'Side', 'Change', 'Aggregate Net Short Positions']])

    # Save the styled DataFrame to an Excel file
    df_style.to_excel(r"C:\Users\33659\Desktop\Net shorts\Daily Report\UK.xlsx", index=False)

    # Return the original DataFrame (or df_style if you want the styled version)
    return df_style