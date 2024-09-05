
import pandas as pd
#-------------------
from selenium import webdriver
#from webdriver_manager.chrome import ChromeDriverManager # pour utiliser Google Chrome
from selenium.webdriver.common.keys import Keys # pour écrire du texte sur une page, utile pour se connecter sur le site
from selenium.webdriver import ActionChains
from selenium.common.exceptions import NoSuchElementException
import win32com.client as win32 # pour se conncecter à Outlook et envoyer le mails
import warnings # permet de lancer le disable.warnings ci-dessous
from selenium.webdriver.common.by import By
#import tia.bbg.datamgr as dm
#-------------------
from datetime import timedelta
from datetime import time
from datetime import date
from datetime import datetime
import time as t
#-------------------
from openpyxl import load_workbook
#-------------------
import matplotlib.pyplot as plt
from matplotlib.pyplot import figure

import os
import glob
from openpyxl import load_workbook

warnings.filterwarnings("ignore") # évite que la console soit polluée par d'inutiles warnings

# proxies = {
#     "http":"http://rduong:Saberqueenbieryo60&@proxy60-2.oddo.fr:8080",
#     "https":"https://rduong:Saberqueenbieryo60&@proxy60-2.oddo.fr:8080"
# }


def color_arrows(dataframe):

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



def german_shorts():
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory' : r"C:\Users\33659\Desktop\Net shorts\Data Bases\GE"}
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument("--disable-search-engine-choice-screen")
    chrome = webdriver.Chrome(options=chrome_options)

    general_link = "https://www.bundesanzeiger.de/pub/en/to_nlp_start?0"
    chrome.get(general_link)
    chrome.find_element(By.ID, "cc_all").click()
    chrome.find_element(By.ID, "select2-id1-container").click()
    chrome.find_element(By.XPATH, "//span[@class='select2-selection select2-selection--single']").send_keys(Keys.DOWN, Keys.DOWN, Keys.ENTER)

    if int(datetime.today().weekday()) == 0:
        yesterday_date = datetime.strftime(datetime.today().date() - timedelta(days=3), "%Y-%m-%d")
    else:
        yesterday_date = datetime.strftime(datetime.today().date() - timedelta(days=1), "%Y-%m-%d")

    rows = []
    for short_row in chrome.find_elements(By.XPATH, "//a[@class='intern']"):
        ActionChains(chrome).key_down(Keys.CONTROL).click(short_row).key_up(Keys.CONTROL).perform()

    tabs_nb = len(chrome.window_handles)
    for tab_num in range(tabs_nb - 1, 1, -1):
        chrome.switch_to.window(chrome.window_handles[tab_num])
        t.sleep(0.5)

        short_date_element = chrome.find_element(By.XPATH, "//div[@class='col-td-5 nlp-datum']")
        short_date = short_date_element.text
        short_date = datetime.strptime(short_date, "%Y-%m-%d")
        if datetime.strptime(yesterday_date, "%Y-%m-%d") != short_date:
            continue

        try:
            cur_row = chrome.find_element(By.XPATH, "//div[@class='row even']")
        except NoSuchElementException:
            cur_row = chrome.find_element(By.XPATH, "//div[@class='row odd history-result']")
        
        isin = cur_row.find_element(By.XPATH, "//div[@class='col-td-3']").text
        company_name = cur_row.find_element(By.XPATH, "//div[@class='col-td-2']").text
        holder = cur_row.find_element(By.XPATH, "//div[@class='col-td-1']").text
        current_pos = cur_row.find_element(By.XPATH, "//div[@class='col-td-4 nlp-position']").text[:-2]
        previous_pos = 0.0
        previous_date = "NEW"

        if cur_row.find_elements(By.XPATH, "//div[@class='col-td-4 nlp-position']"):
            previous_pos_elements = cur_row.find_elements(By.XPATH, "//div[@class='col-td-4 nlp-position']")
            previous_pos = previous_pos_elements[1].text[:-2]
            previous_date_elements = cur_row.find_elements(By.XPATH, "//div[@class='col-td-5 nlp-datum']")
            previous_date = previous_date_elements[1].text
        
        variation = float(current_pos) - float(previous_pos)
        direction = "NEW" if float(previous_pos) == 0 else ("↗" if variation > 0 else "↘")

        name_to_look_for = cur_row.find_element(By.XPATH, "//div[@class='col-td-2']").text
        chrome.switch_to.window(chrome.window_handles[0])
        chrome.refresh()

        chrome.find_element(By.XPATH, "//input[@placeholder='position holder, issuer name, or ISIN']").clear()
        chrome.find_element(By.XPATH, "//input[@placeholder='position holder, issuer name, or ISIN']").send_keys(name_to_look_for)
        chrome.find_element(By.XPATH, "//input[@class='btn btn-green']").click()
        t.sleep(5)

        total_pos = 0.0
        name_of_holders = []
        holder_count = int(chrome.find_element(By.XPATH, '//*[@id="content"]/section[2]/div/div/div/div/div[5]/div[1]').text[0])
        for holder_nb in range(holder_count):
            try:
                total_pos += float(chrome.find_elements(By.XPATH, "//div[@class='col-td-4 nlp-position']")[holder_nb].text[:-2])
            except IndexError:
                pass
            else:
                name = chrome.find_elements(By.XPATH, "//div[@class='col-td-1']")[holder_nb].text
                if name in name_of_holders:
                    total_pos -= float(chrome.find_elements(By.XPATH, "//div[@class='col-td-4 nlp-position']")[holder_nb].text[:-2])
                else:
                    name_of_holders.append(name)

        if float(previous_pos) < 0.5:
            total_pos = float(total_pos)
        
        row_to_add = {
            "Company Name": company_name.upper(),
            "Position Date": datetime.strftime(datetime.strptime(yesterday_date, "%Y-%m-%d"), "%d-%b-%y"),
            "Position Holder": holder.upper(),
            "Previous Position": previous_pos,
            "Net short position": current_pos,
            "Previous Position Date": previous_date,
            "Side": direction,
            "Total short position": total_pos,
            "Variation": variation,
            "ISIN": isin
        }
        rows.append(row_to_add)

    chrome.quit()

    if not rows:
        print("No data found for the specified date.")
        return None

    df = pd.DataFrame(rows)

    try:
        df["Total variation"] = df["Variation"].groupby(df["Position Holder"]).transform("sum").apply(lambda x: round(x, 2))
    except KeyError:
        return df

    columns_in_order = [
        
        "ISIN",
        "Company Name",
        "Position Holder",
        "Previous Position Date",
        "Previous Position",
        "Net short position",
        "Side",
        "Variation",
        "Total short position",
        "Position Date",

    ]

    df = df[columns_in_order].sort_values("Company Name")

    # Optional: apply styling if working in an environment that supports it
    df_style = df.style.apply(color_arrows, subset=pd.IndexSlice[:, ["Side", 'Variation']])
    df_style = df_style.apply(align_center, subset=pd.IndexSlice[:, ["Previous Position Date", 'Previous Position', 'Net short position', 'Side', 'Variation', 'Total short position', 'Position Date']])

    df_style.to_excel(r"C:\Users\33659\Desktop\Net shorts\Daily Report\GE.xlsx", index=False)

    # Returning the DataFrame as output
    return df_style