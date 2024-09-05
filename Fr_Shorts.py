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

# proxies = {
#     "http":"http://LOGIN:PASSWD@proxy60-2.oddo.fr:8080",
#     "https":"https://LOGIN:PASSWD@proxy60-2.oddo.fr:8080"
# }

def get_french_database() :

    print("Launching...")
    files = glob.glob(r"C:\Users\33659\Desktop\Net shorts\Data Bases\FR\\*")
    for f in files:
        os.remove(f) # On enlève tous les fichiers précédents 
    
    lien = f"https://www.amf-france.org/fr/actualites-publications/dossiers-thematiques/ventes-decouvert"
    # Les trois arguments proxies, verify et stream sont indispensables (en tout cas
    # au moment où je rédige cet algo) pour pouvoir effectuer une requête internet
    # malgré le pare-feu d'Oddo BHF.
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory' : r"C:\Users\33659\Desktop\Net shorts\Data Bases\FR"}
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument("--disable-search-engine-choice-screen")
    # chrome_options.add_experimental_option("excludeSwitches", ['enable-automation'])
    # chrome_options.add_argument("disable-infobars")
    # chrome_options.add_argument('--proxy-server=%s' % proxies)
    chrome = webdriver.Chrome(options=chrome_options)
    chrome.get(lien)
    
    t.sleep(2)
    try :
        chrome.find_element(By.XPATH, '//*[@id="tarteaucitronPersonalize2"]').click()
    except :
        pass

    t.sleep(2)
    chrome.find_element(By.XPATH,'//*[@id="block-amf-content"]/div[2]/div/div[2]/div/div[2]/div[3]/div[5]/div/a/div').click()
    t.sleep(2)


    lien_folder_base = r"C:\Users\33659\Desktop\Net shorts\Data Bases\FR"
    for file in os.listdir(lien_folder_base):
        os.rename(lien_folder_base + "\\" + file, lien_folder_base + "\\" + f"base_du_jour_fr.csv" )

    chrome.quit()
    
    
def get_french_shorts_online():
    print("Launching...")
    #os.chdir(r"O:\securities\users\$FR SALES TRADER\STAGIAIRE\Produits Stagiaire\DOCS UTILES POUR LE STAGIAIRE\MACROS_DAILY\Notebooks Python")
    today = datetime.today()
    #today = date(2022, 9, 16)
    today = today.strftime("%d/%m/%Y")[0:2]
    debut = datetime.now()
    mtn = str(datetime.today())
    #mtn = str(date(2022, 9, 16))
    ajd = mtn[:mtn.find(" ")]
    jour = int(mtn[8:10])
    mois = mtn[5:7]
    an = mtn[:4]
    files = glob.glob(r"C:\Users\33659\Desktop\Net shorts\Data Bases\Fr_short_files\\*")
    for f in files:
        os.remove(f) # On enlève tous les fichiers précédents 
    if datetime.today().weekday() == 0:
    #if date(2022, 9, 16).weekday() == 0:
        jour -= 3
    else:
        jour -= 1
    if len(str(jour)) == 1:
        jour = "0" + str(jour)
    
    lien = f"https://bdif.amf-france.org/fr?dateDebut={an}-{mois}-{today}&dateFin={an}-{mois}-{today}&typesInformation=VAD"
       
    

    # Les trois arguments proxies, verify et stream sont indispensables (en tout cas
    # au moment où je rédige cet algo) pour pouvoir effectuer une requête internet
    # malgré le pare-feu d'Oddo BHF.
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory' : r"C:\Users\33659\Desktop\Net shorts\Data Bases\Fr_short_files"}
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument("--disable-search-engine-choice-screen")
    # chrome_options.add_experimental_option("excludeSwitches", ['enable-automation'])
    # chrome_options.add_argument("disable-infobars")
    # chrome_options.add_argument('--proxy-server=%s' % proxies)
    chrome = webdriver.Chrome(options=chrome_options)

    chrome.get(lien)
    
    #on récupère le nombre de Shorts 
    while True:
        try :
            chrome.find_element(By.XPATH,"/html/body/app-root/div/div/main/app-home-container/section/app-results-container/div[1]/h2")
            if chrome.find_element(By.XPATH,"/html/body/app-root/div/div/main/app-home-container/section/app-results-container/div[1]/h2").is_displayed():
                Nombre_de_shorts = chrome.find_element(By.XPATH,"/html/body/app-root/div/div/main/app-home-container/section/app-results-container/div[1]/h2").text
                break
        except BaseException:
            t.sleep(0.5)
    
    #Accept cookies
    try:
        if chrome.find_element(By.XPATH,"/html/body/div[2]/div[3]/button[1]").is_displayed():
            chrome.find_element(By.XPATH,"/html/body/div[2]/div[3]/button[1]").click()
        else :
            t.sleep(1.5)
    except:
        pass


        
    Nombre_de_shorts= Nombre_de_shorts.replace("RÉSULTATS", "")
    print(Nombre_de_shorts)
    Nombre_de_shorts = int(Nombre_de_shorts)

    if Nombre_de_shorts > 20 :
        chrome.find_element(By.XPATH,"//*[@id='home-container']/app-results-container/div[2]/div[2]/div/button").click()

    lien_folder_shorts = r"C:\Users\33659\Desktop\Net shorts\Data Bases\Fr_short_files\\"
    os.chdir(lien_folder_shorts)
    nbr_download =0 
    height = 0
    i=1
   
    
    while nbr_download != Nombre_de_shorts :
        try:
            t.sleep(1)
            chrome.find_element(By.XPATH,"/html/body/app-root/div/div/main/app-home-container/section/app-results-container/div[2]/div/ul/li["+str(i)+"]/div/app-result-list-view/div/mat-card/mat-card-actions/button").click()
            nbr_download+=1
            #print(nbr_download)
            while len(os.listdir(r"C:\Users\33659\Desktop\Net shorts\Data Bases\Fr_short_files")) != nbr_download:
                t.sleep(1)
            i+=1
        except BaseException:
            try :
                chrome.execute_script(f"window.scrollTo(0,{height})")
                height= height + 72
                t.sleep(0.3)
                if chrome.find_element(By.XPATH,"/html/body/app-root/div/div/app-home-container/section/app-results-container/div[2]/div[2]/div/a").is_displayed():
                    chrome.find_element(By.XPATH,"/html/body/app-root/div/div/app-home-container/section/app-results-container/div[2]/div[2]/div/a").click()
                    t.sleep(0.3)
                else:
                    continue
            except NoSuchElementException :
                break  
   
    page_nb = 1
    for file in os.listdir(lien_folder_shorts):
        os.rename(lien_folder_shorts + "\\"+ file, lien_folder_shorts + "\\"+ f"Short n°{page_nb}.pdf" )
        page_nb+=1
    print('Nombre de shorts téléchargés: '+ str(page_nb))
    print(f"TEMPS D'EXEC : {datetime.now() - debut}, les Shorts ont bien été enregistrés")
    chrome.quit()

  


def align_center(df):
    list_of_styles = []
    for elem in df:
        list_of_styles.append('text-align: center')

    return list_of_styles


def color_arrows(dataframe):
# Cette fonction sera utilisée ensuite pour le style du dataframe
# de reporting. Le style de l'objet renvoyé (et notamment la
# syntaxe "c: red", qui rappelle celle d'un dictionnaire
# sans y coller parfaitement) peut sembler étonnant, mais il
# est important de le conserver, cf docu sur pandas Styler
    list_of_styles = []
    for elem in dataframe:
        print(elem)
        if elem == '↗':
            print("fleche haute")
            list_of_styles.append("color: red")
        elif elem == '↘':
            print("fleche basse")
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
    print(list_of_styles)
    return list_of_styles
 
def date_to_us_format(date, asdate=False):
    try:
        if asdate:
            return(datetime.strptime(str(date)[:10], "%d/%m/%Y"))
        else:
            return(datetime.strftime(datetime.strptime(str(date)[:10], "%d/%m/%Y"), "%Y-%m-%d"))
    except ValueError:
        try:
            if asdate:
                return(datetime.strptime(str(date)[:10], "%d-%b-%y"))
            else:
                return(datetime.strftime(datetime.strptime(str(date)[:10], "%d-%b-%y"), "%Y-%m-%d")) 
        except ValueError:
            return(str(date))
    
def date_to_xl_format(date, asdate=False):
    try:
        if asdate:
            return(datetime.strptime(str(date)[:10], "%d-%b-%y"))
        else:
            return(datetime.strftime(datetime.strptime(str(date)[:10], "%Y-%m-%d"), "%d-%b-%y"))
    except ValueError:
        return(str(date))
 
def date_to_fr_format(date, asdate=False):
    try:
        if asdate:
            return(datetime.strptime(str(date)[:10], "%Y-%m-%d"))
        else:
            return(datetime.strftime(datetime.strptime(str(date)[:10], "%Y-%m-%d"), "%d-%m-%Y")) 
    except ValueError:
        try:
            if asdate:
                return(datetime.strptime(str(date)[:10], "%d-%b-%y"))
            else:
                return(datetime.strftime(datetime.strptime(str(date)[:10], "%d-%b-%y"), "%d-%m-%Y")) 
        except ValueError:
            return(str(date))
          
def filter(x):
    x=re.sub(",","",x)
    x=re.sub("  "," ",x)
    x=re.sub("/.","",x)
    return x
 
def position_totale(valeur, N_caractères_noms):
 
    valeur =str(valeur)
    if len(valeur)>N_caractères_noms:
        valeur = valeur[:N_caractères_noms]
    
    df = pd.read_csv(r"C:\Users\33659\Desktop\Net shorts\Data Bases\FR\base_du_jour_fr.csv", sep=";")
    liste_de_detenteurs = df['Emetteur / issuer'].unique()
    dict = {}
    for i in range(len(liste_de_detenteurs)):
        dict[liste_de_detenteurs[i]] = liste_de_detenteurs[i][:N_caractères_noms]
 
    df.replace(dict, inplace = True)
 
    for column in df.columns:
        if "Unnamed" in column:
            df = df.drop(column,axis=1)
 
    print(list(df.columns))
 
    df_valeur=df[df["Emetteur / issuer"]==valeur]
    df_valeur["Detenteur de la position courte nette"]=df_valeur["Detenteur de la position courte nette"].apply( lambda x : filter(x))
    df_valeur=df_valeur.drop_duplicates(subset=['Detenteur de la position courte nette'])
    totale_pos =df_valeur[df_valeur["Emetteur / issuer"]==valeur]["Ratio"].apply(lambda x: float(x) if float(x) >= 0.5 else 0).sum()
    
    return totale_pos




# Cet algo a pour but de rajouter les données de la veille au
# fichier d'archives, lire dans les fichiers des shorts les pdf,
# en extraire les infos, et ajouter les infos supplémentaires
# dont on a besoin pour le reporting quotidien.
 
def shorts_fr(first_time=True, N_caractères_noms = 15):
    
    dict_exceptions = {"BG Master Fund ICAV" : "B&G MASTER FUND PLC", "BG MASTER FUND ICAV" : "B&G MASTER FUND PLC"}
 
    Base_de_données = pd.read_csv(r"C:\Users\33659\Desktop\Net shorts\Data Bases\FR\base_du_jour_fr.csv", sep=";")
 
    Base_de_données["Detenteur de la position courte nette"] = [x.replace(",", "") for x in Base_de_données["Detenteur de la position courte nette"]]
    Base_de_données["Emetteur / issuer"] = [x.replace(",", "") for x in Base_de_données["Emetteur / issuer"]]
 
    Base_de_données.replace(dict_exceptions, inplace = True)
 
    df1 = pd.read_csv(r"C:\Users\33659\Desktop\Net shorts\Data Bases\FR\base_du_jour_fr.csv", sep=";")
    df1["Detenteur de la position courte nette"] = [x.replace(",", "") for x in df1["Detenteur de la position courte nette"]]
    df1["Emetteur / issuer"] = [x.replace(",", "") for x in df1["Emetteur / issuer"]]
    df1.replace(dict_exceptions, inplace = True)
    liste_de_detenteurs = df1['Detenteur de la position courte nette'].unique()
    dict_detenteurs = {}
    dict_detenteurs_invert = {}
    for i in range(len(liste_de_detenteurs)):
        dict_detenteurs[liste_de_detenteurs[i]] = str(liste_de_detenteurs[i][:N_caractères_noms]).upper()
        dict_detenteurs_invert[str(liste_de_detenteurs[i][:N_caractères_noms]).upper()] = liste_de_detenteurs[i]
 
    Base_de_données.replace(dict_detenteurs, inplace = True)
 
    df1 = pd.read_csv(r"C:\Users\33659\Desktop\Net shorts\Data Bases\FR\base_du_jour_fr.csv", sep=";")
    df1["Detenteur de la position courte nette"] = [x.replace(",", "") for x in df1["Detenteur de la position courte nette"]]
    df1["Emetteur / issuer"] = [x.replace(",", "") for x in df1["Emetteur / issuer"]]
    df1.replace(dict_exceptions, inplace = True)
    liste_de_detenteurs = df1['Detenteur de la position courte nette'].unique()
    for i in range(len(liste_de_detenteurs)):
        liste_de_detenteurs[i] = str(liste_de_detenteurs[i][:N_caractères_noms]).upper()




    df2 = pd.read_csv(r"C:\Users\33659\Desktop\Net shorts\Data Bases\FR\base_du_jour_fr.csv", sep=";")
    df2["Detenteur de la position courte nette"] = [x.replace(",", "") for x in df2["Detenteur de la position courte nette"]]
    df2["Emetteur / issuer"] = [x.replace(",", "") for x in df2["Emetteur / issuer"]]
    df2.replace(dict_exceptions, inplace = True)
    liste_de_issuers = df2['Emetteur / issuer'].unique()
    dict_issuers = {}
    dict_issuers_invert = {}
    for i in range(len(liste_de_issuers)):
        dict_issuers[liste_de_issuers[i]] = str(liste_de_issuers[i][:N_caractères_noms]).upper()
        dict_issuers_invert[str(liste_de_issuers[i][:N_caractères_noms]).upper()] = liste_de_issuers[i]
 
    Base_de_données.replace(dict_issuers, inplace = True)
 
    df2 = pd.read_csv(r"C:\Users\33659\Desktop\Net shorts\Data Bases\FR\base_du_jour_fr.csv", sep=";")
    df2["Detenteur de la position courte nette"] = [x.replace(",", "") for x in df2["Detenteur de la position courte nette"]]
    df2["Emetteur / issuer"] = [x.replace(",", "") for x in df2["Emetteur / issuer"]]
    df2.replace(dict_exceptions, inplace = True)
    liste_de_issuers = df2['Emetteur / issuer'].unique()
    for i in range(len(liste_de_issuers)):
        liste_de_issuers[i] = str(liste_de_issuers[i][:N_caractères_noms]).upper()
 
    #mgr = dm.BbgDataManager()
    columns = [
               "Date",
               "Name",
               "Détenteur de la position",
               "Date précédente",
               "Position précédente",
               "Nouvelle position",
               "Sens de variation",
               "Position totale",
               "Variation",
               "Issuer vs Holder"
              ]
    df = pd.DataFrame(data={}, columns=columns)
    directory = "C:\\Users\\33659\\Desktop\\Net shorts\\Data Bases\\Fr_short_files\\"
    for file in os.listdir(directory):
        filename = os.fsdecode(file)
        print(filename)
        
        reader = PyPDF2.PdfReader(str(directory)[:] + filename)#PyPDF2.PdfFileReader(str(directory)[:] + filename)
        texte_utilisable = reader.pages[0].extract_text()#reader.getPage(0).extractText()
        texte_utilisable = texte_utilisable[texte_utilisable.rfind("AMF"):][N_caractères_noms:]
        texte_utilisable = texte_utilisable.replace(",","")
        texte_utilisable = texte_utilisable.replace("\n","")
        
        if "BG Master Fund ICAV" in texte_utilisable:
            texte_utilisable = texte_utilisable.replace("BG Master Fund ICAV", "B&G MASTER FUND PLC")
 
        print(texte_utilisable)
        for detenteur in liste_de_detenteurs:
            if (str(detenteur) in texte_utilisable) & (pd.notnull(detenteur)):
 
                texte_utilise = texte_utilisable[len(detenteur):]
 
                print(texte_utilise)
                
                isin = str(re.search('(FR|GB|US|LU|BE|NL)[0-9]{2}(.*?)[.]', texte_utilise).group(0))[:-2]
                print(isin)
               
                
                try :
                    issuer = Base_de_données[Base_de_données['code ISIN']==isin]["Emetteur / issuer"].unique()
                   # ticker = mgr["/isin/" + isin]["EQY_FUND_TICKER"]
                    #issuer = mgr[str(ticker) + " Equity"]["Name"]
                    print(f"Le nom de l'issuer est:{issuer}")
                except :
                    break
                    #issuer = mgr[str(ticker) + " Equity"]["Name"]
 
                    
                if issuer == "MCPHY ENERGY SA":
                    issuer = "MC PHY ENERGY"
                elif issuer == "CGG SA":
                    issuer = "CGG"
                elif issuer == "SHOWROOMPRIVE":
                    issuer = "SRP GROUPE"
                elif issuer == "CASINO GUICHARD PERRACHON":
                    issuer = "CASINO GUICHARD-PERRACHON"
 
                if len(issuer)>N_caractères_noms:
                    issuer = issuer[:N_caractères_noms]
      
 
                for i in range(Base_de_données.shape[0]) :
                    if Base_de_données.loc[i, "Emetteur / issuer"] in issuer :  # and Base_de_données.loc[i,"Detenteur de la position courte nette"] == detenteur
                        issuer = Base_de_données.loc[i, "Emetteur / issuer"]
                        print("on a bien changé l'issuer")
                        break
                
                print(issuer)
                
                #############
                # if issuer  == "SES":
                #     issuer = "SES IMAGOTAG"
 
                if issuer == "S.O.I.T.E.C.":
                    issuer = "SOITEC"
 
                if issuer ==  "EUROFINS SCIENTIFIC":
                    issuer = "EUROFINS SCIENTIFIC SE"
                ######################
 
                nouvelle_pos = re.search('[0-9]\.[0-9]{2}', texte_utilise).group(0)
                new_date = re.search('[0-9]{4}-[0-9]{2}-[0-9]{2}', texte_utilise[-10:]).group(0)
                print(new_date)
 
                value_df = Base_de_données[Base_de_données["Detenteur de la position courte nette"] == detenteur][Base_de_données["Emetteur / issuer"] == issuer]        
 
                try :
                    Date_fin = str(value_df["Date de fin de publication position"].iloc[0])
                    print(Date_fin)
                    if Date_fin != "nan":
                        try :
                            Date_fin = datetime.strptime(Date_fin, "%d/%m/%Y")
                        except :
                            Date_fin = datetime.strptime(Date_fin, "%Y-%m-%d")
                        if datetime.today()-timedelta(days = 90) > Date_fin :
                            previous_date = "NEW"
                        else :
                            previous_date = str(value_df["Date de debut position"].iloc[0])
                        
                    else :
                        previous_date=str(value_df["Date de debut position"].iloc[0])
                        print(str(previous_date) + " Date histo excel")
 
                except IndexError:
                    try :
                        previous_date= mgr[str(ticker) + " Equity"]["SHORT_HOLDINGS_BY_SECURITY"].set_index("Holder Name").loc[detenteur, "Filing Date"]
                        print(str(previous_date) + " Date bloom")
                    except :         
                        previous_date = "NEW"
            
 
                try:
                    previous_pos = value_df[value_df["Date de debut position"] == previous_date]["Ratio"].iloc[0]
                except BaseException:
                    try :
                        previous_pos= -round(mgr[str(ticker) + " Equity"]["SHORT_HOLDINGS_BY_SECURITY"].set_index("Holder Name").loc[detenteur, "Percent Outstanding"], 2)
                    except BaseException:
                        previous_pos = 0
                
                total_pos = position_totale(issuer, N_caractères_noms)
               
                
 
        
                
                variation = float(nouvelle_pos) - float(previous_pos)
                
                if float(previous_pos) < 0.5:
                    variation = float(nouvelle_pos)
 
                if previous_pos == float(0):
                    sens = "NEW"
                elif variation < 0:
                    sens = "↘"
                else:
                    sens = "↗"
                if previous_date != "NEW":
                    print(previous_date)
                    print("Changing date!")
                    print(previous_date)
                if total_pos =="":
                    total_pos = nouvelle_pos
                
                try :
                    print(dict_detenteurs_invert[detenteur])
                    detenteur_fin = dict_detenteurs_invert[detenteur]
                except KeyError :
                    print("KeyError detenteur, trouver dans base_du_jour_fr.csv")
                    print(detenteur)
                    detenteur_fin = detenteur
                try :
                    print(dict_issuers_invert[issuer])
                    issuer_fin = dict_issuers_invert[issuer]
                except KeyError :
                    print("KeyError issuer, trouver dans base_du_jour_fr.csv")
                    print(issuer)
                    issuer_fin = issuer
 
                issuer_vs_holder = str(issuer) + " vs " + str(detenteur) 
 
                
                ligne_a_ajouter = {#"Ticker": ticker,
                                "Date": new_date,
                                "Isin": isin,
                                "Name": issuer_fin,
                                "Détenteur de la position": detenteur_fin,
                                "Date précédente": previous_date,
                                "Position précédente": previous_pos,
                                "Nouvelle position": nouvelle_pos,
                                "Sens de variation": sens,
                                "Position totale": total_pos,
                                "Variation": variation,
                                "Issuer vs Holder": issuer_vs_holder,
                                }
            
                df = pd.concat([df, pd.DataFrame([ligne_a_ajouter])], ignore_index=True)
 
                break
 
    df["Variation totale"] = df["Variation"].groupby(df["Name"]).transform("sum").apply(lambda x: round(x, 2))
    df["Position totale"] += df["Variation totale"]
 
    final_columns = ['Isin',
                     'Date',
                     'Name',
                     'Détenteur de la position',
                     'Date précédente',
                     'Position précédente',
                     'Nouvelle position',
                     'Sens de variation',
                     'Position totale',
                     'Variation totale',
                     'Variation',
                     'Issuer vs Holder']
    print(df)
    df1 = df[final_columns]
    df1 = df.sort_values("Name")
 
    df1 = df1.drop(["Issuer vs Holder","Isin","Date","Variation"],axis=1)
    df1 = df1.rename(columns={"Name":"Company Name","Détenteur de la position": "Position Holder","Date précédente":"Previous Position Date","Position précédente":"Previous Position","Nouvelle position":"Net Short Position","Sens de variation":"Side","Position totale":"Aggregate Net Short Positions","Variation":"Total var."})
    df1 = df1.sort_values(by='Company Name').reset_index(drop=True)
    df1_style = df1.style.apply(color_arrows, subset=pd.IndexSlice[:, ['Side', 'Variation totale']])
        # Apply align_center styling to specific columns (remove duplicate 'Previous Position Date')
    df1_style = df1_style.apply(align_center, subset=pd.IndexSlice[:, ['Previous Position Date', 'Net Short Position', 'Previous Position', 'Side','Aggregate Net Short Positions']])
    
    
    with pd.ExcelWriter('C:/Users/33659/Desktop/Net shorts/Daily Report/FR.xlsx') as writer:
        df1_style.to_excel(writer, sheet_name='Sheet1')

    return df1_style



