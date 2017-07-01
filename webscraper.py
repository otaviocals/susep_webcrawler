#                                                                                                       #
#   Just Another Webscraper (J.A.W.)                                                                    #
#                                    v0.1                                                               #
#                                                                                                       #
#       written by Otavio Cals                                                                          #
#                                                                                                       #
#   Description: A webscrapper for downloading tables and exporting them to .csv files autonomously.    #
#                                                                                                       #
#########################################################################################################


#Required external modules: Selenium, BeautifoulSoup4, Hashlib

from csv import writer, reader
from unicodedata import normalize
from contextlib import closing
from selenium.webdriver import PhantomJS
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from pathlib import Path
from hashlib import sha512
from os.path import isdir
from os import makedirs,getcwd
from sys import platform
from time import sleep
import lxml
import sys
import gc


######################
# Webscrapping Stage #
######################

def Webscraper(url, folder, print_output = None, visual_output = None, phantom_path = ""):



    rows = []
    row = []
    append_to_csv = False
    current_os = platform

    if current_os.startswith("linux"):
        slash = "/"
    elif current_os.startswith("win32") or current_os.startswith("cygwin"):
        slash = "\\"
    elif current_os.startswith("darwin"):
        slash = "/"
    else:
        slash = "/"

#Getting Raw Data

    with closing(PhantomJS(phantom_path)) as browser:

        browser.implicitly_wait(20)
        table_class = ""
        tries = 0

#Getting Table HTML

        got_table = False
        browser.get(url)
        while tries <= 5:

            page_source = browser.find_elements_by_xpath("//div/table/tbody/tr[13]/td[3]/table")

            got_table = page_source[0].get_attribute("outerHTML").find("Taxas de juros")

            if got_table > 0:
                html_table = page_source[0].get_attribute("outerHTML")
                dates_query = "//div/table/tbody/tr[4]/td[4]/table/tbody/tr/td"
                sub_query = "//div/table/tbody/tr[9]/td[3]/table/tbody/tr/td"
                got_table = True

                break
            else:
                page_source = browser.find_elements_by_xpath("//div/table/tbody/tr[16]/td[3]/table")
                got_table = page_source[0].get_attribute("outerHTML").find("Taxas de juros")
                if got_table > 0:
                    html_table = page_source[0].get_attribute("outerHTML")
                    dates_query = "//div/table/tbody/tr[5]/td[3]/table/tbody/tr/td"
                    sub_query = "//div/table/tbody/tr[12]/td[3]/table/tbody/tr/td"
                    got_table = True

                    break
            tries += 1

#Getting Dates HTML

        dates_list = browser.find_elements_by_xpath(dates_query)
        dates = dates_list[0].text

        if len(dates) > 16:
            dates_class = "1"

        else:
            dates_class = "2"


#Getting Title HTML

        title = browser.find_element_by_class_name("mnuTitulo").text.replace(" ", "_")
        title += "_"
        sub_title_list = browser.find_elements_by_xpath(sub_query)
        sub_title = sub_title_list[0].text.replace(" ", "_").upper()

        title += sub_title
        title = normalize("NFKD",title).encode("ASCII","ignore").decode("ASCII")
        print(title)

        if not print_output == None:
            print_output.write((title+"\n").encode("utf-8").decode("utf-8"))
        if not visual_output == None:
            visual_output.append(title+"\n")

        page_source = browser.page_source


######################
#  Processing Stage  #
######################


#Getting Table Dates

    if dates_class == "1":
        date_string_cut = dates.find(" a ")
        start_date = dates[:date_string_cut]
        end_date = dates[date_string_cut+3:]
    elif dates_class == "2":
        start_date = dates
        end_date = start_date

#Getting Table Digest

    current_hash = sha512(html_table.encode("UTF-8")).hexdigest()

#Parsing HTML to XML

    soup_table = BeautifulSoup(html_table, "lxml")

######################
#    Writing Stage   #
######################

#Check Folder Existence

    if(not isdir(folder)):
        makedirs(folder)

#Creating Document if there is no previous file

    if(not Path(folder+slash+title+".csv").is_file()):
        for tr in [soup_table.find_all("tr")[2]]:
            tds = tr.find_all("td")
            #row = [normalize("NFKD",ele.text).encode("ASCII","ignore").decode("ASCII") for elem in tds]
            #row.insert(0,"Data")
            row = ["Data","Posicao","Instituicao","% a.m.","% a.a."]
            rows.append(row)
            append_to_csv = True
            print("No file detected. Creating table...")
            if not print_output == None:
                print_output.write("No file detected. Creating table...\n")
            if not visual_output == None:
                visual_output.append("No file detected. Creating table...\n")

#Loading previous file if it exists

    else:
        with open(folder+slash+title+".csv","r",encoding="UTF-8",newline="") as f:
            csv_file_read = reader(f,delimiter=";")
            for read_row in csv_file_read:
                rows.append(read_row)

#Checking Digests

            same_table = (current_hash == rows[len(rows)-1][2])

            if(not same_table):
                rows.pop()
                append_to_csv = True
                print("Changes Detected! Appending new data...")
                if not print_output == None:
                    print_output.write("Changes Detected! Appending new data...\n")
                if not visual_output == None:
                    visual_output.append("Changes Detected! Appending new data...\n")
            else:
                append_to_csv = False
                print("No changes detected.")
                if not print_output == None:
                    print_output.write("No changes detected.\n")
                if not visual_output == None:
                    visual_output.append("No changes detected.\n")

#Parsing XML to CSV

    if(append_to_csv == True):
        for tr in soup_table.find_all("tr")[3:]:
            tds = tr.find_all("td")
            row = [elem.text for elem in tds]
            row.insert(0,end_date)
            row[2] = normalize("NFKD",row[2]).encode("ASCII","ignore").decode("ASCII")
            rows.append(row)
        rows.append(["","Current Version:",current_hash])


#Writing current table to file

        with open(folder+slash+title+".csv","w",encoding="utf-8",newline="") as g:
            csv_file = writer(g,delimiter=";")
            csv_file.writerows(rows)

    gc.collect()

######################
#        Main        #
######################

if __name__ == "__main__":
    from sys import argv

    url = input("Enter url to scrap from:\n")
    folder = input("Enter folder to download to:\n")
    driver_path = input("Enter driver path:\n")

    print("Starting...")

    Webscraper(url,folder,phantom_path = driver_path)
