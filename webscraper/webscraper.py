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
from os import makedirs, getcwd, remove
from sys import platform
from time import sleep
from datetime import datetime, timedelta
from dateutil.parser import parse
from urllib.request import urlretrieve
from clint.textui import progress
from progressbar import ProgressBar, Percentage, Bar, ETA, FileTransferSpeed, RotatingMarker
from glob import glob
from zipfile import ZipFile
import requests
import lxml
import sys
import gc


######################
# Webscrapping Stage #
######################

def Webscraper(folder, print_output = None, visual_output = None, phantom_path = ""):


    url="http://www2.susep.gov.br/menuestatistica/ses/principal.aspx"
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

    folder_data = folder+slash+"data"
    folder_logs = folder+slash+"logs"



#Checking Folders Existence

    if(not isdir(folder)):
        makedirs(folder)
    if(not isdir(folder_data)):
        makedirs(folder_data)
    if(not isdir(folder_logs)):
        makedirs(folder_logs)

#Creating date.log if there is no previous file

    if(not Path(folder_logs+slash+"date.log").is_file()):
        recorded_date = datetime(1970,1,1,0,0)
        with open(folder_logs+slash+"date.log","w") as f:
            print(recorded_date,file=f)
        print("No date.log detected. Creating file...")
        if not print_output == None:
            print_output.write("No date.log detected. Creating file...\n")
        if not visual_output == None:
            visual_output.append("No date.log detected. Creating file...\n")

#Loading date.log if it exists

    else:
        with open(folder_logs+slash+"date.log","r") as f:
            date_log_content = f.readlines()
            date_log_content = [x.strip() for x in date_log_content]
            recorded_date = parse(date_log_content[0])

#Getting Raw Data


    with closing(PhantomJS(phantom_path)) as browser:

        browser.implicitly_wait(5)
        tries = 0

    #Getting Newest Data Date

        browser.get(url)
        while tries <= 5:

            date_source = browser.find_elements_by_xpath("//form[@id=\"aspnetForm\"]/div[4]/table/tbody/tr/td/table[2]/tbody/tr[9]/td/font/div/span")
            date_string = date_source[0].text

            if len(date_string) > 0:
                break

            tries += 1
        last_date = parse(date_string)

    #Checking if Data is up to date

        if (recorded_date >= last_date):
            print("Data is up to date.")
            if not print_output == None:
                print_output.write("Data is up to date.\n")
            if not visual_output == None:
                visual_output.append("Data is up to date.\n")
            return False
        else:
            with open(folder_logs+slash+"date.log","w") as f:
                print(last_date,file=f)
            print("New data detected! Downloading data...")
            if not print_output == None:
                print_output.write("New data detected! Downloading data...\n")
            if not visual_output == None:
                visual_output.append("New data detected! Downloading data...\n")

        #Deleting Old Data

            for fl in glob(folder_data+slash+"*"):
                remove(fl)

        #Downloading Data

            data_source = browser.find_elements_by_xpath("//form[@id=\"aspnetForm\"]/div[4]/table/tbody/tr/td/table[2]/tbody/tr[9]/td/font/div/a")
            data_address = data_source[0].get_attribute("href")
            data_path = folder_data+slash+"zipped_data.zip"
            print(data_address)
            prct_achiev = [0,0,0,0,0,0,0,0,0,0]
            def dlProgress(count, blockSize, totalSize):

                if(prct_achiev[0] == 0 and (count*blockSize/totalSize) > 0.009 ):
                    prct_achiev[0] = 1
                    elap_time = datetime.now() - down_start
                    down_speed = int((count*blockSize)/(elap_time.seconds*1000))
                    eta_num = (totalSize)/(down_speed*60000)
                    eta = timedelta(minutes=int(eta_num), seconds = (eta_num - int(eta_num))*(60))
                    print("00% Downloaded. Download Speed: "+str(down_speed)+" kB/s Elapsed time: "+str(elap_time).split(".")[0]+" ETA: "+str(eta).split(".")[0])

                if(prct_achiev[1] == 0 and (count*blockSize/totalSize) > 0.1 ):
                    prct_achiev[1] = 1
                    elap_time = datetime.now() - down_start
                    down_speed = int((count*blockSize)/(elap_time.seconds*1000))
                    eta_num = (totalSize)/(down_speed*60000)
                    eta = timedelta(minutes=int(eta_num), seconds = (eta_num - int(eta_num))*(60))
                    print("10% Downloaded. Download Speed: "+str(down_speed)+" kB/s Elapsed time: "+str(elap_time).split(".")[0]+" ETA: "+str(eta).split(".")[0])

                if(prct_achiev[2] == 0 and (count*blockSize/totalSize) > 0.2 ):
                    prct_achiev[2] = 1
                    elap_time = datetime.now() - down_start
                    down_speed = int((count*blockSize)/(elap_time.seconds*1000))
                    eta_num = (totalSize)/(down_speed*60000)
                    eta = timedelta(minutes=int(eta_num), seconds = (eta_num - int(eta_num))*(60))
                    print("20% Downloaded. Download Speed: "+str(down_speed)+" kB/s Elapsed time: "+str(elap_time).split(".")[0]+" ETA: "+str(eta).split(".")[0])

                if(prct_achiev[3] == 0 and (count*blockSize/totalSize) > 0.3 ):
                    prct_achiev[3] = 1
                    elap_time = datetime.now() - down_start
                    down_speed = int((count*blockSize)/(elap_time.seconds*1000))
                    eta_num = (totalSize)/(down_speed*60000)
                    eta = timedelta(minutes=int(eta_num), seconds = (eta_num - int(eta_num))*(60))
                    print("30% Downloaded. Download Speed: "+str(down_speed)+" kB/s Elapsed time: "+str(elap_time).split(".")[0]+" ETA: "+str(eta).split(".")[0])

                if(prct_achiev[4] == 0 and (count*blockSize/totalSize) > 0.4 ):
                    prct_achiev[4] = 1
                    elap_time = datetime.now() - down_start
                    down_speed = int((count*blockSize)/(elap_time.seconds*1000))
                    eta_num = (totalSize)/(down_speed*60000)
                    eta = timedelta(minutes=int(eta_num), seconds = (eta_num - int(eta_num))*(60))
                    print("40% Downloaded. Download Speed: "+str(down_speed)+" kB/s Elapsed time: "+str(elap_time).split(".")[0]+" ETA: "+str(eta).split(".")[0])

                if(prct_achiev[5] == 0 and (count*blockSize/totalSize) > 0.5 ):
                    prct_achiev[5] = 1
                    elap_time = datetime.now() - down_start
                    down_speed = int((count*blockSize)/(elap_time.seconds*1000))
                    eta_num = (totalSize)/(down_speed*60000)
                    eta = timedelta(minutes=int(eta_num), seconds = (eta_num - int(eta_num))*(60))
                    print("50% Downloaded. Download Speed: "+str(down_speed)+" kB/s Elapsed time: "+str(elap_time).split(".")[0]+" ETA: "+str(eta).split(".")[0])

                if(prct_achiev[6] == 0 and (count*blockSize/totalSize) > 0.6 ):
                    prct_achiev[6] = 1
                    elap_time = datetime.now() - down_start
                    down_speed = int((count*blockSize)/(elap_time.seconds*1000))
                    eta_num = (totalSize)/(down_speed*60000)
                    eta = timedelta(minutes=int(eta_num), seconds = (eta_num - int(eta_num))*(60))
                    print("60% Downloaded. Download Speed: "+str(down_speed)+" kB/s Elapsed time: "+str(elap_time).split(".")[0]+" ETA: "+str(eta).split(".")[0])

                if(prct_achiev[7] == 0 and (count*blockSize/totalSize) > 0.7 ):
                    prct_achiev[7] = 1
                    elap_time = datetime.now() - down_start
                    down_speed = int((count*blockSize)/(elap_time.seconds*1000))
                    eta_num = (totalSize)/(down_speed*60000)
                    eta = timedelta(minutes=int(eta_num), seconds = (eta_num - int(eta_num))*(60))
                    print("70% Downloaded. Download Speed: "+str(down_speed)+" kB/s Elapsed time: "+str(elap_time).split(".")[0]+" ETA: "+str(eta).split(".")[0])

                if(prct_achiev[8] == 0 and (count*blockSize/totalSize) > 0.8 ):
                    prct_achiev[8] = 1
                    elap_time = datetime.now() - down_start
                    down_speed = int((count*blockSize)/(elap_time.seconds*1000))
                    eta_num = (totalSize)/(down_speed*60000)
                    eta = timedelta(minutes=int(eta_num), seconds = (eta_num - int(eta_num))*(60))
                    print("80% Downloaded. Download Speed: "+str(down_speed)+" kB/s Elapsed time: "+str(elap_time).split(".")[0]+" ETA: "+str(eta).split(".")[0])

                if(prct_achiev[9] == 0 and (count*blockSize/totalSize) > 0.9 ):
                    prct_achiev[9] = 1
                    elap_time = datetime.now() - down_start
                    down_speed = int((count*blockSize)/(elap_time.seconds*1000))
                    eta_num = (totalSize)/(down_speed*60000)
                    eta = timedelta(minutes=int(eta_num), seconds = (eta_num - int(eta_num))*(60))
                    print("90% Downloaded. Download Speed: "+str(down_speed)+" kB/s Elapsed time: "+str(elap_time).split(".")[0]+" ETA: "+str(eta).split(".")[0])

            down_start = datetime.now()
            urlretrieve(data_address, data_path, reporthook = dlProgress)
            print("Download finished!")
            if not print_output == None:
                print_output.write("Download finished!\n")
            if not visual_output == None:
                visual_output.append("Download finished!\n")

        #Unzipping Data

            with ZipFile(data_path,"r") as zip_ref:
                zip_ref.extractall(folder_data)
            remove(data_path)
            return True















#    #Getting Dates HTML
#
#        dates_list = browser.find_elements_by_xpath(dates_query)
#        dates = dates_list[0].text
#
#        if len(dates) > 16:
#            dates_class = "1"
#
#        else:
#            dates_class = "2"
#
#
#    #Getting Title HTML
#
#        title = browser.find_element_by_class_name("mnuTitulo").text.replace(" ", "_")
#        title += "_"
#        sub_title_list = browser.find_elements_by_xpath(sub_query)
#        sub_title = sub_title_list[0].text.replace(" ", "_").upper()
#
#        title += sub_title
#        title = normalize("NFKD",title).encode("ASCII","ignore").decode("ASCII")
#        print(title)
#
#        if not print_output == None:
#            print_output.write((title+"\n").encode("utf-8").decode("utf-8"))
#        if not visual_output == None:
#            visual_output.append(title+"\n")
#
#        page_source = browser.page_source


######################
#  Processing Stage  #
######################


##Getting Table Dates
#
#    if dates_class == "1":
#        date_string_cut = dates.find(" a ")
#        start_date = dates[:date_string_cut]
#        end_date = dates[date_string_cut+3:]
#    elif dates_class == "2":
#        start_date = dates
#        end_date = start_date
#
##Getting Table Digest
#
#    current_hash = sha512(html_table.encode("UTF-8")).hexdigest()
#
##Parsing HTML to XML
#
#    soup_table = BeautifulSoup(html_table, "lxml")
#
#######################
##    Writing Stage   #
#######################
#
##Check Folder Existence
#
#    if(not isdir(folder)):
#        makedirs(folder)
#
##Creating Document if there is no previous file
#
#    if(not Path(folder+slash+title+".csv").is_file()):
#        for tr in [soup_table.find_all("tr")[2]]:
#            tds = tr.find_all("td")
#            #row = [normalize("NFKD",ele.text).encode("ASCII","ignore").decode("ASCII") for elem in tds]
#            #row.insert(0,"Data")
#            row = ["Data","Posicao","Instituicao","% a.m.","% a.a."]
#            rows.append(row)
#            append_to_csv = True
#            print("No file detected. Creating table...")
#            if not print_output == None:
#                print_output.write("No file detected. Creating table...\n")
#            if not visual_output == None:
#                visual_output.append("No file detected. Creating table...\n")
#
##Loading previous file if it exists
#
#    else:
#        with open(folder+slash+title+".csv","r",encoding="UTF-8",newline="") as f:
#            csv_file_read = reader(f,delimiter=";")
#            for read_row in csv_file_read:
#                rows.append(read_row)
#
##Checking Digests
#
#            same_table = (current_hash == rows[len(rows)-1][2])
#
#            if(not same_table):
#                rows.pop()
#                append_to_csv = True
#                print("Changes Detected! Appending new data...")
#                if not print_output == None:
#                    print_output.write("Changes Detected! Appending new data...\n")
#                if not visual_output == None:
#                    visual_output.append("Changes Detected! Appending new data...\n")
#            else:
#                append_to_csv = False
#                print("No changes detected.")
#                if not print_output == None:
#                    print_output.write("No changes detected.\n")
#                if not visual_output == None:
#                    visual_output.append("No changes detected.\n")
#
##Parsing XML to CSV
#
#    if(append_to_csv == True):
#        for tr in soup_table.find_all("tr")[3:]:
#            tds = tr.find_all("td")
#            row = [elem.text for elem in tds]
#            row.insert(0,end_date)
#            row[2] = normalize("NFKD",row[2]).encode("ASCII","ignore").decode("ASCII")
#            rows.append(row)
#        rows.append(["","Current Version:",current_hash])
#
#
##Writing current table to file
#
#        with open(folder+slash+title+".csv","w",encoding="utf-8",newline="") as g:
#            csv_file = writer(g,delimiter=";")
#            csv_file.writerows(rows)
#
#    gc.collect()

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
