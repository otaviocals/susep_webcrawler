#########################################
#                                       #
#          SUSEP-Webscraper             #
#                v0.1                   #
#                                       #
#       written by Otavio Cals          #
#                                       #
#    Description: A webscrapper for     #
#    downloading tables and exporting   #
#    them to .csv files autonomously.   #
#                                       #
#########################################

from kivy.app import App
from kivy.uix.label import Label
from kivy.uix.image import Image
from kivy.uix.textinput import TextInput
from kivy.uix.spinner import Spinner
from kivy.uix.button import Button
from kivy.uix.togglebutton import ToggleButton
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.scrollview import ScrollView
from kivy.effects.scroll import ScrollEffect
from kivy.core.window import Window
from kivy.config import Config
from kivy.core.window import Window
from kivy.properties import ObjectProperty
from kivy.clock import Clock
from os.path import join, isdir, expanduser, isfile
from os import makedirs
from sys import platform
import subprocess
from pathlib import Path
from kivy.lang import Builder
from pylibs.file import XFolder
from pylibs.xpopup import XPopup
import pylibs.tools
from time import sleep
from unidecode import unidecode
from webscraper.webscraper import Webscraper
import time
import datetime
import gc
import sys
import os
import shutil
import webbrowser

######################
#    Pre-Settings    #
######################

#Setting configurations

Config.set("kivy","log_enable","0")
Config.set("kivy","log_level","critical")
Config.set("graphics","position","custom")
Config.set("graphics","top","50")
Config.set("graphics","left","10")
Config.set("graphics","width","450")
Config.set("graphics","height","900")
Config.set("graphics","fullscreen","0")
Config.set("graphics","borderless","0")
Config.set("graphics","resizable","0")
Config.write()

current_os = platform

try:
    import win32api
    win32api.SetDllDirectory(sys._MEIPASS)
except:
    pass

if current_os.startswith("linux"):
    slash = "/"
    phantom = "phantomjs/phantomjs"
elif current_os.startswith("win32") or current_os.startswith("cygwin"):
    slash = "\\"
    phantom = "phantomjs\\phantomjs.exe"
    import pylibs.win32timezone
elif current_os.startswith("darwin"):
    slash = "/"
    phantom = "phantomjs/phantomjs"
else:
    slash = "/"
    phantom = "phantomjs/phantomjs"


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
        if(not isdir(base_path+slash+"rlibs")):
            makedirs(base_path+slash+"rlibs")
    return os.path.join(base_path,relative_path)

logo_path = resource_path("logo"+slash+"logo.png")
phantom_path = resource_path(phantom)
r_script_path = resource_path("rscripts")
r_libs_path = resource_path("rlibs")
kivi_app_path = resource_path("kivylibs"+slash+"app_screen.kv")
config_path = os.path.abspath(".")+slash+"config.txt"



#######################
#     Config Data     #
#######################

#Initiating config data

#1st Analysis satrtup

analysis_check = [True]
analysis_config = [False]*18
dates_1=["01/01/2017","01/02/2017"]
ent_1=[True,""]
ram_1=[True,""]
deltatime_1=[True,False,False,False]

#2nd Analysis startup

analysis_2_check = [True]
analysis_2_config = [False]*13
dates_2=["01/01/2017","01/02/2017"]
ent_2=[True,""]
ram_2=[True,""]
grup_2=[True,""]
deltatime_2=[True,False,False,False]

#3rd Analysis startup

analysis_3_check = [True]
dates_3=["01/01/2017","01/02/2017"]
ent_3=[True,""]
ram_3=[True,""]
grup_3=[True,""]
deltatime_3=[True,False,False,False]
text_input_3 = ""

#Loading saved configurations

if(Path(config_path).is_file()):
    with open(config_path,"r") as f:
        configs = f.readlines()
        configs = [x.strip() for x in configs]

    #1st Analysis Config Read

        analysis_config_string = configs[1][1:len(configs[1])-1].split(", ")
        dates_1_string = configs[2][1:len(configs[2])-1].split(", ")
        ent_1_string = configs[3][1:len(configs[3])-1].split(", ")
        ram_1_string = configs[4][1:len(configs[4])-1].split(", ")
        deltatime_1_string = configs[5][1:len(configs[5])-1].split(", ")
        analysis_check_string = configs[6][1:len(configs[6])-1]

        i=0
        for x in analysis_config_string:
            analysis_config[i] = (x == "True")
            i+=1

        dates_1[0]=dates_1_string[0][1:len(dates_1_string[0])-1]
        dates_1[1]=dates_1_string[1][1:len(dates_1_string[1])-1]

        ent_1[0]=(ent_1_string[0]=="True")
        ent_1[1]=ent_1_string[1][1:len(ent_1_string[1])-1]

        ram_1[0]=(ram_1_string[0]=="True")
        ram_1[1]=ram_1_string[1][1:len(ram_1_string[1])-1]

        i=0
        for x in deltatime_1_string:
            deltatime_1[i] = (x == "True")
            i+=1

        analysis_check[0]=(analysis_check_string == "True")

    #2nd Analysis Config Read
        analysis_config_2_string = configs[8][1:len(configs[8])-1].split(", ")
        dates_2_string = configs[9][1:len(configs[9])-1].split(", ")
        ent_2_string = configs[10][1:len(configs[10])-1].split(", ")
        ram_2_string = configs[11][1:len(configs[11])-1].split(", ")
        grup_2_string = configs[12][1:len(configs[12])-1].split(", ")
        deltatime_2_string = configs[13][1:len(configs[13])-1].split(", ")
        analysis_2_check_string = configs[14][1:len(configs[14])-1]

        i=0
        for x in analysis_config_2_string:
            analysis_2_config[i] = (x == "True")
            i+=1

        dates_2[0]=dates_2_string[0][1:len(dates_2_string[0])-1]
        dates_2[1]=dates_2_string[1][1:len(dates_2_string[1])-1]

        ent_2[0]=(ent_2_string[0]=="True")
        ent_2[1]=ent_2_string[1][1:len(ent_2_string[1])-1]

        ram_2[0]=(ram_2_string[0]=="True")
        ram_2[1]=ram_2_string[1][1:len(ram_2_string[1])-1]

        grup_2[0]=(grup_2_string[0]=="True")
        grup_2[1]=grup_2_string[1][1:len(grup_2_string[1])-1]

        i=0
        for x in deltatime_2_string:
            deltatime_2[i] = (x == "True")
            i+=1

        analysis_2_check[0]=(analysis_2_check_string == "True")

    #3rd Analysis Config Read
        analysis_3_check_string = configs[16][1:len(configs[16])-1]
        dates_3_string = configs[17][1:len(configs[17])-1].split(", ")
        ent_3_string = configs[18][1:len(configs[18])-1].split(", ")
        ram_3_string = configs[19][1:len(configs[19])-1].split(", ")
        grup_3_string = configs[20][1:len(configs[20])-1].split(", ")
        deltatime_3_string = configs[21][1:len(configs[21])-1].split(", ")

        dates_3[0]=dates_3_string[0][1:len(dates_3_string[0])-1]
        dates_3[1]=dates_3_string[1][1:len(dates_3_string[1])-1]

        ent_3[0]=(ent_3_string[0]=="True")
        ent_3[1]=ent_3_string[1][1:len(ent_3_string[1])-1]

        ram_3[0]=(ram_3_string[0]=="True")
        ram_3[1]=ram_3_string[1][1:len(ram_3_string[1])-1]

        grup_3[0]=(grup_3_string[0]=="True")
        grup_3[1]=grup_3_string[1][1:len(grup_3_string[1])-1]

        i=0
        for x in deltatime_3_string:
            deltatime_3[i] = (x == "True")
            i+=1

        analysis_3_check[0]=(analysis_3_check_string == "True")
        text_input_3_array = ""
        for i in range(22,len(configs)):
            if i>22:
                text_input_3_array += "\n"
            text_input_3_array += configs[i]
        text_input_3 = text_input_3_array

else:
    with open(config_path,"a") as f:
        f.write("Config Analysis 1\n")
        f.write(str(analysis_config)+"\n")
        f.write(str(dates_1)+"\n")
        f.write(str(ent_1)+"\n")
        f.write(str(ram_1)+"\n")
        f.write(str(deltatime_1)+"\n")
        f.write(str(analysis_check)+"\n")

        f.write("Config Analysis 2\n")
        f.write(str(analysis_2_config)+"\n")
        f.write(str(dates_2)+"\n")
        f.write(str(ent_2)+"\n")
        f.write(str(ram_2)+"\n")
        f.write(str(grup_2)+"\n")
        f.write(str(deltatime_2)+"\n")
        f.write(str(analysis_2_check)+"\n")

        f.write("Config Analysis 3\n")
        f.write(str(analysis_3_check)+"\n")
        f.write(str(dates_3)+"\n")
        f.write(str(ent_3)+"\n")
        f.write(str(ram_3)+"\n")
        f.write(str(grup_3)+"\n")
        f.write(str(deltatime_3)+"\n")
        f.write(str(text_input_3))

#1st Analysis temp values
new_analysis_check = analysis_check[:]
new_analysis_config = analysis_config[:]
new_dates_1 = dates_1[:]
new_ent_1 = ent_1[:]
new_ram_1 = ram_1[:]
new_deltatime_1 = deltatime_1[:]

#2nd Analysis temp values
new_analysis_2_check = analysis_2_check[:]
new_analysis_2_config = analysis_2_config[:]
new_dates_2=dates_2[:]
new_ent_2 = ent_2[:]
new_ram_2 = ram_2[:]
new_grup_2 = grup_2[:]
new_deltatime_2=deltatime_2[:]

#3rd Analysis temp values
new_analysis_3_check = analysis_3_check[:]
new_dates_3=dates_3[:]
new_ent_3 = ent_3[:]
new_ram_3 = ram_3[:]
new_grup_3 = grup_3[:]
new_deltatime_3=deltatime_3[:]
new_text_input_3 = text_input_3[:]

#Building GUI

Builder.load_file(kivi_app_path)


######################
#    App Functions   #
######################

class AppScreen(GridLayout):

    default_folder = expanduser("~") + slash + "Documents"
    global sel_folder
    sel_folder = default_folder

#Calling webscraper

    def scrap(self,event):
        if(not isdir(sel_folder+slash+"logs")):
            makedirs(sel_folder+slash+"logs")
        if(not isdir(sel_folder+slash+"proc_data")):
            makedirs(sel_folder+slash+"proc_data")
        if(not isdir(sel_folder+slash+"proc_data"+slash+"plots")):
            makedirs(sel_folder+slash+"proc_data"+slash+"plots")
        log_file = open((sel_folder+slash+"logs"+slash+"history.log"),"a",encoding="utf-8")

    #Starting scrap

        print("Starting at "+str(datetime.datetime.now()))
        log_file.write("Starting at "+str(datetime.datetime.now())+"\n")
        self.ids.log_output.text += "Starting at "+str(datetime.datetime.now())+"\n"

        print("Webscraper Starting...")
        log_file.write("Webscraper Starting...\n")
        self.ids.log_output.text += "Websccraper Starting...\n"

        start_time_seconds = time.time()
        self.ids.start_button.disabled = True

    #Running webscraper

        str_output_logger = []
        new_data = Webscraper(sel_folder,log_file,str_output_logger,phantom_path)
        self.ids.log_output.text += "".join(str_output_logger)
        gc.collect()

    #Running R scripts

        r_verify = shutil.which("Rscript")

        if(not r_verify == None):
            print("Running Analysis...")
            log_file.write("Running Analysis...\n")
            self.ids.log_output.text += "Running Analysis...\n"

            log_file.close()

            r_output = subprocess.check_output(["Rscript",r_script_path+slash+"analysis.R", slash, sel_folder,config_path, r_script_path,r_libs_path, sel_folder+slash+"logs"+slash+"history.log"],universal_newlines=True)
            print(r_output)

            log_file = open((sel_folder+slash+"logs"+slash+"history.log"),"a",encoding="utf-8")

        else:
            print("R compiler not found.")
            log_file.write("R compiler not found.\n")
            self.ids.log_output.text += "R compiler not found.\n"

    #Ending scrap

        self.ids.start_button.disabled = False
        end_time_seconds = time.time()
        elapsed_time = str(int(round(end_time_seconds - start_time_seconds)))

        print("Ending at "+str(datetime.datetime.now()))
        log_file.write("Ending at "+str(datetime.datetime.now())+"\n")
        self.ids.log_output.text += "Ending at "+str(datetime.datetime.now())+"\n"

        print("Total Elapsed Time: "+elapsed_time+" seconds.")
        log_file.write("Total Elapsed Time: "+elapsed_time+" seconds.\n")
        self.ids.log_output.text += "Total Elapsed Time: "+elapsed_time+" seconds.\n"

        print("Success!")
        log_file.write("Success!\n")
        self.ids.log_output.text += "Success!\n"

        log_file.close()

        self.ids.start_button.state = "normal"


#GUI functions

    def start(self,*args):
        global event
        if args[1] == "down":
            self.ids.check_button.disabled = True
            self.ids.folder_button.disabled = True
            event = Clock.schedule_once(self.scrap)
        if args[1] == "normal":
            self.ids.check_button.disabled = False
            self.ids.folder_button.disabled = False
            Clock.unschedule(event)

    def update_button(self,event):
        self.ToggleButton.text="Stop"

    def _open_link(self):
        webbrowser.open("http://www.calsemporium.com")

#Directory selection

    def _filepopup_callback(self, instance):
            if instance.is_canceled():
                return
            s = 'Path: %s' % instance.path
            s += ('\nSelection: %s' % instance.selection)
            global sel_folder
            self.ids.text_input.text = "Pasta selecionada:    " + instance.path
            sel_folder = instance.path

    def _folder_dialog(self):
        XFolder(on_dismiss=self._filepopup_callback, path=expanduser(u'~'))



    def _check_dialog(self):
        popup = CheckPopup()

        popup.ids.scroll_grid_1.bind(minimum_height=popup.ids.scroll_grid_1.setter("height"))
        popup.ids.scroll_grid_2.bind(minimum_height=popup.ids.scroll_grid_2.setter("height"))
        popup.ids.scroll_grid_3.bind(minimum_height=popup.ids.scroll_grid_3.setter("height"))

        popup.title="Selecao de Consultas"

        global analysis_check
        global new_analysis_check
        global analysis_config
        global new_analysis_config
        global dates_1
        global new_dates_1
        global ent_1
        global new_ent_1
        global ram_1
        global new_ram_1
        global deltatime_1
        global new_deltatime_1

        global analysis_2_check
        global new_analysis_2_check
        global analysis_2_config
        global new_analysis_2_config
        global dates_2
        global new_dates_2
        global ent_2
        global new_ent_2
        global ram_2
        global new_ram_2
        global grup_2
        global new_grup_2
        global deltatime_2
        global new_deltatime_2

        global analysis_3_check
        global new_analysis_3_check
        global dates_3
        global new_dates_3
        global ent_3
        global new_ent_3
        global ram_3
        global new_ram_3
        global grup_3
        global new_grup_3
        global deltatime_3
        global new_deltatime_3
        global text_input_3
        global new_text_input_3

        #Initiating Analysis 1 values

        new_analysis_check = analysis_check[:]
        new_analysis_config = analysis_config[:]
        new_dates_1 = dates_1[:]
        new_ent_1 = ent_1[:]
        new_ram_1 = ram_1[:]
        new_deltatime_1 = deltatime_1[:]

        #Initiating Analysis 2 values

        new_analysis_2_check = analysis_2_check[:]
        new_analysis_2_config = analysis_2_config[:]
        new_dates_2=dates_2[:]
        new_ent_2 = ent_2[:]
        new_ram_2 = ram_2[:]
        new_grup_2 = grup_2[:]
        new_deltatime_2=deltatime_2[:]

        #Initiating Analysis 3 values

        new_analysis_3_check = analysis_3_check[:]
        new_dates_3=dates_3[:]
        new_ent_3 = ent_3[:]
        new_ram_3 = ram_3[:]
        new_grup_3 = grup_3[:]
        new_deltatime_3=deltatime_3[:]
        new_text_input_3 = text_input_3[:]

        #Setting Popup Values

        popup.ids.check_analysis_1.active=new_analysis_check[0]
        for i in range(0,18):
            popup.ids["check_"+str(i)].active=new_analysis_config[i]
        popup.ids.date_1_1.text=new_dates_1[0]
        popup.ids.date_1_2.text=new_dates_1[1]
        popup.ids.check_ent_1.active=new_ent_1[0]
        popup.ids.text_ent_1.text=new_ent_1[1]
        popup.ids.check_ram_1.active=new_ram_1[0]
        popup.ids.text_ram_1.text=new_ram_1[1]
        popup.ids.check_time_1_1.active=new_deltatime_1[0]
        popup.ids.check_time_1_2.active=new_deltatime_1[1]
        popup.ids.check_time_1_3.active=new_deltatime_1[2]
        popup.ids.check_time_1_4.active=new_deltatime_1[3]

        popup.ids.check_analysis_2.active=new_analysis_2_check[0]
        for i in range(0,len(new_analysis_2_config)):
            popup.ids["check2_"+str(i)].active=new_analysis_2_config[i]
        popup.ids.date_2_1.text=new_dates_2[0]
        popup.ids.date_2_2.text=new_dates_2[1]
        popup.ids.check_ent_2.active=new_ent_2[0]
        popup.ids.text_ent_2.text=new_ent_2[1]
        popup.ids.check_ram_2.active=new_ram_2[0]
        popup.ids.text_ram_2.text=new_ram_2[1]
        popup.ids.check_grup_2.active=new_grup_2[0]
        popup.ids.text_grup_2.text=new_grup_2[1]
        popup.ids.check_time_2_1.active=new_deltatime_2[0]
        popup.ids.check_time_2_2.active=new_deltatime_2[1]
        popup.ids.check_time_2_3.active=new_deltatime_2[2]
        popup.ids.check_time_2_4.active=new_deltatime_2[3]

        popup.ids.check_analysis_3.active=new_analysis_3_check[0]
        popup.ids.date_3_1.text=new_dates_3[0]
        popup.ids.date_3_2.text=new_dates_3[1]
        popup.ids.check_ent_3.active=new_ent_3[0]
        popup.ids.text_ent_3.text=new_ent_3[1]
        popup.ids.check_ram_3.active=new_ram_3[0]
        popup.ids.text_ram_3.text=new_ram_3[1]
        popup.ids.check_grup_3.active=new_grup_3[0]
        popup.ids.text_grup_3.text=new_grup_3[1]
        popup.ids.check_time_3_1.active=new_deltatime_3[0]
        popup.ids.check_time_3_2.active=new_deltatime_3[1]
        popup.ids.check_time_3_3.active=new_deltatime_3[2]
        popup.ids.check_time_3_4.active=new_deltatime_3[3]
        popup.ids.text_input_3.text = new_text_input_3

        popup.open()


#Setting window layout

    def __init__(self,**kwargs):
        super(AppScreen, self).__init__(**kwargs)

        Window.size = (900,750)
        Window.set_title("Webscraper SUSEP")


        self.cols = 1
        self.size_hint = (None,None)
        self.width = 900
        self.height = 750
        self.icon = logo_path
        self.ids.ce_logo.source = logo_path


class CheckPopup(XPopup):

#1st column analysis

    #Analysis On/Off

    def _checkanalysis_callback(self,value):
        global new_analysis_check
        new_analysis_check[0] = value

    #Analysis check boxes

    def _checkbox_callback(self,n,value):
        global new_analysis_config
        new_analysis_config[n]=value

    #Analysis Company field

    def _entbox_callback(self,value):
        global new_ent_1
        new_ent_1[0]=value

    #Analysis Field field

    def _rambox_callback(self,value):
        global new_ram_1
        new_ram_1[0]=value

    #Analysis time interval field

    def _deltatime_callback(self,n):
        global deltatime_1
        if n==0:
            new_deltatime_1[0]=True
            new_deltatime_1[1]=False
            new_deltatime_1[2]=False
            new_deltatime_1[3]=False
        elif n==1:
            new_deltatime_1[0]=False
            new_deltatime_1[1]=True
            new_deltatime_1[2]=False
            new_deltatime_1[3]=False
        elif n==2:
            new_deltatime_1[0]=False
            new_deltatime_1[1]=False
            new_deltatime_1[2]=True
            new_deltatime_1[3]=False
        elif n==3:
            new_deltatime_1[0]=False
            new_deltatime_1[1]=False
            new_deltatime_1[2]=False
            new_deltatime_1[3]=True

#2nd column analysis

    #Analysis On/Off

    def _checkanalysis_2_callback(self,value):
        global new_analysis_2_check
        new_analysis_2_check[0] = value

    #Analysis check boxes

    def _checkbox_2_callback(self,n,value):
        global new_analysis_2_config
        new_analysis_2_config[n]=value

    #Analysis Company field

    def _entbox_2_callback(self,value):
        global new_ent_2
        new_ent_2[0]=value

    #Analysis Field field

    def _rambox_2_callback(self,value):
        global new_ram_2
        new_ram_2[0]=value

    #Analysis Group field

    def _groupbox_2_callback(self,value):
        global new_grup_2
        new_grup_2[0]=value


    #Analysis time interval field

    def _deltatime_2_callback(self,n):
        global deltatime_2
        if n==0:
            new_deltatime_2[0]=True
            new_deltatime_2[1]=False
            new_deltatime_2[2]=False
            new_deltatime_2[3]=False
        elif n==1:
            new_deltatime_2[0]=False
            new_deltatime_2[1]=True
            new_deltatime_2[2]=False
            new_deltatime_2[3]=False
        elif n==2:
            new_deltatime_2[0]=False
            new_deltatime_2[1]=False
            new_deltatime_2[2]=True
            new_deltatime_2[3]=False
        elif n==3:
            new_deltatime_2[0]=False
            new_deltatime_2[1]=False
            new_deltatime_2[2]=False
            new_deltatime_2[3]=True

#3rd column analysis

    #Analysis On/Off

    def _checkanalysis_3_callback(self,value):
        global new_analysis_3_check
        new_analysis_3_check[0] = value

    #Analysis Company field

    def _entbox_3_callback(self,value):
        global new_ent_3
        new_ent_3[0]=value

    #Analysis Field field

    def _rambox_3_callback(self,value):
        global new_ram_3
        new_ram_3[0]=value

    #Analysis Group field

    def _groupbox_3_callback(self,value):
        global new_grup_3
        new_grup_3[0]=value


    #Analysis time interval field

    def _deltatime_3_callback(self,n):
        global deltatime_3
        if n==0:
            new_deltatime_3[0]=True
            new_deltatime_3[1]=False
            new_deltatime_3[2]=False
            new_deltatime_3[3]=False
        elif n==1:
            new_deltatime_3[0]=False
            new_deltatime_3[1]=True
            new_deltatime_3[2]=False
            new_deltatime_3[3]=False
        elif n==2:
            new_deltatime_3[0]=False
            new_deltatime_3[1]=False
            new_deltatime_3[2]=True
            new_deltatime_3[3]=False
        elif n==3:
            new_deltatime_3[0]=False
            new_deltatime_3[1]=False
            new_deltatime_3[2]=False
            new_deltatime_3[3]=True

#Saving fields to data

    def _check_popup_callback(self,value):
        global analysis_check
        global new_analysis_check
        global analysis_config
        global new_analysis_config
        global dates_1
        global new_dates_1
        global ent_1
        global new_ent_1
        global ram_1
        global new_ram_1
        global deltatime_1
        global new_deltatime_1

        global analysis_2_check
        global new_analysis_2_check
        global analysis_2_config
        global new_analysis_2_config
        global dates_2
        global new_dates_2
        global ent_2
        global new_ent_2
        global ram_2
        global new_ram_2
        global grup_2
        global new_grup_2
        global deltatime_2
        global new_deltatime_2

        global analysis_3_check
        global new_analysis_3_check
        global dates_3
        global new_dates_3
        global ent_3
        global new_ent_3
        global ram_3
        global new_ram_3
        global grup_3
        global new_grup_3
        global deltatime_3
        global new_deltatime_3
        global text_input_3
        global new_text_input_3

        if value:

            #Analysis 1 Save

            new_dates_1[0] = self.ids.date_1_1.text
            new_dates_1[1] = self.ids.date_1_2.text
            new_ent_1[1] = self.ids.text_ent_1.text
            new_ram_1[1] = self.ids.text_ram_1.text
            analysis_check = new_analysis_check[:]
            analysis_config = new_analysis_config[:]
            dates_1 = new_dates_1[:]
            ent_1 = new_ent_1[:]
            ram_1 = new_ram_1[:]
            deltatime_1 = new_deltatime_1[:]

            #Analysis 2 save

            new_dates_2[0] = self.ids.date_2_1.text
            new_dates_2[1] = self.ids.date_2_2.text
            new_ent_2[1] = self.ids.text_ent_2.text
            new_ram_2[1] = self.ids.text_ram_2.text
            new_grup_2[1] = self.ids.text_grup_2.text
            analysis_2_check = new_analysis_2_check[:]
            analysis_2_config = new_analysis_2_config[:]
            dates_2 = new_dates_2[:]
            ent_2 = new_ent_2[:]
            ram_2 = new_ram_2[:]
            grup_2 = new_grup_2[:]
            deltatime_2 = new_deltatime_2[:]

            #Analysis 3 save

            new_dates_3[0] = self.ids.date_3_1.text
            new_dates_3[1] = self.ids.date_3_2.text
            new_ent_3[1] = self.ids.text_ent_3.text
            new_ram_3[1] = self.ids.text_ram_3.text
            new_grup_3[1] = self.ids.text_grup_3.text
            analysis_3_check = new_analysis_3_check[:]
            dates_3 = new_dates_3[:]
            ent_3 = new_ent_3[:]
            ram_3 = new_ram_3[:]
            grup_3 = new_grup_3[:]
            deltatime_3 = new_deltatime_3[:]
            text_input_3 = unidecode(self.ids.text_input_3.text)
            self.ids.text_input_3.text = text_input_3
            lines_to_write = text_input_3.split("\n")

            #Saving file
            with open(config_path,"r") as f:
                filelines=f.readlines()
            with open(config_path,"w") as f:
                #Analysis 1
                filelines[1]=str(analysis_config)+"\n"
                filelines[2]=str(dates_1)+"\n"
                filelines[3]=str(ent_1)+"\n"
                filelines[4]=str(ram_1)+"\n"
                filelines[5]=str(deltatime_1)+"\n"
                filelines[6]=str(analysis_check)+"\n"

                #Analysis 2
                filelines[8]=str(analysis_2_config)+"\n"
                filelines[9]=str(dates_2)+"\n"
                filelines[10]=str(ent_2)+"\n"
                filelines[11]=str(ram_2)+"\n"
                filelines[12]=str(grup_2)+"\n"
                filelines[13]=str(deltatime_2)+"\n"
                filelines[14]=str(analysis_2_check)+"\n"

                #Analysis 3
                filelines[16]=str(analysis_3_check)+"\n"
                filelines[17]=str(dates_3)+"\n"
                filelines[18]=str(ent_3)+"\n"
                filelines[19]=str(ram_3)+"\n"
                filelines[20]=str(grup_3)+"\n"
                filelines[21]=str(deltatime_3)+"\n"
                for i in range(0,len(lines_to_write)):
                    if 22+i<=len(filelines)-1:
                        filelines[22+i]=lines_to_write[i]+"\n"
                    else:
                        filelines.append(lines_to_write[i]+"\n")
                filelines = filelines[0:len(lines_to_write)+22]


                f.writelines(filelines)
            self.dismiss()
        else:
            #Analysis 1 dismiss
            new_analysis_check = analysis_check[:]
            new_analysis_config = analysis_config[:]
            new_dates_1 = dates_1[:]
            new_ent_1 = ent_1[:]
            new_ram_1 = ram_1[:]
            new_deltatime_1 = deltatime_1[:]

            #Analysis 2 dismiss
            new_analysis_2_check = analysis_2_check[:]
            new_analysis_2_config = analysis_2_config[:]
            new_dates_2=dates_2[:]
            new_ent_2 = ent_2[:]
            new_ram_2 = ram_2[:]
            new_grup_2 = grup_2[:]
            new_deltatime_2=deltatime_2[:]

            #Analysis 3 dismiss
            new_analysis_3_check = analysis_3_check[:]
            new_dates_3=dates_3[:]
            new_ent_3 = ent_3[:]
            new_ram_3 = ram_3[:]
            new_grup_3 = grup_3[:]
            new_deltatime_3=deltatime_3[:]
            new_text_input_3 = text_input_3[:]

            self.dismiss()
    pass

######################
#    Starting App    #
######################


class SUSEP(App):

    def build(self):
        self.icon = logo_path
        self.resizable = 0
        self.title = "Webscraper SUSEP"
        self.log_enable = 0
        return AppScreen()

######################
#        Main        #
######################

if __name__ == "__main__" :
	SUSEP().run()
