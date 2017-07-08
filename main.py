#                                                                                                       #
#   Just Another Webscraper (J.A.W.)                                                                    #
#                                    v0.1                                                               #
#                                                                                                       #
#       written by Otavio Cals                                                                          #
#                                                                                                       #
#   Description: A webscrapper for downloading tables and exporting them to .csv files autonomously.    #
#                                                                                                       #
#########################################################################################################

#Required External Modules: cython, pygame, kivy

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
from webscraper.webscraper import Webscraper
import time
import datetime
import gc
import sys
import os
import shutil

######################
#    Pre-Settings    #
######################

#Setting configurations

Config.set("kivy","log_enable","1")
Config.set("kivy","log_level","debug")
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
	
print("hehe")


if current_os.startswith("linux"):
    slash = "/"
    phantom = "phantomjs/phantomjs"
elif current_os.startswith("win32") or current_os.startswith("cygwin"):
    slash = "\\"
    phantom = "phantomjs\\phantomjs.exe"
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
    return os.path.join(base_path,relative_path)

logo_path = resource_path("logo"+slash+"logo.png")
phantom_path = resource_path(phantom)
r_script_path = resource_path("rscripts")
r_libs_path = resource_path("rlibs")
kivi_app_path = resource_path("kivylibs"+slash+"app_screen.kv")
config_path = os.path.abspath(".")+slash+"config.txt"

#print(os.path.abspath("."))

#Initiating config data

analysis_check = [True]
analysis_config = [False]*18
dates_1=["01/01/2017","01/02/2017"]
ent_1=[True,""]
ram_1=[True,""]
deltatime_1=[True,False,False,False]

#Loading saved configurations

if(Path(config_path).is_file()):
    with open(config_path,"r") as f:
        configs = f.readlines()
        configs = [x.strip() for x in configs]

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
else:
    with open(config_path,"a") as f:
        f.write("Config Analysis 1\n")
        f.write(str(analysis_config)+"\n")
        f.write(str(dates_1)+"\n")
        f.write(str(ent_1)+"\n")
        f.write(str(ram_1)+"\n")
        f.write(str(deltatime_1)+"\n")
        f.write(str(analysis_check)+"\n")

new_analysis_check = analysis_check[:]
new_analysis_config = analysis_config[:]
new_dates_1 = dates_1[:]
new_ent_1 = ent_1[:]
new_ram_1 = ram_1[:]
new_deltatime_1 = deltatime_1[:]

#Setting links to scrap from

links = []


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
#        popup.ids.scroll_grid_2.bind(minimum_height=popup.ids.scroll_grid_2.setter("height"))
#        popup.ids.scroll_grid_3.bind(minimum_height=popup.ids.scroll_grid_3.setter("height"))

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

        new_analysis_check = analysis_check[:]
        new_analysis_config = analysis_config[:]
        new_dates_1 = dates_1[:]
        new_ent_1 = ent_1[:]
        new_ram_1 = ram_1[:]
        new_deltatime_1 = deltatime_1[:]

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
    def _checkanalysis_callback(self,value):
        global new_analysis_check
        new_analysis_check[0] = value

    def _checkbox_callback(self,n,value):
        global new_analysis_config
        new_analysis_config[n]=value

    def _entbox_callback(self,value):
        global new_ent_1
        new_ent_1[0]=value

    def _rambox_callback(self,value):
        global new_ram_1
        new_ram_1[0]=value

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
        if value:
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
            with open(config_path,"r") as f:
                filelines=f.readlines()
            with open(config_path,"w") as f:
                filelines[1]=str(analysis_config)+"\n"
                filelines[2]=str(dates_1)+"\n"
                filelines[3]=str(ent_1)+"\n"
                filelines[4]=str(ram_1)+"\n"
                filelines[5]=str(deltatime_1)+"\n"
                filelines[6]=str(analysis_check)+"\n"
                f.writelines(filelines)
            self.dismiss()
        else:
            new_analysis_check = analysis_check[:]
            new_analysis_config = analysis_config[:]
            new_dates_1 = dates_1[:]
            new_ent_1 = ent_1[:]
            new_ram_1 = ram_1[:]
            new_deltatime_1 = deltatime_1[:]
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
        self.log_enable = 1
        return AppScreen()

######################
#        Main        #
######################

if __name__ == "__main__" :
	print("hehe")
	SUSEP().run()
