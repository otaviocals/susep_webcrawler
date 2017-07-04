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
from sys import platform
import subprocess
from pathlib import Path
from kivy.lang import Builder
from libs.file import XFolder
from libs.xpopup import XPopup
import libs.tools
from time import sleep
from webscraper.webscraper import Webscraper
import time
import datetime
import gc
import sys
import os


######################
#    Pre-Settings    #
######################

#Setting configurations

Config.set("kivy","log_enable","0")
Config.set("kivy","log_level","critical")
Config.set("graphics","width","450")
Config.set("graphics","height","900")
Config.set("graphics","fullscreen","0")
Config.set("graphics","borderless","0")
Config.set("graphics","resizable","0")
Config.write()

current_os = platform


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
r_script_path = resource_path("rscripts"+slash+"analysis.R")
kivi_app_path = resource_path("kivy"+slash+"app_screen.kv")
config_path = resource_path("config.txt")



#global analysis_config
analysis_config = [False]*50

#Loading saved configurations

if(Path(config_path).is_file()):
    with open(config_path,"r") as f:
        configs = f.readlines()
        configs = [x.strip() for x in configs]
        analysis_config_string = configs[0][1:len(configs[0])-1].split(", ")
        i=0
        for x in analysis_config_string:
            analysis_config[i] = (x == "True")
            i+=1
else:
    with open(config_path,"w") as f:
        print(analysis_config,file=f)

#global new_analysis_config
new_analysis_config = analysis_config[:]

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
        log_file = open((sel_folder+slash+"logs"+slash+"history.log"),"a",encoding="utf-8")

    #Starting scrap

        print("Starting at "+str(datetime.datetime.now()))
        log_file.write("Starting at "+str(datetime.datetime.now())+"\n")
        self.ids.log_output.text += "Starting at "+str(datetime.datetime.now())+"\n"

        print("Download Starting...")
        log_file.write("Download Starting...\n")
        self.ids.log_output.text += "Download Starting...\n"

        start_time_seconds = time.time()
        self.ids.start_button.disabled = True

    #Running webscraper

        str_output_logger = []
        new_data = Webscraper(sel_folder,log_file,str_output_logger,phantom_path)
        self.ids.log_output.text += "".join(str_output_logger)
        gc.collect()
        print(new_data)

    #Running R scripts

        r_output = subprocess.check_output(["Rscript",r_script_path],universal_newlines=True)
        print(r_output)

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
        popup.ids.scroll_grid_2.bind(minimum_height=popup.ids.scroll_grid_2.setter("height"))
        popup.ids.scroll_grid_3.bind(minimum_height=popup.ids.scroll_grid_3.setter("height"))

        popup.title="Selecao de Consultas"

        global analysis_config
        global new_analysis_config
        new_analysis_config = analysis_config[:]
        for i in range(0,50):
            popup.ids["check_"+str(i)].active=new_analysis_config[i]

        popup.open()


#Setting window layout

    def __init__(self,**kwargs):
        super(AppScreen, self).__init__(**kwargs)

        Window.size = (450,750)
        Window.set_title("Webscraper de Taxa de Juros")


        self.cols = 1
        self.size_hint = (None,None)
        self.width = 450
        self.height = 750
        self.icon = logo_path
        self.ids.ce_logo.source = logo_path


class CheckPopup(XPopup):
    def _checkbox_callback(self,n,value):
        global new_analysis_config
        new_analysis_config[n]=value

    def _check_popup_callback(self,value):
        global new_analysis_config
        global analysis_config
        if value:
            analysis_config = new_analysis_config[:]
            with open(config_path,"w") as f:
                print(analysis_config,file=f)
            self.dismiss()
        else:
            new_analysis_config = analysis_config[:]
            self.dismiss()
    pass

######################
#    Starting App    #
######################


class SUSEP(App):

    def build(self):
        self.icon = logo_path
        self.resizable = 0
        self.title = "Webscraper de Taxa de Juros"
        self.log_enable = 0
        return AppScreen()

######################
#        Main        #
######################

if __name__ == "__main__" :
    SUSEP().run()
