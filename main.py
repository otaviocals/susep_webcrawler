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
from kivy.core.window import Window
from kivy.config import Config
from kivy.core.window import Window
from kivy.clock import Clock
from os.path import join, isdir, expanduser, isfile
from sys import platform
from kivy.lang import Builder
from file import XFolder
import tools
from time import sleep
from webscraper import Webscraper
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
    phantom = "phantomjs"
elif current_os.startswith("win32") or current_os.startswith("cygwin"):
    slash = "\\"
    phantom = "phantomjs.exe"
elif current_os.startswith("darwin"):
    slash = "/"
    phantom = "phantomjs"
else:
    slash = "/"
    phantom = "phantomjs"


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path,relative_path)

logo_path = resource_path("logo.png")
phantom_path = resource_path(phantom)

#Setting links to scrap from

links = [
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20F%C3%ADsica%20-%20Aquisi%C3%A7%C3%A3o%20de%20outros%20bens&parametros='tipopessoa:1;modalidade:402;encargo:101'",
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20F%C3%ADsica%20-%20Aquisi%C3%A7%C3%A3o%20de%20ve%C3%ADculos&parametros='tipopessoa:1;modalidade:401;encargo:101'",
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20F%C3%ADsica%20-%20Cart%C3%A3o%20de%20cr%C3%A9dito%20parcelado&parametros='tipopessoa:1;modalidade:215;encargo:101'",
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20F%C3%ADsica%20-%20Cart%C3%A3o%20de%20cr%C3%A9dito%20rotativo&parametros='tipopessoa:1;modalidade:204;encargo:101'",
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20F%C3%ADsica%20-%20Cheque%20especial&parametros='tipopessoa:1;modalidade:216;encargo:101'",
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20F%C3%ADsica%20-%20Cr%C3%A9dito%20pessoal%20consignado%20INSS&parametros='tipopessoa:1;modalidade:218;encargo:101'",
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20F%C3%ADsica%20-%20Cr%C3%A9dito%20pessoal%20consignado%20privado&parametros='tipopessoa:1;modalidade:219;encargo:101'", #
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20F%C3%ADsica%20-%20Cr%C3%A9dito%20pessoal%20consignado%20p%C3%BAblico&parametros='tipopessoa:1;modalidade:220;encargo:101'",
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20F%C3%ADsica%20-%20Desconto%20de%20cheques&parametros='tipopessoa:1;modalidade:302;encargo:101'", #
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais-ModalidadeMensal.rdl&nome=Pessoa%20F%C3%ADsica%20-%20Financiamento%20imobili%C3%A1rio%20com%20taxas%20reguladas&parametros='tipopessoa:1;modalidade:905;encargo:101'", ##
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais-ModalidadeMensal.rdl&nome=Pessoa%20F%C3%ADsica%20-%20Financiamento%20imobili%C3%A1rio%20com%20taxas%20de%20mercado&parametros='tipopessoa:1;modalidade:903;encargo:101'", ##
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20F%C3%ADsica%20-%20Cr%C3%A9dito%20pessoal%20n%C3%A3o%20consignado&parametros='tipopessoa:1;modalidade:221;encargo:101'", #
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20F%C3%ADsica%20-%20Leasing%20de%20ve%C3%ADculos&parametros='tipopessoa:1;modalidade:1205;encargo:101'", #
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais-ModalidadeMensal.rdl&nome=Pessoa%20F%C3%ADsica%20-%20Financiamento%20imobili%C3%A1rio%20com%20taxas%20reguladas&parametros='tipopessoa:1;modalidade:905;encargo:201'", ##
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais-ModalidadeMensal.rdl&nome=Pessoa%20F%C3%ADsica%20-%20Financiamento%20imobili%C3%A1rio%20com%20taxas%20de%20mercado&parametros='tipopessoa:1;modalidade:903;encargo:201'", ##
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20jur%C3%ADdica%20-%20Antecipa%C3%A7%C3%A3o%20de%20faturas%20de%20cart%C3%A3o%20de%20cr%C3%A9dito&parametros='tipopessoa:2;modalidade:303;encargo:101'",
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20jur%C3%ADdica%20-%20Capital%20de%20giro%20com%20prazo%20at%C3%A9%20365%20dias&parametros='tipopessoa:2;modalidade:210;encargo:101'",
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20jur%C3%ADdica%20-%20Capital%20de%20giro%20com%20prazo%20superior%20a%20365%20dias&parametros='tipopessoa:2;modalidade:211;encargo:101'",
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20jur%C3%ADdica%20-%20Cheque%20especial&parametros='tipopessoa:2;modalidade:216;encargo:101'",
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20jur%C3%ADdica%20-%20Conta%20garantida&parametros='tipopessoa:2;modalidade:217;encargo:101'",
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20jur%C3%ADdica%20-%20Desconto%20de%20cheques&parametros='tipopessoa:2;modalidade:302;encargo:101'",
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20jur%C3%ADdica%20-%20Desconto%20de%20duplicata&parametros='tipopessoa:2;modalidade:301;encargo:101'",
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20jur%C3%ADdica%20-%20Vendor&parametros='tipopessoa:2;modalidade:404;encargo:101'",
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20jur%C3%ADdica%20-%20Capital%20de%20giro%20com%20prazo%20at%C3%A9%20365%20dias&parametros='tipopessoa:2;modalidade:210;encargo:204'",
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20jur%C3%ADdica%20-%20Capital%20de%20giro%20com%20prazo%20superior%20a%20365%20dias&parametros='tipopessoa:2;modalidade:211;encargo:204'",
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20jur%C3%ADdica%20-%20Conta%20garantida&parametros='tipopessoa:2;modalidade:217;encargo:204'",
    "http://www.bcb.gov.br/pt-br/#!/r/txjuros/?path=conteudo%2Ftxcred%2FReports%2FTaxasCredito-Consolidadas-porTaxasAnuais.rdl&nome=Pessoa%20jur%C3%ADdica%20-%20Adiantamento%20sobre%20contratos%20de%20c%C3%A2mbio&parametros='tipopessoa:2;modalidade:502;encargo:205'"]


#Building GUI

Builder.load_string('''
<AppScreen>
    BoxLayout:
        orientation: "horizontal"
        Image:
            source: ""
            size_hint: None, None
            size: 150,200
            pos_hint: {"center_y": 0.5, "center_x": 0.5}
            id: ce_logo
        Label:
            text:"Webscraper de Taxa de Juros"
            markup: True
            size_hint: None, None
            size: 300, 200
            pos_hint: {"center_y": 0.5, "center_x": 0.5}
    BoxLayout:
        orientation: "vertical"
        id: box_2
        XButton:
            text: 'Selecione a Pasta'
            on_release: root._folder_dialog()
            size_hint: None, None
            size: 150, 50
            pos_hint: {"center_x": 0.5, "center_y": 0.5}
            id: folder_button
        Label:
            text: "Pasta selecionada:    " + root.default_folder
            size: 450, 30
            size_hint: None, None
            id: text_input
    BoxLayout:
        orientation: "vertical"
        Label:
            text:"Frequência de Leitura"
            markup: True
            size_hint: None, None
            size: 450, 100
        Spinner:
            text: "10 minutos"
            values: ("10 minutos","15 minutos","1 hora","6 horas","12 horas")
            size: 150,50
            size_hint: None, None
            pos_hint: {"center_x": 0.5, "center_y": 0.5}
            id: spinner
    FloatLayout:
        ToggleButton:
            text: "Stop" if self.state == "down" else "Start"
            size: 150, 50
            size_hint: None, None
            pos_hint: {"center_y": 0.5, "center_x": 0.5}
            group: "start"
            on_state: root.start(*args)
            id: start_button
    BoxLayout
        orientation: "vertical"
        TextInput:
            text: ""
            hint_text: "Log Output"
            multiline: True
            readonly: True
            id: log_output

''')



######################
#    App Functions   #
######################

class AppScreen(GridLayout):


    default_folder = expanduser("~") + slash + "Documents"

    global sel_folder
    sel_folder = default_folder

#Calling webscraper

    def scrap(self,event):
        log_file = open((sel_folder+slash+"history.log"),"a",encoding="utf-8")

        print("Starting at "+str(datetime.datetime.now()))
        log_file.write("Starting at "+str(datetime.datetime.now())+"\n")
        self.ids.log_output.text += "Starting at "+str(datetime.datetime.now())+"\n"

        print("Download Starting...")
        log_file.write("Download Starting...\n")
        self.ids.log_output.text += "Download Starting...\n"

        start_time_seconds = time.time()
        self.ids.start_button.disabled = True
        for elemen in range(0, len(links)):
            str_output_logger = []

            #print(links[elemen])
            #print(sel_folder)
            #print(log_file)
            #print(str_output_logger)
            #print(phantom_path)

            Webscraper(links[elemen],sel_folder,log_file,str_output_logger,phantom_path)

            self.ids.log_output.text += "".join(str_output_logger)
            gc.collect()
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

#GUI functions

    def start(self,*args):
        global event
        if args[1] == "down":
            self.ids.spinner.disabled = True
            self.ids.folder_button.disabled = True
            if self.ids.spinner.text == "10 minutos":
                time_interval = 600
            elif self.ids.spinner.text == "30 minutos":
                time_interval = 1800
            elif self.ids.spinner.text == "1 hora":
                time_interval = 3600
            elif self.ids.spinner.text == "6 horas":
                time_interval = 21600
            elif self.ids.spinner.text == "12 horas":
                time_interval = 43200
            global event
            event = Clock.schedule_once(self.scrap,0.5)
            global event
            event = Clock.schedule_interval(self.scrap, time_interval)
        if args[1] == "normal":
            self.ids.spinner.disabled = False
            self.ids.folder_button.disabled = False
            global event
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


######################
#    Starting App    #
######################


class Taxa_de_JurosApp(App):

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
    Taxa_de_JurosApp().run()
