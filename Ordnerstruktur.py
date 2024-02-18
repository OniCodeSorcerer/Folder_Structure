import os
import elevate
from datetime import datetime
import kivy
from kivy.core.window import Window
from kivy.app import App
from kivy.uix.dropdown import DropDown
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.switch import Switch

class FolderCreator(App):
    def build(self):

        self.current_number = self.get_current_number()
        self.kwtmva = ""
        self.ueber_ordner = "Aufträge"
           
        layout = GridLayout(cols= 2, spacing=5, padding= 50)
        widget_height = 60

        geber_label = Label(text= "Auftraggeber: ",size_hint_y=None, height=30)
        layout.add_widget(geber_label)

        self.geber_text = TextInput(multiline= False,size_hint_y=None, height=30)
        layout.add_widget(self.geber_text)

        arbeit_label = Label(text= "Auftragsarbeit: ",size_hint_y=None, height=30)
        layout.add_widget(arbeit_label)

        self.arbeit_text = TextInput(multiline= False,size_hint_y=None, height=30)
        layout.add_widget(self.arbeit_text)

        strasse_label = Label(text= "Straße: ",size_hint_y=None, height=30)
        layout.add_widget(strasse_label)

        self.strasse_text = TextInput(multiline= False,size_hint_y=None, height=30)
        layout.add_widget(self.strasse_text)

        plz_label = Label(text= "PLZ / Ort: ",size_hint_y=None, height=30)
        layout.add_widget(plz_label)

        self.plz_text = TextInput(multiline= False,size_hint_y=None, height=30)
        layout.add_widget(self.plz_text)

        telefon_label = Label(text= "Telefonnr.: ",size_hint_y=None, height=30)
        layout.add_widget(telefon_label)

        self.telefon_text = TextInput(multiline= False,size_hint_y=None, height=30)
        layout.add_widget(self.telefon_text)

        dropdown_label = Label(text= "Ordnerauswahl: ")
        layout.add_widget(dropdown_label)        
        
        button_label = Label(text= f"Aktuelle Nummer: {self.current_number}")
        layout.add_widget(button_label)

        self.dropdown_button = Button(text="Ausgewählter Ordner",size_hint_y=None, height=widget_height)
        self.dropdown_button.bind(on_release=self.show_dropdown)

        self.dropdown = DropDown()
        for item in ['Kleinstaufträge', 'Klärwerkstechnik', 'Metallverarbeitung']:
            btn = Button(text=item, size_hint_y=None, height=widget_height)
            btn.bind(on_release=self.select_dropdown_option)
            self.dropdown.add_widget(btn)
        layout.add_widget(self.dropdown_button)

        erstellen_button = Button(text= "Ordner Erstellen", on_press=self.create_folder,size_hint_y=None, height=widget_height)
        layout.add_widget(erstellen_button)

        Window.size = (600,380)
        return layout
    

    def show_dropdown(self, instance):
        self.dropdown.open(instance)


    def select_dropdown_option(self, instance):
        print(f"Ausgewählte Option: {instance.text}")
        self.dropdown_button.text = instance.text
        self.dropdown.dismiss()
        if instance.text == "Kleinstaufträge":
            self.ueber_ordner = "04-Kleinstaufträge"
            self.kwtmva = "KAT_"

        elif instance.text == "Klärwerkstechnik":
            self.ueber_ordner = "03-Aufträge"
            self.kwtmva = "KWT_"


        elif instance.text == "Metallverarbeitung":
            self.ueber_ordner = "03-Aufträge"
            self.kwtmva = "MVA_"


    def get_current_number(self):
        if os.path.exists("current_number.txt"):
            with open("current_number.txt", "r") as file:
                return int(file.read().strip())
        else:
            return 1


    def create_folder(self,instance):
        auftragsname = self.geber_text.text
        arbeit = self.arbeit_text.text
        datum = datetime.now().strftime("%Y%m%d")
        num = self.current_number
        auftrag_ordner_name = f"{num:03d}_{datum}_{self.kwtmva}{auftragsname}-{arbeit}"
        texteingabe = f"Erstellt am: {datum} \nAuftraggeber: {auftragsname} \nAuftragsarbeit: {arbeit} \nStraße: {self.strasse_text.text} \nPLZ / Ort: {self.plz_text.text} \nTelefonnr.: {self.telefon_text.text} \n"
        ueber_ordner = self.ueber_ordner        
        manufactory_ordner = "07-Manufactory"
        ueber_dir = os.path.join(os.getcwd(), ueber_ordner, auftrag_ordner_name)
        manufactory_dir = os.path.join(os.getcwd(), manufactory_ordner, auftrag_ordner_name)
        script_directory = os.path.dirname(__file__)

        if not os.path.exists(ueber_ordner):
            os.makedirs(ueber_ordner)
            print(f"Der Ordner '{ueber_ordner}' wurde erfolgreich erstellt.")

        else:
            print(f"Der Ordner '{manufactory_ordner}' existiert bereits.")

        if not os.path.exists(manufactory_ordner):
            os.makedirs(manufactory_ordner)
            print(f"Der Ordner '{manufactory_ordner}' wurde erfolgreich erstellt.")

        else:
            print(f"Der Ordner '{manufactory_ordner}' existiert bereits.")

        os.chdir(manufactory_ordner)

        if not os.path.exists(auftrag_ordner_name):
            os.makedirs(auftrag_ordner_name)
            print(f"Der Ordner '{auftrag_ordner_name}' wurde erfolgreich im {manufactory_ordner} erstellt.")

        else:
            print(f"Der Ordner '{manufactory_ordner}' existiert bereits.")

        os.chdir(script_directory)
        os.chdir(ueber_ordner)

        if not os.path.exists(auftrag_ordner_name):      
            os.makedirs(auftrag_ordner_name)
            print(f"Der Ordner '{auftrag_ordner_name}' wurde erfolgreich erstellt.")

            try:
                manufactory = f"MFC_{auftrag_ordner_name}"
                manufactory_ver = os.path.join(ueber_dir, manufactory)
                os.symlink(manufactory_dir, manufactory_ver)
                print(f"Die Ordnerverknüpfung {manufactory} wurde erfolgreich erstellt.")

            except PermissionError:
                print("keine Rechte vorhanden. Programm als Admin starten!")

            with open(os.path.join(auftrag_ordner_name, f"{auftragsname}_{arbeit}.txt"), "w") as file:
                file.write(texteingabe)

        else:
            print(f"Der Ordner '{auftrag_ordner_name}' existiert bereits.")

        
        os.chdir(script_directory)
        with open("current_number.txt", "w") as file:
            file.write(str(num + 1))

        self.current_number += 1

elevate.elevate()
FolderCreator().run()
