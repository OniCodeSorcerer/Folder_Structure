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
        self.ober_ordner = "Aufträge"
           
        layout = GridLayout(cols= 2, spacing=10, padding= 50)
        widget_height = 60

        geber_label = Label(text= "Auftraggeber: ",size_hint_y=None, height=30)
        layout.add_widget(geber_label)

        self.geber_text = TextInput(multiline= False,size_hint_y=None, height=30)
        layout.add_widget(self.geber_text)

        arbeit_label = Label(text= "Auftragsarbeit: ",size_hint_y=None, height=30)
        layout.add_widget(arbeit_label)

        self.arbeit_text = TextInput(multiline= False,size_hint_y=None, height=30)
        layout.add_widget(self.arbeit_text)

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
            self.ober_ordner = "Kleinstaufträge"

        elif instance.text == "Klärwerkstechnik":
            self.ober_ordner = "Aufträge"
            self.kwtmva = "KWT_"


        elif instance.text == "Metallverarbeitung":
            self.ober_ordner = "Aufträge"
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
        neuer_ordner_name = f"{num:03d}_{datum}_{self.kwtmva}{auftragsname}-{arbeit}"
        texteingabe = f"Erstellt am: {datum}"
        ober_ordner = self.ober_ordner
        ober_dir = os.path.join(os.getcwd(), ober_ordner, neuer_ordner_name)

        os.chdir(ober_ordner)

        if not os.path.exists(ober_dir):
            
            os.makedirs(neuer_ordner_name)
            print(f"Der Ordner '{neuer_ordner_name}' wurde erfolgreich erstellt.")

            try:
                manufactory = f"MFC_{neuer_ordner_name}"
                manufactory_dir = os.path.join(ober_dir, manufactory)
                os.symlink(ober_dir, manufactory_dir)
                print(f"Die Ordnerverknüpfung 'Manufaktur' wurde erfolgreich erstellt.")

            except PermissionError:
                print("keine Rechte vorhanden. Programm als Admin starten")

            with open(os.path.join(neuer_ordner_name, f"{auftragsname}_{arbeit}.txt"), "w") as file:
                file.write(texteingabe)

        else:
            print(f"Der Ordner '{neuer_ordner_name}' existiert bereits.")

        script_directory = os.path.dirname(__file__)
        os.chdir(script_directory)
        with open("current_number.txt", "w") as file:
            file.write(str(num + 1))

elevate.elevate()
FolderCreator().run()
