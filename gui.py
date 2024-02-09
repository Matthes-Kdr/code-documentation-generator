'''
Stellt eine GUI zur Einstellung von Parametern für die durchzuführende Code-Dokumentation bereit. Alle Inputs werden durch dieses Modul abgefragt und - soweit erforderlich - auch innerhalb von einzelnen Objekten gespeichert.

'''


import os
import tkinter as tk
from tkinter import filedialog, simpledialog





def benutzerdefinierte_texte_on_buttons():
    """
    # TODO: Ansoatz  https://stackoverflow.com/questions/16242782/change-words-on-tkinter-messagebox-buttons
    """

    '''
    # from tkinter import *

    def messageWindow():
        win = Toplevel()
        win.title('warning')
        message = "This will delete stuff"
        Label(win, text=message).pack()
        Button(win, text='Delete', command=win.destroy).pack()
    root = Tk()
    Button(root, text='Bring up Message', command=messageWindow).pack()
    root.mainloop()
    '''




class DocumenterGui:
    """
    ### AUSBLICK: Es sollte für die Anwendung i Code-Documenter noch ein Textfeld geben, in dem beliebiger Freitext uebergeben werden kann - oder ein Input-box (extra win mit allem sichtbaren bereich, also nicht nur eine Zeile... Dieser Freitext soll auch in der DCode-Dokumentation iwo eingebaut werdn oben...)


    ### TODOC : ...


    """


    WIDTH_MAIN_WINDOW = 575
    HEIGHT_MAIN_WINDOW = 435


    TITLE = "Code Documenter"

    DEBUG = 1

    def __init__(self):
        """
        Initialisiert ein Hauptfenster zum Parametrisieren der Einstellungen fuer die vorzunehmende automatisierte Dokumentation.
        Das Layout des GUI-Fensters wird durch eine eigene Methode bestimmt, die hier aufgerufen wird.
        Der Mainloop wird im Anschluss direkt gestartet.
        """

        self.__is_ready = False # Flag, auf 
        
        
        
        if self.DEBUG:
             
             self.debug_skip_main_win()
             
             return
        


        # Generate and initialize root:
        self.root = tk.Tk()
        self.root.title(self.TITLE)
        self.root.geometry(f"{self.WIDTH_MAIN_WINDOW}x{self.HEIGHT_MAIN_WINDOW}")
        self.root.resizable(width=False, height=False)  # Make the window not resizable

        # Initialize variables for entry elements and checkboxes:
        self.convert_to_html_var = tk.BooleanVar(value=True)
        self.show_message_var = tk.BooleanVar(value=True)

        self.show_message:bool = True
        self.convert_checked:bool = True

        self.output_dir:str = "" # inital
        self.input_file:str = "" # inital

        self.optinal_user_defined_text:str = "optionaler user defind text..."


        # Build / fill GUI with  elements
        self.create_widgets()

        # start loop:
        self.root.mainloop()



    def get_is_ready(self):
        """
        Gibt den Wert zurueck, der als Flag dient, ob alle Parameter durch die GUI gesetzt wurden und ob die eigentliche Code-Dokumentation von ausserhalb der Klasse gestartet werden soll.

        Returns:
            bool: True falls gestartet werden kann/soll, False falls Abbruch oder zumindest noch kein Start erfolgen soll.

        """
        return self.__is_ready
    



    def debug_skip_main_win(self):
        """
        ### NUR ZUM DEBUGGEN!
        Dient der Simulation des Start-Windows inkl. Initialisierung aller Parameter als Dummys, 
        damit das Start-Window (Main-window) zwecks Debugging uebersprungen werden kann.
        """
        
        self.input_file = "input_data/beispiel_modul_rekursiv.bas"
        self.input_file = "input_data/beispiel_modul_bauer+liebherr.bas"
        self.input_file = "input_data/beispiel_modul2.bas"
        self.input_file = "input_data/beispiel_modul1.bas"
        self.input_file = "input_data/beispiel_modul.bas"
        self.input_file = "input_data/current_test.bas"


        self.output_dir = "output_data"
        
        self.show_message = True
        self.convert_checked = True
        self.optinal_user_defined_text = "optionaler user defind text..."


        self.__is_ready = True





    def get_input_file_path_entry(self):
            return self.tx_input_file_path.get()
    

    
    def get_output_dir_path_entry(self):
            return self.txt_output_dir_path.get()


    def create_widgets(self):
        """
        Definition des Layouts des Haupt-Windows zur Parametrisierung der Einstellungen fuer die vorzunehmende Dokumentation.
        """
        
        # Label:
        temp_lbl_text = "Durch das vorliegende Programm kann eine automatisiert ein Code-Dokumentation von einzelnen VBA-Modulen generiert werden.\n\nHierzu muss der Dateipfad zur zu dokumentierenden .bas-Datei angegeben werden, sowie das Verzeichnis, in welches die Dokumentation exportiert werden soll."
        lbl = tk.Label(self.root, text=temp_lbl_text, wraplength=540, justify=tk.LEFT)
        lbl.grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=10, pady=5)



        # INput File Path Entry and Browse Button:
        self.tx_input_file_path = tk.Entry(self.root, width=65 + 4, state="readonly")
        self.tx_input_file_path.grid(row=1, column=0, padx=10, pady=5, columnspan=2, sticky=tk.W)

        btn_browse_file = tk.Button(self.root, text="Suche Input-File...", command=self.browse_file)
        btn_browse_file.grid(row=1, column=1, padx=5, sticky=tk.E)


        # Output Directory Path Entry and Browse Button
        self.txt_output_dir_path = tk.Entry(self.root, width=65, state="disabled")
        self.txt_output_dir_path.grid(row=2, column=0, padx=10, pady=5, columnspan=2, sticky=tk.W)

        btn_browse_dir = tk.Button(self.root, text="Suche Zielverzeichnis...", command=self.browse_dir)
        btn_browse_dir.grid(row=2, column=1, padx=5, sticky=tk.E)


        # Label:
        temp_lbl_text = "Durch das Programm wird IMMER eine Markdown-Datei ('.md') generiert.\n\nOptional kann zusätzlich aus dieser Datei direkt im Anschluss eine HTML-Datei erstellt werden. Durch unterschiedliche Interpretationen im Rahmen dieser Konvertierung können unterschiedliche Tools zu anderen optischen Erscheinungen in der so generierten HTML-Datei führen. Übersichtlicher und sauberer wird die Umwandlung unter Nutzung der VSCode-Extension 'Markdown Preview Enhanced'. Nach der Generierung der Dokumentation sollte daher nochmals manuell eine solche Umwandlung erfolgen."
        lbl = tk.Label(self.root, text=temp_lbl_text, wraplength=540, justify=tk.LEFT)
        lbl.grid(row=3, column=0, columnspan=2, sticky=tk.W, padx=10, pady=5)

        # Convert Checkbox and Show Message Checkbox
        temp_cb_text = "Erstelle zusätzliche HTML-Datei"
        cb_convert_to_html = tk.Checkbutton(self.root, text=temp_cb_text, variable=self.convert_to_html_var)
        cb_convert_to_html.grid(row=4, column=0, sticky=tk.W, padx=10, pady=5, columnspan=2)

        temp_cb_text = "Im Rahmen der Dokumentation der Call Sequences (Aufrufreihenfolge von Prozeduren):\nFüge einen Endhinweis für jede Prozedur hinzu, wenn kein weiterer Aufruf zu einer (durch dieses Tool dokumentierten) Prozedur erfolgt."
        cb_show_message = tk.Checkbutton(self.root, text=temp_cb_text, variable=self.show_message_var, wraplength=540, justify=tk.LEFT)
        cb_show_message.grid(row=5, column=0, sticky=tk.W, padx=5, pady=5, columnspan=2)


        # Run and Cancel Buttons
        btn_start = tk.Button(self.root, text="Dokumentation generieren", command=self.start_process)
        btn_start.grid(row=6, column=0, padx=10, pady=10, sticky=tk.W)

        btn_cancel = tk.Button(self.root, text="Abbrechen", command=self.cancel_process)
        btn_cancel.grid(row=6, column=1, padx=5, pady=10, sticky=tk.E)






    @staticmethod
    def update_entry_text(entry_obj, new_text:str):
        """
        Methode modifiziert den Text innerhalb eines tk.Entry-Objektes und schreibt den uebergebenen Text in das Entry.
        Damit kann der Text z. B. in Abhängigkeit einer vorherigen Auswahl aktualisiert werden.
        Nach der Modifikation wird das Objekt wieder disabled gesetzt, um manuelle Fehleranfaellige Korrekturen zu verhindern.

        Args:
            entry_obj (tk.Entry): Entry-Objekt, welches mit neuem Text versehen werden soll
            new_text (str): Text, der im Entry-Objekt angezeigt werden soll.
        """
        entry_obj.config(state="normal")
        entry_obj.delete(0, tk.END)
        entry_obj.insert(0, new_text)
        entry_obj.config(state="disabled")



    def browse_file(self):
        """
        Stellt einen FileDialog zur Auswahl von bestimmten Dateitypen bereit. 
        Nach erfolgreicher Auswahl einer Datei durch den Anwender wird eine MEthode aufgerufen,
        durch die die Entry-Objekte entsprechend angepasst werden, sodass dort der ausgewählte Text angezeigt wird.

        Prueft ausserdem, ob bereits ein Output-Folder angegeben wurde. 
        Falls nicht, wird als Vorschlag das Verzeichnis voreingestellt, in dem sich die Input-File befindet.
        """
        
        # Einschraenkungen der Dateitypen:
        extensions = [("VB", ".bas")]
        default_extension = extensions[0]
        
        # Auswahl durch Anwender:
        file_path = filedialog.askopenfilename(defaultextension=default_extension, filetypes=extensions)

        self.update_entry_text(self.tx_input_file_path, file_path)

        # Sofern nicht bereits ein anderer Output-Folder gewaehlt wurde, nehme diesen der Input-File:
        if not self.output_dir:
            # Setze output wie input dir:
            suggested_output_dir = os.path.dirname(file_path)
            self.update_entry_text(self.txt_output_dir_path, suggested_output_dir)




    def browse_dir(self):
        """
        Stellt einen FileDialog zur Auswahl eines Verzeichnisses bereit. 
        Nach erfolgreicher Auswahl einer Datei durch den Anwender wird eine MEthode aufgerufen,
        durch die das Entry-Objekte entsprechend angepasst wird, sodass dort der ausgewählte Text angezeigt wird.
        """
        self.output_dir = filedialog.askdirectory()
        self.update_entry_text(self.txt_output_dir_path, self.output_dir)





    def start_process(self):
        """
        Methode, die beim Klicken auf den Button 'Start' aufgerufen wird.
        Es werden zunächst alle relevanten Variablen aus der GUI in Objektvariable gespeichert.
        Nach einer erfolgreichen Gueltigkeitspruefung wird eine Messagebox mit einer Infonachricht angezeigt.
        Im Anschluss wird das Hauptfenster destroyed.
        Die GUI ist somit zunaechst nicht mehr auf dem Bildschirm aktiv. 
        Auf die  Variablen kann aber weiterhin von außen zugegriffen werden.
        """
        self.show_message = self.show_message_var.get()
        self.convert_checked = self.convert_to_html_var.get()

        self.output_dir = self.txt_output_dir_path.get()
        self.input_file = self.tx_input_file_path.get()

        if not all([self.output_dir, self.input_file]):
            # Falls nciht Alle relevanten Felder ausgefuellt sind:
            
            simpledialog.messagebox.showinfo("Fehlerhafte Eingabe", "Bitte über die Buttons die Pfade für die Input-Datei und das Output-Verzeichnis angeben.", icon="error")
            # # ALTERNATIVER SHORTCUT: (alias):
            # simpledialog.messagebox.showerror("Fehlerhafte Eingabe", "Bitte über die Buttons die Pfade für die Input-Datei und das Output-Verzeichnis angeben.")

            return


        # Hiinweis fuer den Anwender in Msgbox mit anschliessender Return-value Abfrage:
        val = self.msgbox_btn("Die Dokumentation wird erstellt. Bitte warten, bis eine Abschlussmeldung erscheint.")

        if val:
            # Schliessen des Fensters:
            self.root.destroy()
            self.__is_ready = True # Flag fuer Zugriff von außen setzen






    def cancel_process(self):
        """
        Beende / Verberge das Hauptfenster.
        """

        self.root.destroy()



    def msgbox_btn(self, message):
        """
        Einfache Messagebox mit OK und Cancel Buttons

        Args:
            message (str): Anzuzeigender Text.

        Returns:
            bool: Users return value
        """

        root = tk.Tk()
        root.withdraw()  # Hide the main window


        val = simpledialog.messagebox.askokcancel("Information", message)
        
        root.destroy()

        return val
    










    def show_closing_window(self):
        """
        Erstellt ein neues Fenster, in dem saemtliche Einstellungen nochmals aufgefuehrt werden (read-only).
        Anwendung als Closing-Window nach der Durchfuehrung der eigentlichen Dokumentation.
        """

        self.root = tk.Tk()
        self.root.title(self.TITLE)
        self.root.resizable(width=True, height=True)  # Make the window not resizable


        # Inhalte, die untereinander aufgefuehrt werden sollen: 
        contents = [
            "Die Dokumentation des Codes wurde generiert.",
            "",
            "Verwendete Parameter:",
            "-" * 60,
            "Input-File: {}".format(self.input_file),
            "Output-directory: {}".format(self.output_dir),
            "",
            "{}  :  Erstellung einer HTML-Datei.".format(self.convert_checked),
            "{}  :  Hinzufügen eines Endhinweises für jede Prozedur, in der kein weiterer Aufruf erfolgt.".format(self.show_message),
            "",
            "Dieses Fenster kann nun geschlossen werden."
        ]


        for _ in contents:
             temp_label = tk.Label(self.root, text=_).pack(anchor="nw")


        self.root.mainloop()










def demo_ablauf():
    """
    Zeigt die moegliche Anwendung der DocumenterGui-Klasse, so dass sie selbst nur PASSIV im Gesamt-Ablauf verwendet wird und nicht aktiv etwas von außerhalb der Klasse aufrufen muss.
    """

    gui = DocumenterGui()
    
    
    if gui.get_is_ready():
        print("Dann starte Dokumentation von aussen mit den Parametern durch Zugriff auf die Objektvariablen")
    else:
        print("KEIN Start!")


    print("Abschluss")




if __name__ == "__main__":


    DocumenterGui.DEBUG = 1

    demo_ablauf()


