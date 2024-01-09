'''
### TODO: Dies ist nur ein ANSATZ



Modul stellt einen a-nsatz für eine - teilweise schon funktionale einfache GUI zur Auswahl der Parameter. Vielleicht könnte/sollte dieses Modul wirklich in einem unabhägnigen Modul blebien??!


Es fehlt noch Funktionalität und Prüfung (Gültigkeit von Dateiendungen usw...)

'''


import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog





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






class DocumenterParameterGui:
    """
    ### TODO: Es sollte für die Anwendung i Code-Documenter noch ein Textfeld geben, in dem beliebiger Freitext uebergeben werden kann - oder ein Input-box (extra win mit allem sichtbaren bereich, also nicht nur eine Zeile... Dieser Freitext soll auch in der DCode-Dokumentation iwo eingebaut werdn oben...)


    ### TODOC : ...


    """
    # WIDTH = 650
    # WIDTH = 550
    WIDTH_MAIN_WINDOW = 575
    HEIGHT_MAIN_WINDOW = 435


    TITLE = "Code Documenter"

    # def __init__(self, root):
        # self.root = root
    def __init__(self, DEBUG=0):

        if DEBUG:
             
             self.debug_skip_main_win()
             
             return

        self.root = tk.Tk()
        self.root.title(self.TITLE)
        self.root.geometry(f"{self.WIDTH_MAIN_WINDOW}x{self.HEIGHT_MAIN_WINDOW}")
        self.root.resizable(width=False, height=False)  # Make the window not resizable

        # Initialize variables
        self.convert_var = tk.BooleanVar(value=True)
        self.show_message_var = tk.BooleanVar(value=True)

        self.output_dir = "" # inital
        self.input_file = "" # inital

        # self.output_dir_path = tk.StringVar(value="")
        # self.input_file_path = tk.StringVar(value="")

        # GUI elements
        self.create_widgets()


        self.root.mainloop()


    def debug_skip_main_win(self):
        """
        ### NUR ZUM DEBUGGEN!
        Dient der Simulation des Start-Windows inkl. Initialisierung aller Parameter als Dummys, 
        damit das Start-Window (Main-window) zwecks Debugging uebersprungen werden kann.
        """
        self.output_dir = "output_dir/anywhere...DEBUG-DEMO"
        self.input_file = "input_file/anywhere...input_file_DEBUG-DEMO.bas"


        self.show_message_var = True
        self.convert_var = True

        self.optinal_user_defined_text = "optionaler user defind text..."





    def get_file_path_entry(self):
            return self.file_path_entry.get()
                #    self.file_path_entry.get()
    
    
    def get_dir_path_entry(self):
            return self.dir_path_entry.get()


    def create_widgets(self):
        # Label 1
        explanation_text = "Durch das vorliegende Programm kann eine automatisiert ein Code-Dokumentation von einzelnen VBA-Modulen generiert werden.\n\nHierzu muss der Dateipfad zur zu dokumentierenden .bas-Datei angegeben werden, sowie das Verzeichnis, in welches die Dokumentation exportiert werden soll."
        label1 = tk.Label(self.root, text=explanation_text, wraplength=540, justify=tk.LEFT)
        label1.grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=10, pady=5)



        # File Path Entry and Browse Button
        self.file_path_entry = tk.Entry(self.root, width=65 + 4, state="readonly")
        self.file_path_entry.grid(row=1, column=0, padx=10, pady=5, columnspan=2, sticky=tk.W)
        browse_file_button = tk.Button(self.root, text="Suche Input-File...", command=self.browse_file)
        browse_file_button.grid(row=1, column=1, padx=5, sticky=tk.E)

        # Directory Path Entry and Browse Button
        self.dir_path_entry = tk.Entry(self.root, width=65, state="disabled")
        self.dir_path_entry.grid(row=2, column=0, padx=10, pady=5, columnspan=2, sticky=tk.W)
        browse_dir_button = tk.Button(self.root, text="Suche Zielverzeichnis...", command=self.browse_dir)
        browse_dir_button.grid(row=2, column=1, padx=5, sticky=tk.E)

        # Label 10
        explanation_text10 = "Durch das Programm wird IMMER eine Markdown-Datei ('.md') generiert.\n\nOptional kann zusätzlich aus dieser Datei direkt im Anschluss eine HTML-Datei erstellt werden. Durch unterschiedliche Interpretationen im Rahmen dieser Konvertierung können unterschiedliche Tools zu anderen optischen Erscheinungen in der so generierten HTML-Datei führen. Übersichtlicher und sauberer wird die Umwandlung unter Nutzung der VSCode-Extension 'Markdown Preview Enhanced'. Nach der Generierung der Dokumentation sollte daher nochmals manuell eine solche Umwandlung erfolgen."
        label10 = tk.Label(self.root, text=explanation_text10, wraplength=540, justify=tk.LEFT)
        label10.grid(row=3, column=0, columnspan=2, sticky=tk.W, padx=10, pady=5)

        # Convert Checkbox and Show Message Checkbox
        convert_checkbox = tk.Checkbutton(self.root, text="Erstelle zusätzliche HTML-Datei", variable=self.convert_var)
        convert_checkbox.grid(row=4, column=0, sticky=tk.W, padx=10, pady=5, columnspan=2)

        show_message_checkbox = tk.Checkbutton(self.root, text="Im Rahmen der Dokumentation der Call Sequences (Aufrufreihenfolge von Prozeduren):\nFüge einen Endhinweis für jede Prozedur hinzu, wenn kein weiterer Aufruf zu einer (durch dieses Tool dokumentierten) Prozedur erfolgt.", variable=self.show_message_var, wraplength=540, justify=tk.LEFT)
        show_message_checkbox.grid(row=5, column=0, sticky=tk.W, padx=5, pady=5, columnspan=2)

        # Run and Cancel Buttons
        run_button = tk.Button(self.root, text="Dokumentation generieren", command=self.run_process)
        run_button.grid(row=6, column=0, padx=10, pady=10, sticky=tk.W)

        cancel_button = tk.Button(self.root, text="Abbrechen", command=self.cancel_process)
        cancel_button.grid(row=6, column=1, padx=5, pady=10, sticky=tk.E)






    @staticmethod
    def update_entry_text(entry_obj, new_text:str):
        entry_obj.config(state="normal")
        entry_obj.delete(0, tk.END)
        entry_obj.insert(0, new_text)
        entry_obj.config(state="disabled")


    def browse_file(self):
        file_path = filedialog.askopenfilename()

        self.update_entry_text(self.file_path_entry, file_path)

        if not self.output_dir:
            # Setze output wie input dir:
            suggested_output_dir = os.path.dirname(file_path)
            self.update_entry_text(self.dir_path_entry, suggested_output_dir)


        # self.file_path_entry.delete(0, tk.END)
        # self.file_path_entry.insert(0, file_path)


    def browse_dir(self):

        
        self.output_dir = filedialog.askdirectory()

        self.update_entry_text(self.dir_path_entry, self.output_dir)

        # self.dir_path_entry.config(state="normal")
        # self.dir_path_entry.delete(0, tk.END)
        # self.dir_path_entry.insert(0, dir_path)
        # self.dir_path_entry.config(state="disabled")





    def run_process(self):
        show_message = self.show_message_var.get()
        convert_checked = self.convert_var.get()

        output_dir = self.dir_path_entry.get()
        input_file = self.file_path_entry.get()

        if not all([output_dir, input_file]):
             
            # simpledialog.messagebox.showinfo("Fehlerhafte Eingabe", "Bitte über die Buttons die Pfade für die Input-Datei und das Output-Verzeichnis angeben.", icon="error")
            simpledialog.messagebox.showinfo("Fehlerhafte Eingabe", "Bitte über die Buttons die Pfade für die Input-Datei und das Output-Verzeichnis angeben.", icon="error")
            # # ALTERNATIVER SHORTCUT: (alias):
            # simpledialog.messagebox.showerror("Fehlerhafte Eingabe", "Bitte über die Buttons die Pfade für die Input-Datei und das Output-Verzeichnis angeben.")

            return
        




        val = self.msgbox_btn("Die Dokumentation wird erstellt. Bitte warten, bis eine Abschlussmeldung erscheint.")
        # print("ergebnis:", val)


        if val:
            # save current values:
            self.output_dir =  output_dir
            self.input_file = input_file

            self.root.destroy()

            # Aufruf des eigentlichen Code-Documenters!
            print("TODO: Hier: Aufruf des eigentlichen Code-Documenters!")



    def cancel_process(self):

        # simpledialog.messagebox.showinfo("Fehler", "Abbruch.")




        # val = self.msgbox_btn("Mi das Programm wirklich beendet werden?")
        # return val

        self.root.destroy()



    def msgbox_btn(self, message):

        root = tk.Tk()
        root.withdraw()  # Hide the main window
    

        # def on_ok():
        #     root.destroy()
        #     if callback:
        #         callback()


        val = simpledialog.messagebox.askokcancel("Information", message)
        
        root.destroy()

        # print("RETURNED: ", val)

        # if val:
        #     # Beenden des main roots / wins
            # im übergeordneter prozedur!
        #     self.root.destroy()

             

        return val
    










    def show_closing_window(self):
        """
        Erstellt ein neues Fenster, in dem saemtliche Einstellungen nochmals aufgefuehrt werden (read-only).
        Anwendung als Closing-Window nach der Durchfuehrung der eigentlichen Dokumentation.
        """

        self.root = tk.Tk()
        self.root.title(self.TITLE)
        # self.root.geometry(f"{self.WIDTH_MAIN_WINDOW}x{self.HEIGHT_MAIN_WINDOW}")
        self.root.resizable(width=True, height=True)  # Make the window not resizable



        # tx = "Die Dokumentation des Codes wurde generiert."
        # label1 = tk.Label(self.root, text=tx, wraplength=540, justify=tk.LEFT)
        # label1.pack()

        # tx = "Verwendete Parameter:"
        # label1 = tk.Label(self.root, text=tx, wraplength=540, justify=tk.LEFT)
        # label1.pack()
        # # label1.grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=10, pady=5)


        contents = [
            "Die Dokumentation des Codes wurde generiert.",
            "",
            "Verwendete Parameter:",
            "-" * 60,
            "Input-File: {}".format(self.input_file),
            "Output-directory: {}".format(self.output_dir),
            "",
            "{}  :  Erstellung einer HTML-Datei.".format(self.convert_var),
            "{}  :  Hinzufügen eines Endhinweises für jede Prozedur, in der kein weiterer Aufruf erfolgt.".format(self.show_message_var),
            "",
            "Dieses Fenster kann nun geschlossen werden."
        ]


        for _ in contents:
             temp_label = tk.Label(self.root, text=_).pack(anchor="nw")


        self.root.mainloop()






class OLD_OK_DocumenterParameterGui:
    """
    ### TODO: Es sollte für die Anwendung i Code-Documenter noch ein Textfeld geben, in dem beliebiger Freitext uebergeben werden kann - oder ein Input-box (extra win mit allem sichtbaren bereich, also nicht nur eine Zeile... Dieser Freitext soll auch in der DCode-Dokumentation iwo eingebaut werdn oben...)
    """
    # WIDTH = 650
    # WIDTH = 550
    WIDTH = 575
    HEIGHT = 435

    # def __init__(self, root):
        # self.root = root
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Code Documenter")
        self.root.geometry(f"{self.WIDTH}x{self.HEIGHT}")
        self.root.resizable(width=False, height=False)  # Make the window not resizable

        # Initialize variables
        self.convert_var = tk.BooleanVar(value=True)
        self.show_message_var = tk.BooleanVar(value=True)

        self.output_dir = "" # inital
        self.input_file = "" # inital

        # self.output_dir_path = tk.StringVar(value="")
        # self.input_file_path = tk.StringVar(value="")

        # GUI elements
        self.create_widgets()


        self.root.mainloop()



    def get_file_path_entry(self):
            return self.file_path_entry.get()
                #    self.file_path_entry.get()
    
    
    def get_dir_path_entry(self):
            return self.dir_path_entry.get()


    def create_widgets(self):
        # Label 1
        explanation_text = "Durch das vorliegende Programm kann eine automatisiert ein Code-Dokumentation von einzelnen VBA-Modulen generiert werden.\n\nHierzu muss der Dateipfad zur zu dokumentierenden .bas-Datei angegeben werden, sowie das Verzeichnis, in welches die Dokumentation exportiert werden soll."
        label1 = tk.Label(self.root, text=explanation_text, wraplength=540, justify=tk.LEFT)
        label1.grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=10, pady=5)



        # File Path Entry and Browse Button
        self.file_path_entry = tk.Entry(self.root, width=65 + 4, state="readonly")
        self.file_path_entry.grid(row=1, column=0, padx=10, pady=5, columnspan=2, sticky=tk.W)
        browse_file_button = tk.Button(self.root, text="Suche Input-File...", command=self.browse_file)
        browse_file_button.grid(row=1, column=1, padx=5, sticky=tk.E)

        # Directory Path Entry and Browse Button
        self.dir_path_entry = tk.Entry(self.root, width=65, state="disabled")
        self.dir_path_entry.grid(row=2, column=0, padx=10, pady=5, columnspan=2, sticky=tk.W)
        browse_dir_button = tk.Button(self.root, text="Suche Zielverzeichnis...", command=self.browse_dir)
        browse_dir_button.grid(row=2, column=1, padx=5, sticky=tk.E)

        # Label 10
        explanation_text10 = "Durch das Programm wird IMMER eine Markdown-Datei ('.md') generiert.\n\nOptional kann zusätzlich aus dieser Datei direkt im Anschluss eine HTML-Datei erstellt werden. Durch unterschiedliche Interpretationen im Rahmen dieser Konvertierung können unterschiedliche Tools zu anderen optischen Erscheinungen in der so generierten HTML-Datei führen. Übersichtlicher und sauberer wird die Umwandlung unter Nutzung der VSCode-Extension 'Markdown Preview Enhanced'. Nach der Generierung der Dokumentation sollte daher nochmals manuell eine solche Umwandlung erfolgen."
        label10 = tk.Label(self.root, text=explanation_text10, wraplength=540, justify=tk.LEFT)
        label10.grid(row=3, column=0, columnspan=2, sticky=tk.W, padx=10, pady=5)

        # Convert Checkbox and Show Message Checkbox
        convert_checkbox = tk.Checkbutton(self.root, text="Erstelle zusätzliche HTML-Datei", variable=self.convert_var)
        convert_checkbox.grid(row=4, column=0, sticky=tk.W, padx=10, pady=5, columnspan=2)

        show_message_checkbox = tk.Checkbutton(self.root, text="Im Rahmen der Dokumentation der Call Sequences (Aufrufreihenfolge von Prozeduren):\nFüge einen Endhinweis für jede Prozedur hinzu, wenn kein weiterer Aufruf zu einer (durch dieses Tool dokumentierten) Prozedur erfolgt.", variable=self.show_message_var, wraplength=540, justify=tk.LEFT)
        show_message_checkbox.grid(row=5, column=0, sticky=tk.W, padx=5, pady=5, columnspan=2)

        # Run and Cancel Buttons
        run_button = tk.Button(self.root, text="Dokumentation generieren", command=self.run_process)
        run_button.grid(row=6, column=0, padx=10, pady=10, sticky=tk.W)

        cancel_button = tk.Button(self.root, text="Abbrechen", command=self.cancel_process)
        cancel_button.grid(row=6, column=1, padx=5, pady=10, sticky=tk.E)






    @staticmethod
    def update_entry_text(entry_obj, new_text:str):
        entry_obj.config(state="normal")
        entry_obj.delete(0, tk.END)
        entry_obj.insert(0, new_text)
        entry_obj.config(state="disabled")


    def browse_file(self):
        file_path = filedialog.askopenfilename()

        self.update_entry_text(self.file_path_entry, file_path)

        if not self.output_dir:
            # Setze output wie input dir:
            suggested_output_dir = os.path.dirname(file_path)
            self.update_entry_text(self.dir_path_entry, suggested_output_dir)


        # self.file_path_entry.delete(0, tk.END)
        # self.file_path_entry.insert(0, file_path)


    def browse_dir(self):

        
        self.output_dir = filedialog.askdirectory()

        self.update_entry_text(self.dir_path_entry, self.output_dir)

        # self.dir_path_entry.config(state="normal")
        # self.dir_path_entry.delete(0, tk.END)
        # self.dir_path_entry.insert(0, dir_path)
        # self.dir_path_entry.config(state="disabled")





    def run_process(self):
        show_message = self.show_message_var.get()
        convert_checked = self.convert_var.get()

        output_dir = self.dir_path_entry.get()
        input_file = self.file_path_entry.get()

        if not all([output_dir, input_file]):
             
            # simpledialog.messagebox.showinfo("Fehlerhafte Eingabe", "Bitte über die Buttons die Pfade für die Input-Datei und das Output-Verzeichnis angeben.", icon="error")
            simpledialog.messagebox.showinfo("Fehlerhafte Eingabe", "Bitte über die Buttons die Pfade für die Input-Datei und das Output-Verzeichnis angeben.", icon="error")
            # # ALTERNATIVER SHORTCUT: (alias):
            # simpledialog.messagebox.showerror("Fehlerhafte Eingabe", "Bitte über die Buttons die Pfade für die Input-Datei und das Output-Verzeichnis angeben.")

            return
        




        val = self.msgbox_btn("Die Dokumentation wird erstellt. Bitte warten, bis eine Abschlussmeldung erscheint.")
        # print("ergebnis:", val)


        if val:
            # save current values:
            self.output_dir =  output_dir
            self.input_file = input_file

            self.root.destroy()

            # Aufruf des eigentlichen Code-Documenters!
            print("TODO: Hier: Aufruf des eigentlichen Code-Documenters!")



    def cancel_process(self):

        # simpledialog.messagebox.showinfo("Fehler", "Abbruch.")




        # val = self.msgbox_btn("Mi das Programm wirklich beendet werden?")
        # return val

        self.root.destroy()



    def msgbox_btn(self, message):

        root = tk.Tk()
        root.withdraw()  # Hide the main window
    

        # def on_ok():
        #     root.destroy()
        #     if callback:
        #         callback()


        val = simpledialog.messagebox.askokcancel("Information", message)
        
        root.destroy()

        # print("RETURNED: ", val)

        # if val:
        #     # Beenden des main roots / wins
            # im übergeordneter prozedur!
        #     self.root.destroy()

             

        return val





if __name__ == "__main__":




    # gui = DocumenterParameterGui()
    gui = DocumenterParameterGui(DEBUG=1)





    print("nachher")
    print(gui.output_dir)


    # Erstellt ein neues Window zur abschließennden Zusammenfassung der durchgeführten Tätigkeiten und Parameter:
    gui.show_closing_window()