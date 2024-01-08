



import tkinter as tk
from tkinter import filedialog


from tkinter import simpledialog





def msgbox_btn(message:str):


    def show_message_box(message:str):
        root = tk.Tk()
        root.withdraw()  # Hide the main window


        # result = simpledialog.messagebox.info("Information", message)
        val = simpledialog.messagebox.askokcancel("Information", message)

        return val

    # Trigger the message box
    val = show_message_box(message=message)
    return val





def show_message_box(message:str):

    top = tk.Toplevel()
    top.title("Information")

    label = tk.Label(top, text=message, padx=20, pady=20)
    label.pack()

    # Schedule the closing of the message box after 3000 milliseconds (3 seconds)
    top.after(3000, top.destroy)

# Trigger the message box
# show_message_box()



def browse_file():
    file_path = filedialog.askopenfilename()
    file_path_entry.delete(0, tk.END)
    file_path_entry.insert(0, file_path)

def browse_dir():
    dir_path = filedialog.askdirectory()
    dir_path_entry.delete(0, tk.END)
    dir_path_entry.insert(0, dir_path)





def run_process():
    show_message = show_message_var.get()
    convert_checked = convert_var.get()
    
    # Your logic to run the process goes here
    
    if show_message:
        # simpledialog.messagebox.showinfo("Information", "die Endhinweis sollen gesetzt werden. in 3 sec wird root.destroy !")
        val = msgbox_btn("Die Dokumentation wird erstellt. Bitte warten, bis eine Abschlussmeldung erscheint.")

        print("ergebnis:", val)


        if val:
            # close gui:
            root.destroy()

    # root.after(3000, root.destroy)


def cancel_process():
    # Your logic to handle canceling the process goes here
    simpledialog.messagebox.showinfo("Fehler", "abbruch.")


# GUI Setup
    
WIDTH = 550
HEIGHT = 400
root = tk.Tk()
root.title("Code Documenter")
root.geometry(f"{WIDTH}x{HEIGHT}")  # Set the initial size of the window
# root.resizable(width=False, height=False)  # Make the window not resizable


# Label 1
explanation_text = "Durch das vorliegende Programm kann eine automatisiert ein Code-Dokumentation von einzelnen VBA-Modulen generiert werden.\n\nHierzu muss der Dateipfad zur zu dokumentierenden .bas-Datei angegeben werden, sowie das Verzeichnis, in welches die Dokumentation exportiert werden soll."
label1 = tk.Label(root, text=explanation_text, wraplength=WIDTH-10, justify=tk.LEFT)
label1.grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=10, pady=5)

# File Path Entry and Browse Button
file_path_entry = tk.Entry(root, width=65+4)
file_path_entry.grid(row=1, column=0,  padx=10, pady=5, columnspan=2, sticky=tk.W)
browse_file_button = tk.Button(root, text="Suche Input-File...", command=browse_file)
browse_file_button.grid(row=1, column=1, padx=5, sticky=tk.E)

# Directory Path Entry and Browse Button
dir_path_entry = tk.Entry(root, width=65)
# dir_path_entry.grid(row=2, column=0, padx=10, pady=5)
dir_path_entry.grid(row=2, column=0,  padx=10, pady=5, columnspan=2, sticky=tk.W)
browse_dir_button = tk.Button(root, text="Suche Zielverzeichnis...", command=browse_dir)
browse_dir_button.grid(row=2, column=1, padx=5,  sticky=tk.E)



# Label 10
explanation_text10 = "Durch das Programm wird eine Markdown-Datei ('.md') generiert. Optional wird aus dieser Datei direkt im Anschluss eine HTML-Datei erstellt. Durch unterschiedliche Interpretationen im Rahmen dieser Konvertierung können unterschiedliche Tools zu anderen optischen Erscheinungen in der so generierten HTML-Datei führen. Übersichtlicher und sauberer wird die Umwandlung unter Nutzung der VSCode-Extension 'Markdown Preview Enhanced'. Nach der Generierung der Dokumentation sollte daher nochmals manuell eine solche Umwandlung erfolgen."
label10 = tk.Label(root, text=explanation_text10, wraplength=WIDTH-10, justify=tk.LEFT)
label10.grid(row=3, column=0, columnspan=2, sticky=tk.W, padx=10, pady=5)





# Convert Checkbox and Show Message Checkbox
convert_var = tk.BooleanVar(value=1)
convert_checkbox = tk.Checkbutton(root, text="Erstelle zusätzliche HTML-Datei", variable=convert_var)
convert_checkbox.grid(row=4, column=0, sticky=tk.W, padx=10, pady=5, columnspan=2)



show_message_var = tk.BooleanVar(value=1)
show_message_checkbox = tk.Checkbutton(root, text="Im Rahmen der Dokumentation der Call Sequences (Aufrufreihenfolge von Prozeduren):\nFüge einen Endhinweis für jede Prozedur hinzu, wenn kein weiterer Aufruf zu einer (durch dieses Tool dokumentierten) Prozedur erfolgt.", variable=show_message_var, wraplength=WIDTH-10-10, justify=tk.LEFT)
show_message_checkbox.grid(row=5, column=0, sticky=tk.W, padx=5, pady=5, columnspan=2)

# # # Label 2
# # label2 = tk.Label(root, text="Bla:")
# # label2.grid(row=5, column=0, sticky=tk.W, padx=10, pady=5)

# Run and Cancel Buttons
run_button = tk.Button(root, text="Dokumentation generieren", command=run_process)
run_button.grid(row=6, column=0, padx=10, pady=10, sticky=tk.W)

cancel_button = tk.Button(root, text="Abbrechen", command=cancel_process)
cancel_button.grid(row=6, column=1, padx=5, pady=10, sticky=tk.E)

# # Result Label
# result_label = tk.Label(root, text="")
# result_label.grid(row=7, column=0, columnspan=2, pady=10)




print("vor mainloop")
root.mainloop()
print("nach mainloop")







# use_grid2()

# msgbox()
