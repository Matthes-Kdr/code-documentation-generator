



def use_pack():
    import tkinter as tk
    from tkinter import filedialog

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
        print("run!")
        
        if show_message:
            result_label.config(text="Process completed!")

    def cancel_process():
        # Your logic to handle canceling the process goes here
        print("cancel!")

        result_label.config(text="Process canceled.")









    # GUI Setup
    root = tk.Tk()
    root.title("Code Documenter")

    # Convert Checkbox
    convert_var = tk.BooleanVar()
    convert_checkbox = tk.Checkbutton(root, text="Convert", variable=convert_var)
    convert_checkbox.pack()

    # File Path Entry and Browse Button
    file_path_entry = tk.Entry(root, width=40)
    file_path_entry.pack()
    browse_file_button = tk.Button(root, text="Browse Input-File", command=browse_file)
    browse_file_button.pack()

    # Directory Path Entry and Browse Button
    dir_path_entry = tk.Entry(root, width=40)
    dir_path_entry.pack()
    browse_dir_button = tk.Button(root, text="Browse Directory", command=browse_dir)
    browse_dir_button.pack()

    # Show Message Checkbox
    show_message_var = tk.BooleanVar()
    show_message_checkbox = tk.Checkbutton(root, text="Show End Message", variable=show_message_var)
    show_message_checkbox.pack()

    # Run and Cancel Buttons
    run_button = tk.Button(root, text="Run", command=run_process)
    run_button.pack()
    cancel_button = tk.Button(root, text="Cancel", command=cancel_process)
    cancel_button.pack()

    # Result Label
    result_label = tk.Label(root, text="")
    result_label.pack()

    root.mainloop()












def use_grid1():
        
    import tkinter as tk
    from tkinter import filedialog

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
            result_label.config(text="Process completed!")

    def cancel_process():
        # Your logic to handle canceling the process goes here
        result_label.config(text="Process canceled.")

    # GUI Setup
    root = tk.Tk()
    root.title("Simple GUI")

    # Convert Checkbox
    convert_var = tk.BooleanVar()
    convert_checkbox = tk.Checkbutton(root, text="Convert", variable=convert_var)
    convert_checkbox.grid(row=0, column=0, columnspan=2, sticky=tk.W)

    # File Path Entry and Browse Button
    file_path_entry = tk.Entry(root, width=40)
    file_path_entry.grid(row=1, column=0, columnspan=2, padx=10, pady=5)
    browse_file_button = tk.Button(root, text="Browse File", command=browse_file)
    browse_file_button.grid(row=1, column=2, padx=5)

    # Directory Path Entry and Browse Button
    dir_path_entry = tk.Entry(root, width=40)
    dir_path_entry.grid(row=2, column=0, columnspan=2, padx=10, pady=5)
    browse_dir_button = tk.Button(root, text="Browse Directory", command=browse_dir)
    browse_dir_button.grid(row=2, column=2, padx=5)

    # Show Message Checkbox
    show_message_var = tk.BooleanVar()
    show_message_checkbox = tk.Checkbutton(root, text="Show End Message", variable=show_message_var)
    show_message_checkbox.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=5)

    # Run and Cancel Buttons
    run_button = tk.Button(root, text="Run", command=run_process)
    run_button.grid(row=4, column=0, pady=10)
    cancel_button = tk.Button(root, text="Cancel", command=cancel_process)
    cancel_button.grid(row=4, column=1, pady=10)

    # Result Label
    result_label = tk.Label(root, text="")
    result_label.grid(row=5, column=0, columnspan=3, pady=10)

    root.mainloop()












def use_grid2():

    import tkinter as tk
    from tkinter import filedialog

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
            result_label.config(text="Process completed!")

    def cancel_process():
        # Your logic to handle canceling the process goes here
        result_label.config(text="Process canceled.")


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





    root.mainloop()


# use_pack()
# use_grid1()
use_grid2()







