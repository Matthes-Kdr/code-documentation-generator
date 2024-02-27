'''
Provides a GUI for setting parameters for the code documentation to be carried out. All inputs are queried by this module and - if necessary - also saved within individual objects.

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
    ### OUTLOOK: There should still be a text field for the application i Code-Documenter, in which any free text can be passed - or an input box (extra win with all visible area, so not just one line...). This free text should also be included somewhere in the DCode documentation above...)


    ### TODOC : ...


    """


    WIDTH_MAIN_WINDOW = 575
    HEIGHT_MAIN_WINDOW = 435


    TITLE = "Code Documenter"


    __DEBUG = False # should be set from outside the class in the main script as otherwise there will be multiple definitions! (fixed already...)

    @classmethod
    def set_debug_mode(cls, value:bool):
        """
        Sets the debug mode for this class.

        Args:
            value (bool|int): 0 or 1 or boolean values therefore.
        """
        cls.__DEBUG = bool(value)
        if cls.__DEBUG:
            print("!!!  Use the debug mode of class 'DocumenterGui'  !!!")




    def __init__(self):
        """
        Initializes a main window for parameterizing the settings for the automated documentation to be carried out.
        The layout of the GUI window is determined by a separate method that is called here.
        The main loop is then started directly.
        """

        self.__is_ready = False # Flag
        
        
        if self.__DEBUG:
             
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
        Returns the value that serves as a flag as to whether all parameters have been set by the GUI and whether the actual code documentation should be started from outside the class.

        Returns:
            bool: True if can/should be started, False if abort or at least no start should take place.

        """
        return self.__is_ready
    



    def debug_skip_main_win(self):
        """
        ### FOR DEBUGGING ONLY!

        Used to simulate the start window including initialization of all parameters as dummies, 
        so that the Start-Window (Main-window) can be skipped for debugging purposes.

        If the constant DEBUG from the modul 'code_documenter' is True, 
        then the self.input_file will be set here without the GUI and without any further input of the user.
        """
        
        # Just a few paths to examples - you could use any .bas file with vba code.
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
        Definition of the layout of the main window to parameterize the settings for the documentation to be created.
        """
        
        # Label:
        temp_lbl_text = "This program will automatically generate a code documentation for individual VBA modules.\n\nTo do so, the file path to the '.bas' file to be documented must be specified, as well as the directory to which the documentation should be exported. The code must be saved as .bas files as direct accessing to Excel '.xlsm' files is not supported by this script."
        lbl = tk.Label(self.root, text=temp_lbl_text, wraplength=540, justify=tk.LEFT)
        lbl.grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=10, pady=5)



        # INput File Path Entry and Browse Button:
        self.tx_input_file_path = tk.Entry(self.root, width=65 + 4, state="readonly")
        self.tx_input_file_path.grid(row=1, column=0, padx=10, pady=5, columnspan=2, sticky=tk.W)

        btn_browse_file = tk.Button(self.root, text="Browse Input-File...", command=self.browse_file)
        btn_browse_file.grid(row=1, column=1, padx=5, sticky=tk.E)


        # Output Directory Path Entry and Browse Button
        self.txt_output_dir_path = tk.Entry(self.root, width=65, state="disabled")
        self.txt_output_dir_path.grid(row=2, column=0, padx=10, pady=5, columnspan=2, sticky=tk.W)

        btn_browse_dir = tk.Button(self.root, text="Browse output directory...", command=self.browse_dir)
        btn_browse_dir.grid(row=2, column=1, padx=5, sticky=tk.E)


        # Label:
        temp_lbl_text = "The program ALWAYS generates a Markdown file ('.md').\n\nOptionally, an HTML file can also be created from this file directly afterwards. Due to different interpretations in the context of this conversion, different tools can lead to different visual appearances in the HTML file generated in this way. The conversion is clearer and cleaner when using the VSCode extension 'Markdown Preview Enhanced'. Once the documentation has been generated, it should therefore be converted again manually."
        lbl = tk.Label(self.root, text=temp_lbl_text, wraplength=540, justify=tk.LEFT)
        lbl.grid(row=3, column=0, columnspan=2, sticky=tk.W, padx=10, pady=5)

        # Convert Checkbox and Show Message Checkbox
        temp_cb_text = "Create additional HTML file"
        cb_convert_to_html = tk.Checkbutton(self.root, text=temp_cb_text, variable=self.convert_to_html_var)
        cb_convert_to_html.grid(row=4, column=0, sticky=tk.W, padx=10, pady=5, columnspan=2)

        temp_cb_text = "Regarding the documentation of the call sequences (call sequence of procedures):\nAdd an end note for each procedure if no further call is made to a procedure (documented by this tool)."
        cb_show_message = tk.Checkbutton(self.root, text=temp_cb_text, variable=self.show_message_var, wraplength=540, justify=tk.LEFT)
        cb_show_message.grid(row=5, column=0, sticky=tk.W, padx=5, pady=5, columnspan=2)


        # Run and Cancel Buttons
        btn_start = tk.Button(self.root, text="Generate documentation", command=self.start_process)
        btn_start.grid(row=6, column=0, padx=10, pady=10, sticky=tk.W)

        btn_cancel = tk.Button(self.root, text="Cancel", command=self.cancel_process)
        btn_cancel.grid(row=6, column=1, padx=10, pady=10, sticky=tk.E)






    @staticmethod
    def update_entry_text(entry_obj, new_text:str):
        """
        method modifies the text within a tk.entry object and writes the transferred text to the entry.
        This allows the text to be updated depending on a previous selection, for example.
        After the modification, the object is disabled again to prevent manual error-prone corrections.

        Args:
            entry_obj (tk.Entry): Entry object which is to be provided with new text
            new_text (str): Text to be displayed in the entry object.
        """
        entry_obj.config(state="normal")
        entry_obj.delete(0, tk.END)
        entry_obj.insert(0, new_text)
        entry_obj.config(state="disabled")



    def browse_file(self):
        """
        Provides a file dialog for selecting specific file types. 
        Once the user has successfully selected a file, a MEthode is called,
        which adjusts the entry objects accordingly so that the selected text is displayed there.

        Also checks whether an output folder has already been specified. 
        If not, the directory in which the input file is located is preset as the default.
        """
        
        # Restrictions on file types:
        extensions = [("VB", ".bas")]
        default_extension = extensions[0]
        
        # Selection by user:
        file_path = filedialog.askopenfilename(defaultextension=default_extension, filetypes=extensions)

        self.update_entry_text(self.tx_input_file_path, file_path)

        # If another output folder has not already been selected, use this one from the input file:
        if not self.output_dir:
            # Set output as input dir:
            suggested_output_dir = os.path.dirname(file_path)
            self.update_entry_text(self.txt_output_dir_path, suggested_output_dir)




    def browse_dir(self):
        """
        Provides a file dialog for selecting a directory. 
        Once the user has successfully selected a file, a MEthode is called,
        which adjusts the entry objects accordingly so that the selected text is displayed there.
        """
        self.output_dir = filedialog.askdirectory()
        self.update_entry_text(self.txt_output_dir_path, self.output_dir)





    def start_process(self):
        """
        Method that is called when the 'Start' button is clicked.
        All relevant variables from the GUI are first saved in object variables.
        After a successful validity check, a message box with an info message is displayed.
        The main window is then destroyed.
        The GUI is therefore initially no longer active on the screen. 
        However, the variables can still be accessed from outside.
        """
        self.show_message = self.show_message_var.get()
        self.convert_checked = self.convert_to_html_var.get()

        self.output_dir = self.txt_output_dir_path.get()
        self.input_file = self.tx_input_file_path.get()

        if not all([self.output_dir, self.input_file]):
            # If not all relevant fields are filled in:
            
            simpledialog.messagebox.showinfo("Invalid input!", "Please use the buttons to specify the paths for the input file and the output directory.", icon="error")
            # # ALTERNATIVER SHORTCUT: (alias): (not translated...)
            # simpledialog.messagebox.showerror("Fehlerhafte Eingabe", "Bitte über die Buttons die Pfade für die Input-Datei und das Output-Verzeichnis angeben.")

            return


        # Note for the user in Msgbox with subsequent return value query:
        val = self.msgbox_btn("The documentation is being created. Please wait until a completion message appears.")

        if val:
            # close window:
            self.root.destroy()
            self.__is_ready = True # set Flag for accessing from outside the class






    def cancel_process(self):
        """
        Close / hide the main window.
        """

        self.root.destroy()



    def msgbox_btn(self, message):
        """
        Simple message box with OK and Cancel buttons

        Args:
            message (str): Text to be displayed.

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
        Creates a new window in which all settings are listed again (read-only).
        Used as a closing window after the actual documentation has been executed.
        """

        self.root = tk.Tk()
        self.root.title(self.TITLE)
        self.root.resizable(width=True, height=True)  # Make the window not resizable


        # Contents to be listed among each other: 
        contents = [
            "The documentation of the code has been generated.",
            "",
            "Parameters used:",
            "-" * 60,
            "Input-File: {}".format(self.input_file),
            "Output-directory: {}".format(self.output_dir),
            "",
            "{} : Creation of an HTML file.".format(self.convert_checked),
            "{} : Add an end hint for each procedure in which no further call is made.".format(self.show_message),
            "",
            "This window can now be closed."
        ]


        for _ in contents:
             temp_label = tk.Label(self.root, text=_).pack(anchor="nw")


        self.root.mainloop()










def demo_ablauf():
    """
    Shows the possible use of the DocumenterGui class so that it is only used PASSIVELY in the overall process and does not have to actively call anything from outside the class.
    """

    gui = DocumenterGui()
    
    
    if gui.get_is_ready():
        print("Then start documentation from outside with the parameters by accessing the object variables")
    else:
        print("NO  Start!")


    print("Finished.")




if __name__ == "__main__":


    DocumenterGui.__DEBUG = 1

    demo_ablauf()


