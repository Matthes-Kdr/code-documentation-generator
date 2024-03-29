﻿# -*- coding: utf-8 -*-
'''
Created on: Fri, 2023-12-29 (00:45:39)


@author: Matthias Kader


> For further information, see README.md
----------------------------------------
>  Tool to automatically generate a documentation of the source code - mainly used for function-based flows (currently only supporting single VBA-moduls, support for Python and C++ shall follow).

> **Unlinke many other generic documentation tools this project focus on the documentation of code with a functional-programming approach rather than a class-based approach.** Aim is to **show the program schedule (flow / calls of other procedures)** within the procedures.

**<span style='color:red'>&#9888; As this project was not meant to become a public one in the first place, I did not chore and tidy the code by now! Also I began in German, but I will fix those both issues step by step and while choring I will translate everything in the code :-)</span>**

> **NOTE:** As I started this project privatly and was not very forward-looking a lot of the text (readme + code comments) is written in **German**. Step by step I will change this and translate everything to have a continuous language - sorry for that.

> **<span style='color:red'>&#9888; If there is anything unclear or poor translated - feel free to correct it or to ask (e.g. via a new issue?) for the meaning...</span>**



# =============================================================================
#### Notes on application and use: ####
# =============================================================================

- To generate a docstring from the VBA-Source make sure that the text to shown is located directly below the declaration line of the procedure. The text is considered completed with the first following line in the code which is not an entire comment line. Empty lines that are to be included must also be labelled as comments.

- The script generates an MD file (Markdown), which is then immediately converted to HTML using the markdown library, so that 2 files are created after the script is completed. However, due to different interpretations during the conversion, the display of the HTML file generated in this way differs if it is converted separately via VSCode Extension ("Markdown all in one"). The file generated via VSCode is clearer and more attractive. This should therefore be done separately at the end.





### Implemented ready:

- Generating Table of contents / index
- Overall layout incl. title, subheadings for individual sections
- Listing the module-wide program header docstring in the generated documentation
- Listing the reference searches (Where is the procedure called?) in the generated documentation
- Function to immediate export of the MD file to an HTML file within this script
- Listing the organizational data regarding the code to be documented and the script used for documentation in the generated documentation
- Listing of the calling sequence (calling sequence / calling levels) within each procedure in the generated documentation: Enumeration of the calls of other procedures covered in this documentation. Including recursive nested list of which calls are made in each of the called procedures.
- Providing of a simple GUI / HMI to parameterize input and output paths


> For the latest list of issues / TODOS see Github Issues at https://github.com/Matthes-Kdr/code-documentation-generator/issues








# =============================================================================
#### Important call sequence of the method within this Python script 
for creating the documentation of the call sequence of the VBA procedures to be documented:
# =============================================================================

First, all procedures will be completely analyzed, and only after then all procedures will be completely documented. For both procedures, this takes place in a method at object level, whereby this respective method is called in both cases from a class method in which iteration takes place via the individual procedure objects within this class:

- analyse_call_sequence(cls)
    - analyse_calling_sequence_in_one_proc(self)
- prepare_all_call_sequences_docs(cls)
    - prepare_single_call_sequence_docs(cls)

(Side-note: by the way, the developed tool would have been of great usage to document this in 'nice', if it already had the ability to document Python syntax :-) )












# =============================================================================
#### Unimportant trivialities: Code analysis summary just for fun and for myself: ####
# =============================================================================

version from 2024-01-11 - 00:18:43:
    Details in each case: [count of lines @ code_documenter.py] + [count of lines @ gui] = [Summe]
    - Total of lines: 408 (100%)+2201 (100%)=2609 (100%)
    - thereof empty lines: 168 (41,1764705882353%)+1071 (48,6597001363017%)=1239 (47,489459563051%)
    - thereof single line comments: 20 (4,90196078431373%)+244 (11,0858700590641%)=264 (10,1188194710617%)
    - thereof multiline-comments: 85 (20,8333333333333%)+374 (16,9922762380736%)=459 (17,5929474894596%)

    ==> sum of all comment-lines:: 105 (25,7352941176471%)+618 (28,0781462971377%)=723 (27,7117669605213%)
    ==> actual relevant lines with Code commands: 135 (33,0882352941176%)+512 (23,2621535665607%)=647 (24,7987734764277%)

-----------------------------------------------    

version from 2024-01-07 - 23:37:04:
    - Total of lines: 2164 (100%)
    - thereof empty lines: 1091 (50%)
    - thereof single line comments: 226 (10%)
    - thereof multiline-comments: 364 (17%)

    ==> sum of all comment-lines: 590 (27%)
    ==> actual relevant lines with Code commands: 483 (22%)

-----------------------------------------------

version from 2024-01-07 - 15:26:03:
    - Total of lines: 2771 (100%)
    - thereof empty lines: 1408 (51%)
    - thereof single line comments: 278 (10%)
    - thereof multiline-comments: 550 (20%)

    ==> sum of all comment-lines: 828 (30%)
    ==> actual relevant lines with Code commands: 535 (19%)

-----------------------------------------------



'''















from datetime import datetime
import inspect
import os
import re

import markdown

import subprocess
import gitinfo


from gui import DocumenterGui
from programming_languages import SyntaxVba



from sys import argv




# =============================================================================
#### GLOBALS: ####
# =============================================================================

DEBUG = 0 # Can be changed by passing another argument while starting script from terminal.


# DocumenterGui.set_debug_mode(DEBUG)

# HACK: only Vba by now:
SyntaxClass = SyntaxVba


# CONVERT_TO_HTML = 1



def db(*args):
    """
    Loops through for printing - only for debugging
    # TODO: Replace this method with a logger of the python logging module, see #18 @ Github.
    """
    if DEBUG == False:
        return
    
    print("__DEBUG_PRINT__\n")
    for _ in args:
        print(_)






def inspect_get_current_line_number():
    """
    Returns the line number of the code - suitable for debugging!

    Returns:
        int: number of line  in code file.
    """

    # calling Stack-Informationen:
    stack = inspect.stack()
    
    # get the information for the current function/frame
    current_frame = stack[1]
    
    # extract line number:
    line_number = current_frame[2]
    

    return line_number






# =============================================================================
#### Workaround: Use of MetaClasses: 
# Meta classes for direct calls / implicit calls of a class method 
# directly after implementing a class ####
# =============================================================================


# =============================================================================
#### # WIEDERHOLFUNKTION-KANDIDAT!!! ####
# =============================================================================
class AutoCallMeta(type):
    """
    ### OBSOLETE: Since the script provides an GUI this is not  used anymore.

    If this class is used as a metaclass for another class, the class method listed below is automatically called directly after definition without an additional, explicit call.

    To do this, only the name of the method to be called must be parameterized in the class variable 'class_name_to_call_implicit'.
    """    

    # However, it is more flexible to actually call the MetaClass.initialize_class method explicitly at a point where all relevant attributes have already been assigned values. 
    # Therefore, with the implementation of the GUI, the process was changed accordingly and this superclass was no longer inherited by the MetaClass class


    class_name_to_call_implicit = "initialize_class" # NUR HIER ZU PARAMETRISIEREN!
    
    def __init__(cls, name, bases, attrs):
        super().__init__(name, bases, attrs)
        func_name  = AutoCallMeta.class_name_to_call_implicit

        # Aufruf der relevanten Methode, sofern sie vorhanden ist:
        if hasattr(cls, func_name):
            func = getattr(cls, func_name)
            # Aufruf:
            func()














# =============================================================================
#### CLASSES: ####
# =============================================================================






class MetaData():
# class MetaData(metaclass=AutoCallMeta):
    """
    
    ### CHANGELOG 2024-01-10 - 19:45:14:
    This superclass has ensured that its MetaClass.initialize_class method is automatically called implicitly directly after the declaration of the MetaData class. However, it is more flexible to actually call it explicitly at a point where all relevant attributes have already been assigned values. 
    Therefore, with the implementation of the GUI, the process was changed accordingly and this superclass was no longer inherited by the MetaClass class.



    This class mainly stores data that can later be regarded as a type of metadata.
    The parent class / metaclass ensures that a method parameterized in the metaclass is called directly after the implementation of this (ordinary) class.
    This means that it is not necessary to explicitly call the MetaData.initialize_class class method, as this is done via the metaclass.

    # TODO: The class is not yet finished.

    The data stored in it includes e.g:

    - This documentation tool script
        - Version information, based on the last Git commit
        - TODO: Version number of this script...
    - The module to be documented
        - File path
        - File name
        - Date of the last save / change
    - Current timestamp to indicate the time of the documentation

      
    """


    __input_path:str = None
    __output_dir:str = None

    __convert_to_html = False

    __user_defined_additional_text = ""


    @classmethod
    def set_user_defined_additional_text(cls, val:str):
        cls.__user_defined_additional_text = val
    

    @classmethod
    def get_user_defined_additional_text(cls) -> str:
        return cls.__user_defined_additional_text
    


    @classmethod
    def set_convert_to_html(cls, val:bool):
        cls.__convert_to_html = val
    

    @classmethod
    def get_convert_to_html(cls) -> bool:
        return cls.__convert_to_html
    

    @classmethod
    def get_input_path(cls) -> str:
        return cls.__input_path
    
    
    @classmethod
    def get_input_filename(cls) -> str:
        return cls.__input_filename
    
    
    
    @classmethod
    def get_output_filename(cls) -> str:
        return cls.__output_filename
    
    

    @staticmethod
    def get_last_modified_timestamp(file_path) -> str:
        # Read the timestamp of the last change to the file
        timestamp = os.path.getmtime(file_path)
        
        # Convert the timestamp into a readable date
        last_modified_datetime = datetime.fromtimestamp(timestamp)
        
        # convert the date in the desired format
        formatted_date = last_modified_datetime.strftime('%Y-%m-%d %H:%M')
        
        return formatted_date





    @classmethod
    def extract_date_of_change(cls):
        """
        Reads the last change date of the input file and saves it in the class variable cls.input_file_date_of_change
        """

        date_of_change = cls.get_last_modified_timestamp(cls.__input_path)
        cls.input_file__date_of_change = date_of_change






    @classmethod
    def set_input_path(cls, input_path:str=None):
        """
        Checks the optionally transferred path to see whether it exists and whether there is a .bas file there.
        If this is not the case, or if no path is passed, a new path is requested via input. 
        The method is called recursively until a valid path is found.
        So far there is no possibility for the user to cancel the input (except program abort...)

        Args:
            input_path (str, optional): File path to the .bas file to be documented - as a forward slash and without quotation marks. Defaults to None.

        """

        if input_path != None:
            if os.path.isfile(input_path):
                if input_path.endswith(".bas"):
                    cls.__input_path = input_path
                    cls.extract_date_of_change()
                    return
                
        new_input = input("!!! ERROR !!! The specified file is not a .bas file! Please enter a valid path to the corresponding file (forward slashes! without quotation marks)\n> Your input: ")

        cls.set_input_path(new_input)


    @classmethod
    def set_output_dir(cls, output_dir:str=None):
        """
        Checks the optionally transferred path to see whether it exists and whether this is a directory
        If this is not the case, or if no path is passed, a new path is requested via input. 
        The method is called recursively until a valid path is found.
        So far there is no possibility for the user to cancel the input (except program abort...)

        Args:
            output_dir (str, optional): File path to the folder in which the generated output files are to be exported - as a forward slash and without quotation marks. Defaults to None.


        """

        if output_dir != None:
            if os.path.isdir(output_dir):
                cls.__output_dir = output_dir
                return
        new_output_dir = input("!!! ERROR !!! The specified path is not a valid directory. Please enter a valid path for the export of the output files (forward slashes! without quotation marks)\n> your input: ")

        cls.set_output_dir(new_output_dir)




    @classmethod
    def extract_git_info(cls):
        info:dict = gitinfo.get_git_info()

        # Transfer the information from git to the class:
        for key, value in info.items():


            attr_name =  f"documenter_version__{key}"

            setattr(cls, attr_name, value)





    @classmethod
    def save_current_timestamp(cls):

        current_datetime = datetime.now()

        # convert into a timestamp:
        current_timestamp = int(current_datetime.timestamp())

        db(f"Current Timestamp: {current_timestamp}")

        current_datetime.strftime("%Y-%m-%d %H:%M:%S")

        cls.date_of_process = current_datetime.strftime("%Y-%m-%d %H:%M:%S")







    @classmethod
    def make_output_filename(cls):
        
        cls.__input_filename = os.path.basename(cls.__input_path)
        cls.__output_filename = cls.__input_filename + " - Dokumentation"




    @classmethod
    def get_output_path(cls, extension=".md") -> str:
        """
        Returns the entire path for the new output file to be generated, including the file extension.

        Args:
            extension (str, optional): File extension of the output file. Defaults to ".md". Modifiable e.g. to .html or .txt

        Returns:
            str : File path
        """

        path =  os.path.join(cls.__output_dir, cls.__output_filename + extension)
        return path
    
    




    @classmethod
    def initialize_class(cls):
        """
        This method is called implicitly directly after the implementation of this class.
        It therefore no longer needs to be called from outside.
        It initializes all attributes with their values.
        """




        # =============================================================================
        #### DEBUGGING-PARAMETER SECTION: ####
        # =============================================================================
        
        # # HACK: path for Source-vba-code
        # input_file_path = "input_data/beispiel_modul_rekursiv.bas"
        # input_file_path = "input_data/beispiel_modul_bauer+liebherr.bas"
        # input_file_path = "input_data/beispiel_modul2.bas"
        # input_file_path = "input_data/beispiel_modul1.bas"
        # input_file_path = "input_data/beispiel_modul.bas"

        # # HACK: If you are working without GUI data (debugging:)
        # # Is normally called from outside and filled with data from the GUI. Only run if it is for debugging!
        # cls.set_input_path(input_file_path)

        # output_dir = "output_data"
        # cls.set_output_dir(output_dir)



        # cls.git_info_to_str()
        cls.extract_git_info()
        # cls.count_of_commits = cls.get_count_of_commits()


        cls.save_current_timestamp()


        cls.make_output_filename()








class Procedure():
    """
    General class for providing content that is required for all procedures (subs and functions). This includes
        - Definition of the file path for the template into which the extracted text is transferred
        - Flag variable, whether to search for the start or end of the procedure
        - Regex pattern as a string for the start and end of a procedure - whereby the placeholder for the procedure type in the subclasses must still be replaced within this string. Also

    """

    # TEMPLATE = "templates/prozedur.md"
    # HACK: for further development of the retrieval sequence:
    TEMPLATE = "templates/prozedur_dev.md"

    
    # TODO: Make this parameterizable as an option via the GUI:
    __print_final_calling_sequence_message = True
    # print_final_calling_sequence_message = False


    search_for_begin = True # initialer wErt




    # TODO: Variables should better be privat ones (?):

    regex_begin_pattern = SyntaxClass.get_pattern_start_of_procedure()

    regex_end_pattern = SyntaxClass.get_pattern_end_of_procedure()


    __regex_excluded_comment_pattern = SyntaxClass.get_pattern_single_line_comment()


    # No further concretization by subclasses is required here, therefore compiled directly:
    regex_excluded_comment = re.compile(__regex_excluded_comment_pattern, re.VERBOSE | re.IGNORECASE)

    @classmethod
    def set_print_final_calling_sequence_message(cls, value:bool):
        """
        Value passed for setting whether a completion record should be added after each completion of each procedure in the calling sequence.
        """
        if type(value) == bool:
            cls.__print_final_calling_sequence_message = value
            return
        
        raise Exception("Invalid parameter selected for __print_final_calling_sequence_message!")   




    @classmethod
    def find_multiline_comment(cls, code:str):
        """
        ### TODO: Currently only prepared for the usage of Python!
        
        Returns / stores to cls:
            list[tuple] : Format: [(comment1_start_line:int, comment1_end_line:int, comment1_text:str), (..., ..., ...)]

        """


        # # Forecast: For python OK!
        # pattern = re.compile(r'(""".*?""")|(\'\'\'.*?\'\'\')', re.DOTALL)

        # # TEST: @ #6 : 
        # # Comment at Issue #6:
        # """
        # Next step to test with VBA as in VBA there is no syntax available for multiline-comments:

        # As a hack: Invent some random keyword for simulate a multiline-comment in VBA. eg: "#####" as this string should be used in VBA-Codes (comments) rather rarely. Than implement this fake-syntax in one of the demo-input-files (.bas) to validate the functionality.

        # When this works, it should work with real syntax as well (for python, cpp, ...).
        # """
        # # TEST: @ #6 :  Test works fine! The invented syntax was not documented in the auto-generated-documentation! --> fix OK!
        # pattern = re.compile(r'(#####.*?#####)', re.DOTALL)

        # TODO: Get the pattern from a class (e.g.) of the programming language to analyse.
        # pattern = None # for VBA

        # pattern = SyntaxClass.pattern_multiline_comment
        pattern = SyntaxClass.get_pattern_multiline_comment()
        

        comments = []
        multiline_comment_line_numbers = [] # 0-basiert!

        # If no multiline comments exist in the relevant programming language:
        if pattern == None:

            cls.multiline_comments = tuple(comments)
            cls.multiline_comment_line_numbers = tuple(multiline_comment_line_numbers)
            return


        matches = pattern.finditer(code)

        for match in matches:

            comment = match.group(0)

            start_line = code.count('\n', 0, match.start()) # the 0-based index is saved
            end_line = code.count('\n', 0, match.end()) # the 0-based index is saved

            comments.append((start_line, end_line, comment))

            # TODO: It would make more sense (less memory, but more logic) to save only from and to, and then use a function to check whether the relevant line is within an exclusion range. This would also be relevant for class methods in Python or similar (pattern is exactly the same as Functions but within a class...)
            # Listing of all line numbers of the area:
            for line_no in range(start_line, end_line + 1):
                multiline_comment_line_numbers.append(line_no)


        for start_line, end_line, comment in comments:
            print(f"Found comment starting at line {start_line} and ending at line {end_line}:\n{comment}\n")

        cls.multiline_comments = tuple(comments)
        cls.multiline_comment_line_numbers = tuple(multiline_comment_line_numbers)


        # TODO: Must now be referenced at calling_sequence and references - filter that the line to be checked is NOT in one of these areas with 
            #if line_no in cls.multiline_comment_line_numbers: ...
        
        # OBSOLETE:
        return comments





    @classmethod
    def initialize_input_code(cls, input_path:str):
        """
        Reads in the input source code to be analyzed and documented and saves it within the superclass as an always available list of individual line contents.

        ### TODO: Later, a simple GUI should also be created here for selection!
        If necessary, this GUI should be detached from this class (repeat function!). Therefore as encapsulation with passed input-path!
        """

        """
        # ALTERNATIVE / OBSOLETE (before tackling issue #6):
        with open(input_path, "r") as file:
            cls.raw_source_code = file.readlines()
        """

        # New approach for Issue #6 : 
        # First read in as str, so that you can determine the line numbers in addition to the lines and later ignore block comments (multiline comments) in the references / calling sequences:
            
        with open(input_path, "r") as file:
            raw_source_code_str = file.read()

        # split to lines:
        cls.raw_source_code = raw_source_code_str.splitlines(keepends=True)
        cls.find_multiline_comment(raw_source_code_str)






    @classmethod
    def detail_analyse_procedures(cls):
        """
        Detailed evaluation of all previously identified procedures (regardless of their type).
        The individual components are stored in objects of the subclasses
        """

        for procedure_type_cls in [Sub, Function]:

            for line_begin, line_end in procedure_type_cls.matches_line_ixs:

                # _BUG(?FIX?): if there is an End {procedure_type} in the very last line of the source code and NO SPACE LINE FOLLOWS, line_end = None . This raises a TypeError: unsupported operand type(s) for +: 'NoneType' and 'int'
                lines = cls.raw_source_code[line_begin:line_end + 1]

                sub = procedure_type_cls(tuple(lines))



    @classmethod
    # def identify_procedures(cls, list_of_code_lines:list[str]):
    def identify_procedures(cls):
        """
        Identify all procedures (methods and functions) at class level and save the line numbers of the declaration and end lines within the respective class variables.

        The individual strings of the line contents of the source code passed as a list are searched.

        Args:
            raws (list[str]): List with one entry per line of the source code

        """

        for ix, text in enumerate(cls.raw_source_code):

            if cls.search_for_begin:

                if Sub.check_and_get_match(text, ix):
                    current_procedure_type_cls = Sub
                    continue

                if Function.check_and_get_match(text, ix):
                    current_procedure_type_cls = Function


            else:
                # Search for the appropriate end:
                current_procedure_type_cls.search_end_of_procedure(text, ix)



            













    @classmethod
    def check_and_get_match(cls, text:str, line_no:int) -> bool:
        """
        Checks a single line of code (the transferred text) to see whether it is a match of a declaration line and also whether it is NOT commented out. If this is the case, this match is included in the list of matches.

        Args:
            text (str): Text section to be checked
            line_no (int) : Line number of the list of total lines to be checked

        Returns:
            bool: True if the match was recorded
        """

        if cls.regex_begin.match(text):
            if not cls.regex_excluded_comment.match(text):

                # TEST: consider multiline-comments:
                if line_no not in cls.multiline_comment_line_numbers:

                    # Then add to the list of MAtches! incl. placeholder for end line number:
                    cls.matches_line_ixs.append([line_no, None])
                    Procedure.search_for_begin = False

                    return True
        
        return False






    @classmethod
    def search_end_of_procedure(cls, text:str, line_no:int) -> bool:
        """
        Checks a single line of code (the transferred text) to see whether it is a match of a declaration END line and also whether it is NOT commented out. If this is the case, this MAtch is included in the list of MAtches.

        Args:
            text (str): Text section to be checked
            line_no (int) : Line number of the list of total lines to be checked

        Returns:
            bool: True if the match was recorded
        """

        
        # BUG: The regex also finds something like End sub although it is about end function! Therefore, the program flow in the main was adapted to ensure that the correct class termination is always searched for... If you're bored: See why later...

        if cls.regex_end.match(text):
            if not cls.regex_excluded_comment.match(text):
                # Dann aufnehmen in die Liste der MAtches!
                cls.matches_line_ixs[-1][-1] = line_no

                Procedure.search_for_begin = True

                return True
        
        return False


    


    @classmethod
    def initialize_page_top_text(cls):
        """
        Initializes the text for the Markdown text of the page header and saves it in the cls.head class variable. 
        Access is via the Procedure superclass. 

        
        """

        page_top_text = cls.__read_template("templates/sec_head.md")

        # Insert total number of available procedures per type:
        placeholder_replacer = {
            "@PLACEHOLDER_INPUT_FILE@" : MetaData.get_input_filename(),
            "@PLACEHOLDER_TIMESTAMP_NOW@" : MetaData.date_of_process,
            "@PLACEHOLDER_TIMESTAMP_SOURCEFILE@" : MetaData.input_file__date_of_change,
        }


        for placeholder, replacer in placeholder_replacer.items():
            page_top_text = page_top_text.replace(placeholder, str(replacer))

        cls.page_top_text = page_top_text
        




    @classmethod
    def initialize_toc(cls):
        """
        Initializes the text for the Markdown text of the table of content and saves it in the class variable cls.toc. Access is via the Procedure superclass. 
        Within the resulting text, there is still 1 placeholder for each procedure type. These will be replaced separately later.
        """

        toc = cls.__read_template("templates/sec_toc.md")

        # Insert total number of available procedures per type:
        placeholder_replacer = {
            "@PLACEHOLDER_SUBS_COUNTS@" : len(Sub.instances),
            "@PLACEHOLDER_FUNCTIONS_COUNTS@" : len(Function.instances),
        }


        for placeholder, replacer in placeholder_replacer.items():
            toc = toc.replace(placeholder, str(replacer))

        cls.toc = toc
        






    @classmethod
    def generate_toc_entries(cls, subklasse):
        """
        Generates the individual entries of all procedures of the passed subclass 
        and inserts them into the class variable cls.toc by using placeholders.
        Access at superclass level

        Args:
            subklasse (Sub | Function): Class that is currently to be listed
        """


        toc = cls.toc # shortcut

        keyword_type = subklasse.KEYWORD_TYPE.upper()
        placeholder = f"@PLACEHOLDER_ENTRIES_{keyword_type}S@"
        

        for (procedure_obj, procedure_name, procedure_line_of_declaration) in subklasse.all_procedures_final:

            # Iterate over each instance:

            # TODO: What to do if there is NO instance???! - should not be a problem, because the initial state of the placeholder is always restored after each insertion and it is deleted at the end??!

            # Structure of the Markdown code for this TOC entry:
            new_entry = cls.get_markdown_for_code_line_of_call_entry(procedure_name, procedure_line_of_declaration, line_text="")

            # Fügen Sie den Platzhalter ein, um dieses Kennzeichen für weitere Einträge zu haben:
            new_entry = new_entry + "\n  " + placeholder
            # Insert the placeholder to have this indicator for further entries:


            toc = toc.replace(placeholder, new_entry)


        # After all procedures of 1 type: Delete the leaving placeholder in the toc for this type of procedures.
        toc = toc.replace(placeholder, "")

        # store into class-variable:
        cls.toc = toc

        













    @classmethod
    def analyse_references(cls):
        """
        Access is via the Procedure superclass. Within the method, it will be iteratded over each object of each subclass.

        For simplification, a new class variable is created at superclass level, which contains all elements of the two lists of the same name of the individual subclasses Sub and Function.all_procedures_final. This list is sorted by ascending line number.

        The system then searches for references (calls) of each individual procedure in the entire source code. If a match is found, it is further identified within which higher-level procedure this call was made. 
        An object variable (list) references is created for each object, which is initially empty and is extended with a tuple of the following form when matches are found: (line_no, bezeichnung_ueberordnete_prozedur, zeilentext_des_aufrufes).

        This method also forms the basic database on which the calling_sequences can then be analyzed.

        """

        # Creation of the common list at superclass level:
        all_procedures = Sub.all_procedures_final + Function.all_procedures_final


        # Sort the new list: Based on the line number of the declaration line
        cls.all_procedures_final = sorted(all_procedures, key=lambda stored_tuple: stored_tuple[2])


        for (procedure_obj, procedure_name, procedure_line_of_declaration) in cls.all_procedures_final:

            # Create an object variable: List of all references still to be found:
            procedure_obj.references = []
            
       
            regex_including_patterns = SyntaxClass.get_pattern_references(procedure_name)


            regex_excluding_patterns = [
                cls.regex_excluded_comment,
            ]
            



            # Search the ENTIRE SOURCE TEXT for a call to this procedure:
            for line_no, line_text in enumerate(cls.raw_source_code, 1):


                if line_no == procedure_line_of_declaration:
                    # Then this is the declaration line of the function for which the calls are to be found - so ignore it!
                    continue

                
                # TEST: Has it to be just    line_no    or    line_no - 1 ???
                if line_no - 1 in cls.multiline_comment_line_numbers:

                    # then multiline-comment --> not relevant
                    # TEST: For testing with VBA everything seems to be ok, result of example_module1.md matches before/after.
                    continue


                for regex_pattern in regex_including_patterns:
                    
                    regex = re.compile(regex_pattern)



                    if (match:=regex.search(line_text)):

                        break
                
                if not match:
                    
                    # then all regex have been no match -> next line
                    continue


                # TODO: Backreference to the group with the sub-name (no backref necessary, as it is only searches for this!) necessary for call sequence
                target_procedure_name = procedure_name

                # At this point, match is ALWAYS != None:

                # Check wheather there is a comment-symbol before the procedure name:
                __procedure_name_start_pos = match.span()[0]
                if (cls.regex_excluded_comment.search(line_text[0:__procedure_name_start_pos])):
                    # then: it is a comment
                    continue


                # An actual call has been found here, which must now be documented:
                # Now the call must be documented!
                

                # only for me privat (learning... ;-) )
                # NOTE: IMPORTANT: The following step / the logic and usage (python syntax) was new to me - research again! both with the filtering / by own key, as well as that the result is not THE SINGLE NUMBER, but actually directly the TUPLE IN WHICH THIS SINGLE NUMBER IS STORED!  Extremely practical!!!!

                
                # Logic for finding the calling procedure: 
                # # The declaration line number must be < line_no. 
                # # The next one is the line number of the declaration line
                calling_procedure_tuple_relevant = [tpl for tpl in cls.all_procedures_final if tpl[2] < line_no]

                # The relevant tuple of the list cls.all_procedures_final is stored in the same form in the following variable:
                calling_procedure = min(calling_procedure_tuple_relevant, key=lambda x_tuple: abs(x_tuple[2] - line_no))


                # Adding the name of procedure which is called:
                procedure_obj.references.append(
                    (line_no, calling_procedure[1], line_text, target_procedure_name)
                )



               



    @classmethod
    def finalize(cls, output_file_path:str):
        """
        ### TODO: Whether everything really belongs in this class is questionable! 
        
        Access via the Procedure superclass. Within the method, it is again iterated in loops via the subclasses.
        
        Prepares the final output and calls further methods to write this output (which is the Markdown file).
        This preparation includes
         
            - Sorting the individual procedures within the different procedure types according to the alphabetical order of their identifiers (names)
            
            - Configuring / initializing the subheadings / header texts below the individual section headings (stored in class variable cls.header)


        Args:
            output_file_path (str): File path of the Markdown file to be created
        """

        cls.initialize_page_top_text()


        # =============================================================================
        #### # Extract the module-wide docstring (if there is one provided): ####
        # =============================================================================
        
        # start at line 1 instead of line 0 as in the line 0 are stored codes of decoding
        __modul_docstring = cls.identify_docstring(cls.raw_source_code[1:], trim_empty_rows=True, return_alternativ_text=True)
        __template_docstring = cls.__read_template("templates/sec_modulinfos.md")
        # Replace placeholders in the template with the docstring found:
        cls.modul_docstring = __template_docstring.replace("@PLACEHOLDER_MODUL_DOCSTRING@", __modul_docstring)

        

        # preparation of the table of content (TOC):
        cls.initialize_toc()

        for procedure_type_cls in [Sub, Function]: 

            # Initialize the header below the section heading (from separate template) and save in class variable:
            content = cls.__read_template(procedure_type_cls.TEMPLATE_SECTION_HEAD)
            procedure_type_cls.header = content

            # Sort instances of the procedures alphabetically by name and save them in class variable:
            procedure_type_cls.sort_procedures_by_names()


            # Add a single entry for a method! and attach it to the TOC            
            # NOTE: This works on superclass-level
            cls.generate_toc_entries(procedure_type_cls)

        
        # Search the entire source code for all references for all procedures found 
        # and save them in the respective objects of the individual procedures:
        cls.analyse_references()


        # Analyze each procedure and save each further call of another procedure in the procedure object. 
        # Writing this to file will be done later, as all calling sequences can then be accessed recursively.
        cls.analyse_call_sequences()
        # TODO / TEST -->: in there: implement query wheather line no is in multiline_comment_line_numbers
        # TEST: Actually the analysis bases on analyse_references. \
        # As pattern for references in multiline-comments are not stored by the last change of analyse_reference, \
        # there should be no need to modify analyse_call_sequences.

        cls.prepare_all_call_sequence_docs()



        # Call the method for actually writing the text file:
        cls.write_to_file(output_file_path)







    @classmethod
    def get_procedure_obj_by_name(cls, procedure_name_to_get:str) -> 'Procedure':
        """
        Returns the procedure object with the transferred name
        The object is from the subclass Sub or Function.

        Args:
            procedure_name (str): Name of the procedure searched for
        """

        for (procedure_obj, procedure_name, procedure_line_of_declaration) in Procedure.all_procedures_final:
            if procedure_name == procedure_name_to_get:
                return procedure_obj

        db("not found!")     # only for debugging

        return None


    @staticmethod
    def indent_str(text:str, count_of_indents:int=0) -> str:
        """
        Precedes each line with the corresponding indention symbol.

        Args:
            text (str): Text
            count_of_indents (int, optional): Level of indention. Defaults to 0.

        Returns:
            str: indented text.
        """
        CHARS_PER_INDENT = "  "

        indendet_text = ""
        
        pre_line:str = CHARS_PER_INDENT * count_of_indents

        list_of_lines = text.split("\n")
    
        for line in list_of_lines:
            if line != "":
                # continue
                # indendet_text = indendet_text + "\n"
            # else:
                indendet_text = indendet_text + pre_line + line
            
            indendet_text = indendet_text + "\n"
    


        return indendet_text
    


    @staticmethod
    def get_markdown_for_code_line_of_call_entry(proc_name:str, line_no:str, line_text:str) -> str:
        """
        Generates a string that can be inserted into a Markdown file with the corresponding syntax.
        The string contains the name of a procedure, a relevant line number and the relevant text of the code line.

        Within the method, the string is structured in such a way that the procedure name is later linked to this procedure.

        The method can be used both for calling_sequences (calls) and for references (superordinate / calling procedures) and ensures that the output format is always identical.
        
        Args:
            proc_name (str): Name of the procedure
            line_no (str): Relevant line number in the code (as str)
            line_text (str): Text line in the code


        Returns:
            str: String in Markdown syntax that provides an interactive overview of the parameters in readable form.
        """

        # proc_name:str, line_no:str, line_text:str
    

        # replacer_placeholder_reference = "\n" * 3 + f"- [`{target_procedure_name}`](#{target_procedure_name}) <small> : [Zeile {line_no_reference}] : `{line_code}` </small>".replace(line_code, line_code.rstrip("\n")) + "\n"
        
        if line_text == "":
            __optional_code_line = ""
        else:
            __optional_code_line = f": `{line_text}`"


        markdown_entry = f"- [`{proc_name}`](#{proc_name}) : <small>  [Zeile {line_no}] {__optional_code_line} </small>"

        markdown_entry = markdown_entry.replace(line_text, line_text.rstrip("\n")) 
        
        markdown_entry = markdown_entry + "\n" * 2


        return markdown_entry





    def prepare_single_call_sequence_docs(self, level=0):
        """

        ### TODOC: docs Perhaps not up to date!
        NOTE: This method has been changed fundamentally, and I don't think that this docstring is up to date!


        
        Generates the complete documentation of the calling sequences in the individual objects and saves it in the obj.calling_sequences_doc attribute. After completion, obj.calling_sequences_state is set to True.

        Access is at object level.
        """
        db(f"name of the procedure which should be analysed : {self.name}")

        if self.calling_sequences_state:

            end_text_per_procedure = ""

            text_to_return = self.calling_sequences_doc +  end_text_per_procedure
            
            
            # NOTE: do not insert  any indentions, as this will be done at the very end before writing
            text_to_return = self.indent_str(text_to_return, 0)


            return text_to_return


        # If there is no text yet, initializes it from the template:
        if not self.calling_sequences_doc:
            self.calling_sequences_doc = ""


        for (line_no_reference, uebergeordnetes_sub, line_code, target_procedure_name) in self.calling_sequences:
            
            replacer_placeholder_reference =  self.get_markdown_for_code_line_of_call_entry(target_procedure_name, line_no_reference, line_code)

            # NOTE: do not insert  any indentions, as this will be done at the very end before writing
            replacer_placeholder_reference = self.indent_str(replacer_placeholder_reference, 0)
            
            self.calling_sequences_doc = self.calling_sequences_doc  + replacer_placeholder_reference 

            level = level + 1

            # Recursive calls for subordinate calling sequence:
            if target_procedure_name == self.name:
                # If finding a recursive call in the procedure, stop the further documentation, as this will be endless recursivly, too.
                # Instead, add a sentence to document this  to the the user later.
                further_calls_doc = "- <small> *... recursivly calls itself under certain conditions ...* </small> \n\n"

            else:

                # get the object from target_procedure_name:
                target_procedure_obj = self.get_procedure_obj_by_name(target_procedure_name)

                # Recursive: fill the next placeholder with the next documentation of the called call:
                # further_calls_doc = target_procedure_obj.prepare_single_call_sequence_docs(level + 1)
                further_calls_doc = target_procedure_obj.prepare_single_call_sequence_docs(level + 0)

            
            # NOTE: **HERE** the indentations will be inserted at last:
            self.calling_sequences_doc = self.calling_sequences_doc  + self.indent_str(further_calls_doc, count_of_indents=level)

            # Decrease the level:
            level = level - 1


        # Increase the level:
        level = level + 1


        ending_note = "\n- <small>*No further calls to other procedures that are documented here.*</small>"
        if Procedure.__print_final_calling_sequence_message == False:
            ending_note = ""
        
        # NOTE: do not insert  any indentions, as this will be done at the very end before writing
        ending_note = self.indent_str(ending_note, 0)


        self.calling_sequences_doc = self.calling_sequences_doc + ending_note

        self.calling_sequences_state = True
        text_to_return = self.calling_sequences_doc

        return text_to_return
























    @classmethod
    def prepare_all_call_sequence_docs(cls):
        """
        Generates the complete documentation of the calling sequences in the individual objects and saves it in the obj.calling_sequences_doc attribute. After completion, obj.calling_sequences_state is set to True.

        Access takes place at superclass level, whereby the instances of the class are referenced internally

        
        # TODOC: Why is it labeld with 'NO' ? Should it be so but it is not yet implemented?
        # NO: After actually iterating and documenting the calling sequence, the placeholders for the overview are replaced within this class method for the object (number and introductory record).


        """


        for procedure in cls.all_procedures_final:
            procedure_obj:Procedure = procedure[0]
            db(procedure_obj.name)

            procedure_obj.prepare_single_call_sequence_docs(level=0)

            db(len(procedure_obj.calling_sequences), procedure_obj.calling_sequences)

            db("next procedure...")
            

        db("all call sequences are ready prepared.")

























    @classmethod
    def analyse_call_sequences(cls):
        """
        It iterates over all procedure objects to be documented and calls the method analyze_calling_sequences_in_one_proc for each object, in which all calling sequences within this one procedure are determined.

        Access takes place at superclass level, whereby the instances of the class are referenced internally and these are passed through to the object method!
        """
        cls.all_procedures_final

        # Use the existing object method for each procedure:
        for procedure in cls.all_procedures_final:

            procedure_obj:Procedure = procedure[0]

            procedure_obj.analyse_calling_sequences_in_one_proc()



            # Sortiere generierte Liste aufsteigend nach den Aufrufzeilen:
            procedure_obj.calling_sequences = sorted(procedure_obj.calling_sequences, key=lambda stored_tuple: stored_tuple[0])


        db("All procedures analyzed, not yet documented!")




	








    
    def analyse_calling_sequences_in_one_proc(self) -> None:
        """
        Determines the calling sequence of external (but documented by this documenter tool) procedures within this procedure self.
        The result is stored in the object variable self.calling_sequences (list of tuples, where each tuple represents a call).
        Within this method, the self.documentation is NOT changed, as this should only be done after this method analyze_calling_sequences_in_one_proc has been run for all VBA procedures to be documented.

        ### NOTE:
        This method is based on the previously used method OLD_generate_calling_entries, which was the starting point of the development and, in addition to saving the findings, also documented them directly. 
        """

        self.calling_sequences = []

        # Suche für alle vorhandenen Prozeduren jeweils ihre Referenzen heraus, und prüfe, ob eine Referenz zu der jetzt gerade untersuchten Prozedur (self.name) führt. Wenn ja, ist dies ein Aufruf, der der calling_sequences liste angehangen werden soll.
        for (procedure_obj, procedure_name, procedure_declaration_line_no) in Procedure.all_procedures_final:
            
            # Filter the list so that parent_sub == name
            db("Search call sequence for procedure >> {}\nCurrently: check, wheather the following procedure is be called: >> {}".format(self.name, procedure_name)) # is ALWAYS EQUAL in a single call!

            
            relevant_references = [_ for _ in procedure_obj.references if _[1] == self.name]
            
            
            db("----> Number of relevant references: {}".format(len(relevant_references)))



            # =============================================================================
            ### # Add the relevant references as calls within the procedure to be analyzed: ####
            # =============================================================================
            
            for reference in relevant_references:
                self.calling_sequences.append(reference)


            # NOTE: Access later if required:

        db("Calling Sequences for procedure {}:\n{}".format(self.name, self.calling_sequences))










    def generate_reference_entries(self) -> None:
        """
        Generates the documentation of the referencing from the object of the Function or Sub subclass.
        Saves the result in the existing object variable procedure_obj.documentation
        """



        # TODO: Structure...? that belongs to the references... is actually constant, would not have to be redone every time the function is called, but that's how it belongs together... (I think poor structure)
        _PLACEHOLDER_REFERENCE = "@PLACEHOLDER_PROCEDURE_REFERENCES_ENTRY@" 



        # shortcut:
        doc:str = self.documentation 

        # All referencing is stored in the object variable procedure_obj.references:
        count_of_references = len(self.references)

        # Set the placeholder for the number of references in the previous documentation:
        doc = doc.replace("@PLACEHOLDER_PROCEDURE_COUNT_OF_REFERENCES@", str(count_of_references))
        

        # Initialization and parameterization of the introductory record:
        einleitungssatz = "Kein Aufruf gefunden." # default

        if count_of_references > 0:

            einleitungssatz = "The procedure is called in the following higher-level procedures:"

            
            # Iterate over each referencing to document it:
            for (line_no, calling_procedure_name, line_code, target_procedure_name) in self.references:

                # Construct the replacement value for the placeholder incl. appending the placeholder for further replacements:
                
                # In order not to break MArkdown, the last line break of the line_code must be removed:
                
                # replacer_placeholder_reference = f"* [`{calling_procedure_name}`](#{calling_procedure_name}) : <small>  [Zeile {line_no}] : `{line_code}` </small>"


                replacer_placeholder_reference = self.get_markdown_for_code_line_of_call_entry(calling_procedure_name, line_no, line_code)



                # # FORECAST: more collapse details : works technically, but the collapsable is always in a new line and therefore it will be big - maybe make it nice later...

                # __replacer_placeholder_reference = f"*   [`{calling_procedure_name}`](#{calling_procedure_name})  <details> <summary>: <small>Zeile {line_no}</small> </summary> `{line_code}` </details>"



                replacer_placeholder_reference = replacer_placeholder_reference  + _PLACEHOLDER_REFERENCE
                
                # Replacing:
                doc = doc.replace(_PLACEHOLDER_REFERENCE, replacer_placeholder_reference)


        # Delete the remaining placeholder for inserting individual references:
        doc = doc.replace(_PLACEHOLDER_REFERENCE, "")



        # Insert the introduction sentence:
        doc = doc.replace("@PLACEHOLDER_PROCEDURE_REFERENCES_INTRODUCTION@", einleitungssatz)


        # resubstitution of the shortcut:
        self.documentation = doc
        







    @classmethod
    def write_to_file(cls, output_file_path:str):
        """
        Writes the documentation of all procedures (of all types) to the target file (path) passed as the file path.
        """

        # initialize_toc

        with open(output_file_path, "w", ) as file:

            # TODO: This doesn't really belong in the Procedure class anymore, but it still belongs here to have everything together. 
            # Possibly refactor later...
            
            # Initialize the page with title and organizational notes:
            content = cls.__read_template("templates/sec_head.md")

            file.write(cls.page_top_text)

            # Insert Index / TOCs:
            file.write(cls.toc)


            # Write the Modulinfos / Docstrings:
            file.write(cls.modul_docstring)


            # # TODO: Poor structure...? actually this regards to the references... Maybe refactor later...
            # _PLACEHOLDER_REFERENCE = "@PLACEHOLDER_PROCEDURE_REFERENCES_ENTRY@" 


            for procedure_type_cls in [Sub, Function]: # Order is essential for the order of the documented sections!


                # Insert headline of the section (from individual template file)
                file.write(procedure_type_cls.header)



                # Document all individual procedures:
                for procedur_infos in procedure_type_cls.all_procedures_final:
                # for (procedure_obj, procedure_name, procedure_line) in procedure_type_cls.all_procedures_final:

                    procedure_obj:Procedure = procedur_infos[0]





                    procedure_obj.generate_reference_entries()

                    db("name der aktuell zu dokumentierenden prozedur: {}".format(procedure_obj.name))
                    # Checking + Dokumentation Calling Sequences:
                    # procedure_obj.generate_calling_entries()
                    if not procedure_obj.calling_sequences_state:
                        db(f"PROBLEM! with {procedure_obj.name}")
                        db(f"PROBLEM! with {procedure_obj.name}")
                        # BUG: Problem! (wow! this does not help me anymore, and I tagged it myself :0: )
                    procedure_obj.documentation = procedure_obj.documentation.replace("@PLACEHOLDER_PROCEDURE_CALLING_SEQUENCES_BLOCK@", procedure_obj.calling_sequences_doc)



                    # =============================================================================
                    #### Enter overview parameters for the object: / the procedure: ####
                    # =============================================================================
                    
                    # Initialization and parameterization of the introductory record:

                    if procedure_obj.calling_sequences:

                        introduction_note = "The following subordinate procedures are called within the procedure:"

                    else:

                        introduction_note = "No further calls found for procedures documented here." # default


                    # Insert introduction sentence: 
                    procedure_obj.documentation = procedure_obj.documentation.replace("@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@", introduction_note)


                    # Set the overview count of calls:
                    procedure_obj.documentation = procedure_obj.documentation.replace("@PLACEHOLDER_PROCEDURE_COUNT_OF_ABRUFFOLGE@", str(len(procedure_obj.calling_sequences)))







                    # Write the modified procedure documentation to the file:
                    file.write(procedure_obj.documentation)


                # Print summary per documented section:
                print("Es wurden {} {}s identifiziert und dokumentiert.".format(len(procedure_type_cls.instances), procedure_type_cls.KEYWORD_TYPE))



            ## TODO: This doesn't really belong in the Procedure class anymore, but it still belongs here to have everything together. Possibly refactoring later...
            # Finalize the page with closing remarks:
            content = cls.__read_template("templates/sec_tail.md")
            # HACK: Actually for debugging usage... (or not?)
            content = content.replace("@PLACEHOLDER_TIMESTAMP_NOW@", MetaData.date_of_process)
            content = content.replace("@PLACEHOLDER_DOC_PYTHON@", __doc__.replace("\n", "<br>"))
            content = content.replace("@PLACEHOLDER_DOCUMENTER_VERSION__AUTHOR@", MetaData.documenter_version__author.rstrip(" <_>"))
            content = content.replace("@PLACEHOLDER_DOCUMENTER_VERSION__COMMIT@", MetaData.documenter_version__commit)
            content = content.replace("@PLACEHOLDER_DOCUMENTER_VERSION__DATE@", MetaData.documenter_version__author_date)

            file.write(content)

        print("MD-file has been created at: {}".format(MetaData.get_output_path(".md")))


    @classmethod
    def sort_procedures_by_names(cls):
        """
        Filling the class variable of the **SUBCLASS** cls.all_procedures_final with one tuple per procedure, which contains the individual object as well as its name and the line number of the declaration in the form [(object:Procedure, object.name:str, object.line_begin)]. 
        This list is sorted after completion based on the alphabetical identifiers, so that the connection between the objects and the names is still given, but at the same time the objects can be documented in the alphabetical order of their names.

        """

        all_procedures = [] 


        for procedure_obj in cls.instances:
            
            all_procedures.append((procedure_obj, procedure_obj.name, procedure_obj.line_begin))



        # After filling: Sort based on the identifier names:
        cls.all_procedures_final = sorted(all_procedures, key=lambda stored_tuple: stored_tuple[1])



        
            













    @staticmethod
    def __read_template(template_file_path) -> str:
        """
        Reads the passed Template-MArkdown file and returns the content as a string.
        
        Args:
            template (str) : File path to the template

        Return:
            str : Text content of the template
        """
        with open(template_file_path, "r") as file:
            content = file.read()
            return content





    def read_template(self, template="einzelprozedur") -> str:
        """
        Reads the relevant Template-MArkdown file and returns the content as a string. 
        If this is the default template for any procedure, the self.documentation
        the self.documentation attribute is also initialized with this text.

        Args:
            template (str) : File path to the template - Exception: If the default template stored in the class 
                            default template self.TEMPLATE stored in the class is to be used for a procedure, 
                            then the keyword "single procedure" must be passed. 
                            This is the default case --> no specification required.

        Return:
            str : Content of the template
        """


        if template == "einzelprozedur":
        
            # Read out and save in object variable
            self.documentation = self.__read_template(self.TEMPLATE)
            return self.documentation


        # For non-standard procedure template: Read out this template and return the text content:
        content = self.__read_template(template)
        return content





    def generate_documentation(self):
        """
        Documents the parameters found using a template markdown text file and saves the text to be written in the self.documentation attribute
        """

        self.read_template()

        placeholder_replacer = {
            "@PLACEHOLDER_PROCEDURE_TYPE@" : self.KEYWORD_TYPE,
            "@PLACEHOLDER_PROCEDURE_MODIFIER@" : self.modifier,
            "@PLACEHOLDER_PROCEDURE_NAME@" : self.name,
            "@PLACEHOLDER_PROCEDURE_LINE_BEGIN@" : self.line_begin,
            "@PLACEHOLDER_PROCEDURE_DOCSTRING@" : self.docstring,
            "@PLACEHOLDER_PROCEDURE_SOURCE_CODE@" : self.source_code,
        }


        for placeholder, replacer in placeholder_replacer.items():
            self.documentation = self.documentation.replace(placeholder, str(replacer))













    def extract_source_code(self):
        """
        Reads the source_code and saves it in the self.source_code attribute
        """

        source_code = ""
        for line in self.lines:
            source_code = source_code + line

        self.source_code = source_code





    def extract_modifier(self):
        """
        Reads the modifier and saves it in the self.modifier attribute
        """
        # Reuse the declaration line regex:
        match = self.regex_begin.match(self.lines[0])

        # Remove the string up to the name:
        pos_name = match.string.find(self.name + "(")
        potential_modifier = match.string[0:pos_name]

        # Identify the modifier from this:
        
        # # Actually, it should already work on the basis of the actual regex, but it doesn't work! It always returns something else, therefore, new regex! (If there is time, see why!)
        # # Extraction of the group with the name: (actual approach)
        # db(match.groups())
        # modifier = match.group(1)
        
        
        regex_modifier = re.compile(r"((?:Private|Public|Friend)?)")
        match = regex_modifier.match(potential_modifier)
        modifier = match.group(1)


        if modifier == "":
            modifier = "Public"

        self.modifier = modifier




    def extract_name(self):
        """
        Reads out the identifier name and saves it in the self.name attribute
        """
        # Reuse the declaration line regex:
        
        
        match = self.regex_begin.match(self.lines[0])

        # Extraction of the group with the name:
        self.name = match.group(1)



    def extract_line_numbers(self):
        """
        Reads out the relevant start and end line numbers and saves them in the self.line_begin and self.line_end attributes
        """
        # Get index of the instance list, this is equal to the index of the matches_line_ixs list:
        ix = self.instances.index(self) 

        match_lines_ix = self.matches_line_ixs[ix]
        
        self.line_begin = match_lines_ix[0] + 1
        self.line_end = match_lines_ix[1] + 1




    @staticmethod
    def identify_docstring(text_lines:tuple[str], trim_empty_rows=False, return_alternativ_text=False) -> str:
        """
        ### NOTE: Developed from the method extract_docstring(self), which is to be used at object level. 
        To be able to use the same logic for the module-wide docstring, this construct is introduced.


        Reads the first block comment / docstring and returns it.
        Every comment line that is DIRECTLY AND WITHOUT PREVIOUS EMPTY LINES BELOW THE FIRST LINE SENT is counted as a docstring. 
        As soon as an empty line follows, the docstring is considered to have ended.

        Example:
                ' This is a comment directly below the declaration line. It is therefore considered a docstring.
                ' This also
                '
                ' As the previous line ALSO contains a comment (an empty one), this is still part of the docstring.

                ' There was a NOT comment before this line, so this is no longer part of the docstring
                MsgBox("This belongs to the program")
            End Sub

            
        Args:
            text_lines (tuple[str]): Tuple with the individual text lines belonging to the relevant procedure.
            trim_empty_rows (bool) : False, if the algorithm is to be aborted directly at an empty line ( = DEFAULT), 
                                        or True, if the algorithm should continue to run with preceding empty lines until there is NO more empty line or another termination condition is reached.
            return_alternativ_text (bool | str) : Default = False. An empty string is then returned if it is not found. 
                                                    If True, the alternative record text defined in the MEthode is returned in such cases. 
                                                    If a string is passed, this string is returned in these cases.
        Returns:
            str : Docstring / block comment
        """


        docstring  = ""


        for line in text_lines:
            # line:str
            content = line.lstrip(" ")

            if(content[0]) == "'":
               # belongs to Docstring
                docstring = docstring + content.lstrip("'")
                continue # next line


            if trim_empty_rows == False:

                break

            # ELSE: Then trim the empty row
            if docstring != "":
                # Then there is already a docstring -> Cancel
                break


            if re.match(r"^\s*\n?$", content):
                # Then the line consists only of whitespaces --> Trim!
                # Recursive call of this method with one frontmost line less!
                docstring = Procedure.identify_docstring(text_lines[1:], trim_empty_rows=True, return_alternativ_text=False) 
                # TODO: Error avoidance / borderline case: IndexOutOfRange (end arrived, --> docstring exists --> will raises error!
                break



        # =============================================================================
        #### # Alternative text for empty docstrings: ####
        # =============================================================================
        
        if return_alternativ_text == True:

            if docstring == "":
                docstring = "*No information availible. For more information expand source code.*" #  The asterisks cause an italic print in MArkdown

        elif isinstance(return_alternativ_text, str):
            if docstring == "":
                docstring = return_alternativ_text


        return docstring
  




    def extract_docstring(self):
        """
        # CHANGELOG: 2023-12-30 - 03:17:04 Until V. 0.0.5, this function was 'sole authoritative'. 
        Only then was the pass-through to the static method identify_docstring introduced,
        in order to be able to find module-wide docstrings.


        Reads the docstring and saves it in the self.docstring attribute
        If no docstring has been identified in the code, a corresponding info is written in the text.

        Every comment line that is DIRECTLY AND WITHOUT PREVIOUS SPACE LINE BELOW THE DECLARATION LINE of the procedure is considered a docstring. As soon as an empty line follows, the docstring is regarded as terminated.

        Example:
            Private Sub exampleProgram() ' Declaration line
                ' This is a comment directly below the declaration line. It is therefore evaluated as a docstring.
                ' This also
                '
                ' Since the previous line ALSO contains a comment (an empty one), this is still part of the docstring.

                ' There was a NOT comment before this line, so this is no longer part of the docstring
                MsgBox("This belongs to the program")
            End Sub
        """


        # Filter out docstring via generalized method:
        docstring = self.identify_docstring(self.lines[1:], return_alternativ_text=True)

        # Save in object:
        self.docstring = docstring



    

    def __init__(self, text_lines:tuple[str]) -> None:
        """
        Creates a procedure object, whereby the individual components of the code are searched for and saved.
        An arbitrarily large tuple must be passed, each element of which contains the string of a single line of this procedure.

        Args:
            text_lines (tuple[str]): Tuple with the individual lines of text belonging to the relevant procedure.
        """

        # TODO: Ist  dies notwendig???! Eigentlich nicht! Dies ist schon die superklasse!
        super().__init__()

        # Remember the instanc for later iteration: Class assignment dynamic!
        type(self).instances.append(self)
        
        # Save all text lines:
        self.lines = text_lines


        # =============================================================================
        #### # Filter out individual components: ####
        # =============================================================================
        
        self.extract_line_numbers()
        self.extract_source_code()
        self.extract_name()
        self.extract_modifier()
        
        self.extract_docstring()



        # =============================================================================
        #### Summarizing and writing the documentation: ####
        # =============================================================================
        self.calling_sequences_doc = None
        self.calling_sequences_state = False


        self.generate_documentation()




















class Sub(Procedure):
    """
    Subclass for a VBA sub-procedure.
    Among other things, the regex expression pre-initialized by the superclass, which is required in VBA syntax for the start and end of the procedure type, is also specified and compiled here.

    Most of the methods are stored in the superordinate superclass, as they are identical for VBA subs and VBA methods.

    """

    matches_line_ixs = []
    instances = []

    all_procedures_final = [] # List is only created after all procedures have been identified. It is used to sort the instances based on the alphabetical order of their procedure names

    
    KEYWORD_TYPE = "Sub" 

    TEMPLATE_SECTION_HEAD = "templates/sec_subs.md"
    

    # Concretization and compilation of the regex expressions:

    regex_begin = re.compile(Procedure.regex_begin_pattern.replace("PLACEHOLDER_PROCEDURE_TYPE", KEYWORD_TYPE), re.VERBOSE | re.IGNORECASE)


    regex_end = re.compile(Procedure.regex_end_pattern.replace("PLACEHOLDER_PROCEDURE_TYPE", KEYWORD_TYPE), re.VERBOSE | re.IGNORECASE)















class Function(Procedure):
    """
    Subclass for a VBA sub-procedure.
    Among other things, the regex expression pre-initialized by the superclass, which is required in VBA syntax for the start and end of the procedure type, is also specified and compiled here.

    Most of the methods are stored in the superordinate superclass, as they are identical for VBA subs and VBA methods.

    """

    matches_line_ixs = []
    instances = []
    
    all_procedures_final = [] # List is only created after all procedures have been identified. It is used to sort the instances based on the alphabetical order of their procedure names
    
    KEYWORD_TYPE = "Function" 
    
    TEMPLATE_SECTION_HEAD = "templates/sec_functions.md"
    
    
    # Concretization and compilation of the regex expressions:

    regex_begin = re.compile(Procedure.regex_begin_pattern.replace("PLACEHOLDER_PROCEDURE_TYPE", KEYWORD_TYPE), re.VERBOSE | re.IGNORECASE)


    regex_end = re.compile(Procedure.regex_end_pattern.replace("PLACEHOLDER_PROCEDURE_TYPE", KEYWORD_TYPE), re.VERBOSE | re.IGNORECASE)









# =============================================================================
#### FUNCTIONS: ####
# =============================================================================






def load_parameter(documenter_gui_obj:DocumenterGui):
    """
    Loads all relevant attributes (set parameters) of the GUI into the MetaData (or Procedure) class and then initializes the MetaData class.

    Args:
        documenter_gui_obj (DocumenterGui): Object with parameters for the documentation to be created.
    """
    
    MetaData.set_output_dir(documenter_gui_obj.output_dir)
    MetaData.set_input_path(documenter_gui_obj.input_file)
    MetaData.set_convert_to_html(documenter_gui_obj.convert_checked)
    MetaData.set_user_defined_additional_text(documenter_gui_obj.optinal_user_defined_text)

    Procedure.set_print_final_calling_sequence_message(documenter_gui_obj.show_message)


    # Explicit call of the initialization method of the MetaData.
    # The class attributes are already filled with the data from the GUI by calling individual setters.
    MetaData.initialize_class()




def convert_markdown_to_html():
    """
    Converts the generated .md Markdown file into an HTML file.
    However, the formatting is not nearly as useful and clear due to the library used,
    as when the .md file is subsequently converted manually using VSCode (Markdown all in one extension)
    """

    if MetaData.get_convert_to_html():
        markdown.markdownFromFile(
            input=MetaData.get_output_path(extension=".md"),
            output=MetaData.get_output_path(extension=".html"), 
            encoding="utf8"
        )

        print("HTML file was generated from MD file: {}".format(MetaData.get_output_path(".html")))
        









def evaluate_debug_mode():
    """
    Checks wheather a parameter has been passed to start the script.
    If so, check wheather it is the flag for turning on the DEBUG-Mode
    """
    # Set debug mode if starting from terminal with arg:
    if len(argv) < 2:
        return
    
    if (argv[1] == '-d' or argv[1] == '--debug'):
        
        # Set the module constant DEBUG to true:
        print("set DEBUG MODE!")
        global DEBUG
        DEBUG = 1

    else:
        raise Exception("Invalid argument while starting script: '{}'\nPermitted is only the param '-d' or '--debug' to set the DEBUG mode.".format(argv[1]))









# =============================================================================
#### MAIN: 
# =============================================================================
def main():
    """
    Main program. Controls the overall flow of the script. 
    Most methods are defined within the Procedure superclass.
    """

    evaluate_debug_mode()



    DocumenterGui.set_debug_mode(DEBUG)

    gui = DocumenterGui()
    
    if not gui.get_is_ready():
        db("NO start! -> Abort")
        return

    db("Then start documentation from outside with the parameters by accessing the object variables")


    # Save PArameter of the GUI object in MetaData class and initialize the MetaDate class
    load_parameter(gui)


    # initialize the class including reading the input file:
    Procedure.initialize_input_code(MetaData.get_input_path())


    # Identify and save the declaration lines of procedures:
    Procedure.identify_procedures()


    # Detailed analysis and storage of individual components in objects:
    Procedure.detail_analyse_procedures()


    Procedure.finalize(MetaData.get_output_path())


    convert_markdown_to_html()

    if DEBUG:
        print("\n\nExecution finished.")
    else:
        gui.show_closing_window()











if __name__ == '__main__':

    print("START OF MAIN....")

    main()

    print(".... END OF MAIN.")
