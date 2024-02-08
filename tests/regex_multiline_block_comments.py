# -*- coding: utf-8 -*-
'''
Created on: Sun, 2024-01-21 (22:14:48)

@author: Matthias Kader


Test zu Issue #6 and #2:

Referenzenierungen innerhalb von Multiline-Comments dürfen nicht als solche gewertet werden!




'''















import re














class Procedure:
    """
    ### ACHTUNG: Nachbau - nur von den relevanten zu testzwecken erforderlichen Inhalten der Klasse. Diese sind z.T. modifizeirt ohne weiteren Kommentar!!
    """


    @classmethod
    def initialize_input_code(cls, input_path:str):
        """
        Liesst den zu analysierenden und zu dokumentierenden Input-Quellcode ein und speichert ihn innerhalb dder Superklasse als immer verfuegbare Liste einzelner Zeileninhalte ab.

        ### TODO: Später sollte hier auch noch eine einfache GUI erstellt werden zur Auswahl!
        ggf. sollte diese GUI aber losgelöst von dieser Klasse sein (Wiederholfunktion!). Daher als Kapselung mit übergebenen input-path!
        """

        with open(input_path, "r") as file:

            cls.raw_source_code_str = file.read()



        cls.raw_source_code_list = cls.raw_source_code_str.split("\n")













def find_lines_with_blockcomment(text:str) -> list[int]:


    pattern = re.compile(r'(""".*?""")|(\'\'\'.*?\'\'\')', re.DOTALL)

    matches = pattern.findall(text)


    comments = [match[0] or match[1] for match in matches]
    return comments



    pass





def find_multiline_comment_incl_range_numbers(code:str):
        """
        ### TODO: Aktuell nur fuer Python!
        
        Returns / stores to cls:
            list[tuple] : Format: [(comment1_start_line:int, comment1_end_line:int, comment1_text:str), (..., ..., ...)]

        """

        # HACK: nur python bisher!
        pattern = re.compile(r'(""".*?""")|(\'\'\'.*?\'\'\')', re.DOTALL)
        
        
        matches = pattern.finditer(code)
        

        comments = []
        multiline_comment_line_numbers = [] # 0-basiert!

        for match in matches:

            comment = match.group(0)

            start_line = code.count('\n', 0, match.start()) # + 1 # es wird der 0-basierte Index gespeichert!

            end_line = code.count('\n', 0, match.end()) # + 1 # es wird der 0-basierte Index gespeichert!


            comments.append((start_line, end_line, comment))

            # Auflistung aller Zeilennummern des Bereiches:
            for line_no in range(start_line, end_line + 1):
                multiline_comment_line_numbers.append(line_no)




        for start_line, end_line, comment in comments:
            print(f"Found comment starting at line {start_line} and ending at line {end_line}:\n{comment}\n")


        multiline_comments = tuple(comments)

        multiline_comment_line_numbers = tuple(multiline_comment_line_numbers)



        
        # TODO: Muss jetzt noch referenziert werden bei calling_sequence und references - filtern, dass die zu pruefnde Zeile NICHT in einem dieser Bereiche liegt
        # OBSOLET:
        return comments












def find_multiline_comment(code):
    pattern = re.compile(r'(""".*?""")|(\'\'\'.*?\'\'\')', re.DOTALL)
    matches = pattern.finditer(code)
    
    comments = []
    for match in matches:
        comment = match.group(0)
        start_line = code.count('\n', 0, match.start()) + 1
        end_line = code.count('\n', 0, match.end()) + 1
        comments.append((start_line, end_line, comment))
    
    return comments







Procedure.initialize_input_code("tests/regex_multiline_block_comments_beispieltext.py")


print(Procedure.raw_source_code_str)






# # OBSOLET / ALTERNATIVE:
# lines_block = find_lines_with_blockcomment(Procedure.raw_source_code_str)


multiline_comments = find_multiline_comment(Procedure.raw_source_code_str)



multiline_comments = find_multiline_comment_incl_range_numbers(Procedure.raw_source_code_str)





for start_line, end_line, comment in multiline_comments:
    print(f"Found comment starting at line {start_line} and ending at line {end_line}:\n{comment}\n")


















print("ende")



