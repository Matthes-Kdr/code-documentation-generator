# -*- coding: utf-8 -*-
'''
Created on: Fri, 2023-12-29 (00:45:39)


@author: Matthias Kader


Für generelles Ziel und Ablauf des Scriptes siehe MArkdown im Verzeichnis ../Tests/Programmablauf.html

Wichtige Details siehe am Ende dieses docstrings.




### Fertig implementiert:

- Inhaltsverzeichnis / Index

- Gesamtlayout inkl. Titel, Zwischenüberschriften für einzelne Sections

- Aufführen  des modulweiten Programmkopf-Docstring in der generierten Dokumentation

- Aufführen der References-Durchsuchungen (Wo wird die Prozedur aufgerufen?) in der generierten Dokumentation

- Sofortiger Export der MD-Datei in eine  HTML-Datei

- Aufführen der organisatorischer Daten bzgl. des zu dokumentierenden Codes und des verwendeten Skripts zum Dokumentieren in der generierten Dokumentation

- Aufführen der Calling Sequence (Aufrufabfolge / Aufrufebenen) innerhalb jeder Prozedur in der generierten Dokumentation: Aufzählung der Aufrufe anderer, in dieser Dokumentation behandelten Prozeduren. Inklusive rekursive geschachtelte Liste, welche Aufrufe jeweils in den aufgerufenen Prozeduren erfolgen.


- Bereitstellung einer einfachen GUI / HMI, um Input- und Output Pfade zu parametrisieren







### TODOS:


- Chore: Aufräumen des Quellcodes

- Refactor: ggfs. modifizieren von write_content

- BUGFIX: modul 1 aufrufe

- "Help... " Button in GUI, in dem Erklärungen stehen! --> BESSER, universeller, einfacher und weniger duplizierend: ERstelle eine README.md im Repository, und beim Klick auf "help-btn" wird diese Datei in eine HTML umgewandelt und im Browser angezeigt...




### AUSBLICK für später und in schön:


- Zusatzmöglichkeit in GUI einen benutzerdefinierten Text einzugeben (Prio sehr gering!!). Dieser würde dann in einre eigenen Section angezeigt werden.

- Index an der Seite wie eine NavBar zum einzelnd scrollen

- Ermöglichung von Berücksichtigung weiterer Module innerhalb der Dokumentation
    
    - z. B. 2 VBA-Module innerhalb eines Projektes, wobei Prozeduren von Modul1  andere Prozeduren aus Modul2 aufrufen.

        - Erstmal nur als Verweis  (Mögl. Ansatz included = "Modul1.*" ohne rekursive Auflistung derer Aufrufe... oder eben mit... bestenfalls auch das parametrisierbar)

- Dokumentation von weiteren PRogrammiersprachen

    - OK --> VBA
    - Nächste Prio: C++ / Arduino
    - Letzte Prio: Python (v.a. für den Ablaufsequence sehr hilfreich, für den rest gibt es pdoc...)








# =============================================================================
#### Wichtige Aufrufreihenfolge der Methode innerhalb dieses Python-Scriptes zur Erstellung der Dokumentation der Aufrufreihenfolge der zu dokumentierenden VBA-Prozeduren: ####
# =============================================================================

Es werden zunächst alle Prozeduren komplett analysiert, erst danach werden wiederum alle Prozeduren komplett dokumentiert. Für beide Vorgänge erfolgt dies in einer Methode auf Objektebene, wobei diese jeweilige MEthode in beiden Fällen aus einer Klassenmethode aufgerufen wird, in der über die einzelnen Prozedur-Objekte innerhalb dieser Klasse iteriert wird:

- analyse_call_sequence(cls)
    - analyse_calling_sequence_in_one_proc(self)
- prepare_all_call_sequences_docs(cls)
    - prepare_single_call_sequence_docs(cls)

(hierfür wäre das entwickelte Tool  übrigens eine tolle Anwendung gewesen, sofern sie später auch mal Python-Syntax dokumentieren könnte :-) )





# =============================================================================
#### Hinweise zur Anwendung und Benutzung: ####
# =============================================================================

- To generate a docstring from the VBA-Source make sure that the text to shown is located directly below the declaration line of the procedure. The text is considered completed with the first following line in the code which is not an entire comment line. Empty lines that are to be included must also be labelled as comments.

- Durch das Script wird eine MD-Datei (Markdown) erzeugt, die anschließend über die Library markdown sofort in eine HTML umgewandelt wird, sodass nach Abschluss des Scriptes 2 Dateien erstellt wurden. Durch unterschiedliche Interpretationen im Rahmen der Konvertierung unterscheidet sich die Darstellung der so generierten HTML-Datei allerdings, wenn sie über VSCode Extension gesondert konvertiert wird. Die über VSCode generierte Datei ist übersichtlicher und schöner. Das sollte also am Ende nochmals gesondert erfolgen.







# =============================================================================
#### Unwichtige Nebensächlichkeiten: Code-Analyse Zusammenfassung: ####
# =============================================================================

In der Version vom 2024-01-11 - 00:18:43:
    Angaben jeweils: [Zeilen @ code_documenter.py] + [Zeilen @ gui] = [Summe]
    - Gesamtanzahl der Zeilen: 408 (100%)+2201 (100%)=2609 (100%)
    - davon Leerzeilen: 168 (41,1764705882353%)+1071 (48,6597001363017%)=1239 (47,489459563051%)
    - davon Einzelkommentarzeilen: 20 (4,90196078431373%)+244 (11,0858700590641%)=264 (10,1188194710617%)
    - davon Blockkommentarzeilen: 85 (20,8333333333333%)+374 (16,9922762380736%)=459 (17,5929474894596%)

    ==> Summe aller Kommentarzeilen: 105 (25,7352941176471%)+618 (28,0781462971377%)=723 (27,7117669605213%)
    ==> Code-relevante Zeilen: 135 (33,0882352941176%)+512 (23,2621535665607%)=647 (24,7987734764277%)

-----------------------------------------------    

In der Version vom 2024-01-07 - 23:37:04:
    - Gesamtanzahl der Zeilen: 2164 (100%)
    - davon Leerzeilen: 1091 (50%)
    - davon Einzelkommentarzeilen: 226 (10%)
    - davon Blockkommentarzeilen: 364 (17%)

    ==> Summe aller Kommentarzeilen 590 (27%)
    ==> Code-relevante Zeilen: 483 (22%)

-----------------------------------------------

In der Version vom 2024-01-07 - 15:26:03:
    - Gesamtanzahl der Zeilen: 2771 (100%)
    - davon Leerzeilen: 1408 (51%)
    - davon Einzelkommentarzeilen: 278 (10%)
    - davon Blockkommentarzeilen: 550 (20%)

    ==> Summe aller Kommentarzeilen 828 (30%)
    ==> Code-relevante Zeilen: 535 (19%)

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





# =============================================================================
#### GLOBALS: ####
# =============================================================================

DEBUG = 1

# CONVERT_TO_HTML = 1





def db(*args):
    """
    Schleust zum printen durch - nur zum Debuggen
    """
    if DEBUG == False:
        return
    
    print("__DEBUG_PRINT__\n")
    for _ in args:
        print(_)






def inspect_get_current_line_number():
    """
    Gibt die Zeilennummer des Codes zurück - geeignet zum Debuggen!

    Returns:
        int: number of line  in code file.
    """

    # Die Stack-Informationen abrufen
    stack = inspect.stack()
    
    # Die Informationen für die aktuelle Funktion/Frame erhalten
    aktueller_frame = stack[1]
    
    # Die Zeilennummer extrahieren
    zeilennummer = aktueller_frame[2]
    

    return zeilennummer






# =============================================================================
#### Workaround: Verwendung von MetaClasses: Metaklassen für Direktaufrufe / implizite Aufrufe einer Classmethod direkt nach Implementierung einer Klasse ####
# =============================================================================


# =============================================================================
#### # WIEDERHOLFUNKTION-KANDIDAT!!! ####
# =============================================================================
class AutoCallMeta(type):
    """
    Sofern diese Klasse als metaclass für eine andere Klasse verwendet wird, wird die unten aufgeführte Klassenmethode direkt nach Definition ohne ein zusätzlichen, expliziten Aufruf automatisch aufgerufen.

    Dazu muss nur der Name der aufzurufenden Methode in der Klassenvariable 'class_name_to_call_implicit' parametrisiert werden.
    """    

    # ACHTUNG: Bis zur Implementierung der GUI wurde es wie unten gemacht!
    # Flexibler ist es aber, die Methode MetaClass.initialize_class tatsächlich selbst explizit aufzurufen an einer Stelle, wo bereits alle relevanten Attibute mit Werten belegt wurden. 
    # Daher wurde mit Implementierung der GUI der Ablauf dementsprechend geändert  und diese Superklasse auch nicht mehr beerbt von der Klasse MetaClass
    # pass


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
    ### ACHTUNG: Bis zur Implementierung der GUI wurde geerbt von der metaclass=AutoCallMeta.
    Diese Superklasse hat dafür gesortt, dass automatisch direkt nach  der Deklaration der Klasse MetaData ihre Methode MetaClass.initialize_class implizit aufgerufen wurde. Es ist aber flexibler tatsächlich selbst explizit aufzurufen an einer Stelle, wo bereits alle relevanten Attibute mit Werten belegt wurden. 
    Daher wurde mit Implementierung der GUI der Ablauf dementsprechend geändert  und diese Superklasse auch nicht mehr beerbt von der Klasse MetaClass.



    In dieser Klasse werden hauptsächlich Daten gespeichert, die später als Art MetaDaten angesehen werden können.
    Der Parent-class / metaclass sorgt dafür, dass direkt nach Implementierung dieser (gewöhnlichen) Klasse eine in der metaclass parametrisierte Methode aufgerufen wird.
    Somit ist kein expliziter Aufruf der Klassenmethode MetaData.initialize_class erforderlich, da dies über die metaclass erledigt wird.

    # TODO: Die Klasse ist noch nicht fertig.

    Zu den hierin gespeicherten Daten gehören z. B.:

    - Dieses Dokumentations-Tool-Script
        - Versionsinformationen, basierend auf dem letzten Git-Commit
        - TODO: Versionsnummer dieses Scriptes...
    - Das zu dokumentierende Modul
        - Dateipfad
        - Dateiname
        - Datum der letzten Speicherung / Änderung
    - Aktuellen Zeitstempel zur Angabe des Zeitpunktes der Dokumentation
      
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
        # Den Zeitstempel der letzten Änderung der Datei auslesen
        timestamp = os.path.getmtime(file_path)
        
        # Den Zeitstempel in ein lesbares Datum umwandeln
        last_modified_datetime = datetime.fromtimestamp(timestamp)
        
        # Das Datum im gewünschten Format ausgeben
        formatted_date = last_modified_datetime.strftime('%Y-%m-%d %H:%M')
        
        return formatted_date





    @classmethod
    def extract_date_of_change(cls):
        """
        Liesst das letzte Aenderungsdatum der Input-Datei aus und speichert diese in der Klassenvariable cls.input_file_date_of_change
        """

        date_of_change = cls.get_last_modified_timestamp(cls.__input_path)
        cls.input_file__date_of_change = date_of_change






    @classmethod
    def set_input_path(cls, input_path:str=None):
        """
        Prüft den optional übergebenen Pfad, ob er existiert und dort eine .bas Datei vorliegt.
        Ist dies nicht der Fall, oder wird kein Pfad übergeben, wird per Input ein neuer Pfad abgefragt. 
        Bis zu einem gültigen Pfad wird die Methode rekursiv aufgerufen.
        Bislang gibt es noch keine Möglichkeit für den Nutzer, die Eingabe abzubrechen (außer Programmabbruch...)

        Args:
            input_path (str, optional): Dateipfad zur .bas-Datei, die dokumenteirt werden soll - als Foreward-Slash und ohne Anführungszeichen. Defaults to None.

        """

        if input_path != None:
            if os.path.isfile(input_path):
                if input_path.endswith(".bas"):
                    cls.__input_path = input_path
                    cls.extract_date_of_change()
                    return
                
        neuer_input = input("!!! FEHLER !!! Die Angegebene Datei ist keine .bas Datei! Bitte einen gueltigen Pfad zur entsprechenden Datei eingeben (Foreward-Slashes! ohne Anfuehrungszeichen)\n> Ihre Eingabe: ")

        cls.set_input_path(neuer_input)


    @classmethod
    def set_output_dir(cls, output_dir:str=None):
        """
        Prüft den optional übergebenen Pfad, ob er existiert und ob dies ein Verzeichnis ist
        Ist dies nicht der Fall, oder wird kein Pfad übergeben, wird per Input ein neuer Pfad abgefragt. 
        Bis zu einem gültigen Pfad wird die Methode rekursiv aufgerufen.
        Bislang gibt es noch keine Möglichkeit für den Nutzer, die Eingabe abzubrechen (außer Programmabbruch...)

        Args:
            output_dir (str, optional): Dateipfad zum Ordner, in dem die generierten Output-Dateien exportiert werden sollen - als Foreward-Slash und ohne Anführungszeichen. Defaults to None.

        """

        if output_dir != None:
            if os.path.isdir(output_dir):
                cls.__output_dir = output_dir
                return
        neuer_input = input("!!! FEHLER !!! Der  angegebene Pfad ist kein gueltiges Verzeichnis. Bitte einen gueltigen Pfad fuer den Export der Output-Dateien  eingeben (Foreward-Slashes! ohne Anfuehrungszeichen)\n> Ihre Eingabe: ")

        cls.set_output_dir(neuer_input)




    @classmethod
    def extract_git_info(cls):
        info:dict = gitinfo.get_git_info()

        # Übernehme die Infos aus git in die Klasse:
        for key, value in info.items():


            attr_name =  f"documenter_version__{key}"

            setattr(cls, attr_name, value)





    @classmethod
    def save_current_timestamp(cls):

        # Aktuelles Datum und Uhrzeit
        current_datetime = datetime.now()

        # Wandele das Datum und die Uhrzeit in einen Zeitstempel um
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
        Gibt den gesamten Pfad fuer die neu zu generierende Output-Datei zurueck, inkl. Dateierweiterung.

        Args:
            extension (str, optional): Dateiendung der Output-Dtaei. Defaults to ".md". Modifizierbar z. B. zu .html oder .txt

        Returns:
            str : Dateipfad
        """

        path =  os.path.join(cls.__output_dir, cls.__output_filename + extension)
        return path
    
    




    @classmethod
    def initialize_class(cls):
        """
        Diese Methode wird implizit direkt nach Implementierung dieser Klasse aufgerufen.
        Somit muss sie nicht mehr von außen aufgerufen werden.
        Sie initialisiert alle Attribute mit ihren WErten.
        """



        # # HACK: path for Source-vba-code
        # input_file_path = "input_data/beispiel_modul_rekursiv.bas"
        # input_file_path = "input_data/beispiel_modul_bauer+liebherr.bas"
        # input_file_path = "input_data/beispiel_modul2.bas"
        # input_file_path = "input_data/beispiel_modul1.bas"
        # input_file_path = "input_data/beispiel_modul.bas"

        # # HACK: Falls ohne GUI-Daten gearbeitet wird (debugging:)
        # # Wird normalerweise von aussen aufgerufen und mit Daten aus der GUI gefuellt. Nur durchlaufen, wenn es zum debuggen ist!
        # cls.set_input_path(input_file_path)

        # # HACK: Falls ohne GUI-Daten gearbeitet wird (debugging:)
        # # Wird normalerweise von aussen aufgerufen und mit Daten aus der GUI gefuellt
        # output_dir = "output_data"
        # cls.set_output_dir(output_dir)



        # cls.git_info_to_str()
        cls.extract_git_info()
        # cls.count_of_commits = cls.get_count_of_commits()


        cls.save_current_timestamp()


        cls.make_output_filename()








class Procedure():
    """
    Allgemeine Klasse zur Bereitstellung von Inhalten, die fuer alle Prozeduren (Subs und Functions) erforderlich sind. Dazu gehoert:
        
        - Definition des Dateipfades fuer Template, in die der extrahierte Text übernommen wird
        
        - Flag-Variable, ob nach Beginn oder Ende der Prozedur gesucht wird
        
        - Regex-Muster als String für den Beginn und das Ende einer Prozedur - wobei innerhalb dieses Strings der Platzhalter für die Prozedurart in den Subklassen noch ersetzt werden muss. Ebenfalls
    """

    # TEMPLATE = "templates/prozedur.md"
    # HACK: für weiterentwicklung bzgl abruffolge:
    TEMPLATE = "templates/prozedur_dev.md"

    
    # TODO: Mache  dies via GUI als Option parametrisierbar:
    __print_final_calling_sequence_message = True
    # print_final_calling_sequence_message = False


    search_for_begin = True # initialer wErt




    # TODO: Vars besser privatisieren:

    # Das folgenden Regex-Muster berücksichtigt nicht das Auskommentieren dieser Zeile
    regex_begin_pattern = r""".*     # Start mit beliebigen Zeichen
                        (?:Private|Public|Friend)?
                        (?:PLACEHOLDER_PROCEDURE_TYPE)    # Beinhaltet das KEyword
                        \s+        # mind. 1 bis n Leerzeichen
                        (\w+)        # mind. 1 bis n Wortzeichen
                        \(         # Geöffnete Klammer
                        """



    # _BUGFIX: Besonders beim Beispielmmodul 1 wird zu früh beendet! Neue Regex > V. 0.1.2
    regex_end_pattern = r""".*?     # Start mit beliebigen Zeichen
                        (?:End)    # Beinhaltet das KEyword
                        \s+        # mind. 1 bis n Leerzeichen
                        (?:PLACEHOLDER_PROCEDURE_TYPE)    # Beinhaltet das KEyword
                        \s+         # mind. 1 bis n Leerzeichen
                        .*"""
    


    '''
    # ALT: Bis Version 0.1.1:
    regex_end_pattern = r""".*?     # Start mit beliebigen Zeichen
                        (?:End)    # Beinhaltet das KEyword
                        \s+        # mind. 1 bis n Leerzeichen
                        (?:PLACEHOLDER_PROCEDURE_TYPE)?    # Beinhaltet das KEyword
                        .*"""
    '''






    # Zum späteren Prüfen bzgl. Auskommentierung extra Regex:
    __regex_ausschlus_kommentar_pattern = r"""'    # KommentarApostroph
                                        .?       # Beliebiges Zeichen
                                        """

    # Hier ist keine weitere Konkretisierung durch Subklassen erforderlich, daher direkt kompiliert:
    regex_ausschluss_kommentar = re.compile(__regex_ausschlus_kommentar_pattern, re.VERBOSE | re.IGNORECASE)

    @classmethod
    def set_print_final_calling_sequence_message(cls, value:bool):
        """
        Wert uebergeben fuer die Einstellung, ob nach jedem Abschluss jeder Prozedur in der Calling Sequence ein Abschlusssatz ergänzt werden soll.
        """
        if type(value) == bool:
            cls.__print_final_calling_sequence_message = value
            return
        
        raise Exception("Ungueltiger Parameter fuer __print_final_calling_sequence_message gewaehlt!")   



    @classmethod
    def initialize_input_code(cls, input_path:str):
        """
        Liesst den zu analysierenden und zu dokumentierenden Input-Quellcode ein und speichert ihn innerhalb dder Superklasse als immer verfuegbare Liste einzelner Zeileninhalte ab.

        ### TODO: Später sollte hier auch noch eine einfache GUI erstellt werden zur Auswahl!
        ggf. sollte diese GUI aber losgelöst von dieser Klasse sein (Wiederholfunktion!). Daher als Kapselung mit übergebenen input-path!
        """

        with open(input_path, "r") as file:
            cls.raw_source_code = file.readlines()






    @classmethod
    def detail_analyse_procedures(cls):
        """
        Detail-Auswertung aller zuvor identifizierten Prozeduren (unabhängig ihrer Art).
        Die einzelnen Komponenten werden dabei in Objekte der Subklassen gespeichert
        """

        for procedure_type_cls in [Sub, Function]:

            for line_begin, line_end in procedure_type_cls.matches_line_ixs:

                # _BUG(?FIX?): sofern nach ein End {procedure_type} in der allerletzten Zeile des Quellcodes steht undn KEINE LEERZEILE FOLGT, isst line_end = None . Das raised einen TypeError: unsupported operand type(s) for +: 'NoneType' and 'int'
                lines = cls.raw_source_code[line_begin:line_end + 1]

                sub = procedure_type_cls(tuple(lines))



    @classmethod
    # def identify_procedures(cls, list_of_code_lines:list[str]):
    def identify_procedures(cls):
        """
        Identifizieren  aller Prozeduren (Methoden und Funktione) auf Klassenebene und Speichern der Zeilennummern der Deklarations- und End-Zeilen innerhalb der jeweiligen Klassenvariablen.

        Durchsucht werden die einzelnen Strings der als Liste übergebenen Zeileninhalte des Quellcodes.

        Args:
            raws (list[str]): Liste mit einem Eintrag pro Zeile des Quellcodes

        """

        for ix, text in enumerate(cls.raw_source_code):

            # db(text)

            if cls.search_for_begin:

                if Sub.check_and_get_match(text, ix):
                    current_procedure_type_cls = Sub
                    continue

                if Function.check_and_get_match(text, ix):
                    current_procedure_type_cls = Function


            else:
                # Suche nach dem entsprechend passenden Ende
                current_procedure_type_cls.search_end_of_procedure(text, ix)



            













    @classmethod
    def check_and_get_match(cls, text:str, line_no:int) -> bool:
        """
        Prüft eine einzelne Zeile eines Codes (den übergebenen text), ob es sich um ein Match einer Deklarationszeile handelt, und auch, ob diese NICHT-auskommentiert ist. Sofern dies der Fall ist, wird dieses MAtch in die Liste der MAtsches aufgenommen.

        Args:
            text (str): Zu prüfender Text ausschnitt
            line_no (int) : Zeilennummer der Liste an gesamt zu untersuchenden Zeilen

        Returns:
            bool: True falls das match aufgenommen wurde
        """

        if cls.regex_begin.match(text):
            if not cls.regex_ausschluss_kommentar.match(text):
                # Dann aufnehmen in die Liste der MAtches! inkl. Platzhalter für Endzeilennummer:
                cls.matches_line_ixs.append([line_no, None])
                Procedure.search_for_begin = False

                return True
        
        return False






    @classmethod
    def search_end_of_procedure(cls, text:str, line_no:int) -> bool:
        """
        Prüft eine einzelne Zeile eines Codes (den übergebenen text), ob es sich um ein Match einer Deklarations-END-zeile handelt, und auch, ob diese NICHT-auskommentiert ist. Sofern dies der Fall ist, wird dieses MAtch in die Liste der MAtsches aufgenommen.

        Args:
            text (str): Zu prüfender Text ausschnitt
            line_no (int) : Zeilennummer der Liste an gesamt zu untersuchenden Zeilen

        Returns:
            bool: True falls das match aufgenommen wurde
        """

        
        # BUG: Durch die Regex wird aber auch etwas wie End sub  gefunden obwohl es um end function geht! Daher wurde der PRogrammablauf im main angepasst, um sicherzustellen, dass immer nach der richtigen Klassen-Beendigung gesucht wird... Falls Langeweile: Später mal schauen warum...
        if cls.regex_end.match(text):
            if not cls.regex_ausschluss_kommentar.match(text):
                # Dann aufnehmen in die Liste der MAtches!
                cls.matches_line_ixs[-1][-1] = line_no

                Procedure.search_for_begin = True

                return True
        
        return False


    


    @classmethod
    def initialize_page_top_text(cls):
        """
        Initialisiert den Text für den Markdown-Text des  Headers der Seite und speichert es in der Klassenvariable cls.head. 
        Zugriff erfolgt über die Superklasse Procedure. 

        
        """

        page_top_text = cls.__read_template("templates/sec_head.md")

        # Gesamtanzahl an verfügbaren prozeduren je Art einsetzen:
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
        Initialisiert den Text für den Markdown-Text des  Table of content und speichert es in der Klassenvariable cls.toc. Zugriff erfolgt über die Superklasse Procedure. 
        Innerhalb des resultierenden Textes sind weiterhin jeweils 1 Platzhalter fuer jede Prozedur-Art vorhanden. Diese werden spaeter gesondert ersetzt.
        """

        toc = cls.__read_template("templates/sec_toc.md")

        # Gesamtanzahl an verfügbaren prozeduren je Art einsetzen:
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
        
        Generiert die Einzelnen Eintraege aller Prozeduren der uebergebenen subklasse und fuegt diese in die Klassenvariable cls.toc ein, indem Platzhalter verwendet werden.
        
        Zugriff auf Superklassen-Ebene


        Args:
            subklasse (Sub | Function): Klasse, die aktuell gelistet werden soll
        """


        toc = cls.toc # shortcut

        keyword_type = subklasse.KEYWORD_TYPE.upper()
        platzhalter = f"@PLACEHOLDER_ENTRIES_{keyword_type}S@"
        

        for (prozedur_obj, prozedur_name, prozedur_initialisierungszeile) in subklasse.all_procedures_final:

            # Iterieren ueber jede Instanz:

            # TODO: Vorgehen wenn es KEINE instanz gibt???! - sollte kein Problem sein, dadurch dass nach jedem Einfügen immer wieder der Ausgangszustand bzgl. des Platzhalters wiederhergestellt wird un dieser am Ende gelöscht wird??!

            # Aufbau des Markdown.-Codes fuer diesen TOC-Eintrag:
            new_entry = cls.get_markdown_for_code_line_of_call_entry(prozedur_name, prozedur_initialisierungszeile, line_text="")

            # Append the placeholder to have this flag for insert further entries:
            new_entry = new_entry + "\n  " + platzhalter
            # ACHTUNG: Zusatz-Zeile vorweg sind relevant fuer korrekte Einrueckung im MArkdown!


            toc = toc.replace(platzhalter, new_entry)


        # After all procedures of 1 type: Delete the leaving placeholder in the toc for this type of procedures.
        toc = toc.replace(platzhalter, "")

        # store into class-variable:
        cls.toc = toc

        













    @classmethod
    def analyse_references(cls):
        """
        Zugriff erfolgt über die Superklasse Procedure. Innerhalb der Methode wird ueber jedes Objekt jeder Subklasse iterriert.

        Zur Vereinfachung wird eine neue Klassenvariable auf Superklassen-Ebene erstellt, in der alle Elemente der beiden gleichnamigen Listen der einzelnen Subklassen Sub und Function.all_procedures_final enthalten sind. Diese Liste ist sortiert nach aufsteigender Zeilennummer.

        Es wird dann nach Referenzierungen (Aufrufen) jeder Einzelnen Prozedur im gesamten Quelltext gesucht. Bei einem gefundenen Match wird weiter identifiziert, innerhalb welcher uebergeordneten Prozedur dieser Aufruf erfolgte. 
        Fuer jedes Objekt wird eine Objektvariable (Liste) references erstellt, die initial leer ist und bei gefundenen Matches jeweils mit einem Tuple der folgenden Form erweitert wird: (line_no, bezeichnung_uebergeordnete_prozedur, zeilentext_des_aufrufes).

        """

        # ERstellung der gemeinsamen Liste auf Superklassen-Ebene:
        all_procedures = Sub.all_procedures_final + Function.all_procedures_final


        # Sortierung der neuen Liste: Basierend auf der Zeilen-Nummer der Deklarationszeile
        cls.all_procedures_final = sorted(all_procedures, key=lambda stored_tuple: stored_tuple[2])


        for (prozedur_obj, prozedur_name, prozedur_initialisierungszeile) in cls.all_procedures_final:

            # Erstellen einer Objektvariable: Liste fuer alle noch zu findenen Referenzierungen:
            prozedur_obj.references = []
            
       
            
            # ACHTUNG: Die folgende regex matcht zwar NICHT etwas wie ' Prozedurname, aber ohne das Leerzeichen ('Prozedurname) wird fälschlicherweise immer noch gematcht! Daher nach dem matschen einfach nochmal prüfen, ob VOR dem ProcName noch ein ' steht... dann entnehme das Matsch wieder (ähnliches Vorgehensweise wie bei der Identifizierung von Procedures-Deklarationen...). 
            # regex_ansatz = re.compile(r"(?<!('.*))(\W{___PROC_NAME___}\W)".format(___PROC_NAME___=prozedur_name))
            # ACHTUNG: Durch diese Einschränkung brauche ich EH diese Kontrolle, daher wird jetzt doch eine regex verwendet, die das Auskommentieren GAR NICHT berücksichtigt, somit werden auch alle auskommentierten Aufrufe gematcht... Dafür wird stattdessen direkt mit berücksichtigt, dass Deklarationszeilen NICHT gematcht werden sollen! Eine einzelne Abfrage hierzu ist also nicht mehr erforderlich!
            # regex = re.compile(r"(?<!(Function|Sub))\W{___PROC_NAME___}\W".format(___PROC_NAME___=prozedur_name))
            # regex = re.compile(r"\b(!Function|Sub\b)\W{___PROC_NAME___}\W".format(___PROC_NAME___=prozedur_name))
            
            
            
            regex_call = re.compile(r"\bcall\s+{___PROC_NAME___}\b".format(___PROC_NAME___=prozedur_name))



            # Folgende macht leider einen erroer .  im "normalen " regex geht es, in python.re nicht:
            # regex_brackets = re.compile(r"(?<!((Function|Sub).*))({___PROC_NAME___}\()".format(___PROC_NAME___=prozedur_name))

            # ACHTUNG: es wird hiermit DOCH NICHT ABGEFANGEN, dass die Deklarationszeile nicht gematcht wird! die muss also später noch raus kommen!
            regex_brackets = re.compile(r"\b{___PROC_NAME___}\(".format(___PROC_NAME___=prozedur_name))



            # Durchsuche den GESAMTEN QUELLTEXT nach einem Aufruf dieser Prozedur:
            for line_no, line_text in enumerate(cls.raw_source_code, 1):



                # db("gesuchte Methode:", prozedur_name, "Line-Nr = ", line_no, "  :  ", line_text)


                if line_no == prozedur_initialisierungszeile:
                    # Dann ist dies die  Deklarationszeile der Funktion, zu der die Aufrufe gefunden werden sollen - also ignorieren!
                    continue



                if not (match:=regex_brackets.search(line_text)):
                    # Kein Aufruf mittels ...PROZEDURNAME(...
                    if not (match:=regex_call.search(line_text)):
                        # Kein Aufruf mittels ...Call PROZEDURNAME...

                        continue
                
                
                # TODO: Backreference auf die Gruppe mit dem Sub-Namen (kenie backref nötig, da es ja EH nur nach dieser gesucht wird!) notwendig für aufruf abfolge
                ziel_prozedur_name = prozedur_name
                # db(match)
                    


                # An dieser Stelle ist match IMMER != None:
                # db(match)

                # Check wheather there is a comment-symbol before the procedure name:
                __prozedur_name_start_pos = match.span()[0]
                if (cls.regex_ausschluss_kommentar.search(line_text[0:__prozedur_name_start_pos])):
                    # Dann kommentar
                    continue


                # Hier ist eine tatsächlicher Aufruf gefunden worden, der jetzt dokumentiert werden muss:
                # JEtzt muss der Aufruf dokumentiert werdn!
                


                # ACHTUNG:  WICHTIG:  Der folgende Schritt / die Logik und Anwendung war mir neu  -  nochmal recherchieren! sowohl mit dem filtern / nach eigenem Key, als auch, dass das Ergebnis nicht DIE EINZELNE ZAHL  ist, sondern tatsächlich direkt das TUPLE, IN DEM DIESE EINZELNE ZAHL GESPEICHERT IST!  Extrem praktich!!!!

                
                # Logik zum Finden der aufrufenden Prozedur: Die Deklarationszeilennummer muss < line_no sein. Die nächste dran ist die Zeilennummer der Deklarationszeile
                aufrufende_prozedur_tuple_relevant = [tpl for tpl in cls.all_procedures_final if tpl[2] < line_no]

                # Das relevante Tuple der Liste cls.all_procedures_final wird in der selben Form gespeichert in der folgenden Variable:
                aufrufende_prozedur_tuple = min(aufrufende_prozedur_tuple_relevant, key=lambda x_tuple: abs(x_tuple[2] - line_no))


                #Erweiterung um den Prozedurnamen, der aufgerufen wird:
                prozedur_obj.references.append(
                    (line_no, aufrufende_prozedur_tuple[1], line_text, ziel_prozedur_name)
                )



               



    @classmethod
    def finalize(cls, output_file_path:str):
        """
        ### TODO: Ob wirklich alles hier in die Klasse gehört ist fraglich! 
        
        Zugriff erfolgt über die Superklasse Procedure. Innerhalb der MEthode wird wiederum in Schleifen itteriert ueber die Subklassen.
        
        Bereitet die finale Ausgabe vor und ruft weitere MEthoden zum Schreiben dieser Ausgabe / der Markdown-Datei auf.
        Zu dieser Vorbereitung gehört:
         
            - Sortierung der einzelnen Prozeduren innerhalb der verschiedenen Prozedur-Arten gemaess der alphabetischen Reihenfolge ihrer Bezeichner (Namen)
            
            - Konfigurieren / Initilalisierung der Zwischenüberschriften / Header-Texte unterhalb der einzelnen Section-Überschriften (gespeichert in Klassenvariable cls.header)


        Args:
            output_file_path (str): Dateipfad der zu erstellenden Markdown-Datei
        """

        cls.initialize_page_top_text()


        # =============================================================================
        #### # Extrahiere den Modulweiten Docstring (sofern es einen gibt): ####
        # =============================================================================
        
        # Beginne erst ab Zeile 1 wegen der Codierungszeichen!
        __modul_docstring = cls.identify_docstring(cls.raw_source_code[1:], trim_empty_rows=True, return_alternativ_text=True)
        __template_docstring = cls.__read_template("templates/sec_modulinfos.md")
        # Ersetze Platzhalter in der Template durch gefundenen Docstring:
        cls.modul_docstring = __template_docstring.replace("@PLACEHOLDER_MODUL_DOCSTRING@", __modul_docstring)

        

        # Vorbereitung des TOCS:
        cls.initialize_toc()

        for procedure_type_cls in [Sub, Function]: 

            # Initialisieren des Headers unterhalb der Section-Überschrift  (aus gesonderter Template) und in Klassenvariable speichern:
            content = cls.__read_template(procedure_type_cls.TEMPLATE_SECTION_HEAD)

            # _BUGFIX: cls.header = content
            # _BUGFIX: Durch das Überschreiben von Sub durch Function bevor die Template überschrieben wird, gibt es 2x die Überschrift "Functions"
            procedure_type_cls.header = content



            # Instanzen der Prozeduren  Alphabetisch  sortieren nach den Namen! und in Klassenvariable speichern:
            procedure_type_cls.sort_procedures_by_names()


            # Ergänze einen einzigen Eintrag für eine Methode! und haenge sie dem TOC an
            # WICHTIG: Auf Superklassen-Ebene!
            cls.generate_toc_entries(procedure_type_cls)

        
        # Durchsuche gesamten Quelltext nach allen Referenzierungen für jeweils alle gefundenen Prozeduren und speichere sie in den jeweiligen Objekten der einzelnen Prozeduren:
        cls.analyse_references()


        # TODO:  Analysiere jede Prozedur und speichere jeden weiteren Aufruf einer weiteren Prozedur in dem Prozedur-Objekt. Geschrieben wird es erst später, da dann rekursiv auf alle Calling-Sequences zugegriffen werden kann
        cls.analyse_call_sequences()

        cls.prepare_all_call_sequence_docs()



        # Aufruf der Methode zum tatsächlichen Schreiben der Textdatei:
        cls.write_to_file(output_file_path)







    @classmethod
    def get_procedure_obj_by_name(cls, procedure_name:str) -> 'Procedure':
        """
        Gibt das Prozedur-Objetk mit dem übergebenen Namen zurück
        Das Objekt ist aus der Subklasse Sub oder Function.

        Args:
            procedure_name (str): Name der gesuchten Procedur
        """

        # Procedure.all_procedures_final
        
        for (prozedur_obj, prozedur_name, prozedur_initialisierungszeile) in Procedure.all_procedures_final:
            if prozedur_name == procedure_name:
                return prozedur_obj

        db("nicht gefunden!!!")    
        return None


    @staticmethod
    def indent_str(text:str, count_of_indents:int=0) -> str:
        """
        Stellt jeder Zeile entsprechende Indention-Symbols vorweg.

        Args:
            text (str): Text
            count_of_indents (int, optional): Level of indention. Defaults to 0.

        Returns:
            str: Text eingerückt.
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
        Generiert einen String, der in eine MArkdown-Datei mit entsprechendem Syntax eingefügt werden kann.
        Der String enthält den Namen einer Prozedur, eine relevante Zeilennummer und den relevanten Text der Code-Zeile.

        Innerhalb der Methode wird der String so aufgebaut, dass der Prozedurname später verlinkt ist zu dieser Prozedur.

        Die Methode kann sowohl für calling_sequences (Aufrufe), als auch für references (übergeordnete / aufrufende Prozeduren) verwendet werden und stellt sicher, dass das Ausgabeformat immer identisch ist.

        # ACHTUNG: Weitere Darstellungmöglchkeiten siehe OLD_self !
        
        Args:
            proc_name (str): Name der Prozedur
            line_no (str): Relevante Zeilennummer im Code (als str)
            line_text (str): Textzeile im Code


        Returns:
            str: String im Markdown-Syntax, der eine interaktive Übersicht über die Parameter in leserlicher Form gibt.
        """

        # proc_name:str, line_no:str, line_text:str
    

        # replacer_placeholder_reference = "\n" * 3 + f"- [```{target_procedure_name}```](#{target_procedure_name}) <small> : [Zeile {line_no_reference}] : ```{line_code}``` </small>".replace(line_code, line_code.rstrip("\n")) + "\n"
        
        if line_text == "":
            __optional_code_line = ""
        else:
            __optional_code_line = f": ```{line_text}```"


        markdown_entry = f"- [```{proc_name}```](#{proc_name}) : <small>  [Zeile {line_no}] {__optional_code_line} </small>"

        markdown_entry = markdown_entry.replace(line_text, line_text.rstrip("\n")) 
        
        markdown_entry = markdown_entry + "\n" * 2


        return markdown_entry





    def prepare_single_call_sequence_docs(self, level=0):
        """
        ### TODOC: Siehe prepare_all_call_sequence_docs... (gelöscht)
        ### CHANGELOG:

        - 2024-01-07 - 04:16:47 Beginn neuafbau
        - 2024-01-07 - 23:07:07 läuft jetzt

        
        Generiert die vollständige Dokumentation der Aufrufsequenzen in den einzelnen Objekten und speichert sie im Attribut obj.calling_sequences_doc. Nach vervollständigung wird obj.calling_sequences_state = True  gesetzt.

        Zugriff erfolgt auf Objekt-Ebene.
        """
        db(f"name der ZU ANALYSIERENDEN prozedur : {self.name}")

        if self.calling_sequences_state:
            db("diese proz ist fertig!")

            end_text_per_procedure = ""

            text_to_return = self.calling_sequences_doc +  end_text_per_procedure
            
            
            # ACHTUNG: keine indentions einfuegen!
            text_to_return = self.indent_str(text_to_return, 0)


            return text_to_return


        # Initialisiert den Initialtext aus der template, falls es noch keinen Text gibt:
        if not self.calling_sequences_doc:
            self.calling_sequences_doc = ""


        for (line_no_reference, uebergeordnetes_sub, line_code, target_procedure_name) in self.calling_sequences:
            
            replacer_placeholder_reference =  self.get_markdown_for_code_line_of_call_entry(target_procedure_name, line_no_reference, line_code)

            # ACHTUNG: keine indentions einfuegen!
            replacer_placeholder_reference = self.indent_str(replacer_placeholder_reference, 0)
            
            self.calling_sequences_doc = self.calling_sequences_doc  + replacer_placeholder_reference 

            level = level + 1

            # Rekursive Aufrufe für untergeordnete Calling-sequenzabfolge:
            if target_procedure_name == self.name:
                # Abbruch gegen Endlosrekusrion:
                further_calls_doc = "- <small> *... recursivly calls itself under certain conditions ...* </small> \n\n"

            else:

                # get the object from target_procedure_name:
                db(f"neuer ziel-name =  {target_procedure_name}")
                target_procedure_obj = self.get_procedure_obj_by_name(target_procedure_name)

                # Rekursiv: den nächsten Platzhalter mit der nächsten Dokumentation des aufgerufenen calls füllen:
                # further_calls_doc = target_procedure_obj.prepare_single_call_sequence_docs(level + 1)
                further_calls_doc = target_procedure_obj.prepare_single_call_sequence_docs(level + 0)
                # TEST 2024-01-07 - 22:25:28: hier nur level + 0 ?! scheint zu fixxen!!!! :-) JA @ 2024-01-07 - 23:09:23


            
            # ACHTUNG: HIER JA:  indentions einfuegen!
            self.calling_sequences_doc = self.calling_sequences_doc  + self.indent_str(further_calls_doc, count_of_indents=level)

            # Verringerung des Levels:
            level = level - 1


        level = level + 1


        abschlusstext = "\n- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>"
        if Procedure.__print_final_calling_sequence_message == False:
            abschlusstext = ""
        
        # ACHTUNG: keine indentions einfuegen!
        abschlusstext = self.indent_str(abschlusstext, 0)


        self.calling_sequences_doc = self.calling_sequences_doc + abschlusstext

        self.calling_sequences_state = True
        text_to_return = self.calling_sequences_doc

        return text_to_return
























    @classmethod
    def prepare_all_call_sequence_docs(cls):
        """
        Generiert die vollständige Dokumentation der Aufrufsequenzen in den einzelnen Objekten und speichert sie im Attribut obj.calling_sequences_doc. Nach vervollständigung wird obj.calling_sequences_state = True  gesetzt.

        Zugriff erfolgt auf Superklassen-Ebene, wobei intern die Instanzen der Klasse referenziert werden

        # NEIN: Nach der eigentlichem Itterieren und Dokumentieren der Abruffolge werden innerhalb dieser Klassenmethode für das Objekt die Platzhalter für die Übersicht ersetzt (Anzahl und Einleitungssatz).


        """


        for prozedur in cls.all_procedures_final:
            prozedur_obj:Procedure = prozedur[0]
            db(prozedur_obj.name)

            prozedur_obj.prepare_single_call_sequence_docs(level=0)

            db(len(prozedur_obj.calling_sequences), prozedur_obj.calling_sequences)

            db("weiter")
            

        db("alle fertig prepared.")

























    @classmethod
    def analyse_call_sequences(cls):
        """
        Itteriert über sämtliche zu dokumentierenden Prozedur-Objekte und ruft für jedes Objekt die Methode analyse_calling_sequences_in_one_proc auf, in der sämtzliche Aufruf-Abfolgen innerhalb dieser einen Prozedur ermittelt werden.

        Zugriff erfolgt auf Superklassen-Ebene, wobei intern die Instanzen der Klasse referenziert werden und diese durchgeschleust werden zur objekt method!
        """
        cls.all_procedures_final

        # Anwendung der bereits bestehenden Objekt-Methode fuer jede Prozedur:
        for prozedur in cls.all_procedures_final:

            prozedur_obj:Procedure = prozedur[0]

            prozedur_obj.analyse_calling_sequences_in_one_proc()



            # Sortiere generierte Liste aufsteigend nach den Aufrufzeilen:
            prozedur_obj.calling_sequences = sorted(prozedur_obj.calling_sequences, key=lambda stored_tuple: stored_tuple[0])


        db("Alle Prozeduren analysiert, noch nicht dokumentiert!")




	








    
    def analyse_calling_sequences_in_one_proc(self) -> None:
        """
        Ermittelt die Aufruf-Abfolge fremder (aber durch dieses Documenter-Tool dokumentierte) Prozeduren innerhalb dieser Prozedur self.
        Das Ergebnis wird gespeichert in der Objektvariable self.calling_sequences (Liste aus Tuples, wobei jedes Tuple einen Aufruf darstellt.)
        Innerhalb dieser MEthode wird NICHT die self.documentation verändert, da dies erst erfolgen sollte, nachdem diese Methode analyse_calling_sequences_in_one_proc für sämtliche zu dokumentierenden VBA-Prozeduren durchlaufen wurde.

        ### NOTE:
        Diese Methode basiert auf der zuvor verwendeten Methode OLD_generate_calling_entries, die den Startpunkt der Entwicklung darstellte und zusätzlich zu dem Speichern der Erkenntnisse diese auch direkt dokumentierte. 
        """

        self.calling_sequences = []

        # Suche für alle vorhandenen Prozeduren jeweils ihre Referenzen heraus, und prüfe, ob eine Referenz zu der jetzt gerade untersuchten Prozedur (self.name) führt. Wenn ja, ist dies ein Aufruf, der der calling_sequences liste angehangen werden soll.
        for (procedure_obj, prozedur_name, procedure_declaration_line_no) in Procedure.all_procedures_final:
            
            # Filtern der liste, so dass uebergeordnetes_sub == name
            db("Suche Aufrufabfolge für Prozedur >> {}\nAktuell: check ob die folgende Prozedur aufgerufen wird: >> {}".format(self.name, prozedur_name)) # ist IMMER GLEICH in einem einzigen Aufruf!

            
            relevante_references = [_ for _ in procedure_obj.references if _[1] == self.name]
            
            
            db("----> Anzahl Relevanter Referenzierungen: {}".format(len(relevante_references)))



            # =============================================================================
            #### Füge die relevanten Referencen als Aufrufe innerhalb der zu analysierenden Prozedur hinzu: ####
            # =============================================================================
            
            for reference in relevante_references:
                self.calling_sequences.append(reference)

            # NOTE: Zugriff später bei Bedarf:
            # for (line_no_reference, uebergeordnetes_sub, code, ziel_prozedur_name) in self.calling_sequence:


        db("Calling Sequences for procedure {}:\n{}".format(self.name, self.calling_sequences))










    # @classmethod # eigentlich bezieht sich das auf ein einzelnes Objekt!!!
    def generate_reference_entries(self) -> None:
    # def generate_reference_entries(cls, procedure_obj) -> None:
        """
        Generiert von dem Objekt der Subklasse Function oder Sub die Dokumentation der Referenzierungen.
        Speichert das Ergebnis in der bereits existierenden Objektvariable procedure_obj.documentation
        """



        # TODO: Struktur...? das gehört zu den references... ist eig. statisch, müsste nicht jedes mal beim funktionsaufruf neu gemacht werden, aber so ist es zusammen...
        _PLACEHOLDER_REFERENCE = "@PLACEHOLDER_PROCEDURE_REFERENCES_ENTRY@" 



        # shortcut:
        doc:str = self.documentation 

        # Alle Referenzierungen sind in der Objektvariablen  procedure_obj.references gespeichert:
        count_of_references = len(self.references)

        # ERsetzen des Platzhalters für die Anzahl der Referenzierungen in der bisherigen Dokumentation:
        doc = doc.replace("@PLACEHOLDER_PROCEDURE_COUNT_OF_REFERENCES@", str(count_of_references))
        

        # Initialisierung und  Parametrisierung  des Einleitungssatzes:
        einleitungssatz = "Kein Aufruf gefunden." # default

        if count_of_references > 0:

            einleitungssatz = "Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:"

            
            # Iterrieren ueber jede Referenzierung, um diese zu dokumentieren:
            for (line_no, calling_procedure_name, line_code, target_procedure_name) in self.references:

                # Zusammenbau des Ersatzwertes für den Platzhalter inkl. Anhängen des Platzhalters für weitere Ersetzungen:
                
                # Um MArkdown nicht zu zerschiessen muss der letzte Zeilenumbruch des line_codes entfernt werrden:
                
                # OBSOLET: wird in Methode get_markdown_for_code_line_of_call_entry erledigt
                # line_code:str = line_code.rstrip("\n")

                # replacer_placeholder_reference = f"* [```{calling_procedure_name}```](#{calling_procedure_name}) : <small>  [Zeile {line_no}] : ```{line_code}``` </small>"


                replacer_placeholder_reference = self.get_markdown_for_code_line_of_call_entry(calling_procedure_name, line_no, line_code)



                # # AUSBLICK: weitere collapse details : funktioniert technisch, allerdings steht das collapsable immer in neuer Zeile und daher wird es groß - vielleicht später in schön machen...
                # __replacer_placeholder_reference = f"*   [```{calling_procedure_name}```](#{calling_procedure_name})  <details> <summary>: <small>Zeile {line_no}</small> </summary> ```{line_code}``` </details>"



                replacer_placeholder_reference = replacer_placeholder_reference  + _PLACEHOLDER_REFERENCE
                
                # Ersetzen:
                doc = doc.replace(_PLACEHOLDER_REFERENCE, replacer_placeholder_reference)


        # Loeschen des verbliebenen Platzhalters zum Einfuegen einzelner Referenzen:
        doc = doc.replace(_PLACEHOLDER_REFERENCE, "")



        # Einsetzen des Einleitungssatzes:
        doc = doc.replace("@PLACEHOLDER_PROCEDURE_REFERENCES_INTRODUCTION@", einleitungssatz)


        # shortcut / resubstitution:
        self.documentation = doc
        







    @classmethod
    def write_to_file(cls, output_file_path:str):
        """
        Schreibt die Dokumentation aller Prozeduren (aller Art) in die als Dateipfad übergebene Zieldatei(pfad).
        """

        # initialize_toc

        with open(output_file_path, "w", ) as file:

            ## TODO: Dies gehört eigentlich nicht mehr in die Klasse Procedure, aber aktuell trotzdem hierher, um alles zusammen zu haben. Ggf.s später refactoring...
            # Initialisiere die Seite mit Titel und organisatorischen Hinweisen:
            content = cls.__read_template("templates/sec_head.md")

            file.write(cls.page_top_text)

            #Einfügen des Index / TOCs /  Inhaltsverzeichnis:
            file.write(cls.toc)


            # schreiben der Modulinfos / Docstrings:
            file.write(cls.modul_docstring)


            # # TODO: Struktur...? das gehört zu den references...
            # _PLACEHOLDER_REFERENCE = "@PLACEHOLDER_PROCEDURE_REFERENCES_ENTRY@" 


            for procedure_type_cls in [Sub, Function]: # Reihenfolge wichtig fuer die Reihenfolge der Dokumentierten Sections!


                # Überschrift der Section einfügen (aus gesonderter Template):
                file.write(procedure_type_cls.header)



                # Dokumentieren aller einzelnen Prozeduren:
                for procedur_infos in procedure_type_cls.all_procedures_final:
                # for (procedure_obj, procedure_name, procedure_line) in procedure_type_cls.all_procedures_final:

                    procedure_obj:Procedure = procedur_infos[0]





                    procedure_obj.generate_reference_entries()

                    db("name der aktuell zu dokumentierenden prozedur: {}".format(procedure_obj.name))
                    # Ermittlung + Dokumentation Calling Sequences:
                    # procedure_obj.generate_calling_entries()
                    if not procedure_obj.calling_sequences_state:
                        db(f"PROBLEM! mit {procedure_obj.name}")
                        db(f"PROBLEM! mit {procedure_obj.name}")
                    # BUG: Problem!
                    procedure_obj.documentation = procedure_obj.documentation.replace("@PLACEHOLDER_PROCEDURE_CALLING_SEQUENCES_BLOCK@", procedure_obj.calling_sequences_doc)



                    # =============================================================================
                    #### Übersichts-Parameter eintragen für das Objekt: / die Prozedur: ####
                    # =============================================================================
                    
                    # Initialisierung und  Parametrisierung  des Einleitungssatzes:

                    if procedure_obj.calling_sequences:

                        einleitungssatz = "Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:"

                    else:

                        einleitungssatz = "Keine weiteren Aufrufe zu hier dokumentierten Prozeduren gefunden." # default


                    # Einsetzen des Einleitungssatzes:
                    procedure_obj.documentation = procedure_obj.documentation.replace("@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@", einleitungssatz)


                    # Einsetzen der ÜBersichtsanzahl an Aufrufen:
                    procedure_obj.documentation = procedure_obj.documentation.replace("@PLACEHOLDER_PROCEDURE_COUNT_OF_ABRUFFOLGE@", str(len(procedure_obj.calling_sequences)))







                    # Schreibe die modifizierte Prozedur-Dokumentation in die Datei:
                    file.write(procedure_obj.documentation)


                # Ausgabe der Zusammenfassung pro Section:
                print("Es wurden {} {}s identifiziert und dokumentiert.".format(len(procedure_type_cls.instances), procedure_type_cls.KEYWORD_TYPE))



            ## TODO: Dies gehört eigentlich nicht mehr in die Klasse Procedure, aber aktuell trotzdem hierher, um alles zusammen zu haben. Ggf.s später refactoring...
            # Finalisiere die Seite mit Schlusbemerkungen:
            content = cls.__read_template("templates/sec_tail.md")
            # HACK: Nur fuer ENtwicklungsstadium!
            content = content.replace("@PLACEHOLDER_TIMESTAMP_NOW@", MetaData.date_of_process)
            content = content.replace("@PLACEHOLDER_DOC_PYTHON@", __doc__.replace("\n", "<br>"))
            content = content.replace("@PLACEHOLDER_DOCUMENTER_VERSION__AUTHOR@", MetaData.documenter_version__author.rstrip(" <_>"))
            content = content.replace("@PLACEHOLDER_DOCUMENTER_VERSION__COMMIT@", MetaData.documenter_version__commit)
            content = content.replace("@PLACEHOLDER_DOCUMENTER_VERSION__DATE@", MetaData.documenter_version__author_date)

            file.write(content)

        print("MD-Datei wurde aus MD-Datei generiert: {}".format(MetaData.get_output_path(".md")))


    @classmethod
    def sort_procedures_by_names(cls):
        """
        Füllen der Klassenvariable der **SUBKLASSE** cls.all_procedures_final mit einem Tupel pro Prozedur, in dem sowohl das einzelne Objekt, als auch seine Bezeichnung und die Zeilennummer der Deklaratrion enthalten ist in der Form [(object:Procedure, object.name:str, object.line_begin)]. 
        Diese Liste wird nach Fertigstellung sortiert basierend auf den alphabetischen Bezeichnern, sodass der Zusammenhang zwischen den Objekten und den Namen weiterhin gegeben ist, gleichzeitig aber die Objekte in der alphabetischen Reihenfolge ihrer Namen dokumentiert werden können.

        """

        all_procedures = [] 


        for procedure_obj in cls.instances:
            
            all_procedures.append((procedure_obj, procedure_obj.name, procedure_obj.line_begin))



        # Nach dem Füllen: Sortieren basierend auf den Bezeichner-Namen:
        cls.all_procedures_final = sorted(all_procedures, key=lambda stored_tuple: stored_tuple[1])



        
            













    @staticmethod
    def __read_template(template_file_path) -> str:
        """
        Ließt die uebergebene Template-MArkdown Datei ein und gibt den Inhalt als String zurück.
        
        Args:
            template (str) : Dateipfad zur Template

        Return:
            str : Textinhalt der Template
        """
        with open(template_file_path, "r") as file:
            content = file.read()
            return content





    def read_template(self, template="einzelprozedur") -> str:
        """
        Ließt die relevante Template-MArkdown Datei ein und gibt den Inhalt als String zurück. 
        Sofern es sich um die Default-Template für eine beliebige Prozedur handelt
        wird außerdem auch das Attribut self.documentation mit diesem Text initialisiert.

        Args:
            template (str) : Dateipfad zur Template - Ausnahme: Falls die in der Klasse gespeicherte 
                            Standardtemplate self.TEMPLATE fuer eine Prozedur verwendet werden soll, 
                            dann muss das Stichwort "einzelprozedur" übergeben werden. 
                            Dies ist default der Fall --> keine Angabe erforderlich.

        Return:
            str : Inhalt der Template
        """


        if template == "einzelprozedur":
        
            # Auslesen und Speichern in Objektvariable
            self.documentation = self.__read_template(self.TEMPLATE)
            return self.documentation


        # Bei NICHT-Standard-Prozedur-Template: Auslesen dieser Template und Rückgabe des Textinhaltes:
        content = self.__read_template(template)
        return content





    def generate_documentation(self):
        """
        Dokumentiert die gefundenen Parameter unter Nutzung einer Template-Markdown-Textdatei und speichert den zu schreibenden Text im Attribut self.documentation
        """

        self.read_template()

        placeholder_replacer = {
            "@PLACEHOLDER_PROCEDURE_TYPE@" : self.KEYWORD_TYPE,
            "@PLACEHOLDER_PROCEDURE_MODIFIER@" : self.modifier,
            "@PLACEHOLDER_PROCEDURE_NAME@" : self.name,
            "@PLACEHOLDER_PROCEDURE_LINE_BEGIN@" : self.line_begin,
            "@PLACEHOLDER_PROCEDURE_DOCSTRING@" : self.docstring,
            "@PLACEHOLDER_PROCEDURE_SOURCE_CODE@" : self.source_code,

            # # TODO:  Kommen die References extra??! Wahrscheinlich schon stand jetzt (Version> 0.1.3)
            # "@PLACEHOLDER_PROCEDURE_REFERENCES@" : self.references,
            # "@PLACEHOLDER_PROCEDURE_COUNT_OF_REFERENCES@" : self.count_of_references,
        }


        for placeholder, replacer in placeholder_replacer.items():
            self.documentation = self.documentation.replace(placeholder, str(replacer))













    def extract_source_code(self):
        """
        Ließt den source_code aus und speichert diesen im Attribut self.source_code
        """

        source_code = ""
        for line in self.lines:
            source_code = source_code + line

        self.source_code = source_code





    def extract_modifier(self):
        """
        Ließt den Modifier aus und speichert diesen im Attribut self.modifier
        """
        # Wiederverwendung der Deklarationszeilen-Regex:
        match = self.regex_begin.match(self.lines[0])

        # Entnehme den String bis vor den NAmen:
        pos_name = match.string.find(self.name + "(")
        potentieller_modifier = match.string[0:pos_name]

        # Identifiziere hieraus den modifier:
        
        # # Eigentlich sollte es schon auf Grundlage der eigentlichen regex gehen, funktioniert aber nicht! Es wird immer was leres zurückgegeben, daher, neue regex! (Bei Zeit mal schauen warum!)
        # # Extraktion der Gruppe mit dem Name: (eigentlicher Ansatz)
        # db(match.groups())
        # modifier = match.group(1)
        
        
        regex_modifier = re.compile(r"((?:Private|Public|Friend)?)")
        match = regex_modifier.match(potentieller_modifier)
        modifier = match.group(1)


        if modifier == "":
            modifier = "Public"

        # db(modifier)

        self.modifier = modifier




    def extract_name(self):
        """
        Ließt den Bezeichnungsnamen aus und speichert diesen im Attribut self.name
        """
        # Wiederverwendung der Deklarationszeilen-Regex:
        
        
        match = self.regex_begin.match(self.lines[0])

        # db(match.groups())
        # Extraktion der Gruppe mit dem Name:
        name = match.group(1)

        # db(name)

        self.name = name



    def extract_line_numbers(self):
        """
        Ließt die relevanten Start- und End-Zeilennummern aus und speichert diesen im Attribut self.line_begin und self.line_end
        """
        # Hole Index der Instanzenliste, dieser ist gleich dem Index der matches_line_ixs Liste:
        ix = self.instances.index(self) 

        match_lines_ix = self.matches_line_ixs[ix]
        
        self.line_begin = match_lines_ix[0] + 1
        self.line_end = match_lines_ix[1] + 1




    @staticmethod
    def identify_docstring(text_lines:tuple[str], trim_empty_rows=False, return_alternativ_text=False) -> str:
        """
        ### ACHTUNG: Entstanden aus der Methode extract_docstring(self), die auf Objektebene anzuwenden ist. 
        Um die gleiche Logik aber auch fuer den modulweiten Docstring nutzen zu koennen, wird dieses Konstrukt eingefuehrt.


        Ließt den ersten Block-KOmmentar / docstring aus und gibt ihn zurueck.
        Als Docstring wird jede Kommentarzeile gewertet, die DIREKT UND OHNE VORHERIGE LEERZEEILE UNTERHALB DER ERSTEN UEBERGEBENEN ZEILE steht. 
        Sobald eine Leerzeile folgt, wird der Docstring als beendet angesehen.

        Beispiel:
                ' Dies ist ein Kommentar direkt unter der Deklarationszeile. Somit wird es als Docstring gewertet.
                ' Dies auch
                '
                ' Da die vorherige Zeile AUCH einen Kommentar (einen leeren) enthält, ist dies hier immer noch Bestandteil des Docstrings.

                ' Vor dieser Zeile war ein NICHT-Kommentar, daher gehört das hier nicht mehr zum Docstring
                MsgBox("Das gehört zum Programm")
            End Sub

            
        Args:
            text_lines (tuple[str]): Tuple mit den einzelnen Textzeilen, die zu der relevaten  Prozedur gehoeren.
            trim_empty_rows (bool) : False, sofern der Algorithmus direkt bei einer Leerzeile abgebrochen werden soll ( = DEFAULT), 
                                        oder True, falls der Algorithmus bei vorangehenden Leerzeilen solange weiterlaufen soll, bis dass KEINE Leerzeile mehr vorliegt oder eine andere Abbruchsbedingung erreicht ist.
            return_alternativ_text (bool | str) : Default = False. Dann wird beim Nicht-Finden ein leerer String zurückgegeben. 
                                                    Bei True wird in solchen Fällen der in der MEthode definierte Alternativ-ERsatztext zurückgegeben. 
                                                    Bei Übergabe eines Strings wird in diesen Fällen dieser String zurückgegeben.
        Returns:
            str : Docstring / Blockkommentar
        """


        docstring  = ""


        for line in text_lines:
            # line:str
            content = line.lstrip(" ")

            if(content[0]) == "'":
               # gehört zum Docstring
                docstring = docstring + content.lstrip("'")
                continue # nächste Zeile


            if trim_empty_rows == False:

                break

            # ELSE: Dann trimme die leere Reihe
            if docstring != "":
                # Dann gibt es bereits einen Docstring -> Abbruch
                break


            if re.match(r"^\s*\n?$", content):
                # Dann besteht die Zeile nur aus Whitespaces --> Trimmen!
                # Rekursiver Aufruf dieser Methode mit jeweils einer vordersten Zeile weniger!
                docstring = Procedure.identify_docstring(text_lines[1:], trim_empty_rows=True, return_alternativ_text=False) 
                # TODO: Fehlervermeidung / Grenzfall:  IndexOutOfRange (Ende angekommen, --> docstring vorhanden --> Wird Error geben!
                break



        # =============================================================================
        #### #  Alternativtext bei leeren Docstrings: ####
        # =============================================================================
        
        if return_alternativ_text == True:

            if docstring == "":
                docstring = "*No information availible. For more information expand source code.*" #  Die Sternchen bewirken im MArkdown ein Kursivdruck

        elif isinstance(return_alternativ_text, str):
            if docstring == "":
                docstring = return_alternativ_text


        return docstring
  




    def extract_docstring(self):
        """
        # CHANGELOG: 2023-12-30 - 03:17:04 Bis V. 0.0.5 war diese Funktion 'alleinherschend'. 
        Erst danach  wurde die Durchschleusung zur static method identify_docstring eingeführt,
        um damit auch modulweite Docstrings finden zu koennen.


        Ließt den docstring aus und speichert diesen im Attribut self.docstring
        Sofern kein Docstring im Code identifiziert wurde, wird eine entsprechende Info in den Text geschrieben.

        Als Docstring wird jede Kommentarzeile gewertet, die DIREKT UND OHNE VORHERIGE LEERZEEILE UNTERHALB DER DEKLARIERUNGSZEILE der Prozedur steht. Sobald eine Leerzeile folgt, wird der Docstring als beendet angesehen.

        Beispiel:
            Private Sub beispielProgramm() ' Deklarationszeile
                ' Dies ist ein Kommentar direkt unter der Deklarationszeile. Somit wird es als Docstring gewertet.
                ' Dies auch
                '
                ' Da die vorherige Zeile AUCH einen Kommentar (einen leeren) enthält, ist dies hier immer noch Bestandteil des Docstrings.

                ' Vor dieser Zeile war ein NICHT-Kommentar, daher gehört das hier nicht mehr zum Docstring
                MsgBox("Das gehört zum Programm")
            End Sub
        """


        # Docstring ueber generalisierte Methode herausfiltern:
        docstring = self.identify_docstring(self.lines[1:], return_alternativ_text=True)

        # Speichern im Objekt:
        self.docstring = docstring





    '''

    def extract_references(self):
        """
        ### TODO: Wird aktuell (Version > 0.1.2) ganz wo anders erledigt, ist dort noch nicht optimal - aber um Redundanzen und Verewchslungen vorzubeugen, wird diese Methode erst mal platt gemacht und nicht mehr aufgerufen!

        Später wäre es shcön...
        
        ### ALT: 
        Durchsucht den gesamten Quelltext nach Referenzierungen (Aufrufen) dieser Prozedur und speichert diese im Attribut self.references
        # TODO ALLES
        """
        # self.references = " # TODO ... self.references"
        # self.count_of_references = " # TODO ... self.count_of_references"
        pass
    '''


    def __init__(self, text_lines:tuple[str]) -> None:
        """
        Erstellt ein Prozedur-Objekt, wodurch die einzelnen Komponenten des Codes gesucht und gespeichert werden.
        Uebergeben muss ein beliebig grosses Tuple, wobei jedes Element davon den String einre einzelnen Zeile dieser Prozedur enthaelt.

        Args:
            text_lines (tuple[str]): Tuple mit den einzelnen Textzeilen, die zu der relevaten  Prozedur gehoeren.
        """

        # TODO: Ist  dies notwendig???! Eigentlich nicht! Dies ist schon die superklasse!
        super().__init__()

        # Merken der instanc zum späteren Iterieren: Klassenzuweisung dynamisch!
        type(self).instances.append(self)
        
        # Speichern aller Textzeilen:
        self.lines = text_lines


        # =============================================================================
        #### # HErausfiltern einzelner KOmponenten: ####
        # =============================================================================
        
        self.extract_line_numbers()
        self.extract_source_code()
        self.extract_name()
        self.extract_modifier()
        
        self.extract_docstring()



        # =============================================================================
        #### Zusammenfassen und Schreiben der Dokumentation: ####
        # =============================================================================
        # TODOC: -... was passiert darin?
        self.calling_sequences_doc = None
        self.calling_sequences_state = False
        # self.doc_of_calling_sequences = (calling_sequences_doc, calling_sequences_state)



        self.generate_documentation()




















class Sub(Procedure):
    """
    Subklasse für eine VBA-Sub-Prozedur.
    U.a. wird hier auch der von der Superklasse vor-initialisierte Regex-Ausdruck konkretisiert und kompiliert, der im VBA-Syntax für den Beginn und das Ende der Prozedurart erforderlich ist.

    Die meisten Methoden sind in der übergeordneten Superklasse gelagert, da sie vom Ablauf fuer VBA-Subs und VBA-Methoden identisch sind.

    """

    matches_line_ixs = []
    instances = []

    all_procedures_final = [] # Liste wird erst nach Identifizierung aller Prozeduren erstellt. Sie dient der Sortierung  der Instanzen basierend auf der alphabetischen Reihenfolge ihrer Prozedur-Bezeichnungen

    
    KEYWORD_TYPE = "Sub" 

    TEMPLATE_SECTION_HEAD = "templates/sec_subs.md"
    

    # Konkretisierung und Kompilierung der Regex-Ausdrücke:

    regex_begin = re.compile(Procedure.regex_begin_pattern.replace("PLACEHOLDER_PROCEDURE_TYPE", KEYWORD_TYPE), re.VERBOSE | re.IGNORECASE)


    regex_end = re.compile(Procedure.regex_end_pattern.replace("PLACEHOLDER_PROCEDURE_TYPE", KEYWORD_TYPE), re.VERBOSE | re.IGNORECASE)















class Function(Procedure):
    """
    Subklasse für eine VBA-Sub-Prozedur.
    U.a. wird hier auch der von der Superklasse vor-initialisierte Regex-Ausdruck konkretisiert und kompiliert, der im VBA-Syntax für den Beginn und das Ende der Prozedurart erforderlich ist.

    Die meisten Methoden sind in der übergeordneten Superklasse gelagert, da sie vom Ablauf fuer VBA-Subs und VBA-Methoden identisch sind.

    """

    matches_line_ixs = []
    instances = []
    
    all_procedures_final = [] # Liste wird erst nach Identifizierung aller Prozeduren erstellt. Sie dient der Sortierung  der Instanzen basierend auf der alphabetischen Reihenfolge ihrer Prozedur-Bezeichnungen
    
    KEYWORD_TYPE = "Function" 
    
    TEMPLATE_SECTION_HEAD = "templates/sec_functions.md"
    
    
    # Konkretisierung und Kompilierung der Regex-Ausdrücke:

    regex_begin = re.compile(Procedure.regex_begin_pattern.replace("PLACEHOLDER_PROCEDURE_TYPE", KEYWORD_TYPE), re.VERBOSE | re.IGNORECASE)


    regex_end = re.compile(Procedure.regex_end_pattern.replace("PLACEHOLDER_PROCEDURE_TYPE", KEYWORD_TYPE), re.VERBOSE | re.IGNORECASE)









# =============================================================================
#### FUNCTIONS: ####
# =============================================================================






def load_parameter(documenter_gui_obj:DocumenterGui):
    """
    Lädt alle relevanten Attribute (eingestellten Parameter) der GUI in die Klasse MetaData (bzw. Procedure) und initialisiert anschließend die Klasse MetaData.

    Args:
        documenter_gui_obj (DocumenterGui): Objekt mit Parametern fuer die zu erstellende Dokumentation.
    """
    
    MetaData.set_output_dir(documenter_gui_obj.output_dir)
    MetaData.set_input_path(documenter_gui_obj.input_file)
    MetaData.set_convert_to_html(documenter_gui_obj.convert_checked)
    MetaData.set_user_defined_additional_text(documenter_gui_obj.optinal_user_defined_text)

    Procedure.set_print_final_calling_sequence_message(documenter_gui_obj.show_message)


    # Expliziter Aufurf  der Initialisierungsmethode der MetaData.
    # Die Klassenattribute werden vorher bereits durch die Aufruf einzelner Setter mit den Daten aus der GUI belegt.
    MetaData.initialize_class()




def convert_markdown_to_html():
    """
    Wandelt die generierte .md Markdown-Datei in eine HTML-Datei um.
    Die Formatierung ist durch die verwendete Bibliothek allerdings bei weitem nicht so sinnvoll und uebersichtlich,
    wie wenn die .md Datei im Anschluss manuell durch VSCode umgewandelt wird (Extension Markdown all in one)
    """

    if MetaData.get_convert_to_html():
        markdown.markdownFromFile(
            input=MetaData.get_output_path(extension=".md"),
            output=MetaData.get_output_path(extension=".html"), 
            encoding="utf8"
        )

        print("HTML-Datei wurde aus MD-Datei generiert: {}".format(MetaData.get_output_path(".html")))
        




















# =============================================================================
#### MAIN: 
# =============================================================================
def main():
    """
    Hauptprogramm. Steuert den Gesamt-Ablauf des Scripts. 
    Die meisten Methoden sind innerhalb der Superklasse Procedure definiert.
    """

    gui = DocumenterGui()
    
    if not gui.get_is_ready():
        db("KEIN Start! -> Abbruch")
        return

    db("Dann starte Dokumentation von aussen mit den Parametern durch Zugriff auf die Objektvariablen")


    # PArameter des GUI-Objektes in MetaData-Class speichern und Initialisieren der MetaDate-Klasse
    load_parameter(gui)


    # initialize the class including reading the input file:
    Procedure.initialize_input_code(MetaData.get_input_path())


    #  Identifizieren und Speichern der Deklarationszeilen von Prozeduren:
    Procedure.identify_procedures()


    # Detail-Analyse und Speicherung einzelner Bestandteile in Objekten:
    Procedure.detail_analyse_procedures()


    Procedure.finalize(MetaData.get_output_path())


    convert_markdown_to_html()

    gui.show_closing_window()











if __name__ == '__main__':

    print("START....")


    # DocumenterGui.DEBUG = DEBUG

    main()

    print(".... ENDE.")
