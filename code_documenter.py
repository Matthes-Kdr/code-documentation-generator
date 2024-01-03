# -*- coding: utf-8 -*-
'''
Created on: Fri, 2023-12-29 (00:45:39)


@author: Matthias Kader


Für Ziel und Ablauf des Scriptes siehe MArkdown im Verzeichnis ../Tests/Programmablauf.html




### Fertig implementiert:

• Implementierung Inhaltsverzeichnis / Index

• Gesamtlayout inkl. Titel, Zwischenüberschriften für einzelne Sections

• Einbindung vom Programmkopf-Docstring

• Implementierung von References-Durchsuchungen

• Implementierung eines Exportes zu HTML


### TODO: Größere TODOS:

• Einbindung organisatorischer Daten bzgl. des zu dokumentierenden Codes und des verwendeten Skripts zum Dokumentieren




### AUSBLICK für später und in schön:

• Index an der Seite wie eine NavBar zum einzelnd scrollen



• Call Sequenz / Calling Sequence:
Schön (Ausblick) wäre auch ein weiterer Unterpunkt pro Prozedur, in der die Aufrufabfolge hervorgeht.
Idee ist etwas wie die Aufrufebenen-Auflistung beim Noten-Converter-Programm, d.h. ausgehend von einer Prozedur soll eine Liste stehen der Aufrufe von weiteren Prozeduren die aufgerufen werden (und die in diesem Dokument auch dokumentiert werden... also keine Builtins o.ä.). Im Idealfall kann jeder Punkt dieser Liste wiederum erweitert/expanded werden, darin ist dann wiederum die Liste von DIESER AUFGERUFENEN Funktion drin usw... Rekursiv. Jede Methode, die einmal so dokumentiert wurde kann weiter verwendet werden per Direktzugriff....


'''














from datetime import datetime
import os
import re

import markdown

import subprocess
import gitinfo






# =============================================================================
#### GLOBALS: ####
# =============================================================================

DEBUG = 1

CONVERT_TO_HTML = 1





def db(*args):
    """
    Schleust zum printen durch - nur zum Debuggen
    """
    if DEBUG == False:
        return
    
    print("__DEBUG_PRINT__\n")
    for _ in args:
        print(_)





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






class MetaData(metaclass=AutoCallMeta):
    """
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



    # OBSOLET: nicht erforderlich - zur Vorbeugung von Verwendchslung daher auskommenteirt:
    # @classmethod
    # def get_output_dir(cls) -> str:
    #     return cls.__output_dir
    
    


    @staticmethod
    def get_count_of_commits():
        """
        ### ACHTUNG: Nicht funktional!!!
        """


        cmd = "git rev-list --count HEAD"
        return_ =  subprocess.run(cmd)
        
        db(return_)

        db(str(return_).split(" ")[0])
        # return str(return_).split(" ")[0]
        return str(return_)



    @classmethod
    def extract_git_info(cls):
        info:dict = gitinfo.get_git_info()

        # Übernehme die Infos aus git in die Klasse:
        for key, value in info.items():

            # if not hasattr(cls, key):
            #     # Wenn nicht, erstelle es und weise den Wert zu
            #     setattr(cls, key, value)

            attr_name =  f"documenter_version__{key}"

            setattr(cls, attr_name, value)
            db(getattr(cls, attr_name))




            # bold = "**" if key in keys_bold else ""
            # indent_message = "** \n> **" if key == "message" else " "

            # cls.gitinfo = cls.gitinfo + "- " + bold + key + ":" + indent_message + value + bold  + "\n"





    # @classmethod
    # def git_info_to_str(cls):

    #     info:dict = gitinfo.get_git_info()
    #     cls.gitinfo = "## Infos zum Script, welches für die Erstellung dieser Dokumentation verwendet wurde:\n\n"

    #     keys_bold = ["commit", "message", "refs", "author_date"]


    #     for key, value in info.items():
    #         bold = "**" if key in keys_bold else ""
    #         indent_message = "** \n> **" if key == "message" else " "

    #         cls.gitinfo = cls.gitinfo + "- " + bold + key + ":" + indent_message + value + bold  + "\n"





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




        # HACK: path for Source-vba-code
        input_file_path = "input_data/beispiel_modul.bas"
        # input_file_path = "input_data/beispiel_modul1.bas"
        # input_file_path = "input_data/beispiel_modul1.bas"
        
        cls.set_input_path(input_file_path)






        # HACK: DIR FOR OAUTPUT
        # Schreibe die Datei erst in eine Markdown-Datei:
        output_dir = "output_data"
        cls.set_output_dir(output_dir)










        # cls.git_info_to_str()
        cls.extract_git_info()
        # cls.count_of_commits = cls.get_count_of_commits()




        cls.save_current_timestamp()


        cls.make_output_filename()








class Procedure():
    """
    Allgemeine Klasse zur Bereitstellung von Inhalten, die fuer alle Prozeduren (Subs und Functions) erforderlich sind. Dazu gehoert:
        
        • Definition des Dateipfades fuer Template, in die der extrahierte Text übernommen wird
        
        • Flag-Variable, ob nach Beginn oder Ende der Prozedur gesucht wird
        
        • Regex-Muster als String für den Beginn und das Ende einer Prozedur - wobei innerhalb dieses Strings der Platzhalter für die Prozedurart in den Subklassen noch ersetzt werden muss. Ebenfalls
    """

    # TEMPLATE = "templates/prozedur.md"
    # HACK: für weiterentwicklung bzgl abruffolge:
    TEMPLATE = "templates/prozedur_dev.md"


    search_for_begin = True # initialer wErt




    # TODO: Das wäre eig. schöner wenn sie private vars sind...

    # Das folgenden Regex-Muster berücksichtigt nicht das Auskommentieren dieser Zeile
    regex_begin_pattern = r""".*     # Start mit beliebigen Zeichen
                        (?:Private|Public|Friend)?
                        (?:PLACEHOLDER_PROCEDURE_TYPE)    # Beinhaltet das KEyword
                        \s+        # mind. 1 bis n Leerzeichen
                        (\w+)        # mind. 1 bis n Wortzeichen
                        \(         # Geöffnete Klammer
                        """



    # BUGFIX: Besonders beim Beispielmmodul 1 wird zu früh beendet! Neue Regex > V. 0.1.2
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

                # db(line_begin, line_end)
                
                # TODO: Bug ausmisten!
                # BUG: sofern nach ein End {procedure_type} in der allerletzten Zeile des Quellcodes steht undn KEINE LEERZEILE FOLGT, isst line_end = None . Das raised einen TypeError: unsupported operand type(s) for +: 'NoneType' and 'int'
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
            new_entry = "* [```{temp_prozedur_name}```](#{temp_prozedur_name}) <small>(Zeile {temp_prozedur_zeile})</small>".format(temp_prozedur_name=prozedur_name, temp_prozedur_zeile=prozedur_initialisierungszeile)

            # Append the placeholder to have this flag for insert further entries:
            new_entry = new_entry + "\n  " + platzhalter
            # ACHTUNG: Leerzeichen vorweg sind relevant fuer korrekte Einrueckung im MArkdown!


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
        Fuer jedes Objekt wird eine Objektvariable (Liste) references erstellt, die initial leer ist und bei gefundenen Matches jeweils mit einem Tuple der folgenden Form erweitert wird: (line_no, bezeichnung_uebergeordnete_prozedur, zeilentext_des_aufrufes)

        """

        # ERstellung der gemeinsamen Liste auf Superklassen-Ebene:
        all_procedures = Sub.all_procedures_final + Function.all_procedures_final


        # Sortierung der neuen Liste: Basierend auf der Zeilen-Nummer der Deklarationszeile
        cls.all_procedures_final = sorted(all_procedures, key=lambda stored_tuple: stored_tuple[2])


        for (prozedur_obj, prozedur_name, prozedur_initialisierungszeile) in cls.all_procedures_final:

            # Erstellen einer Objektvariable: Liste fuer alle noch zu findenen Referenzierungen:
            prozedur_obj.references = []
            
            """
            # DEBUG:
            # BEISPIELERGEBNIS:
                # prozedur_obj = <__main__.Function object at 0x0000018180C603A0>
                # prozedur_name = "addieren"
                # prozedur_initialisierungszeile = 24

            """
            
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
                
                
                # TODO: Backreference auf die Gruppe mit dem Sub-Namen (kenie backref nötig, da es ja EH nur nach dieser gesucht wird!)
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


                # prozedur_obj.references.append((line_no, aufrufende_prozedur_tuple[1], line_text))

                #Erweiterung um den Prozedurnamen, der aufgerufen wird:
                # ziel_prozedur_name = ""

                prozedur_obj.references.append((line_no, aufrufende_prozedur_tuple[1], line_text, ziel_prozedur_name))



                
                """
                # _BUGFIX: Fehlerhafte Zuordnung! ERLEDIGT + OK IN V. >= 0.2.0

                Aktuelle Parameter im scope:
                    line_no = 68
                    aufrufende_prozedur_tuple[1] = "bauer" # _BUGFIX: SOLLTE MAIN SEIN!!!!!


                """

                # db(prozedur_obj.references)





    @classmethod
    def finilize(cls, output_file_path:str):
        """
        ### TODO: Ob wirklich alles hier in die Klasse gehört ist fraglich! 
        
        Zugriff erfolgt über die Superklasse Procedure. Innerhalb der MEthode wird wiederum in Schleifen itteriert ueber die Subklassen.
        
        Bereitet die finale Ausgabe vor und ruft weitere MEthoden zum Schreiben dieser Ausgabe / der Markdown-Datei auf.
        Zu dieser Vorbereitung gehört:
         
            • Sortierung der einzelnen Prozeduren innerhalb der verschiedenen Prozedur-Arten gemaess der alphabetischen Reihenfolge ihrer Bezeichner (Namen)
            
            • Konfigurieren / Initilalisierung der Zwischenüberschriften / Header-Texte unterhalb der einzelnen Section-Überschriften (gespeichert in Klassenvariable cls.header)


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

            cls.header = content





            # Instanzen der Prozeduren  Alphabetisch  sortieren nach den Namen! und in Klassenvariable speichern:
            procedure_type_cls.sort_procedures_by_names()


            # Ergänze einen einzigen Eintrag für eine Methode! und haenge sie dem TOC an
            # WICHTIG: Auf Superklassen-Ebene!
            cls.generate_toc_entries(procedure_type_cls)

        
        # Durchsuche gesamten Quelltext nach allen Referenzierungen für jeweils alle gefundenen Prozeduren und speichere sie in den jeweiligen Objekten der einzelnen Prozeduren:
        cls.analyse_references()

        # JULIA: Diese Referenzierungen müssen jetzt noch in die Template prozedur.md eingebaut werden! Aktuell erfolgt das alles inkl. der Auswertung in der Methode write_to_file

        # TODO: Diese Referenzierungen müssen jetzt noch in die Template prozedur.md eingebaut werden! Aktuell erfolgt das alles inkl. der Auswertung in der Methode write_to_file










        # Aufruf der Methode zum tatsächlichen Schreiben der Textdatei:
        cls.write_to_file(output_file_path)



    
    def generate_calling_entries(self) -> None:
        """
        # TODO: # JULIA                                         
        mitten in dev.

        ansatz vorhanden



        Generiert die Dokumentatoin der Abrufsequenzen / Abruffolgen.
        Speichert das Ergebnis in der bereits existierenden Objektvariable procedure_obj.documentation
        """


        # TODO                                                                      
        # JULIA                                                                     
        """
        
        MÖGLICHE HERANGEHENSWEISEN


        - Vorgehensweise A:
            - Durchsuche gesamten Code DIESER Prozedur nach einem Aufruf einre Prozedur, die hier bereits bekannt ist (via regex)

        - Vorgehensweise B:
            -  Durchsuche alle OBJEKT-INSTANZEN aller Prozeduren beim key=References und schaue, ob eine zeilennummer der referenzierung innerhalb der zeilennummer dieser prozedur hier liegt.
            - scheint mir der entwicklungs-technisch schnellere weg zu sein...
            - bei langen programmen mit vielen prozeduren ist es besonders ineffizient, je kürzer die HIER TATSÄCHLICH GERADE zu dokumentierende funktion ist, aber das ist mir mal egal....
            - Vorteil: keine regex entwicklung / duplikat wie bei references!


        """

        """
        # Referenzierungs-zeile muss hierzwischen liegen:
        self.line_begin
        self.line_end







        """

        # TODO: Extrahiere alle Referenzierungen dazwischen:
        db("...")

        self.calling_sequences = []

        # bei hauptfunc1 --> soll 2 einträge: unterfunkntionA + B
        
        Procedure.all_procedures_final
        # liste mit exemplarischen einzel-Eintrag: 
        # (procedure_obj, procedure_name, procedure_line )



        for (procedure_obj, uebergeordnete_prozedur_name, procedure_declaration_line_no) in Procedure.all_procedures_final:

            """
            # iterater = 6 # schleifenzähler pro prozedur
            # einzelprozedur = Procedure.all_procedures_final[iterater][0]
            # ENTSPRICHT JETZT:
            procedure_obj

            procedure_obj.references
            # liste mit exemplarischen einzel-Eintrag: 
            # (line_nr_aufruf, "uebergeordnetes_sub", code)
            """
            
            # SCHNELLERER ANSATZ:
            # Filtern der liste, so dass uebergeordnetes_sub == name
            # dann fertig!

            relevante_references = [_ for _ in procedure_obj.references if uebergeordnete_prozedur_name == self.name]

            for (line_no_reference, uebergeordnetes_sub, code, ziel_prozedur_name) in relevante_references:

                self.calling_sequences.append((line_no_reference, uebergeordnetes_sub, code, ziel_prozedur_name)) # relevant ist nur der name aus der regex im cod
                # TODO: ??? ggf doch nicht nötig? Bei erstellung dieses tupels bei analyse der referenzen --> speichere auch das ziel, aleso den subname
            

            """
            # ALTER ALTERNATIVER ANSATZ:

            # ist line_nr zwischen self.line.begin + self.line.end der ZU DOKUMENTIERENDEN Prozedur??
            for (line_no_reference, uebergeordnetes_sub, code) in procedure_obj.references:
                if line_no_reference < self.line_begin:
                    continue
                if line_no_reference > self.line_end:
                    continue
                # SONST: liegt es darin!!! AUFNEHMEN IN LISTE!!!!

                self.calling_sequences.append((line_no_reference, uebergeordnetes_sub, code))
            """



                # TODO:  Am Ende vor der dokumentation die liste noch sortieren aufsteigend nach line_no_reference konzeptionell identisch wie im index/toc

                # JULIA: LOGIK FALSCH????
            db(self.calling_sequences)


            _PLACEHOLDER_REFERENCE = "@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_ENTRY@"

            # TODO: ersetze im doc:
            doc = self.documentation

    


            # Initialisierung und  Parametrisierung  des Einleitungssatzes:
            einleitungssatz = "Keine weiteren Aufrufe zu hier dokumentierten Prozeduren gefunden." # default

            if len(self.calling_sequences) > 0:

                einleitungssatz = "Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:"

            for (line_no_reference, uebergeordnetes_sub, line_code, target_procedure_name) in self.calling_sequences:


                # TODO:
                # HACK:
                target_procedure_name = "main"



                replacer_placeholder_reference = f"<small>  Zeile {line_no_reference} </small> : ```{line_code}``` ODER: {target_procedure_name}"
                # replacer_placeholder_reference = f"* [```{code}```](#{target_procedure}) : <small>  Zeile {line_no_reference} : ```{line_code}``` </small>"

                replacer_placeholder_reference = replacer_placeholder_reference  + f"\n{_PLACEHOLDER_REFERENCE}"
                    
                # Ersetzen:
                doc = doc.replace(_PLACEHOLDER_REFERENCE, replacer_placeholder_reference)

                    
            # Loeschen des verbliebenen Platzhalters zum Einfuegen einzelner Referenzen:
            doc = doc.replace(_PLACEHOLDER_REFERENCE, "")



            # Einsetzen des Einleitungssatzes:
            doc = doc.replace("@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@", einleitungssatz)


            # shortcut / resubstitution:
            self.documentation = doc




    
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
                line_code:str = line_code.rstrip("\n")

                replacer_placeholder_reference = f"* [```{calling_procedure_name}```](#{calling_procedure_name}) : <small>  Zeile {line_no} : ```{line_code}``` </small>"


                # # AUSBLICK: weitere collapse details : funktioniert technisch, allerdings steht das collapsable immer in neuer Zeile und daher wird es groß - vielleicht später in schön machen...
                # __replacer_placeholder_reference = f"*   [```{calling_procedure_name}```](#{calling_procedure_name})  <details> <summary>: <small>Zeile {line_no}</small> </summary> ```{line_code}``` </details>"



                replacer_placeholder_reference = replacer_placeholder_reference  + f"\n{_PLACEHOLDER_REFERENCE}"
                
                # Ersetzen:
                doc = doc.replace(_PLACEHOLDER_REFERENCE, replacer_placeholder_reference)


        # Loeschen des verbliebenen Platzhalters zum Einfuegen einzelner Referenzen:
        doc = doc.replace(_PLACEHOLDER_REFERENCE, "")



        # Einsetzen des Einleitungssatzes:
        doc = doc.replace("@PLACEHOLDER_PROCEDURE_REFERENCES_INTRODUCTION@", einleitungssatz)


        # shortcut / resubstitution:
        self.documentation = doc
        
        # TODO: Auslagern in eigene Funktion!     ENDE                       








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


                                
            # # TEST                                                                                                              
            # # Extrahiere Daten von der Version DIESES DOKUMENTIER-TOOLS:
            # content = MetaData.gitinfo 
            # content = content + "\n"*3 + "Datum der Umwandlung: " + MetaData.date_of_process
                          


            # TODO: Hier muss noch Placeholders ersetzt werden!!'
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

                    """
                    # TODO: Auslagern in eigene Funktion!     START                        erledigt ab V. 0.2.2

                    # shortcut:
                    doc:str = procedure_obj.documentation 

                    # Alle Referenzierungen sind in der Objektvariablen  procedure_obj.references gespeichert:
                    count_of_references = len(procedure_obj.references)

                    # ERsetzen des Platzhalters für die Anzahl der Referenzierungen in der bisherigen Dokumentation:
                    doc = doc.replace("@PLACEHOLDER_PROCEDURE_COUNT_OF_REFERENCES@", str(count_of_references))
                    

                    # Initialisierung und  Parametrisierung  des Einleitungssatzes:
                    einleitungssatz = "Kein Aufruf gefunden." # default

                    if count_of_references > 0:

                        einleitungssatz = "Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:"

                        
                        # Iterrieren ueber jede Referenzierung, um diese zu dokumentieren:
                        for (line_no, calling_procedure_name, line_code) in procedure_obj.references:

                            # Zusammenbau des Ersatzwertes für den Platzhalter inkl. Anhängen des Platzhalters für weitere Ersetzungen:
                            
                            # Um MArkdown nicht zu zerschiessen muss der letzte Zeilenumbruch des line_codes entfernt werrden:
                            line_code:str = line_code.rstrip("\n")

                            replacer_placeholder_reference = f"* [```{calling_procedure_name}```](#{calling_procedure_name}) : <small>  Zeile {line_no} : ```{line_code}``` </small>"


                            # # AUSBLICK: weitere collapse details : funktioniert technisch, allerdings steht das collapsable immer in neuer Zeile und daher wird es groß - vielleicht später in schön machen...
                            # __replacer_placeholder_reference = f"*   [```{calling_procedure_name}```](#{calling_procedure_name})  <details> <summary>: <small>Zeile {line_no}</small> </summary> ```{line_code}``` </details>"



                            replacer_placeholder_reference = replacer_placeholder_reference  + f"\n{_PLACEHOLDER_REFERENCE}"
                            
                            # Ersetzen:
                            doc = doc.replace(_PLACEHOLDER_REFERENCE, replacer_placeholder_reference)


                    # Loeschen des verbliebenen Platzhalters zum Einfuegen einzelner Referenzen:
                    doc = doc.replace(_PLACEHOLDER_REFERENCE, "")



                    # Einsetzen des Einleitungssatzes:
                    doc = doc.replace("@PLACEHOLDER_PROCEDURE_REFERENCES_INTRODUCTION@", einleitungssatz)


                    # shortcut / resubstitution:
                    procedure_obj.documentation = doc
                    
                    # TODO: Auslagern in eigene Funktion!     ENDE                       


                    """


                    # JULIA: Versuche Wiederherstollung des Tails!!!!
                    # procedure_obj.generate_calling_entries()

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




        # TODO:        Muss noch irgendwoghin!                                     
        # self.extract_references()



        # =============================================================================
        #### Zusammenfassen und Schreiben der Dokumentation: ####
        # =============================================================================
        
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































# =============================================================================
#### MAIN: 
# =============================================================================
def main():
    """
    Hauptprogramm. Steuert den Gesamt-Ablauf des Scripts. 
    Die meisten Methoden sind innerhalb der Superklasse Procedure definiert.
    """

    """
    # REFACTORING: In MetaData.initialize_class

    # HACK: path for Source-vba-code
    # input_file_path = "input_data/beispiel_modul.bas"
    # input_file_path = "input_data/beispiel_modul1.bas"
    input_file_path = "input_data/beispiel_modul2.bas"

    MetaData.set_input_path(input_file_path)

    """
    

    # initialize the class including reading the input file:
    Procedure.initialize_input_code(MetaData.get_input_path())


    #  Identifizieren und Speichern der Deklarationszeilen von Prozeduren:
    Procedure.identify_procedures()



    # Detail-Analyse und Speicherung einzelner Bestandteile in Objekten:
    Procedure.detail_analyse_procedures()



    """
    # REFACTORING: In MetaData.initialize_class
    # 
    # Schreibe die Datei erst in eine Markdown-Datei:
    output_dir = "output_data/demo_output.md"
    MetaData.set_output_dir(output_dir)

    """


    Procedure.finilize(MetaData.get_output_path())




    # Generiere HTML Datei from Markdown:
    if CONVERT_TO_HTML:
        markdown.markdownFromFile(
            input=MetaData.get_output_path(extension=".md"),
            output=MetaData.get_output_path(extension=".html"), 
            encoding="utf8"
        )

        print("HTML-Datei wurde aus MD-Datei generiert.")
        



    print(chr(13), '>>>>>>> ENDE.')









if __name__ == '__main__':
    main()



