# -*- coding: utf-8 -*-
'''
Created on: Wed, 2024-01-03 (16:43:36)

@author: Matthias Kader


# # WICHTIG:
# # IDEE / Programmablauf: für versionierer: 

# - major_change, minor_change, patch_change = 0, 0, 0
# - get git log
# - SCHLEIFE:
#     - from newest to oldest log iterate:
#     - hat das Commit einen Tag ? (eine Versionsnummer?)?
#         - Ja: 
#             - speichere diesen tag als letzte versionsnummer, per regex aufteilen zu last_major, last_minor, last_patch
#             - Dann beende die Schleife
#         - Nein: Dann lese commit message aus
#             - Beginnt es mit regex r"\w*!:" dann major_change += True und abbruch der schleife
#             - Beginnt es mit regex r"feat:" dann minor_change += True
#             - Beginnt es mit regex r"fix:" dann minor_change += True
#     - nächstes commit
# - if major_change:
#     - neue_version = f"{last_major + 1 }.0.0"
# - elif minor_change:
#     - neue_version = f"{last_major}.{last_minor + 1}.0"
# - elif patch_change:
#     - neue_version = f"{last_major}.{last_minor}.{last_patch + 1}"

# - git tag neue_version
                                  
# - 



'''





'''
# OBSOLET:
import subprocess





def get_git_commit_info():
    """
    Durchsucht gesamtes Git-Log des Projektes und gibt eine Liste mit dem Aufbau (<strCommitMessage>, )

    Returns:
        _type_: _description_
    """
    try:
        # Führen Sie den 'git log'-Befehl aus und erfassen Sie die Ausgabe
        git_log = subprocess.check_output(['git', 'log', '--pretty=format:%s']).decode('utf-8').strip()

        # Teilen Sie die Ausgabe in einzelne Commit-Nachrichten auf
        git_commit_messages = git_log.split('\n')




        # Führen Sie den 'git tag'-Befehl aus und erfassen Sie die Ausgabe
        git_tags = subprocess.check_output(['git', 'tag']).decode('utf-8').strip()

        # Teilen Sie die Ausgabe in einzelne Tags auf
        git_tags_list = git_tags.split('\n')

        # Erstellen Sie eine Liste, um die Verfügbarkeit von Tags für jeden Commit zu speichern
        tags_available = []

        for message in git_commit_messages:
            # Überprüfen Sie, ob Tags für die Commit-Nachricht verfügbar sind
            tags_for_commit = [tag for tag in git_tags_list if tag in message]
            tags_available.append(len(tags_for_commit) > 0)

        return list(zip(git_commit_messages, tags_available))
    

    except subprocess.CalledProcessError:
        # Fehlerbehandlung, wenn der Befehl nicht erfolgreich ausgeführt wird
        print("Fehler beim Ausführen des 'git log'-Befehls.")
        return None







def ansatz():

    # Beispielaufruf
    commit_messages = get_git_commit_info()

    if commit_messages:
        for message in commit_messages:
            # Hier können Sie die Commit-Nachrichten im Skript weiterverarbeiten
            print(f"Commit-Nachricht: {message}")

    else:
        print("Fehler beim Abrufen der Git-Commit-Nachrichten.")







def test_chatgpt():

    import subprocess

    def get_git_commit_info():
        try:
            # Führen Sie den 'git log'-Befehl aus und erfassen Sie die Ausgabe
            git_log = subprocess.check_output(['git', 'log', '--pretty=format:%H|%s']).decode('utf-8').strip()

            # Teilen Sie die Ausgabe in einzelne Commits auf
            git_commits = git_log.split('\n')

            # Erstellen Sie eine Liste, um die Verfügbarkeit von Tags für jeden Commit zu speichern
            tags_available = []

            for commit in git_commits:
                # Teilen Sie den Commit in Teile auf
                commit_parts = commit.split('|')
                commit_sha = commit_parts[0]

                # Führen Sie den 'git describe'-Befehl aus und erfassen Sie die Ausgabe
                try:
                    git_describe = subprocess.check_output(['git', 'describe', '--tags', commit_sha]).decode('utf-8').strip()
                    tags_available.append(True)
                except subprocess.CalledProcessError:
                    # 'git describe' schlägt fehl, wenn kein Tag vorhanden ist
                    tags_available.append(False)

            return list(zip(git_commits, tags_available))

        except subprocess.CalledProcessError:
            # Fehlerbehandlung, wenn der Befehl nicht erfolgreich ausgeführt wird
            print("Fehler beim Ausführen des 'git log' oder 'git describe'-Befehls.")
            return None

    # Beispielaufruf
    commit_info = get_git_commit_info()

    if commit_info:
        for commit, tags_available in commit_info:
            # Hier können Sie die Commit-Informationen und die Verfügbarkeit von Tags im Skript weiterverarbeiten
            print(f"Commit: {commit}")
            print(f"Tags verfügbar: {tags_available}")
            print("------")

    else:
        print("Fehler beim Abrufen der Git-Commit-Informationen.")

'''






import re
from typing import Union


def test_chatgpt2_ansatz_OK():

    import subprocess

    def get_git_commit_info():
        try:
            # Führen Sie den 'git log'-Befehl aus und erfassen Sie die Ausgabe
            git_log = subprocess.check_output(['git', 'log', '--pretty=format:%H|%s']).decode('utf-8').strip()

            # Teilen Sie die Ausgabe in einzelne Commits auf
            git_commits = git_log.split('\n')

            # Erstellen Sie eine Liste, um die Verfügbarkeit von Tags und den Text des Tags für jeden Commit zu speichern
            tag_info = []

            for commit in git_commits:
                # Teilen Sie den Commit in Teile auf
                commit_parts = commit.split('|')
                commit_sha = commit_parts[0]

                try:
                    # Führen Sie den 'git describe'-Befehl aus und erfassen Sie die Ausgabe
                    git_describe_output = subprocess.check_output(['git', 'describe', '--tags', commit_sha]).decode('utf-8').strip()
                    tags_available = True
                    tag_info.append((tags_available, git_describe_output))

                except subprocess.CalledProcessError:
                    # 'git describe' schlägt fehl, wenn kein Tag vorhanden ist
                    tags_available = False
                    tag_info.append((tags_available, None))

            return list(zip(git_commits, tag_info))


        except subprocess.CalledProcessError:
            # Fehlerbehandlung, wenn der Befehl nicht erfolgreich ausgeführt wird
            print("Fehler beim Ausführen des 'git log' oder 'git describe'-Befehls.")
            return None





    # Beispielaufruf
    commit_info = get_git_commit_info()

    print("\n"*6, "=" * 160, "\n", "=" * 160)


    if commit_info:
        for commit, (tags_available, tag_text) in commit_info:
            # Hier können Sie die Commit-Informationen, die Verfügbarkeit von Tags und den Text des Tags im Skript weiterverarbeiten
            print(f"Commit: {commit}")
            print(f"Tags verfügbar: {tags_available}")
            print(f"Tag-Text: {tag_text}")
            print("------")

    else:
        print("Fehler beim Abrufen der Git-Commit-Informationen.")










DEBUG = 1





def db(*args):
    """
    Schleust zum printen durch - nur zum Debuggen
    """
    if DEBUG == False:
        return
    
    print("__DEBUG_PRINT__\n")
    for _ in args:
        print(_)







class Commit:

    instances = []

    def __init__(self, commit:str, message:str, tag:str) -> None:
        """
        Erstellt ein neues Objekt der Klasse Commit.
        Als Attribute werden die uebergebenen Parameter gespeichert.

        Args:
            commit (str): Bezeichnung SHA des Commits
            message (str): Commit-MEssage
            tag (str): Tag(s) des Commits
        """
        Commit.instances.append(self)
        self.index = Commit.instances.index(self)

        
        self.commit = commit
        self.message = message
        self.tag = tag

        # None for unknown Version numbers: (default)
        self.major:int = None
        self.minor:int = None
        self.patch:int = None
        is_new_version:bool = None

    @property
    def version(self):
        # Dies entspricht de Getter für diese Property. 
        if all([ self.major, self.minor, self.patch]):  
            return "{}.{}.{}".format(self.major, self.minor, self.patch)
        else:
            return None

    
    @classmethod
    def get_last_commit(cls):
        """
        Gibt das neuste Commit als Objekt zurück.
        Dies entspricht dem Nullten Element der Liste cls.instances

        Returns:
            Commit: Objekt
        """
        return cls.instances[0]
    



    # @version.setter
    # def version(self, pNeueVersion):
    # nicht erforderlich, da sie nicht gesettet wird.


    @classmethod
    def new_obj_from_dict(cls, commit_dict: dict) -> 'Commit':
        """
        Generiert ein neues Commit-Objekt basierend aus einem Dictionary als Parameter.

        Args:
            commit_dict (dict): Dictionary mit den Keys commit, message und tag.

        Returns:
            Commit: Commit-Objekt
        """

        obj =  cls(
            commit=commit_dict['commit'], 
            message=commit_dict['message'], 
            tag=commit_dict['tag']
        )

        return obj

    @staticmethod                
    def add_tag(tag_text:str):
        """
        Führt einen git-befehl aus, um für DAS NEUSTE Commit einen Tag hinzuzufügen - also eig. völlig unabhängig von self!

        ### AUSBLICK: Könnte so gemacht werden, dass wirklcih für einen speziellen Commit ein tag angehangen wird...

        Args:
            tag (str): Text für den Tag
        """


        # if input("ECHT WEITER MACHEN????! [y]") != "y":
        #     db("kein tag hinzugefügt!!")
        #     return

        
        # Git-Befehl zum Erstellen eines Tags
        git_command = f"git tag {tag_text}"

        # Ausführen des Git-Befehls im Terminal
        try:
            subprocess.run(git_command, shell=True, check=True)
            print(f"Git Tag '{tag_text}' wurde erfolgreich erstellt.")

        except subprocess.CalledProcessError as e:
            print(f"Fehler beim Erstellen des Git Tags: {e}")
        



    @classmethod
    def initialize(cls) -> list[dict]:
        """
        Liest sämtliche verfügbaren Commits aus und speichert jedes als einzelnes Objekt innerhalb dieser Klasse. 
        Zusätzlich wird die Liste der Dictionaries, die die einzelnen Commits beschreiben zurückgegeben.
        """

        git_commits_infos = []

        try:
            # Führen Sie den 'git log'-Befehl aus und erfassen Sie die Ausgabe
            git_log = subprocess.check_output(['git', 'log', '--pretty=format:%H|%s']).decode('utf-8').strip()

            # Teilen Sie die Ausgabe in einzelne Commits auf
            git_commits = git_log.split('\n')
            # Erstellen Sie eine Liste, um die Verfügbarkeit von Tags und den Text des Tags für jeden Commit zu speichern
            tag_info = []
            for commit in git_commits:
                
                git_single_commit_infos = {}
                git_commits_infos.append(git_single_commit_infos)
                
                # Teilen Sie den Commit in Teile auf
                commit_parts = commit.split('|')
                commit_sha = commit_parts[0]
                git_single_commit_infos["commit"] = commit_sha
                git_single_commit_infos["message"] = commit_parts[1]
                git_single_commit_infos["tag"] = "" # leer initial
                try:
                    # Führen Sie den 'git describe'-Befehl aus und erfassen Sie die Ausgabe
                    git_describe_output = subprocess.check_output(['git', 'describe', '--tags', commit_sha]).decode('utf-8').strip()

                    # prüfe, ob ein explizit erstellter tag vorhanden ist, oder ob er generisch ist:

                    # Commit.is_valid_version_pattern_generic(git_describe_output, allow_generic_tags=True)

                    tags_available = True
                    tag_info.append((tags_available, git_describe_output))
                    git_single_commit_infos["tag"] = git_describe_output

                except subprocess.CalledProcessError:
                    # 'git describe' schlägt fehl, wenn kein Tag vorhanden ist
                    tags_available = False
                    tag_info.append((tags_available, None))

                # Erstellung eins Commit-Objektes mittels alternativen Konstruktors:
                cls.new_obj_from_dict(git_single_commit_infos)

                # db(cls.instances[0].version)

            print("\n"*10)

            # return list(zip(git_commits, tag_info))
            return git_commits_infos
            # return 





        except subprocess.CalledProcessError:
            # Fehlerbehandlung, wenn der Befehl nicht erfolgreich ausgeführt wird
            print("Fehler beim Ausführen des 'git log' oder 'git describe'-Befehls.")
            return None



    @classmethod
    def is_valid_version_pattern(cls, version:str) -> bool:
        """
        Prüft mittels regex, ob der text der semantischen Versionierung entspreicht.

        Args:
            version (str): zu überprüfender Text

        Returns:
            bool: True falls OK, False falls kein match
        """


        semanticVar = re.compile(r"^\d{1,}\.\d{1,}\.\d{1,}$")   
        # semanticVar = re.compile(f"\d{1,}\.\d{1,}\.\d{1,}")   
        # 0.4.0
        if semanticVar.match(version):
            return True
        else:
            return False
        








    @staticmethod
    def extract_params_from_version(version:str) -> tuple:
        """
        Liesst die Versionstext aus, prueft ihn auf Gueltigkeit nach dem Muster \d*\.\d*\.\d* und extrahiert bei Gültigkeit die numerischen Werte für major, minor und patch und gibt sie als tuple in dieser Reihenfolge zurück.
        Sofern das Muster nicht korrekt ist, werden alle Parameter mit None belegt
        """   

        # # siehe extra func is_valid_version_pattern
        # semanticVar = re.compile(f"\d{1,}\.\d{1,}\.\d{1,}")     
        
        # if semanticVar.match(version):
        if Commit.is_valid_version_pattern(version):
            major, minor, patch = map(int, version.split("."))

        else:
            print("!!! ERROR !!! Der uebergebene Versionstext {} entspricht nicht der KOnvention! Daher werden major, minor, patch alle mit None belegt!".format(version))
            major, minor, patch = (None, None, None)

        return (major, minor, patch)



    def suggest_version(self) -> str:
        """
        Sucht ausgehend von einem konkreten Commit-Objekt so lange in der Historie der Commits, bis ein Commit mit tag gefunden wurde. 
        Aktuell wird vorausgesetzt, dass die Versionsnummer die einzigen Tags sind!
        """

        index = self.index
        # ref_tag = None
        ref_obj:Commit = None




        # =============================================================================
        #### Suche Referenz-Commit, die bereits einen tag hat:  ####
        # =============================================================================
        

        # REFACTOR-TODO: Auslagern als Func!
        # Suche ab diesem Commit nach dem nächstfrüheren Commit, welches mit einem tag versehen ist:
        while (index < len(self.instances)):
                
            temp_obj:Commit = self.instances[index]

            # if temp_obj.tag:
            if self.is_valid_version_pattern(temp_obj.tag):
                # BUG: Da das aktuelle einen generischen tag erhält, wird dies nie abbrechen!!!
                # # found commit with tag!
                # ref_tag = temp_obj.tag
                # ref_index = temp_obj.index
                ref_obj = temp_obj
                break

            index += 1








        if not ref_obj:
            # no tag founded. --> suggest 0.0.0
            suggestion = "0.0.0"

            # break whole thing because there is nothing to analyse when there isn't an initial version number tag!
            return suggestion
        


        # BEISPIEL: ref_obj.tag = "1.2.3"

        # =============================================================================
        #### Durchsuche alle vergangenen Commit-MEssages nach Stichwörtern, für die eine Versionierung durchgeführt werden müssten ####
        # =============================================================================
        
            
        # ON this point there is definitely a previous version number.
        # So check all commit messages from the commit which follows of the founded ref_commit_object, wheather the keywords in their message is one for major, minor, patch or no version increments:



        # Flags for new changes:
        major_change = False
        minor_change = False
        patch_change = False




        
        # for index in range(self.index, ref_index - 1):
        for index, commit_obj in enumerate(self.instances):
            
            commit_obj:Commit # nur für typehints
         


            if commit_obj == ref_obj:
                #  Angekommen an dem versionierten, tagbehafteten Commit.
                # weitere änderungen nicht notwenig, weil alle sübersprüungen wird
                break

            type_of_commit = commit_obj.message.split(":")[0].rstrip(" ").lstrip(" ").lower()
            
            if type_of_commit.endswith("!"):
                # includes "feat!" AND "fix!"
                # BREAKING CHANGE --> SHOULD BE a NEW MAJOR VERSION!
                major_change = True
                # There is no need to search for more changes as major changes are the most top-leveled changes which can be represented in the semantic version.
                break

            
            elif type_of_commit == "feat":
                minor_change = True

            elif type_of_commit == "fix":
                patch_change = True





        # BEISPIEL: ref_obj.tag = "1.2.3"
        previous_version = ref_obj.tag

        major, minor, patch = self.extract_params_from_version(previous_version)
        
        # major, minor, patch = map(int, previous_version.split("."))
        # FIX: ACHTUNG: Fehlervermeidung fehlt! Gültigkeit der versionsnummern werden nicht geprüft!






        # =============================================================================
        #### # Evaluation of the previous changes since the last tagged commit: ####
        # =============================================================================
        
        self.is_new_version = True # default / initial
        # noch nicht herausspringen jeweils, da noch ein weiteres Attibut gesetzt /korrigiert werden muss!
        if major_change:
            suggestion = "{}.0.0".format(major + 1)
            # return suggestion
        
        elif minor_change:
            suggestion = "{}.{}.0".format(major, minor + 1)
            # return suggestion
        
        elif patch_change:
            suggestion = "{}.{}.{}".format(major, minor, patch + 1)
            # return suggestion
        
        else:
            # SONST: Keine bedeutenden Modifikationen. Die letzte  Versionsnummer bleibt bestehen.
            suggestion = previous_version
            self.is_new_version = False

        # als flag:

        return suggestion
        



        
                


        db("debug.....")
            




































import subprocess







def XX_OLD_get_git_commits_infos() -> list[dict]:

    git_commits_infos = []

    try:
        # Führen Sie den 'git log'-Befehl aus und erfassen Sie die Ausgabe
        git_log = subprocess.check_output(['git', 'log', '--pretty=format:%H|%s']).decode('utf-8').strip()

        # Teilen Sie die Ausgabe in einzelne Commits auf
        git_commits = git_log.split('\n')
        # Erstellen Sie eine Liste, um die Verfügbarkeit von Tags und den Text des Tags für jeden Commit zu speichern
        tag_info = []
        for commit in git_commits:
            
            git_single_commit_infos = {}
            git_commits_infos.append(git_single_commit_infos)
            
            # Teilen Sie den Commit in Teile auf
            commit_parts = commit.split('|')
            commit_sha = commit_parts[0]
            git_single_commit_infos["commit"] = commit_sha
            git_single_commit_infos["message"] = commit_parts[1]
            git_single_commit_infos["tag"] = "" # leer initial
            try:
                # Führen Sie den 'git describe'-Befehl aus und erfassen Sie die Ausgabe
                git_describe_output = subprocess.check_output(['git', 'describe', '--tags', commit_sha]).decode('utf-8').strip()
                tags_available = True
                tag_info.append((tags_available, git_describe_output))
                git_single_commit_infos["tag"] = git_describe_output

            except subprocess.CalledProcessError:
                # 'git describe' schlägt fehl, wenn kein Tag vorhanden ist
                tags_available = False
                tag_info.append((tags_available, None))

        # return list(zip(git_commits, tag_info))
        return git_commits_infos
        # return 





    except subprocess.CalledProcessError:
        # Fehlerbehandlung, wenn der Befehl nicht erfolgreich ausgeführt wird
        print("Fehler beim Ausführen des 'git log' oder 'git describe'-Befehls.")
        return None


'''
# OBSOLET: ALT:

def demo_use_git_log():

  # Beispielaufruf
    commits_infos = get_git_commits_infos()

    print("\n"*6, "=" * 160, "\n", "=" * 160)


    if not commits_infos:
        print("Fehler beim Abrufen der Git-Commit-Informationen.")
        return

    for commit in commits_infos:
        # commit:dict
        # for commit, (tags_available, tag_text) in commits_infos:
        for key, value in commit.items():
            # Hier können Sie die Commit-Informationen, die Verfügbarkeit von Tags und den Text des Tags im Skript weiterverarbeiten
            
            db(key, ": ", value)

            # print(f"Commit: {commit}")
            # print(f"Tags verfügbar: {tags_available}")
            # print(f"Tag-Text: {tag_text}")
        db("------")

'''

def main():





    # Bisherige Commits einlesen:
    Commit.initialize()


    # Neue Versionsnumer vorschlagen:
    commit:Commit = Commit.get_last_commit()
    suggested_version = commit.suggest_version()



    if not commit.is_new_version:

        print("Es wurden keine relevanten Modifikationen seit dem letzten Release vorgenommen, die eine neue Version erforderlich machen. Die bisherige Versionsnummer {} bleibt bestehen.{}".format(suggested_version, "" if commit.tag == "" else "(Generischer tag: " +  commit.tag + ")"))

        return False


    # ELSE: tag der vorgeschlagenen Version fuer den aktuellen commit hinzufuegen


    version = suggested_version
    while  not Commit.is_valid_version_pattern(version):
    
    
    
    # if Commit.is_valid_version_pattern(suggested_version):
    # #     # dann gültige Version!


    # # commit.major, commit.minor, commit.tag = commit.extract_params_from_version(suggested_version)
    # # if commit.version:

    # else:
        version = input("!!! ACHTUNG: Die vorgeschlagene Version ist nicht gueltig. Bitte eine gueltige Versionsnummer eingeben oder 'x' fuer Abbruch.\n > Ihre Eingabe: ")
        
        if version.lower() == "x":
            db("Programmabbruch durch Benutzter")

            return False


    # tag setzten:
    commit.add_tag(version)


    # db(Commit)

    # initialize_commits()



    # demo_use_git_log()
  


main()