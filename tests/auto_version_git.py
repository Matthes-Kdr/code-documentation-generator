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
            return "UNKNOWN"

    

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

            # return list(zip(git_commits, tag_info))
            return git_commits_infos
            # return 





        except subprocess.CalledProcessError:
            # Fehlerbehandlung, wenn der Befehl nicht erfolgreich ausgeführt wird
            print("Fehler beim Ausführen des 'git log' oder 'git describe'-Befehls.")
            return None



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



    # Flags for new changes:
    major_change = False
    minor_change = False
    patch_change = False



    Commit.initialize()

    # initialize_commits()



    # demo_use_git_log()
  


main()