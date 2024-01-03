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








# Beispielaufruf
commit_messages = get_git_commit_info()

if commit_messages:
    for message in commit_messages:
        # Hier können Sie die Commit-Nachrichten im Skript weiterverarbeiten
        print(f"Commit-Nachricht: {message}")

else:
    print("Fehler beim Abrufen der Git-Commit-Nachrichten.")



