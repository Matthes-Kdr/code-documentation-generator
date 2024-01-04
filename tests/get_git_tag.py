import subprocess

def get_git_tag():
    try:
        # Führen Sie den 'git tag'-Befehl aus und erfassen Sie die Ausgabe
        git_tag = subprocess.check_output(['git', 'tag']).decode('utf-8').strip()

        # Teilen Sie die Ausgabe, falls mehrere Tags vorhanden sind (nehmen Sie den ersten)
        git_tags = git_tag.split('\n')
        if git_tags:
            return git_tags[0]
        else:
            return None

    except subprocess.CalledProcessError:
        # Fehlerbehandlung, wenn der Befehl nicht erfolgreich ausgeführt wird
        print("Fehler beim Ausführen des 'git tag'-Befehls.")
        return None





# Beispielaufruf
current_git_tag = get_git_tag()

if current_git_tag:
    print("Aktueller Git-Tag: ", current_git_tag)
    # Hier können Sie den Git-Tag im Programm weiterverarbeiten
else:
    print("Fehler beim Abrufen des Git-Tags.")







import subprocess

def get_git_commit_messages():
    try:
        # Führen Sie den 'git log'-Befehl aus und erfassen Sie die Ausgabe
        git_log = subprocess.check_output(['git', 'log', '--pretty=format:%s']).decode('utf-8').strip()

        # Teilen Sie die Ausgabe in einzelne Commit-Nachrichten auf
        git_commit_messages = git_log.split('\n')

        return git_commit_messages

    except subprocess.CalledProcessError:
        # Fehlerbehandlung, wenn der Befehl nicht erfolgreich ausgeführt wird
        print("Fehler beim Ausführen des 'git log'-Befehls.")
        return None


# Beispielaufruf
commit_messages = get_git_commit_messages()

if commit_messages:
    for message in commit_messages:
        # Hier können Sie die Commit-Nachrichten im Skript weiterverarbeiten
        print(f"Commit-Nachricht: {message}")

else:
    print("Fehler beim Abrufen der Git-Commit-Nachrichten.")








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

if commit_info:
    for commit, (tags_available, tag_text) in commit_info:
        # Hier können Sie die Commit-Informationen, die Verfügbarkeit von Tags und den Text des Tags im Skript weiterverarbeiten
        print(f"Commit: {commit}")
        print(f"Tags verfügbar: {tags_available}")
        print(f"Tag-Text: {tag_text}")
        print("------")

else:
    print("Fehler beim Abrufen der Git-Commit-Informationen.")






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