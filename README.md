# Code Documentation Generator



> Tool for generate a documentation of sourcecode (currently only for single VBA-moduls) 

.... (TODOC!)


# Möglicher Programmablauf zur Dokumentierung von VBA-Code


Abarbeiten mit Python, v.a. unter Nutzung von regulärer Ausdrücke (regex, package re)


Hier ein beispielhafter, möglicher Programmablauf


# Grundsätzliche Basis-Struktur zur Generierung erster Ergebnisse



1. Einfache GUI zur Auswahl der Quell-Datei(en)  und Ziel-Speicherorte
1. .bas-Textdatei einlesen
2. Regex-Suche nach sämtlichen Subs, die NICHT auskommentiert sind. Zeilennummern werden gespeichert in Liste
2. Regex-Suche nach sämtlichen Functions, die NICHT auskommentiert sind. Zeilennummern werden gespeichert in Liste.
3. Auflistung aller gefundener Subs in der Kategorie Subs unter Angabe von Namen, ggf. auch von Zeilennummern zusätzlich, und Scope (Private / public...)
3. Auflistung aller gefundener Functions in der Kategorie Functions unter Angabe von Namen, ggf. auch von Zeilennummern zusätzlich, und Scope (Private / public...)



Beim Auflisten wird jeweils eine template verwendet (vgl. Code-Generator / PAP-Designer), sodass mit Platzhaltern gearbeitet werden kann. Es wird ein Anker gesetzt, sodass zusätzlich ein Inhaltsverzeichnis am Ende erstellt werden kann.

Damit hätte man mit relativ wenig Aufwand eine Auflistung aller verfügbarer Subs + Functions als Dokumentation inkl. Index. Was dann noch fehlt, sind die Details zu diesen einzelnen Prozeduren.




# Inhalte der einzelnen Prozeduren identifizieren und dokumentieren


Vorgehensweise dann ist relativ easy, sofern doppelte Arbeit okay wäre - und das ist es:
* Verbinde Listen der Zeilennummern der einzelnen Subs und Functions zu einer großen Liste, unabhängig von Art der Prozedur
* Sortiere diese Liste aufsteigend

Dann ist klar, von welcher Zeile bis zu welcher Zeile jeweils die Prozedur im Quellcode steht. Die Abzüge aufgrund von langen auskommentierten Zeilen oder Leertasten (oder auch nicht verwendeter Prozeduren) erfolgen später

WEitere Vorgehensweise:
1) Extrahiere den Text ab Startzeile einre Prozedur bis zur Startzeile - 1 der nächsten Prozedur
2) Fange von hinten an: Suche von hinten die erste Zeile, in der wirklich Code-Relevantes steht (keine Leerzeilen, kein ausschließliche Kommentarzeile)
3) Von dieser Zeile dann wiederum vorwärts richtung hinten: Sofern diese Zeile + 1 leer ist: Lösche alles dahinter bis zum Ende dieser Prozedur. (Hintergrund: Nach einem Code-Befehl darf noch eine Zeile Kommentar stehen, aber nicht 100 Zeilen kOmmentar + leerzeilen, das könnte nämlcih einfach eine veraltete und nciht mehr gebrauchte, auskommentierte anderer Methode sein...)
4) Damit ist die Startzeile und Endzeile geklärt
5) Als nächsts wird geschaut, ob in Zeile 1 unterhalb der Deklarierung ein Kommentar steht. Falls ja: Werte sie als Docstring. Falls nein (code ODR leezeile: kein Beschreibung vorhanden!).Falls Docstring gewertet wurde: Suche auch in allen folgenden Zeilen danach und erweiterer den docstring, bis dass ein eLeerzeile ODRE eine nicht komplett auskommentierte Zeile gefunden wird.
6) Damit ist auch die Beschreibung verfügbar, die in die Template eingefügt werden kann
7) ebenfalls verfügbar sind hierdurch nun die Zeilennummern, von und bis, in der die Prozedur stattfndt. Dies kann ebenfalls in die Template gesetzt werden.


# Aufrufe / Referenzierungen der Prozeduren
Es fehlt nur noch die liste der Aufrufe der Prozedur.
Hierzu: Durchsuche den gesamten Quelltext nach einem Regex-Muster, der einen NICHT-AUSKOMMENTIERTEN Aufruf der entsprechenden Prozedur hat. 
Extrahiere die Zeile hieraus, und die aufrufende Prozedur (dies ist über Vergleich zwischen ZEilennummern und der Zeilennummernlisten möglich), dann kann auch noch der klartext aufgeschrieben werden.


# Call Sequenz / Calling Sequence

Schön (Ausblick) wäre auch ein weiterer Unterpunkt pro Prozedur, in der die Aufrufabfolge hervorgeht.
Idee ist etwas wie die Aufrufebenen-Auflistung beim Noten-Converter-Programm, d.h. ausgehend von einer Prozedur soll eine Liste stehen der Aufrufe von weiteren Prozeduren die aufgerufen werden (und die in diesem Dokument auch dokumentiert werden... also keine Builtins o.ä.). Im Idealfall kann jeder Punkt dieser Liste wiederum erweitert/expanded werden, darin ist dann wiederum die Liste von DIESER AUFGERUFENEN Funktion drin usw... Rekursiv. Jede Methode, die einmal so dokumentiert wurde kann weiter verwendet werden per Direktzugriff....



# Abschlussgedanke - wichtig vor dem Start
Der Ablauf des Programmes wäre auch ähnlich, wenn statt VBA eine andere PRogrammiersprache  dokumentiert werden sollte.  Lediglich die KEywords und Strukturen der zu dokumentierenden Programmierug weichen ab, dementsprechend müssten andere Regex-Mustern verwendet werden. Etwas komplizierter wäre das, wenn die Sprachen stark von voneinander abweichen, oder falls verschiedene Symbole gleiche Bedeutungen haben (z. B. die verschiednenen String-Einleiter bei Python ' oder "). Trotzdem sollte vor Beginn der Arbeit zumindest überlegt werden, in wie weit es sinnvoll oder möglich wäre, das Programm so aufzubauen, dass es ggf. auf verschiedene PRogrammiersprachen erweitert werden könnte.

Das Python-Modul pdoc liefert zwar ohne eigene Programmierung eine ausgereifte, sehr übersichtliche, gut druckbare und interaktive Übersicht von Python-Modulen, allerdings gehen Referenzierungen der einzelnen PRozeduren nicht aus dieser Dokumentation hervor.
Ein Alternativer Ansatz wäre, obo man das o.g. Zusatzmodul selbst anpassen und erweitern könnte durch solche Funktionalitäten.


Relevant wäre eine solche Dokumentationsmöglichkeiten für die folgenden Programmiersprachen:
* VBA
* PYthon
* C++ (µC)