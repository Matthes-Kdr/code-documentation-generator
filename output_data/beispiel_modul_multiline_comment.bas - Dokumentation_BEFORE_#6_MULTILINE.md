﻿# Code-Dokumentation: Modul 'beispiel_modul_multiline_comment.bas'



**Letzte Änderung** der Quelldatei 'beispiel_modul_multiline_comment.bas' vor der Generierung dieser automatischen Dokumentation: **2024-02-08 21:52**


Generierungsdatum dieser Dokumentation: **2024-02-08 21:53:57**





﻿


<!-- --------------------------------------------------------------- -->
<!-- Index / TOC -->
<!-- --------------------------------------------------------------- -->

## Index / Table Of Content

Alphabetische und verlinkte Auflistung aller Subs und Functions, die in diesem Modul verwendet werden:

* [**Modulinformationen / Modulkopf**](#sec_modulinfos)
  

  
  <!-- ---------- SUBS: --------------- -->

* [**Subs**](#sec_subs) (4)
  
  - [```bauer```](#bauer) : <small>  [Zeile 97]  </small>


  - [```casio```](#casio) : <small>  [Zeile 143]  </small>


  - [```liebherr```](#liebherr) : <small>  [Zeile 126]  </small>


  - [```main```](#main) : <small>  [Zeile 51]  </small>


  




  <!-- ---------- FUNCTIONS: --------------- -->


* [**Functions**](#sec_functions) (2)
  
  
  - [```addieren```](#addieren) : <small>  [Zeile 13]  </small>


  - [```subtrahieren```](#subtrahieren) : <small>  [Zeile 31]  </small>


  




  <!-- ---------- TAIL: --------------- -->


* [**Schlussbemerkungen** (inkl. Angaben zum Entwicklungszustand des Code-Dokumenter-Tools)](#sec_tail)




﻿


<a name="sec_modulinfos"></a>

## Modulbeschreibung

  
 Beispiel Modul zum Testen der Dokumentation der Abruffolge.

﻿
<!-- -------------------------------------------------- -->
<!-- SECTION-START : SUBS -->
<!-- -------------------------------------------------- -->

<a name="sec_subs"></a>

## Subs


﻿





<!-- --------------------------------------------------------------- -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- --------------------------------------------------------------- -->




<!-- PLACEHOLDER: Initialisierungszeile: -->



<a name="bauer"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Sub bauer```</span><small>(Zeile 97)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Anzahl der Referenzierungen im Modul: 0
 Anzahl weiterer internen Aufrufe : 5

 Ruft MEHRFACH die Prozedur 'liebherr' auf (insgesamt 5 mal nacheinander in verschiedenen Kontexten)








<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (0)</summary>

<div style="padding-left:1em;">



Kein Aufruf gefunden.







</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (5)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```liebherr```](#liebherr) : <small>  [Zeile 107] : ```    call liebherr``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```liebherr```](#liebherr) : <small>  [Zeile 108] : ```    call liebherr ' Aufruf``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```liebherr```](#liebherr) : <small>  [Zeile 110] : ```    call liebherr("ERROR") ' Aufruf waere zwar ungültig, aber Prozedur könnte ja anders aussehen!``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```liebherr```](#liebherr) : <small>  [Zeile 113] : ```    var = liebherr``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```liebherr```](#liebherr) : <small>  [Zeile 114] : ```    var = liebherr("gvkil")``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
Public Sub bauer()
' Anzahl der Referenzierungen im Modul: 0
' Anzahl weiterer internen Aufrufe : 5
'
''' Ruft MEHRFACH die Prozedur 'liebherr' auf (insgesamt 5 mal nacheinander in verschiedenen Kontexten)
'

    MsgBox("Dies ist ein explizit als public gekennzeichnetes Sub.")

    ' Aufruf:
    call liebherr
    call liebherr ' Aufruf
    
    call liebherr("ERROR") ' Aufruf waere zwar ungültig, aber Prozedur könnte ja anders aussehen!

    ' Wiederum unügltig:
    var = liebherr
    var = liebherr("gvkil")





End Sub

```

</details>


</div>


---


<!-- --------------------------------------------------------------- -->
<!-- --------------------------------------------------------------- -->


























﻿





<!-- --------------------------------------------------------------- -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- --------------------------------------------------------------- -->




<!-- PLACEHOLDER: Initialisierungszeile: -->



<a name="casio"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Sub casio```</span><small>(Zeile 143)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Anzahl der Referenzierungen im Modul: 0
 Anzahl weiterer internen Aufrufe : 0

 ' Ruft keine weitere Prozedur auf.







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (0)</summary>

<div style="padding-left:1em;">



Kein Aufruf gefunden.







</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (0)</summary>

---


Keine weiteren Aufrufe zu hier dokumentierten Prozeduren gefunden.






- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
   Sub casio()
    ' Anzahl der Referenzierungen im Modul: 0
    ' Anzahl weiterer internen Aufrufe : 0
    '
    ''' ' Ruft keine weitere Prozedur auf.

    MsgBox("Dies ist ein implizit als public gekennzeichnetes Sub.")


End Sub

```

</details>


</div>


---


<!-- --------------------------------------------------------------- -->
<!-- --------------------------------------------------------------- -->


























﻿





<!-- --------------------------------------------------------------- -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- --------------------------------------------------------------- -->




<!-- PLACEHOLDER: Initialisierungszeile: -->



<a name="liebherr"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Sub liebherr```</span><small>(Zeile 126)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Anzahl der Referenzierungen im Modul: 6
 Anzahl weiterer internen Aufrufe : 0

 ' Ruft keine weitere Prozedur auf.







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (6)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```main```](#main) : <small>  [Zeile 87] : ```    call liebherr``` </small>

- [```bauer```](#bauer) : <small>  [Zeile 107] : ```    call liebherr``` </small>

- [```bauer```](#bauer) : <small>  [Zeile 108] : ```    call liebherr ' Aufruf``` </small>

- [```bauer```](#bauer) : <small>  [Zeile 110] : ```    call liebherr("ERROR") ' Aufruf waere zwar ungültig, aber Prozedur könnte ja anders aussehen!``` </small>

- [```bauer```](#bauer) : <small>  [Zeile 113] : ```    var = liebherr``` </small>

- [```bauer```](#bauer) : <small>  [Zeile 114] : ```    var = liebherr("gvkil")``` </small>





</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (0)</summary>

---


Keine weiteren Aufrufe zu hier dokumentierten Prozeduren gefunden.






- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
   Sub liebherr()
   ' Anzahl der Referenzierungen im Modul: 6
    ' Anzahl weiterer internen Aufrufe : 0
    '
    ''' ' Ruft keine weitere Prozedur auf.


    MsgBox("Dies ist ein implizit als public gekennzeichnetes Sub.")


End Sub

```

</details>


</div>


---


<!-- --------------------------------------------------------------- -->
<!-- --------------------------------------------------------------- -->


























﻿





<!-- --------------------------------------------------------------- -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- --------------------------------------------------------------- -->




<!-- PLACEHOLDER: Initialisierungszeile: -->



<a name="main"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Sub main```</span><small>(Zeile 51)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Anzahl der Referenzierungen im Modul: 0
 Anzahl weiterer internen Aufrufe : 3

 Ruft die MEthode 'addieren' auf

 Ruft die MEthode 'subtrahieren' auf (und darin dann wieder addieren)

 Ruft die MEthode 'liebherr' auf







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (0)</summary>

<div style="padding-left:1em;">



Kein Aufruf gefunden.







</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (4)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```addieren```](#addieren) : <small>  [Zeile 66] : ```wert = addieren(i, i) ' Dies sollte NICHT dokumentiert werden, sofern das Multiline-Feature implementiert ist!``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```addieren```](#addieren) : <small>  [Zeile 81] : ```        wert = addieren(i, i)``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```subtrahieren```](#subtrahieren) : <small>  [Zeile 82] : ```        wert = subtrahieren(i, i - 1) ' Erklärung siehe @ Func!``` </small>


  - [```addieren```](#addieren) : <small>  [Zeile 41] : ```    subtrahieren = addieren(a, -b) ' Parameter b wird mit -1 multipliziert übergeben``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```liebherr```](#liebherr) : <small>  [Zeile 87] : ```    call liebherr``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
Private Sub main()
' Anzahl der Referenzierungen im Modul: 0
' Anzahl weiterer internen Aufrufe : 3
'
''' Ruft die MEthode 'addieren' auf
'
''' Ruft die MEthode 'subtrahieren' auf (und darin dann wieder addieren)
'
''' Ruft die MEthode 'liebherr' auf


' HACK: Durch 5-maliges Wiederholen von "#" wird simuliert, als ob es einen Syntax Multiline-Comment gäbe
#####
Dies hier soll ein Block kommentar-Demo sein, den es in VBA gar nicht gibt, aber egal
Hier ist noch eine Zeile davon!
wert = addieren(i, i) ' Dies sollte NICHT dokumentiert werden, sofern das Multiline-Feature implementiert ist!
#####



' Das hier soll nirgendwo stehen.

    dim i as integer

    i = 10

    for i = 0 to 10
        
        msgbox(i)
        ' Ausgabe:
        wert = addieren(i, i)
        wert = subtrahieren(i, i - 1) ' Erklärung siehe @ Func!

    next i


    call liebherr


End Sub

```

</details>


</div>


---


<!-- --------------------------------------------------------------- -->
<!-- --------------------------------------------------------------- -->


























﻿
<!-- -------------------------------------------------- -->
<!-- SECTION-START : FUNCTIONS -->
<!-- -------------------------------------------------- -->

<a name="sec_functions"></a>

## Functions


﻿





<!-- --------------------------------------------------------------- -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- --------------------------------------------------------------- -->




<!-- PLACEHOLDER: Initialisierungszeile: -->



<a name="addieren"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function addieren```</span><small>(Zeile 13)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Anzahl der Referenzierungen im Modul: 2
 Anzahl weiterer internen Aufrufe : 0

 Diese Funktion addiert beide Zahlen miteinander und übergibt das Ergebnis zurück.

 Ruft keine weitere Prozedur auf.







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (3)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```subtrahieren```](#subtrahieren) : <small>  [Zeile 41] : ```    subtrahieren = addieren(a, -b) ' Parameter b wird mit -1 multipliziert übergeben``` </small>

- [```main```](#main) : <small>  [Zeile 66] : ```wert = addieren(i, i) ' Dies sollte NICHT dokumentiert werden, sofern das Multiline-Feature implementiert ist!``` </small>

- [```main```](#main) : <small>  [Zeile 81] : ```        wert = addieren(i, i)``` </small>





</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (0)</summary>

---


Keine weiteren Aufrufe zu hier dokumentierten Prozeduren gefunden.






- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
Private Function addieren(a as integer, b as integer) as integer
' Anzahl der Referenzierungen im Modul: 2
' Anzahl weiterer internen Aufrufe : 0
'
''' Diese Funktion addiert beide Zahlen miteinander und übergibt das Ergebnis zurück.
'
' Ruft keine weitere Prozedur auf.


    ' Addieren:
    addieren = a + b

End Function

```

</details>


</div>


---


<!-- --------------------------------------------------------------- -->
<!-- --------------------------------------------------------------- -->


























﻿





<!-- --------------------------------------------------------------- -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- --------------------------------------------------------------- -->




<!-- PLACEHOLDER: Initialisierungszeile: -->



<a name="subtrahieren"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function subtrahieren```</span><small>(Zeile 31)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Anzahl der Referenzierungen im Modul: 1
 Anzahl weiterer internen Aufrufe : 1

 Diese Funktion subtrahiert b von a und übergibt das Ergebnis zurück.


 Ruft die  Prozedur 'adddieren' auf.







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```main```](#main) : <small>  [Zeile 82] : ```        wert = subtrahieren(i, i - 1) ' Erklärung siehe @ Func!``` </small>





</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (1)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```addieren```](#addieren) : <small>  [Zeile 41] : ```    subtrahieren = addieren(a, -b) ' Parameter b wird mit -1 multipliziert übergeben``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
Private Function subtrahieren(a as integer, b as integer) as integer
' Anzahl der Referenzierungen im Modul: 1
' Anzahl weiterer internen Aufrufe : 1
'
''' Diese Funktion subtrahiert b von a und übergibt das Ergebnis zurück.
'
'
' Ruft die  Prozedur 'adddieren' auf.

    ' Benutze die addieren Funktion:
    subtrahieren = addieren(a, -b) ' Parameter b wird mit -1 multipliziert übergeben


end Function

```

</details>


</div>


---


<!-- --------------------------------------------------------------- -->
<!-- --------------------------------------------------------------- -->


























﻿




---

<a name="sec_tail"></a>

## Schlussbemerkungen



<!-- 
**Notice:**

*To generate a docstring from the VBA-Source, make sure that the text to shown is located directly below the declaration line of the procedure. The text is considered completed with the first following line in the code which is not an entire comment line.  Empty lines that are to be included must also be labelled as comments.*



 **TODO:** Erstellt am (Datum) durch das  automatisierte Code-Dokumentationstool von .... in der Version ....







---



**ODER:** -->

> Diese Dokumentation wurde automatisiert durch ein entsprechendes Programm generiert. Im Einzelfall, insbesondere wenn im zu dokumentierenden VBA-Code von  bestimmten Konventionen abgewichen wird, können Unstimmigkeiten auftreten. Im Zweifel ist der Originalcode heranzuziehen.


Im folgenden werden die Modulinformationen des PYTHON-SCRIPTES aufgeführt, durch welches diese Dokumentation generiert wurde.

<details>

<summary> Modulinformationen anzeigen/verbergen.
</summary>

  <br>Created on: Fri, 2023-12-29 (00:45:39)<br><br><br>@author: Matthias Kader<br><br><br>Für generelles Ziel und Ablauf des Scriptes siehe MArkdown im Verzeichnis ../Tests/Programmablauf.html<br><br>Wichtige Details siehe am Ende dieses docstrings.<br><br><br><br><br>### Fertig implementiert:<br><br>- Inhaltsverzeichnis / Index<br><br>- Gesamtlayout inkl. Titel, Zwischenüberschriften für einzelne Sections<br><br>- Aufführen  des modulweiten Programmkopf-Docstring in der generierten Dokumentation<br><br>- Aufführen der References-Durchsuchungen (Wo wird die Prozedur aufgerufen?) in der generierten Dokumentation<br><br>- Sofortiger Export der MD-Datei in eine  HTML-Datei<br><br>- Aufführen der organisatorischer Daten bzgl. des zu dokumentierenden Codes und des verwendeten Skripts zum Dokumentieren in der generierten Dokumentation<br><br>- Aufführen der Calling Sequence (Aufrufabfolge / Aufrufebenen) innerhalb jeder Prozedur in der generierten Dokumentation: Aufzählung der Aufrufe anderer, in dieser Dokumentation behandelten Prozeduren. Inklusive rekursive geschachtelte Liste, welche Aufrufe jeweils in den aufgerufenen Prozeduren erfolgen.<br><br><br>- Bereitstellung einer einfachen GUI / HMI, um Input- und Output Pfade zu parametrisieren<br><br><br><br><br><br><br><br>### TODOS:<br><br><br>- Chore: Aufräumen des Quellcodes<br><br>- Refactor: ggfs. modifizieren von write_content<br><br>- BUGFIX: Modul 1 aufrufe<br><br>- "Help... " Button in GUI, in dem Erklärungen stehen! --> BESSER, universeller, einfacher und weniger duplizierend: ERstelle eine README.md im Repository, und beim Klick auf "help-btn" wird diese Datei in eine HTML umgewandelt und im Browser angezeigt...<br><br><br><br><br>### AUSBLICK für später und in schön:<br><br><br>- Zusatzmöglichkeit in GUI einen benutzerdefinierten Text einzugeben (Prio sehr gering!!). Dieser würde dann in einre eigenen Section angezeigt werden.<br><br>- Index an der Seite wie eine NavBar zum einzelnd scrollen<br><br>- Ermöglichung von Berücksichtigung weiterer Module innerhalb der Dokumentation<br>    <br>    - z. B. 2 VBA-Module innerhalb eines Projektes, wobei Prozeduren von Modul1  andere Prozeduren aus Modul2 aufrufen.<br><br>        - Erstmal nur als Verweis  (Mögl. Ansatz included = "Modul1.*" ohne rekursive Auflistung derer Aufrufe... oder eben mit... bestenfalls auch das parametrisierbar)<br><br>- Dokumentation von weiteren PRogrammiersprachen<br><br>    - OK --> VBA<br>    - Nächste Prio: C++ / Arduino<br>    - Letzte Prio: Python (v.a. für den Ablaufsequence sehr hilfreich, für den rest gibt es pdoc...)<br><br><br><br><br><br><br><br><br># =============================================================================<br>#### Wichtige Aufrufreihenfolge der Methode innerhalb dieses Python-Scriptes zur Erstellung der Dokumentation der Aufrufreihenfolge der zu dokumentierenden VBA-Prozeduren: ####<br># =============================================================================<br><br>Es werden zunächst alle Prozeduren komplett analysiert, erst danach werden wiederum alle Prozeduren komplett dokumentiert. Für beide Vorgänge erfolgt dies in einer Methode auf Objektebene, wobei diese jeweilige MEthode in beiden Fällen aus einer Klassenmethode aufgerufen wird, in der über die einzelnen Prozedur-Objekte innerhalb dieser Klasse iteriert wird:<br><br>- analyse_call_sequence(cls)<br>    - analyse_calling_sequence_in_one_proc(self)<br>- prepare_all_call_sequences_docs(cls)<br>    - prepare_single_call_sequence_docs(cls)<br><br>(hierfür wäre das entwickelte Tool  übrigens eine tolle Anwendung gewesen, sofern sie später auch mal Python-Syntax dokumentieren könnte :-) )<br><br><br><br><br><br># =============================================================================<br>#### Hinweise zur Anwendung und Benutzung: ####<br># =============================================================================<br><br>- To generate a docstring from the VBA-Source make sure that the text to shown is located directly below the declaration line of the procedure. The text is considered completed with the first following line in the code which is not an entire comment line. Empty lines that are to be included must also be labelled as comments.<br><br>- Durch das Script wird eine MD-Datei (Markdown) erzeugt, die anschließend über die Library markdown sofort in eine HTML umgewandelt wird, sodass nach Abschluss des Scriptes 2 Dateien erstellt wurden. Durch unterschiedliche Interpretationen im Rahmen der Konvertierung unterscheidet sich die Darstellung der so generierten HTML-Datei allerdings, wenn sie über VSCode Extension gesondert konvertiert wird. Die über VSCode generierte Datei ist übersichtlicher und schöner. Das sollte also am Ende nochmals gesondert erfolgen.<br><br><br><br><br><br><br><br># =============================================================================<br>#### Unwichtige Nebensächlichkeiten: Code-Analyse Zusammenfassung: ####<br># =============================================================================<br><br>In der Version vom 2024-01-11 - 00:18:43:<br>    Angaben jeweils: [Zeilen @ code_documenter.py] + [Zeilen @ gui] = [Summe]<br>    - Gesamtanzahl der Zeilen: 408 (100%)+2201 (100%)=2609 (100%)<br>    - davon Leerzeilen: 168 (41,1764705882353%)+1071 (48,6597001363017%)=1239 (47,489459563051%)<br>    - davon Einzelkommentarzeilen: 20 (4,90196078431373%)+244 (11,0858700590641%)=264 (10,1188194710617%)<br>    - davon Blockkommentarzeilen: 85 (20,8333333333333%)+374 (16,9922762380736%)=459 (17,5929474894596%)<br><br>    ==> Summe aller Kommentarzeilen: 105 (25,7352941176471%)+618 (28,0781462971377%)=723 (27,7117669605213%)<br>    ==> Code-relevante Zeilen: 135 (33,0882352941176%)+512 (23,2621535665607%)=647 (24,7987734764277%)<br><br>-----------------------------------------------    <br><br>In der Version vom 2024-01-07 - 23:37:04:<br>    - Gesamtanzahl der Zeilen: 2164 (100%)<br>    - davon Leerzeilen: 1091 (50%)<br>    - davon Einzelkommentarzeilen: 226 (10%)<br>    - davon Blockkommentarzeilen: 364 (17%)<br><br>    ==> Summe aller Kommentarzeilen 590 (27%)<br>    ==> Code-relevante Zeilen: 483 (22%)<br><br>-----------------------------------------------<br><br>In der Version vom 2024-01-07 - 15:26:03:<br>    - Gesamtanzahl der Zeilen: 2771 (100%)<br>    - davon Leerzeilen: 1408 (51%)<br>    - davon Einzelkommentarzeilen: 278 (10%)<br>    - davon Blockkommentarzeilen: 550 (20%)<br><br>    ==> Summe aller Kommentarzeilen 828 (30%)<br>    ==> Code-relevante Zeilen: 535 (19%)<br><br>-----------------------------------------------<br><br><br><br>

</details>

---

<small>

**Notice, Convnentions:**

*To generate a docstring from the VBA-Source make sure that the text to shown is located directly below the declaration line of the procedure. The text is considered completed with the first following line in the code which is not an entire comment line.  Empty lines that are to be included must also be labelled as comments.*

</small> 

---

<small>Dokumentation generiert am 2024-02-08 21:53:57 durch das  automatisierte Code-Dokumentationstool von Matthias Kader (Commit vom 2024-02-08 21:45:31: '856be9332dd54848ebbfbaa5f146bd959aa61b5e')</small> 