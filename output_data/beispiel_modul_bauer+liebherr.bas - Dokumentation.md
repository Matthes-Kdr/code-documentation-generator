# Code-Dokumentation: Modul 'beispiel_modul_bauer+liebherr.bas'



**Letzte Änderung** der Quelldatei 'beispiel_modul_bauer+liebherr.bas' vor der Generierung dieser automatischen Dokumentation: **2024-01-07 01:00**


Generierungsdatum dieser Dokumentation: **2024-01-07 02:06:15**









<!-- TODO: nur temporrary!  -->
# ZWISCHENGELAGERT ALS ZIEL-VORGABE FÜR ABRUFSEQUENZ:


**Aktuelle Bugs:**

- Aufrufabfolge wird noch nicht korrekt dargestellt 
    - (Einrückungen sind falsch!)






<details>
    <summary>      Interne Aufrufabfolge (@PLACEHOLDER_PROCEDURE_COUNT_OF_ABRUFFOLGE@)</summary>

---





@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@





@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_ENTRY@








<!-- TODO: Platzhalter platz -->
<br>
<br>
<br>
<br>
<br>
STATIC  - EXEMPLARISCHES ZIEL- OUTPUT für MAIN:

<!-- TODO: Links einfügen! gleiches prinzip wie bei  references!-->




* ```hauptfunc1```
  * ```unterfunktionA```
    * ```wiederholungsfunktion```
  * ```unterfunktionB```
* ```hauptfunc2```
* ```hauptfunc3```
* ```wiederholungsfunktion```
  * ```wiederholungsfunktion```




</details>


﻿


<!-- --------------------------------------------------------------- -->
<!-- Index / TOC -->
<!-- --------------------------------------------------------------- -->

## Index / Table Of Content

Alphabetische und verlinkte Auflistung aller Subs und Functions, die in diesem Modul verwendet werden:

* [**Modulinformationen / Modulkopf**](#sec_modulinfos)
  

  
  <!-- ---------- SUBS: --------------- -->

* [**Subs**](#sec_subs) (2)
  
  * [```bauer```](#bauer) <small>(Zeile 12)</small>
  * [```liebherr```](#liebherr) <small>(Zeile 40)</small>
  




  <!-- ---------- FUNCTIONS: --------------- -->


* [**Functions**](#sec_functions) (0)
  
  
  




  <!-- ---------- TAIL: --------------- -->


* [**Schlussbemerkungen** (inkl. Angaben zum Entwicklungszustand des Code-Dokumenter-Tools)](#sec_tail)




﻿


<a name="sec_modulinfos"></a>

## Modulbeschreibung

  
 Nur bauer + liebherr

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




<a name="bauer"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Sub bauer```</span><small>(Zeile 12)</small>

<div style="padding-left:2em;">

>  Anzahl der Referenzierungen im Modul: 0
 Anzahl weiterer internen Aufrufe : 4

 Ruft MEHRFACH die Prozedur 'liebherr' auf (insgesamt 4 mal nacheinander in verschiedenen Kontexten)





<details>

<summary> Referenzierungen dieser Prozedur (0)</summary>

<div style="padding-left:1em;">



Kein Aufruf gefunden.







</details

</div>











<!-- TODO: ABRUFABFOLGE (DEV) -->

<details>
    <summary>      Interne Aufrufabfolge (@PLACEHOLDER_PROCEDURE_COUNT_OF_ABRUFFOLGE@)</summary>

---


@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@


<!-- <div style="padding-left:1em;"> -->








- ```liebherr``` <small> : [Zeile 22] : ```    call liebherr``` </small>
! Abschlusstext beim else @ line : 1299


- ```liebherr``` <small> : [Zeile 23] : ```    call liebherr ' Aufruf``` </small>
! Abschlusstext beim else @ line : 1299(nichts weiter... Hier KEINE indentions! @ line :1125)


- ```liebherr``` <small> : [Zeile 25] : ```    call liebherr("ERROR") ' Aufruf waere zwar ungültig, aber Prozedur könnte ja anders aussehen!``` </small>
! Abschlusstext beim else @ line : 1299(nichts weiter... Hier KEINE indentions! @ line :1125)


- ```liebherr``` <small> : [Zeile 28] : ```    var = liebherr("gvkil")``` </small>
! Abschlusstext beim else @ line : 1299(nichts weiter... Hier KEINE indentions! @ line :1125)
 --- Abschluss dieser Doc / keine weiteren Aufrufe @ line :1242






<!-- </div> -->








</details>







<details>
    <summary>      Source Code</summary>

---

```
Public Sub bauer()
' Anzahl der Referenzierungen im Modul: 0
' Anzahl weiterer internen Aufrufe : 4
'
''' Ruft MEHRFACH die Prozedur 'liebherr' auf (insgesamt 4 mal nacheinander in verschiedenen Kontexten)
'

    MsgBox("Dies ist ein explizit als public gekennzeichnetes Sub.")

    ' Aufruf:
    call liebherr
    call liebherr ' Aufruf
    
    call liebherr("ERROR") ' Aufruf waere zwar ungültig, aber Prozedur könnte ja anders aussehen!

    ' Wiederum unügltig:
    var = liebherr("gvkil")





End Sub

```

</details>


</div>


---


<!-- --------------------------------------------------------------- -->


























﻿





<!-- --------------------------------------------------------------- -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- --------------------------------------------------------------- -->




<a name="liebherr"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Sub liebherr```</span><small>(Zeile 40)</small>

<div style="padding-left:2em;">

>  Anzahl der Referenzierungen im Modul: 6
 Anzahl weiterer internen Aufrufe : 0

 ' Ruft keine weitere Prozedur auf.




<details>

<summary> Referenzierungen dieser Prozedur (4)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



* [```bauer```](#bauer) : <small>  Zeile 22 : ```    call liebherr``` </small>
* [```bauer```](#bauer) : <small>  Zeile 23 : ```    call liebherr ' Aufruf``` </small>
* [```bauer```](#bauer) : <small>  Zeile 25 : ```    call liebherr("ERROR") ' Aufruf waere zwar ungültig, aber Prozedur könnte ja anders aussehen!``` </small>
* [```bauer```](#bauer) : <small>  Zeile 28 : ```    var = liebherr("gvkil")``` </small>




</details

</div>











<!-- TODO: ABRUFABFOLGE (DEV) -->

<details>
    <summary>      Interne Aufrufabfolge (@PLACEHOLDER_PROCEDURE_COUNT_OF_ABRUFFOLGE@)</summary>

---


@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@


<!-- <div style="padding-left:1em;"> -->





! Abschlusstext beim else @ line : 1299






<!-- </div> -->








</details>







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


























﻿
<!-- -------------------------------------------------- -->
<!-- SECTION-START : FUNCTIONS -->
<!-- -------------------------------------------------- -->

<a name="sec_functions"></a>

## Functions


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

  <br>Created on: Fri, 2023-12-29 (00:45:39)<br><br><br>@author: Matthias Kader<br><br><br>Für Ziel und Ablauf des Scriptes siehe MArkdown im Verzeichnis ../Tests/Programmablauf.html<br><br><br><br><br>### Fertig implementiert:<br><br>• Implementierung Inhaltsverzeichnis / Index<br><br>• Gesamtlayout inkl. Titel, Zwischenüberschriften für einzelne Sections<br><br>• Einbindung vom Programmkopf-Docstring<br><br>• Implementierung von References-Durchsuchungen<br><br>• Implementierung eines Exportes zu HTML<br><br><br><br>• Einbindung organisatorischer Daten bzgl. des zu dokumentierenden Codes und des verwendeten Skripts zum Dokumentieren<br><br>• Implementierung der Calling Sequence: Für jede Prozedur: Aufzählung der Aufrufe anderer, in dieser Dokumentation behandelten Prozeduren.<br>Bislang wird eine einfache Auflistung gegeben. Perspektivisch wäre eine rekursiver Ansatz denkbar, sodass je Aufruf wieder alle Aufrufe innerhabl dieser Prozedur gelistet werden können usw...<br><br><br><br>### TODO: Größere TODOS:<br><br><br>• Call Sequenz / Calling Sequence:<br><br>Schön (Ausblick) wäre auch ein weiterer Unterpunkt pro Prozedur, in der die Aufrufabfolge hervorgeht.<br>Idee ist etwas wie die Aufrufebenen-Auflistung beim Noten-Converter-Programm, d.h. ausgehend von einer Prozedur soll eine Liste stehen der Aufrufe von weiteren Prozeduren die aufgerufen werden (und die in diesem Dokument auch dokumentiert werden... also keine Builtins o.ä.). Im Idealfall kann jeder Punkt dieser Liste wiederum erweitert/expanded werden, darin ist dann wiederum die Liste von DIESER AUFGERUFENEN Funktion drin usw... Rekursiv. Jede Methode, die einmal so dokumentiert wurde kann weiter verwendet werden per Direktzugriff....<br><br><br><br>### AUSBLICK für später und in schön:<br><br>• Index an der Seite wie eine NavBar zum einzelnd scrollen<br><br><br><br># TODO / CURRENT DEV:<br>    Aufrufebenen im 'main' untersuchen, inkl. Rekursive Auflistung aller Calls.<br><br>

</details>

---

<small>

**Notice, Convnentions:**

*To generate a docstring from the VBA-Source make sure that the text to shown is located directly below the declaration line of the procedure. The text is considered completed with the first following line in the code which is not an entire comment line.  Empty lines that are to be included must also be labelled as comments.*

</small> 

---

<small>Dokumentation generiert am 2024-01-07 02:06:15 durch das  automatisierte Code-Dokumentationstool von Matthias Kader (Commit vom 2024-01-07 01:50:52: '715b0b6485123fbdf7797b023caf84aefe35490c')</small> 
