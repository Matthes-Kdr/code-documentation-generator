# Code-Dokumentation: Modul 'beispiel_modul2.bas'



**Letzte Änderung** der Quelldatei 'beispiel_modul2.bas' vor der Generierung dieser automatischen Dokumentation: **2024-01-03 00:06**


Generierungsdatum dieser Dokumentation: **2024-01-07 06:59:43**









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

* [**Subs**](#sec_subs) (7)
  
  * [```hauptfunc1```](#hauptfunc1) <small>(Zeile 32)</small>
  * [```hauptfunc2```](#hauptfunc2) <small>(Zeile 40)</small>
  * [```hauptfunc3```](#hauptfunc3) <small>(Zeile 49)</small>
  * [```main```](#main) <small>(Zeile 60)</small>
  * [```unterfunktionA```](#unterfunktionA) <small>(Zeile 19)</small>
  * [```unterfunktionB```](#unterfunktionB) <small>(Zeile 25)</small>
  * [```wiederholungsfunktion```](#wiederholungsfunktion) <small>(Zeile 13)</small>
  




  <!-- ---------- FUNCTIONS: --------------- -->


* [**Functions**](#sec_functions) (0)
  
  
  




  <!-- ---------- TAIL: --------------- -->


* [**Schlussbemerkungen** (inkl. Angaben zum Entwicklungszustand des Code-Dokumenter-Tools)](#sec_tail)




﻿


<a name="sec_modulinfos"></a>

## Modulbeschreibung

 Demo Modul zum Testen der Aufrufebenen / Aufruf-Abfolge / Ablauf

 Das Modul besteht aus insgesamt 7 Subs in folgender Aufrufhierarchie:
 ' 1 x main
 ' 3 x hauptfunc
 ' 2 x unterfuncktion
 ' 1 x wiederholungsfunktion



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




<a name="hauptfunc1"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Sub hauptfunc1```</span><small>(Zeile 32)</small>

<div style="padding-left:2em;">

>  weiterere 2 Aufrufe in dieser Funktion




<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



* [```main```](#main) : <small>  Zeile 62 : ```    call hauptfunc1``` </small>




</details

</div>











<!-- TODO: ABRUFABFOLGE (DEV) -->

<details>
    <summary>      Interne Aufrufabfolge (@PLACEHOLDER_PROCEDURE_COUNT_OF_ABRUFFOLGE@)</summary>

---


@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@


<!-- <div style="padding-left:1em;"> -->








- ```unterfunktionA``` <small> : [Zeile 34] : ```    call unterfunktionA()``` </small>




  - ```wiederholungsfunktion``` <small> : [Zeile 21] : ```    call wiederholungsfunktion()``` </small>


    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>





- ```unterfunktionB``` <small> : [Zeile 35] : ```    call unterfunktionB()``` </small>


  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>






<!-- </div> -->








</details>







<details>
    <summary>      Source Code</summary>

---

```
Sub hauptfunc1()
    ' weiterere 2 Aufrufe in dieser Funktion
    call unterfunktionA()
    call unterfunktionB()
end sub

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




<a name="hauptfunc2"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Sub hauptfunc2```</span><small>(Zeile 40)</small>

<div style="padding-left:2em;">

>  weiterer Aufruf von unterfunktionA in dieser Funktion




<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



* [```main```](#main) : <small>  Zeile 63 : ```    call hauptfunc2``` </small>




</details

</div>











<!-- TODO: ABRUFABFOLGE (DEV) -->

<details>
    <summary>      Interne Aufrufabfolge (@PLACEHOLDER_PROCEDURE_COUNT_OF_ABRUFFOLGE@)</summary>

---


@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@


<!-- <div style="padding-left:1em;"> -->








- ```unterfunktionA``` <small> : [Zeile 42] : ```    call unterfunktionA()``` </small>




  - ```wiederholungsfunktion``` <small> : [Zeile 21] : ```    call wiederholungsfunktion()``` </small>


    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>






<!-- </div> -->








</details>







<details>
    <summary>      Source Code</summary>

---

```
Sub hauptfunc2()
    ' weiterer Aufruf von unterfunktionA in dieser Funktion
    call unterfunktionA()
    
    msgbox("standalone")
end sub

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




<a name="hauptfunc3"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Sub hauptfunc3```</span><small>(Zeile 49)</small>

<div style="padding-left:2em;">

>  kein weiterer Aufruf in dieser Funktion




<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



* [```main```](#main) : <small>  Zeile 64 : ```    call hauptfunc3``` </small>




</details

</div>











<!-- TODO: ABRUFABFOLGE (DEV) -->

<details>
    <summary>      Interne Aufrufabfolge (@PLACEHOLDER_PROCEDURE_COUNT_OF_ABRUFFOLGE@)</summary>

---


@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@


<!-- <div style="padding-left:1em;"> -->






- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>






<!-- </div> -->








</details>







<details>
    <summary>      Source Code</summary>

---

```
Sub hauptfunc3()
    ' kein weiterer Aufruf in dieser Funktion
    msgbox("standalone")
end sub

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




<a name="main"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Sub main```</span><small>(Zeile 60)</small>

<div style="padding-left:2em;">

>  Sequentieller Aufruf aller 3 Hauptfunktionen




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








- ```hauptfunc1``` <small> : [Zeile 62] : ```    call hauptfunc1``` </small>




  - ```unterfunktionA``` <small> : [Zeile 34] : ```    call unterfunktionA()``` </small>




    - ```wiederholungsfunktion``` <small> : [Zeile 21] : ```    call wiederholungsfunktion()``` </small>


      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>





  - ```unterfunktionB``` <small> : [Zeile 35] : ```    call unterfunktionB()``` </small>


    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>





- ```hauptfunc2``` <small> : [Zeile 63] : ```    call hauptfunc2``` </small>




  - ```unterfunktionA``` <small> : [Zeile 42] : ```    call unterfunktionA()``` </small>




    - ```wiederholungsfunktion``` <small> : [Zeile 21] : ```    call wiederholungsfunktion()``` </small>


      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>





- ```hauptfunc3``` <small> : [Zeile 64] : ```    call hauptfunc3``` </small>


  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>





- ```wiederholungsfunktion``` <small> : [Zeile 65] : ```    call wiederholungsfunktion``` </small>


  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>






<!-- </div> -->








</details>







<details>
    <summary>      Source Code</summary>

---

```
Sub main()
    ' Sequentieller Aufruf aller 3 Hauptfunktionen
    call hauptfunc1
    call hauptfunc2
    call hauptfunc3
    call wiederholungsfunktion
end sub

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




<a name="unterfunktionA"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Sub unterfunktionA```</span><small>(Zeile 19)</small>

<div style="padding-left:2em;">

>  weiterer Aufruf einer wiederholungsfunktion:




<details>

<summary> Referenzierungen dieser Prozedur (2)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



* [```hauptfunc1```](#hauptfunc1) : <small>  Zeile 34 : ```    call unterfunktionA()``` </small>
* [```hauptfunc2```](#hauptfunc2) : <small>  Zeile 42 : ```    call unterfunktionA()``` </small>




</details

</div>











<!-- TODO: ABRUFABFOLGE (DEV) -->

<details>
    <summary>      Interne Aufrufabfolge (@PLACEHOLDER_PROCEDURE_COUNT_OF_ABRUFFOLGE@)</summary>

---


@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@


<!-- <div style="padding-left:1em;"> -->








- ```wiederholungsfunktion``` <small> : [Zeile 21] : ```    call wiederholungsfunktion()``` </small>


  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>






<!-- </div> -->








</details>







<details>
    <summary>      Source Code</summary>

---

```
Sub unterfunktionA()
    ' weiterer Aufruf einer wiederholungsfunktion:
    call wiederholungsfunktion()
end sub

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




<a name="unterfunktionB"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Sub unterfunktionB```</span><small>(Zeile 25)</small>

<div style="padding-left:2em;">

>  kein externer aufruf




<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



* [```hauptfunc1```](#hauptfunc1) : <small>  Zeile 35 : ```    call unterfunktionB()``` </small>




</details

</div>











<!-- TODO: ABRUFABFOLGE (DEV) -->

<details>
    <summary>      Interne Aufrufabfolge (@PLACEHOLDER_PROCEDURE_COUNT_OF_ABRUFFOLGE@)</summary>

---


@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@


<!-- <div style="padding-left:1em;"> -->






- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>






<!-- </div> -->








</details>







<details>
    <summary>      Source Code</summary>

---

```
Sub unterfunktionB()
    ' kein externer aufruf
    x = 42
end sub

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




<a name="wiederholungsfunktion"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Sub wiederholungsfunktion```</span><small>(Zeile 13)</small>

<div style="padding-left:2em;">

>  Func für standard-aufgaben




<details>

<summary> Referenzierungen dieser Prozedur (2)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



* [```unterfunktionA```](#unterfunktionA) : <small>  Zeile 21 : ```    call wiederholungsfunktion()``` </small>
* [```main```](#main) : <small>  Zeile 65 : ```    call wiederholungsfunktion``` </small>




</details

</div>











<!-- TODO: ABRUFABFOLGE (DEV) -->

<details>
    <summary>      Interne Aufrufabfolge (@PLACEHOLDER_PROCEDURE_COUNT_OF_ABRUFFOLGE@)</summary>

---


@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@


<!-- <div style="padding-left:1em;"> -->






- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>






<!-- </div> -->








</details>







<details>
    <summary>      Source Code</summary>

---

```
Sub wiederholungsfunktion()
    ' Func für standard-aufgaben
    msgbox("Standardfunktion")
end sub

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

<small>Dokumentation generiert am 2024-01-07 06:59:43 durch das  automatisierte Code-Dokumentationstool von Matthias Kader (Commit vom 2024-01-07 05:01:04: '9deb09e301d564a4c3b451128fc9eab5d2d78122')</small> 
