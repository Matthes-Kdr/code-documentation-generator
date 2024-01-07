# Code-Dokumentation: Modul 'beispiel_modul_bauer+liebherr.bas'



**Letzte Änderung** der Quelldatei 'beispiel_modul_bauer+liebherr.bas' vor der Generierung dieser automatischen Dokumentation: **2024-01-07 01:00**


Generierungsdatum dieser Dokumentation: **2024-01-07 11:23:31**









<!-- TODO: nur temporrary!  -->
# ZWISCHENGELAGERT ALS ZIEL-VORGABE FÜR ABRUFSEQUENZ:


**Aktuelle Bugs:**

- Aufrufabfolge wird in manchen Fällen noch nicht ganz korrekt dargestellt 
  - siehe beispiel_modul1.bas --> notengriffe_erzeugen --> getFilePath
  - Scheint v.a. nach  vielen Unteraufrufen aufzutreten...
  - Inhaltlich aber nicht falsch... (ODER????!)



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


  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>




- ```liebherr``` <small> : [Zeile 23] : ```    call liebherr ' Aufruf``` </small>


  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>





- ```liebherr``` <small> : [Zeile 25] : ```    call liebherr("ERROR") ' Aufruf waere zwar ungültig, aber Prozedur könnte ja anders aussehen!``` </small>


  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>





- ```liebherr``` <small> : [Zeile 28] : ```    var = liebherr("gvkil")``` </small>


  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>






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






- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>






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

  <br>Created on: Fri, 2023-12-29 (00:45:39)<br><br><br>@author: Matthias Kader<br><br><br>Für generelles Ziel und Ablauf des Scriptes siehe MArkdown im Verzeichnis ../Tests/Programmablauf.html<br><br>Wichtige Details siehe am Ende dieses docstrings.<br><br><br><br><br>### Fertig implementiert:<br><br>• Inhaltsverzeichnis / Index<br><br>• Gesamtlayout inkl. Titel, Zwischenüberschriften für einzelne Sections<br><br>• Aufführen  des modulweiten Programmkopf-Docstring in der generierten Dokumentation<br><br>• Aufführen der References-Durchsuchungen (Wo wird die Prozedur aufgerufen?) in der generierten Dokumentation<br><br>• Sofortiger Export der MD-Datei in eine  HTML-Datei<br><br>• Aufführen der organisatorischer Daten bzgl. des zu dokumentierenden Codes und des verwendeten Skripts zum Dokumentieren in der generierten Dokumentation<br><br>• Aufführen der Calling Sequence (Aufrufabfolge / Aufrufebenen) innerhalb jeder Prozedur in der generierten Dokumentation: Aufzählung der Aufrufe anderer, in dieser Dokumentation behandelten Prozeduren. Inklusive rekursive geschachtelte Liste, welche Aufrufe jeweils in den aufgerufenen Prozeduren erfolgen.<br><br><br><br>### AUSBLICK für später und in schön:<br><br>• Optimierung der Darstellung der Aufrufebenen: Verlinkung der PRozeduren, genau wie bei den References<br><br>• Index an der Seite wie eine NavBar zum einzelnd scrollen<br><br>• Bugfix: Aufrufebenen ab Unterebene x: Behebung der Formatierungsprobleme (siehe beispiel_modul1.bas --> notengriffe_erzeugen --> getFilePath)<br><br><br><br><br><br># =============================================================================<br>#### Wichtige Aufrufreihenfolge der Methode innerhalb dieses Python-Scriptes zur Erstellung der Dokumentation der Aufrufreihenfolge der zu dokumentierenden VBA-Prozeduren: ####<br># =============================================================================<br><br>Es werden zunächst alle Prozeduren komplett analysiert, erst danach werden wiederum alle Prozeduren komplett dokumentiert. Für beide Vorgänge erfolgt dies in einer Methode auf Objektebene, wobei diese jeweilige MEthode in beiden Fällen aus einer Klassenmethode aufgerufen wird, in der über die einzelnen Prozedur-Objekte innerhalb dieser Klasse iteriert wird:<br><br>- analyse_call_sequence(cls)<br>    - analyse_calling_sequence_in_one_proc(self)<br>- prepare_all_call_sequences_docs(cls)<br>    - prepare_single_call_sequence_docs(cls)<br><br>(hierfür wäre das entwickelte Tool  übrigens eine tolle Anwendung gewesen, sofern sie später auch mal Python-Syntax dokumentieren könnte :-) )<br><br><br><br><br>

</details>

---

<small>

**Notice, Convnentions:**

*To generate a docstring from the VBA-Source make sure that the text to shown is located directly below the declaration line of the procedure. The text is considered completed with the first following line in the code which is not an entire comment line.  Empty lines that are to be included must also be labelled as comments.*

</small> 

---

<small>Dokumentation generiert am 2024-01-07 11:23:31 durch das  automatisierte Code-Dokumentationstool von Matthias Kader (Commit vom 2024-01-07 11:13:09: '54c86f1314d7e8ea2ef2dd5da97908caf2bfcb80')</small> 
