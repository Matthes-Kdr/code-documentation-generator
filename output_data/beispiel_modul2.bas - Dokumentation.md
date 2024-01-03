## Infos zum Script, welches für die Erstellung dieser Dokumentation verwendet wurde:

- **commit: 8ff71792281afdca59395d339dee12fa8902f63e**
- gitdir: D:\Desktop\WORKSPACE\2023-12-28_-_00_51_57_VBA-Code-Dokumentation-Generator_Selbstbau\Programmentwicklung\__DEV__CURRENT_VERSION\.git
- **message:** 
> **BUGFIX: Inputs+Outputs sind jetzt parametrisierbar und anwendbar**
- **refs: Versionsinfo_Documenter**
- tree: 78a136975972fc19a3c62f923f7fc149af326a74
- parent: 4c77e5abcc39ccd5bf24a4e0e21ee5bd39490d4b
- author: Matthias Kader <_>
- **author_date: 2024-01-03 03:29:40**



Datum der Umwandlung: 2024-01-03 03:29:59﻿


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


<!-- <div style="padding-left:1em;"> -->


<details>
    <summary>      Interne Aufrufabfolge (@PLACEHOLDER_PROCEDURE_COUNT_OF_ABRUFFOLGE@)</summary>

---


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






@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@





@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_ENTRY@


</details>


<!-- </div> -->













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


<!-- <div style="padding-left:1em;"> -->


<details>
    <summary>      Interne Aufrufabfolge (@PLACEHOLDER_PROCEDURE_COUNT_OF_ABRUFFOLGE@)</summary>

---


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






@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@





@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_ENTRY@


</details>


<!-- </div> -->













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


<!-- <div style="padding-left:1em;"> -->


<details>
    <summary>      Interne Aufrufabfolge (@PLACEHOLDER_PROCEDURE_COUNT_OF_ABRUFFOLGE@)</summary>

---


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






@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@





@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_ENTRY@


</details>


<!-- </div> -->













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


<!-- <div style="padding-left:1em;"> -->


<details>
    <summary>      Interne Aufrufabfolge (@PLACEHOLDER_PROCEDURE_COUNT_OF_ABRUFFOLGE@)</summary>

---


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






@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@





@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_ENTRY@


</details>


<!-- </div> -->













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


<!-- <div style="padding-left:1em;"> -->


<details>
    <summary>      Interne Aufrufabfolge (@PLACEHOLDER_PROCEDURE_COUNT_OF_ABRUFFOLGE@)</summary>

---


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






@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@





@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_ENTRY@


</details>


<!-- </div> -->













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


<!-- <div style="padding-left:1em;"> -->


<details>
    <summary>      Interne Aufrufabfolge (@PLACEHOLDER_PROCEDURE_COUNT_OF_ABRUFFOLGE@)</summary>

---


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






@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@





@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_ENTRY@


</details>


<!-- </div> -->













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


<!-- <div style="padding-left:1em;"> -->


<details>
    <summary>      Interne Aufrufabfolge (@PLACEHOLDER_PROCEDURE_COUNT_OF_ABRUFFOLGE@)</summary>

---


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






@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_INTRODUCTION@





@PLACEHOLDER_PROCEDURE_ABRUFFOLGE_ENTRY@


</details>


<!-- </div> -->













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

> Diese Dokumentation wurde automatisch generiert durch ein Programm, welches sich noch im Entwicklungsstadium befindet. Im folgenden werden die Modulinformationen des PYTHON-SCRIPTES aufgeführt, durch welches diese Dokumentation generiert wurde.

<details>

<summary> Modulinformationen anzeigen/verbergen.
</summary>

  @
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


@

</details>



<small>

**Notice:**

*To generate a docstring from the VBA-Source make sure that the text to shown is located directly below the declaration line of the procedure. The text is considered completed with the first following line in the code which is not an entire comment line.  Empty lines that are to be included must also be labelled as comments.*

</small> 

<small> **TODO:** Erstellt am (Datum) durch das  automatisierte Code-Dokumentationstool von .... in der Version ....</small> 
