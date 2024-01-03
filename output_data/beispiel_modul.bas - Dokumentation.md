# Code-Dokumentation: Modul '@PLACEHOLDER_INPUT_FILE@'

<!-- TODO: Platzhalter ersetzen! -->




**TODO:** Organisatorische Hinweise zur Verwendeten bzw. dokumentierten Datei.

Erstellungsdatum dieser Dokumentation: @PLACEHOLDER_TIMESTAMP_NOW@
<!-- TODO: Platzhalter ersetzen! --> 
(DAS KOMMT IN DEN TAIL!!! SIEHE FUSSZEILE!)




Die hier dokumentierte Quelldatei wurde vor dieser automatischen Dokumentationserstellung zuletzt modifiziert am @PLACEHOLDER_TIMESTAMP_SOURCEFILE@
<!-- TODO: Platzhalter ersetzen! -->








>  **ACHTUNG.** 
> Das Tool für die automatisierte Erstellung der Dokumentation ist noch nicht fertig! 
> Für weitere Infos: Siehe Schlussbemerkungen!


﻿


<!-- --------------------------------------------------------------- -->
<!-- Index / TOC -->
<!-- --------------------------------------------------------------- -->

## Index / Table Of Content

Alphabetische und verlinkte Auflistung aller Subs und Functions, die in diesem Modul verwendet werden:

* [**Modulinformationen / Modulkopf**](#sec_modulinfos)
  

  
  <!-- ---------- SUBS: --------------- -->

* [**Subs**](#sec_subs) (4)
  
  * [```bauer```](#bauer) <small>(Zeile 78)</small>
  * [```casio```](#casio) <small>(Zeile 116)</small>
  * [```liebherr```](#liebherr) <small>(Zeile 103)</small>
  * [```main```](#main) <small>(Zeile 48)</small>
  




  <!-- ---------- FUNCTIONS: --------------- -->


* [**Functions**](#sec_functions) (2)
  
  
  * [```addieren```](#addieren) <small>(Zeile 24)</small>
  * [```subtrahieren```](#subtrahieren) <small>(Zeile 34)</small>
  




  <!-- ---------- TAIL: --------------- -->


* [**Schlussbemerkungen** (inkl. Angaben zum Entwicklungszustand des Code-Dokumenter-Tools)](#sec_tail)




﻿


<a name="sec_modulinfos"></a>

## Modulbeschreibung

  Nach 3 unnoetigen Leerzeichen:




 HAllo!
 Dieses Modul beinhaltet einige Prozeduren, die für nichts sinnvoll sind...
 Aber es hat immerhin einen Programmkopf.

 Wichtige Prozeduren: Keine


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




<a name="bauer"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Sub bauer```</span><small>(Zeile 78)</small>

<div style="padding-left:2em;">

>  Deklaration erfolgte via public sub.




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
Public Sub bauer()
''' Deklaration erfolgte via public sub.

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


























﻿





<!-- --------------------------------------------------------------- -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- --------------------------------------------------------------- -->




<a name="casio"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Sub casio```</span><small>(Zeile 116)</small>

<div style="padding-left:2em;">

> *No information availible. For more information expand source code.*



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
   Sub casio()


    ''' Hier gibt skeinen docstirng

    MsgBox("Dies ist ein implizit als public gekennzeichnetes Sub.")


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
<span style="background-color: lightgrey; padding: 2px;">```Public Sub liebherr```</span><small>(Zeile 103)</small>

<div style="padding-left:2em;">

>  Deklaration erfolgte via sub.




<details>

<summary> Referenzierungen dieser Prozedur (5)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



* [```main```](#main) : <small>  Zeile 68 : ```    call liebherr``` </small>
* [```bauer```](#bauer) : <small>  Zeile 84 : ```    call liebherr``` </small>
* [```bauer```](#bauer) : <small>  Zeile 85 : ```    call liebherr ' Aufruf``` </small>
* [```bauer```](#bauer) : <small>  Zeile 87 : ```    call liebherr("ERROR") ' Aufruf waere zwar ungültig, aber Prozedur könnte ja anders aussehen!``` </small>
* [```bauer```](#bauer) : <small>  Zeile 91 : ```    var = liebherr("gvkil")``` </small>




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
   Sub liebherr()
''' Deklaration erfolgte via sub.

    MsgBox("Dies ist ein implizit als public gekennzeichnetes Sub.")


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




<a name="main"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Sub main```</span><small>(Zeile 48)</small>

<div style="padding-left:2em;">

>  Hier soll das HAuptprogramm stehen.
 Alles was als KOMMENTAR hier unter der Definitionszeile einer Funktion steht, BEVOR EINE LEERZEILE folgt, soll später als Zusammenfassung angezeigt werden in der Code-Dokumentation - also ähnlich wie im docstring bei python.




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
Private Sub main()
''' Hier soll das HAuptprogramm stehen.
''' Alles was als KOMMENTAR hier unter der Definitionszeile einer Funktion steht, BEVOR EINE LEERZEILE folgt, soll später als Zusammenfassung angezeigt werden in der Code-Dokumentation - also ähnlich wie im docstring bei python.

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




<a name="addieren"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function addieren```</span><small>(Zeile 24)</small>

<div style="padding-left:2em;">

>  Diese Funktion addiert beide Zahlen miteinander und übergibt das Ergebnis zurück.




<details>

<summary> Referenzierungen dieser Prozedur (2)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



* [```subtrahieren```](#subtrahieren) : <small>  Zeile 38 : ```    subtrahieren = addieren(a, -b) ' Parameter b wird mit -1 multipliziert übergeben``` </small>
* [```main```](#main) : <small>  Zeile 62 : ```        wert = addieren(i, i)``` </small>




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
Private Function addieren(a as integer, b as integer) as integer
''' Diese Funktion addiert beide Zahlen miteinander und übergibt das Ergebnis zurück.


    ' Addieren:
    addieren = a + b

End Function

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




<a name="subtrahieren"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function subtrahieren```</span><small>(Zeile 34)</small>

<div style="padding-left:2em;">

>  Diese Funktion subtrahiert b von a und übergibt das Ergebnis zurück.




<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



* [```main```](#main) : <small>  Zeile 63 : ```        wert = subtrahieren(i, i - 1) ' Erklärung siehe @ Func!``` </small>




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
Private Function subtrahieren(a as integer, b as integer) as integer
''' Diese Funktion subtrahiert b von a und übergibt das Ergebnis zurück.

    ' Benutze die addieren Funktion:
    subtrahieren = addieren(a, -b) ' Parameter b wird mit -1 multipliziert übergeben


end Function

```

</details>


</div>


---


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
