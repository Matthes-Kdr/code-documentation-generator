# Code-Dokumentation: Modul 'beispiel_modul.bas'



**Letzte Änderung** der Quelldatei 'beispiel_modul.bas' vor der Generierung dieser automatischen Dokumentation: **2024-01-07 11:28**


Generierungsdatum dieser Dokumentation: **2024-01-07 23:16:28**









<!-- TODO: nur temporrary!  -->
# ZWISCHENGELAGERT ALS ZIEL-VORGABE FÜR ABRUFSEQUENZ:


**Aktuelle Bugs:**

- Probleme mit inkorrekter Einrückungen nach diversen Aufrufebenen scheint behoben zu sein 2024-01-07 - 22:47:41 (Teste nochmal!)
  
- Es werden nicht alle Aufrufe erkannt (ODER??!)
    
  - siehe beispiel_modul.bas --> liebherr : sollte 6 referenzierungen haben, es werden nur 5 dokumentiert... Zeile 106 fehlt:  ```var = liebherr```

  - wäre nicht tragisch, weil liebherr in diesem Syntax bei VBA nur eine Funktion mit Rückgabewert sein kann, und dann könnte man auch liebherr() schreiben, das würde erkannt werden. Allerdings funktioniert auch die Schreibweise ohne KLammern, weshalb sie auch erkannt werden sollte (auch wenn ich es nie so schreiben wollen würde...)


﻿


<!-- --------------------------------------------------------------- -->
<!-- Index / TOC -->
<!-- --------------------------------------------------------------- -->

## Index / Table Of Content

Alphabetische und verlinkte Auflistung aller Subs und Functions, die in diesem Modul verwendet werden:

* [**Modulinformationen / Modulkopf**](#sec_modulinfos)
  

  
  <!-- ---------- SUBS: --------------- -->

* [**Subs**](#sec_subs) (4)
  
  - [```bauer```](#bauer) : <small>  [Zeile 90]  </small>


  - [```casio```](#casio) : <small>  [Zeile 136]  </small>


  - [```liebherr```](#liebherr) : <small>  [Zeile 119]  </small>


  - [```main```](#main) : <small>  [Zeile 54]  </small>


  




  <!-- ---------- FUNCTIONS: --------------- -->


* [**Functions**](#sec_functions) (2)
  
  
  - [```addieren```](#addieren) : <small>  [Zeile 16]  </small>


  - [```subtrahieren```](#subtrahieren) : <small>  [Zeile 34]  </small>


  




  <!-- ---------- TAIL: --------------- -->


* [**Schlussbemerkungen** (inkl. Angaben zum Entwicklungszustand des Code-Dokumenter-Tools)](#sec_tail)




﻿


<a name="sec_modulinfos"></a>

## Modulbeschreibung

  
 Beispiel Modul zum Testen der Dokumentation der Abruffolge.

 ## !!! # ACHTUNG BUGS: !!!
 > Sub 'liebherr' wird insgesamt 6x referenziert, nicht 5x, wie aktuell dokumentiert!

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
<span style="background-color: lightgrey; padding: 2px;">```Public Sub bauer```</span><small>(Zeile 90)</small>






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
    <summary>      Interne Aufrufabfolge (4)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```liebherr```](#liebherr) : <small>  [Zeile 100] : ```    call liebherr``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```liebherr```](#liebherr) : <small>  [Zeile 101] : ```    call liebherr ' Aufruf``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```liebherr```](#liebherr) : <small>  [Zeile 103] : ```    call liebherr("ERROR") ' Aufruf waere zwar ungültig, aber Prozedur könnte ja anders aussehen!``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```liebherr```](#liebherr) : <small>  [Zeile 107] : ```    var = liebherr("gvkil")``` </small>



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
<span style="background-color: lightgrey; padding: 2px;">```Public Sub casio```</span><small>(Zeile 136)</small>






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
<span style="background-color: lightgrey; padding: 2px;">```Public Sub liebherr```</span><small>(Zeile 119)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Anzahl der Referenzierungen im Modul: 6
 Anzahl weiterer internen Aufrufe : 0

 ' Ruft keine weitere Prozedur auf.







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (5)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```main```](#main) : <small>  [Zeile 80] : ```    call liebherr``` </small>

- [```bauer```](#bauer) : <small>  [Zeile 100] : ```    call liebherr``` </small>

- [```bauer```](#bauer) : <small>  [Zeile 101] : ```    call liebherr ' Aufruf``` </small>

- [```bauer```](#bauer) : <small>  [Zeile 103] : ```    call liebherr("ERROR") ' Aufruf waere zwar ungültig, aber Prozedur könnte ja anders aussehen!``` </small>

- [```bauer```](#bauer) : <small>  [Zeile 107] : ```    var = liebherr("gvkil")``` </small>





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
<span style="background-color: lightgrey; padding: 2px;">```Private Sub main```</span><small>(Zeile 54)</small>






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
    <summary>      Interne Aufrufabfolge (3)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```addieren```](#addieren) : <small>  [Zeile 74] : ```        wert = addieren(i, i)``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```subtrahieren```](#subtrahieren) : <small>  [Zeile 75] : ```        wert = subtrahieren(i, i - 1) ' Erklärung siehe @ Func!``` </small>


  - [```addieren```](#addieren) : <small>  [Zeile 44] : ```    subtrahieren = addieren(a, -b) ' Parameter b wird mit -1 multipliziert übergeben``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```liebherr```](#liebherr) : <small>  [Zeile 80] : ```    call liebherr``` </small>



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
<span style="background-color: lightgrey; padding: 2px;">```Private Function addieren```</span><small>(Zeile 16)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Anzahl der Referenzierungen im Modul: 2
 Anzahl weiterer internen Aufrufe : 0

 Diese Funktion addiert beide Zahlen miteinander und übergibt das Ergebnis zurück.

 Ruft keine weitere Prozedur auf.







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (2)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```subtrahieren```](#subtrahieren) : <small>  [Zeile 44] : ```    subtrahieren = addieren(a, -b) ' Parameter b wird mit -1 multipliziert übergeben``` </small>

- [```main```](#main) : <small>  [Zeile 74] : ```        wert = addieren(i, i)``` </small>





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
<span style="background-color: lightgrey; padding: 2px;">```Private Function subtrahieren```</span><small>(Zeile 34)</small>






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



- [```main```](#main) : <small>  [Zeile 75] : ```        wert = subtrahieren(i, i - 1) ' Erklärung siehe @ Func!``` </small>





</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (1)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```addieren```](#addieren) : <small>  [Zeile 44] : ```    subtrahieren = addieren(a, -b) ' Parameter b wird mit -1 multipliziert übergeben``` </small>



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

  <br>Created on: Fri, 2023-12-29 (00:45:39)<br><br><br>@author: Matthias Kader<br><br><br>Für generelles Ziel und Ablauf des Scriptes siehe MArkdown im Verzeichnis ../Tests/Programmablauf.html<br><br>Wichtige Details siehe am Ende dieses docstrings.<br><br><br><br><br>### Fertig implementiert:<br><br>• Inhaltsverzeichnis / Index<br><br>• Gesamtlayout inkl. Titel, Zwischenüberschriften für einzelne Sections<br><br>• Aufführen  des modulweiten Programmkopf-Docstring in der generierten Dokumentation<br><br>• Aufführen der References-Durchsuchungen (Wo wird die Prozedur aufgerufen?) in der generierten Dokumentation<br><br>• Sofortiger Export der MD-Datei in eine  HTML-Datei<br><br>• Aufführen der organisatorischer Daten bzgl. des zu dokumentierenden Codes und des verwendeten Skripts zum Dokumentieren in der generierten Dokumentation<br><br>• Aufführen der Calling Sequence (Aufrufabfolge / Aufrufebenen) innerhalb jeder Prozedur in der generierten Dokumentation: Aufzählung der Aufrufe anderer, in dieser Dokumentation behandelten Prozeduren. Inklusive rekursive geschachtelte Liste, welche Aufrufe jeweils in den aufgerufenen Prozeduren erfolgen.<br><br><br><br>### TODOS:<br><br>• Bereitstellung einer einfachen GUI / HMI, um Input- und Output Pfade zu parametrisieren<br><br>• Bugfix: Aufrufebenen ab Unterebene x: Behebung der Formatierungsprobleme (siehe beispiel_modul1.bas --> notengriffe_erzeugen --> getFilePath)<br><br><br><br>### AUSBLICK für später und in schön:<br><br>• Index an der Seite wie eine NavBar zum einzelnd scrollen<br><br>• Ermöglichung von Berücksichtigung weiterer Module innerhalb der Dokumentation<br>    <br>    • z. B. 2 VBA-Module innerhalb eines Projektes, wobei Prozeduren von Modul1  andere Prozeduren aus Modul2 aufrufen.<br><br>        • Erstmal nur als Verweis  (Mögl. Ansatz included = "Modul1.*" ohne rekursive Auflistung derer Aufrufe... oder eben mit... bestenfalls auch das parametrisierbar)<br><br>• Dokumentation von weiteren PRogrammiersprachen<br><br>    • OK --> VBA<br>    • Nächste Prio: C++ / Arduino<br>    • Letzte Prio: Python (v.a. für den Ablaufsequence sehr hilfreich, für den rest gibt es pdoc...)<br><br><br><br><br><br><br><br><br># =============================================================================<br>#### Wichtige Aufrufreihenfolge der Methode innerhalb dieses Python-Scriptes zur Erstellung der Dokumentation der Aufrufreihenfolge der zu dokumentierenden VBA-Prozeduren: ####<br># =============================================================================<br><br>Es werden zunächst alle Prozeduren komplett analysiert, erst danach werden wiederum alle Prozeduren komplett dokumentiert. Für beide Vorgänge erfolgt dies in einer Methode auf Objektebene, wobei diese jeweilige MEthode in beiden Fällen aus einer Klassenmethode aufgerufen wird, in der über die einzelnen Prozedur-Objekte innerhalb dieser Klasse iteriert wird:<br><br>- analyse_call_sequence(cls)<br>    - analyse_calling_sequence_in_one_proc(self)<br>- prepare_all_call_sequences_docs(cls)<br>    - prepare_single_call_sequence_docs(cls)<br><br>(hierfür wäre das entwickelte Tool  übrigens eine tolle Anwendung gewesen, sofern sie später auch mal Python-Syntax dokumentieren könnte :-) )<br><br><br><br><br><br><br><br># =============================================================================<br>#### Unwichtige Nebensächlichkeiten: ####<br># =============================================================================<br><br>In der Version vom 2024-01-07 - 15:26:03:<br>    - Gesamtanzahl der Zeilen: 2771 (100%)<br>    - davon Leerzeilen: 1408 (51%)<br>    - davon Einzelkommentarzeilen: 278 (10%)<br>    - davon Blockkommentarzeilen: 550 (20%)<br><br>    ==> Summe aller Kommentarzeilen 828 (30%)<br>    ==> Code-relevante Zeilen: 535 (19%)<br><br><br><br><br>

</details>

---

<small>

**Notice, Convnentions:**

*To generate a docstring from the VBA-Source make sure that the text to shown is located directly below the declaration line of the procedure. The text is considered completed with the first following line in the code which is not an entire comment line.  Empty lines that are to be included must also be labelled as comments.*

</small> 

---

<small>Dokumentation generiert am 2024-01-07 23:16:28 durch das  automatisierte Code-Dokumentationstool von Matthias Kader (Commit vom 2024-01-07 15:46:48: 'b0fc982eeee2e7e5a2d7bbc29900ebd1033ed48d')</small> 
