# Code-Dokumentation: Modul 'beispiel_modul1.bas'



**Letzte Änderung** der Quelldatei 'beispiel_modul1.bas' vor der Generierung dieser automatischen Dokumentation: **2023-11-17 00:52**


Generierungsdatum dieser Dokumentation: **2024-02-07 21:06:14**





﻿


<!-- --------------------------------------------------------------- -->
<!-- Index / TOC -->
<!-- --------------------------------------------------------------- -->

## Index / Table Of Content

Alphabetische und verlinkte Auflistung aller Subs und Functions, die in diesem Modul verwendet werden:

* [**Modulinformationen / Modulkopf**](#sec_modulinfos)
  

  
  <!-- ---------- SUBS: --------------- -->

* [**Subs**](#sec_subs) (16)
  
  - [```InsertImageFromFileToCell```](#InsertImageFromFileToCell) : <small>  [Zeile 833]  </small>


  - [```SeitenformatNextPosition```](#SeitenformatNextPosition) : <small>  [Zeile 775]  </small>


  - [```THIS_VERSION__1_17```](#THIS_VERSION__1_17) : <small>  [Zeile 1220]  </small>


  - [```applyLayoutPrintArea```](#applyLayoutPrintArea) : <small>  [Zeile 247]  </small>


  - [```applyLayoutToAllPages```](#applyLayoutToAllPages) : <small>  [Zeile 136]  </small>


  - [```applyLayoutToSinglePage```](#applyLayoutToSinglePage) : <small>  [Zeile 161]  </small>


  - [```erstelleKonvertierteSeite```](#erstelleKonvertierteSeite) : <small>  [Zeile 655]  </small>


  - [```init```](#init) : <small>  [Zeile 119]  </small>


  - [```insertFrameLineBelow```](#insertFrameLineBelow) : <small>  [Zeile 1011]  </small>


  - [```insertSheet```](#insertSheet) : <small>  [Zeile 274]  </small>


  - [```mergeCells```](#mergeCells) : <small>  [Zeile 320]  </small>


  - [```modifyReplacerIfSonderzeichen```](#modifyReplacerIfSonderzeichen) : <small>  [Zeile 535]  </small>


  - [```modifyReplacerIfTaktwechsel```](#modifyReplacerIfTaktwechsel) : <small>  [Zeile 578]  </small>


  - [```notengriffeErzeugen```](#notengriffeErzeugen) : <small>  [Zeile 1169]  </small>


  - [```printLastSheetsAsPdf```](#printLastSheetsAsPdf) : <small>  [Zeile 1081]  </small>


  - [```resetWorkbook```](#resetWorkbook) : <small>  [Zeile 87]  </small>


  




  <!-- ---------- FUNCTIONS: --------------- -->


* [**Functions**](#sec_functions) (17)
  
  
  - [```addSeitenformat```](#addSeitenformat) : <small>  [Zeile 964]  </small>


  - [```areaEmpty```](#areaEmpty) : <small>  [Zeile 800]  </small>


  - [```existsSheetname```](#existsSheetname) : <small>  [Zeile 300]  </small>


  - [```getCountsOfColumnsToPrint```](#getCountsOfColumnsToPrint) : <small>  [Zeile 235]  </small>


  - [```getFilePath```](#getFilePath) : <small>  [Zeile 338]  </small>


  - [```getModifiedKeyIfTaktwechsel```](#getModifiedKeyIfTaktwechsel) : <small>  [Zeile 622]  </small>


  - [```getProjectTitle```](#getProjectTitle) : <small>  [Zeile 79]  </small>


  - [```getReplacerCount```](#getReplacerCount) : <small>  [Zeile 482]  </small>


  - [```getReplacerEmpty```](#getReplacerEmpty) : <small>  [Zeile 516]  </small>


  - [```getReplacerNote```](#getReplacerNote) : <small>  [Zeile 451]  </small>


  - [```getReplacerPost```](#getReplacerPost) : <small>  [Zeile 500]  </small>


  - [```getReplacerPre1```](#getReplacerPre1) : <small>  [Zeile 378]  </small>


  - [```getReplacerPre2```](#getReplacerPre2) : <small>  [Zeile 441]  </small>


  - [```getTypeOfFollowingSeperatorLine```](#getTypeOfFollowingSeperatorLine) : <small>  [Zeile 1044]  </small>


  - [```identifyLastRow```](#identifyLastRow) : <small>  [Zeile 867]  </small>


  - [```konvertiereZeile```](#konvertiereZeile) : <small>  [Zeile 696]  </small>


  - [```newSeitenformat```](#newSeitenformat) : <small>  [Zeile 896]  </small>


  




  <!-- ---------- TAIL: --------------- -->


* [**Schlussbemerkungen** (inkl. Angaben zum Entwicklungszustand des Code-Dokumenter-Tools)](#sec_tail)




﻿


<a name="sec_modulinfos"></a>

## Modulbeschreibung

 @author: Matthias Kader

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



<a name="InsertImageFromFileToCell"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Sub InsertImageFromFileToCell```</span><small>(Zeile 833)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Private Sub InsertImageFromFileToCell(zeile As Integer, spalte As Integer, image_path As String)







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```konvertiereZeile```](#konvertiereZeile) : <small>  [Zeile 744] : ```                Call InsertImageFromFileToCell(oSeite.excelZeilenNr, zielspalte, pathImage, oReplacer.cellFitTo)``` </small>





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
Private Sub InsertImageFromFileToCell(zeile As Integer, spalte As Integer, image_path As String, Optional fitTo As String = "Width")
' Private Sub InsertImageFromFileToCell(zeile As Integer, spalte As Integer, image_path As String)

    'Dim image_path As Variant
    Dim oPic As Shape
    Dim img As Variant
    wsResult.Cells(zeile, spalte).Select  ' Zielzelle markieren
    Set img = wsResult.Shapes.AddPicture(Filename:=image_path, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=Selection.Left, Top:=Selection.Top, Width:=-1, Height:=-1)

    

    ' Groesse zu Zelle ziehen
    With img
        .LockAspectRatio = msoTrue

        If fitTo = "Width" Then
            .Width = .TopLeftCell.Width ' Breite der Bild-Zelle
        Else
            .Height = .TopLeftCell.Height ' Breite der Bild-Zelle
        End If

        .Placement = xlMoveAndSize ' weiss nicht genau wofuer?
    End With




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



<a name="SeitenformatNextPosition"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Sub SeitenformatNextPosition```</span><small>(Zeile 775)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Modifiziert die Attribute des Objektes so, dass die naechste Position in den Attributen gespeichert wird







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```konvertiereZeile```](#konvertiereZeile) : <small>  [Zeile 767] : ```    Call SeitenformatNextPosition(oSeite)``` </small>





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
Private Sub SeitenformatNextPosition(oSeite As Seitenformat)
' Modifiziert die Attribute des Objektes so, dass die naechste Position in den Attributen gespeichert wird
    
    If oSeite.zeilengruppe + 1 > oSeite.anzahlZeilen Then
        ' Reset zeilen fuer neue Spalte
        oSeite.zeilengruppe = 1
        oSeite.spaltengruppe = oSeite.spaltengruppe + 1

        If oSeite.spaltengruppe > oSeite.anzahlSpalten Then
            oSeite.finished = True
        End If
    Else
        oSeite.zeilengruppe = oSeite.zeilengruppe + 1
    End If
    
    oSeite.excelZeilenNr = oSeite.zeilengruppe + oSeite.excelStartZeilenNr - 1

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



<a name="THIS_VERSION__1_17"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Sub THIS_VERSION__1_17```</span><small>(Zeile 1220)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Dummy-Methode, um Versionsnummer in MacroList anzeigen zu lassen...







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
Private Sub THIS_VERSION__1_17()
' Dummy-Methode, um Versionsnummer in MacroList anzeigen zu lassen...

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



<a name="applyLayoutPrintArea"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Sub applyLayoutPrintArea```</span><small>(Zeile 247)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Legt den Druckbereich der einzelnen Seite fest, in Abhängigkeit der tatsaechlich vorhandenen Anzahl von Spalten:







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```notengriffeErzeugen```](#notengriffeErzeugen) : <small>  [Zeile 1192] : ```        Call applyLayoutPrintArea(oSeite)``` </small>





</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (1)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```getCountsOfColumnsToPrint```](#getCountsOfColumnsToPrint) : <small>  [Zeile 254] : ```    anzahlColumnsToPrint = getCountsOfColumnsToPrint(oSeite)``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
Private Sub applyLayoutPrintArea(oSeite As Seitenformat)
    ' Legt den Druckbereich der einzelnen Seite fest, in Abhängigkeit der tatsaechlich vorhandenen Anzahl von Spalten:


    Dim anzahlColumnsToPrint As Integer
    Dim lastRowToPrint As Integer

    anzahlColumnsToPrint = getCountsOfColumnsToPrint(oSeite)
    lastRowToPrint = oSeite.excelStartZeilenNr + oSeite.maxAnzahlZeilen - 1

    With wsResult.PageSetup
    
        If oSeite.ausrichtung = "Querformat" Then
            .Orientation = xlLandscape
        Else
            .Orientation = xlPortrait
        End If

        .PrintArea = Range(Cells(1, 1), Cells(lastRowToPrint, anzahlColumnsToPrint)).Address

    End With

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



<a name="applyLayoutToAllPages"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Sub applyLayoutToAllPages```</span><small>(Zeile 136)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Durchlaeuft alle erstellten Tabellenblaetter und schleust durch zu einer Anwendungsfunktion, die wiederum  den Titel und andere Dinge in diese Blaetter  ein. Als oSeite wird das zuletzt-verwendete Objekt verwendet - benoetigt werden neben den ALLGEMEINEN UND KONSTANTEN Attributen nur die Angabe der pageNum als Gesamtseitenzahl.
 Zusaetzlich wird AUSSERHALB DES DRUCKBEREICHES auch ein Hyperlink zur Steuertabelle eingefuegt...







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```notengriffeErzeugen```](#notengriffeErzeugen) : <small>  [Zeile 1201] : ```    Call applyLayoutToAllPages(oSeite)``` </small>





</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (1)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```applyLayoutToSinglePage```](#applyLayoutToSinglePage) : <small>  [Zeile 150] : ```        Call applyLayoutToSinglePage(ws, page, oSeite)``` </small>


  - [```getCountsOfColumnsToPrint```](#getCountsOfColumnsToPrint) : <small>  [Zeile 179] : ```    anzahlColumnsToMerge = getCountsOfColumnsToPrint(oSeite)``` </small>



      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

  - [```mergeCells```](#mergeCells) : <small>  [Zeile 195] : ```    Call mergeCells(wsResult, titelZeile, titelSpalte, titelZeile, anzahlColumnsToMerge)``` </small>



      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

  - [```mergeCells```](#mergeCells) : <small>  [Zeile 211] : ```    Call mergeCells(wsResult, seitenangabeZeile, seitenangabeSpalte, seitenangabeZeile, anzahlColumnsToMerge)``` </small>



      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
Private Sub applyLayoutToAllPages(oSeite As Seitenformat)
    ' Durchlaeuft alle erstellten Tabellenblaetter und schleust durch zu einer Anwendungsfunktion, die wiederum  den Titel und andere Dinge in diese Blaetter  ein. Als oSeite wird das zuletzt-verwendete Objekt verwendet - benoetigt werden neben den ALLGEMEINEN UND KONSTANTEN Attributen nur die Angabe der pageNum als Gesamtseitenzahl.
    ' Zusaetzlich wird AUSSERHALB DES DRUCKBEREICHES auch ein Hyperlink zur Steuertabelle eingefuegt...

    Dim page As Integer
    Dim pagesTotal As Integer
    Dim ws As Worksheet

    pagesTotal = pageNum ' Abgreifen aus dem letzten Durchlauf.

    For page = 1 To pagesTotal
        
        Set ws = Sheets(Sheets.Count - pagesTotal + page)
        ' Anwenden der Layout-Methode, in der auch die Positionierung der Elemente definiert ist:
        Call applyLayoutToSinglePage(ws, page, oSeite)
    
    Next page

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



<a name="applyLayoutToSinglePage"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Sub applyLayoutToSinglePage```</span><small>(Zeile 161)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Zusaetzlich wird AUSSERHALB DES DRUCKBEREICHES auch ein Hyperlink zur Steuertabelle eingefuegt...







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```applyLayoutToAllPages```](#applyLayoutToAllPages) : <small>  [Zeile 150] : ```        Call applyLayoutToSinglePage(ws, page, oSeite)``` </small>





</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (3)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```getCountsOfColumnsToPrint```](#getCountsOfColumnsToPrint) : <small>  [Zeile 179] : ```    anzahlColumnsToMerge = getCountsOfColumnsToPrint(oSeite)``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

- [```mergeCells```](#mergeCells) : <small>  [Zeile 195] : ```    Call mergeCells(wsResult, titelZeile, titelSpalte, titelZeile, anzahlColumnsToMerge)``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

- [```mergeCells```](#mergeCells) : <small>  [Zeile 211] : ```    Call mergeCells(wsResult, seitenangabeZeile, seitenangabeSpalte, seitenangabeZeile, anzahlColumnsToMerge)``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
Private Sub applyLayoutToSinglePage(wsResult As Worksheet, thisPageNum As Integer, oSeite As Seitenformat)
    ' Zusaetzlich wird AUSSERHALB DES DRUCKBEREICHES auch ein Hyperlink zur Steuertabelle eingefuegt...

    Dim zielspalteHyperlink As Integer

    zielspalteHyperlink = oSeite.excelStartSpaltenNr + oSeite.maxAnzahlSpalten * oSeite.anzahlSubSpaltenProSpalte + 2

    ' Eintragung des Hyperlinkes zur Steuertabelle - ausserhalb des Sdruckbereiches:
    wsResult.Hyperlinks.Add Anchor:=wsResult.Cells(1, zielspalteHyperlink), Address:="", SubAddress:= _
        "STEUERTABELLE!A1", TextToDisplay:="<-- Zurueck zur Steuertabelle..."


    Dim pagesTotal As Integer
    pagesTotal = pageNum

    ' Berechnung der Gesamtspalten, ueber die zentrierte Texte verbunden + zentriert werden soll:
    Dim anzahlColumnsToMerge As Integer
    ' Die letzte "luft"-Spalte wird abgezogen, weil nach der letzten Spalte keine Luft notwendigt ist! (Blattrand!)
    anzahlColumnsToMerge = getCountsOfColumnsToPrint(oSeite)


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' TITEL EINTRAGEN:
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Der Titel soll in ZEile 1 mittig ueber dem Blatt stehen.
    ' Um genau mittig zu sein, wird der Text in Zelle A1 eingetragen und alle Zellen der ersten Spalte anschliessend verbunden werden.
    Dim titelZeile As Integer, titelSpalte As Integer, titel As String

    titelZeile = 1
    titelSpalte = 1
    titel = oSeite.titel

    wsResult.Cells(titelZeile, titelSpalte) = titel

    Call mergeCells(wsResult, titelZeile, titelSpalte, titelZeile, anzahlColumnsToMerge)


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' SEITE EINTRAGEN:
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Die Seitenzahl soll in Zeile 2 mittig ueber dem Blatt stehen.
    ' Um genau mittig zu sein, wird der Text in Zelle A2 eingetragen und alle Zellen der ersten Spalte anschliessend verbunden werden.
    Dim seitenangabeZeile As Integer, seitenangabeSpalte As Integer, seitenangabe As String

    seitenangabeZeile = 2
    seitenangabeSpalte = 1
    seitenangabe = "Seite " & thisPageNum & " von " & pagesTotal

    wsResult.Cells(seitenangabeZeile, seitenangabeSpalte) = seitenangabe

    Call mergeCells(wsResult, seitenangabeZeile, seitenangabeSpalte, seitenangabeZeile, anzahlColumnsToMerge)


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ZUSATZTEXT EINTRAGEN:
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Der Zusatztext soll in Zeile 3 linksbuendig
    Dim zusatztextZeile As Integer, zusatztextSpalte As Integer, zusatztext As String
    zusatztext = oSeite.zusatztext

    zusatztextZeile = 3
    zusatztextSpalte = 1
    zusatztext = oSeite.zusatztext

    wsResult.Cells(zusatztextZeile, zusatztextSpalte) = zusatztext

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



<a name="erstelleKonvertierteSeite"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Sub erstelleKonvertierteSeite```</span><small>(Zeile 655)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Erstelle neues TBB und wende Atrriubte an







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```notengriffeErzeugen```](#notengriffeErzeugen) : <small>  [Zeile 1191] : ```        Call erstelleKonvertierteSeite(startzeileSteuertabelle, oSeite)``` </small>





</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (2)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```insertSheet```](#insertSheet) : <small>  [Zeile 662] : ```        Call insertSheet(oSeite.blattname & pageNum)``` </small>


  - [```existsSheetname```](#existsSheetname) : <small>  [Zeile 283] : ```        If existsSheetname(sheetname) = False Then``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```konvertiereZeile```](#konvertiereZeile) : <small>  [Zeile 677] : ```            If konvertiereZeile(zeileSteuertabelle, oSeite) = False Then``` </small>


  - [```getReplacerPre1```](#getReplacerPre1) : <small>  [Zeile 723] : ```                    oReplacer = getReplacerPre1(stichwort)``` </small>


      - [```modifyReplacerIfSonderzeichen```](#modifyReplacerIfSonderzeichen) : <small>  [Zeile 419] : ```    Call modifyReplacerIfSonderzeichen(oErsatz)``` </small>



        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

      - [```modifyReplacerIfTaktwechsel```](#modifyReplacerIfTaktwechsel) : <small>  [Zeile 422] : ```    Call modifyReplacerIfTaktwechsel(oErsatz)``` </small>



        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


  - [```getReplacerPre2```](#getReplacerPre2) : <small>  [Zeile 725] : ```                    oReplacer = getReplacerPre2(stichwort)``` </small>


      - [```getReplacerPre1```](#getReplacerPre1) : <small>  [Zeile 445] : ```    getReplacerPre2 = getReplacerPre1(stichwort)``` </small>


        - [```modifyReplacerIfSonderzeichen```](#modifyReplacerIfSonderzeichen) : <small>  [Zeile 419] : ```    Call modifyReplacerIfSonderzeichen(oErsatz)``` </small>



          - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

        - [```modifyReplacerIfTaktwechsel```](#modifyReplacerIfTaktwechsel) : <small>  [Zeile 422] : ```    Call modifyReplacerIfTaktwechsel(oErsatz)``` </small>



          - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


  - [```getReplacerNote```](#getReplacerNote) : <small>  [Zeile 727] : ```                    oReplacer = getReplacerNote(stichwort)``` </small>



      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


  - [```getReplacerCount```](#getReplacerCount) : <small>  [Zeile 729] : ```                    oReplacer = getReplacerCount(stichwort)``` </small>



      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


  - [```getReplacerPost```](#getReplacerPost) : <small>  [Zeile 731] : ```                    oReplacer = getReplacerPost(stichwort)``` </small>



      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


  - [```getReplacerEmpty```](#getReplacerEmpty) : <small>  [Zeile 733] : ```                    oReplacer = getReplacerEmpty()``` </small>



      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


  - [```getFilePath```](#getFilePath) : <small>  [Zeile 742] : ```                pathImage = getFilePath(oReplacer.value)``` </small>


      - [```getFilePath```](#getFilePath) : <small>  [Zeile 366] : ```    getFilePath = getFilePath(ERROR_FILENAME)``` </small>


        - <small> *... recursivly calls itself under certain conditions ...* </small> 



      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


  - [```InsertImageFromFileToCell```](#InsertImageFromFileToCell) : <small>  [Zeile 744] : ```                Call InsertImageFromFileToCell(oSeite.excelZeilenNr, zielspalte, pathImage, oReplacer.cellFitTo)``` </small>



      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

  - [```getTypeOfFollowingSeperatorLine```](#getTypeOfFollowingSeperatorLine) : <small>  [Zeile 757] : ```    followingLineStyle = getTypeOfFollowingSeperatorLine(zeileSteuertabelle, oSeite)``` </small>



      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

  - [```insertFrameLineBelow```](#insertFrameLineBelow) : <small>  [Zeile 762] : ```        Call insertFrameLineBelow(zeileSteuertabelle, oSeite, CBool(followingLineStyle - 1))``` </small>



      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

  - [```SeitenformatNextPosition```](#SeitenformatNextPosition) : <small>  [Zeile 767] : ```    Call SeitenformatNextPosition(oSeite)``` </small>



      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
Private Sub erstelleKonvertierteSeite(startzeileSteuertabelle As Integer, oSeite As Seitenformat)
    ' Erstelle neues TBB und wende Atrriubte an

    Dim zeileSteuertabelle As Integer
    'Dim ignoreLine  As Integer
        
        pageNum = pageNum + 1
        Call insertSheet(oSeite.blattname & pageNum)
        
        ' '  BUGFIX: In VBA werden die Bedingungen einer For-Schleife immer nur einmalig beim Beginn der Schleife ausgewertet! Somit wird nicht reagiert, sofern der "BIS-Wert" sich veraendert!
        ' ' BUGFIX: Daher ForSchleife in While-Schleife umgewandelt:
        'For zeileSteuertabelle = startzeileSteuertabelle To startzeileSteuertabelle + oSeite.maxAnzahlEintraege + countOfSkippedLinesTodo
        zeileSteuertabelle = startzeileSteuertabelle
        Do While zeileSteuertabelle <= startzeileSteuertabelle + oSeite.maxAnzahlEintraege + oSeite.countOfSkippedLinesTodo
        
            If oSeite.finished = True Then
                
                Exit Sub ' springe zurueck in die Loop des mains: noengriffeErzeugen"
            
            End If

            
            If konvertiereZeile(zeileSteuertabelle, oSeite) = False Then
            
                ' dann ist naechste Zeile ist leer, daher wird die Laufvariable zusaetzlich um eins erhoeht, ehe sie ohnehin durch den Schleifendurchlauf inkrementiert wird
                zeileSteuertabelle = zeileSteuertabelle + 1
                
            End If
            
            zeileSteuertabelle = zeileSteuertabelle + 1

        'Next zeileSteuertabelle ' BUGFIX: Siehe oben...
        Loop


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



<a name="init"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Sub init```</span><small>(Zeile 119)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Parameter fuer modulweite Variablen laden







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```notengriffeErzeugen```](#notengriffeErzeugen) : <small>  [Zeile 1172] : ```    Call init``` </small>





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
Private Sub init()
    ' Parameter fuer modulweite Variablen laden

    ' Vorlage fuer Format:
    Set wsTemplate = Worksheets("TEMPLATE_LAYOUT")


    ' TODO-Tabellenblatt:
    Set wsTodo = Worksheets("STEUERTABELLE")
    
    ' Seitenzahl:
    pageNum = 0

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



<a name="insertFrameLineBelow"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Sub insertFrameLineBelow```</span><small>(Zeile 1011)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Fuegt die unteren Rahmenlinien aller Zellen dieser Zeile im Subspaltenbereich ein







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```konvertiereZeile```](#konvertiereZeile) : <small>  [Zeile 762] : ```        Call insertFrameLineBelow(zeileSteuertabelle, oSeite, CBool(followingLineStyle - 1))``` </small>





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
Private Sub insertFrameLineBelow(rowAbove As Integer, oSeite As Seitenformat, dashed As Boolean)
    ' Fuegt die unteren Rahmenlinien aller Zellen dieser Zeile im Subspaltenbereich ein
    
    Dim column As Integer
    Dim columnStart As Integer
    columnStart = (oSeite.spaltengruppe - 1) * oSeite.anzahlSubSpaltenProSpalte + 1
    
    For column = columnStart To columnStart + oSeite.anzahlSubSpaltenProSpalte - 2
        
        With wsResult.Cells(oSeite.excelZeilenNr, column).Borders(xlEdgeBottom)

            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
            .LineStyle = xlContinuous ' durchgezogen
            If dashed = True Then
                ' .LineStyle = xlDot ' gestrichelt
                .LineStyle = xlDashDotDot ' Strich-Punkt-Punkt
            End If
            
        End With
    
    Next column
    

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



<a name="insertSheet"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Sub insertSheet```</span><small>(Zeile 274)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Fuegt ein neues Tabellenblatt hinzu mit dem uebergebenen Namen und modifiziert den Tabellenblattnamen, sofern dieser bereits existiert







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```erstelleKonvertierteSeite```](#erstelleKonvertierteSeite) : <small>  [Zeile 662] : ```        Call insertSheet(oSeite.blattname & pageNum)``` </small>





</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (1)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```existsSheetname```](#existsSheetname) : <small>  [Zeile 283] : ```        If existsSheetname(sheetname) = False Then``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
Private Sub insertSheet(sheetname As String)
    ' Fuegt ein neues Tabellenblatt hinzu mit dem uebergebenen Namen und modifiziert den Tabellenblattnamen, sofern dieser bereits existiert

    wsTemplate.Copy After:=Sheets(Sheets.Count)
    Set wsResult = Sheets(Sheets.Count)

    Dim i As Integer
    i = 1
    Do
        If existsSheetname(sheetname) = False Then
            Exit Do
        End If
        
        sheetname = sheetname & i
        i = i + 1
    
    Loop
    
    ' # AUSBLICK: Fehlervermeidung bzgl. Sonderzeichen und max. Zeichenanzahl
    wsResult.name = sheetname
    
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



<a name="mergeCells"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Sub mergeCells```</span><small>(Zeile 320)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Verbindet den uebergebenen Zellbereich







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (2)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```applyLayoutToSinglePage```](#applyLayoutToSinglePage) : <small>  [Zeile 195] : ```    Call mergeCells(wsResult, titelZeile, titelSpalte, titelZeile, anzahlColumnsToMerge)``` </small>

- [```applyLayoutToSinglePage```](#applyLayoutToSinglePage) : <small>  [Zeile 211] : ```    Call mergeCells(wsResult, seitenangabeZeile, seitenangabeSpalte, seitenangabeZeile, anzahlColumnsToMerge)``` </small>





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
Private Sub mergeCells(ws As Worksheet, lineFrom As Integer, columnFrom As Integer, lineTo As Integer, columnTo As Integer)
    ' Verbindet den uebergebenen Zellbereich

    Dim cellTopLeft As Range
    Dim cellBottomRight As Range

    Set cellTopLeft = ws.Cells(lineFrom, columnFrom)
    Set cellBottomRight = ws.Cells(lineTo, columnTo)
   
    ' Verbinden der Zellen:
    Range(cellTopLeft, cellBottomRight).Merge

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



<a name="modifyReplacerIfSonderzeichen"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Sub modifyReplacerIfSonderzeichen```</span><small>(Zeile 535)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Prueft, ob das Stichwort einem definierten Sonderzeichen entspricht.
 Jedes definierte Sonderzeichen erhaelt ein expliziten Bild-Dateiname als Verknuepfung (ohne Dateiendung hier.)
 ACHTUNG: Die keys sind case-INSENSITIV , es gelten also immer gross und kleinscheibung fuer die keys!







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```getReplacerPre1```](#getReplacerPre1) : <small>  [Zeile 419] : ```    Call modifyReplacerIfSonderzeichen(oErsatz)``` </small>





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
Private Sub modifyReplacerIfSonderzeichen(oReplacer As Replacer)
    ' Prueft, ob das Stichwort einem definierten Sonderzeichen entspricht.
    ' Jedes definierte Sonderzeichen erhaelt ein expliziten Bild-Dateiname als Verknuepfung (ohne Dateiendung hier.)
    ' ACHTUNG: Die keys sind case-INSENSITIV , es gelten also immer gross und kleinscheibung fuer die keys!
    

    ' modifyReplacerIfSonderzeichen = False 'default initial: Fuer den Fall dass es KEIN Sonderzeichen ist.

    Dim keysAndReplacers() As Variant
    ' Aufbau des Arrays: ("Key1", "imageNameForKey1, "key2", "imageNameForKey2", ...)

    keysAndReplacers = Array("X", "paukenschlag", _
        "signum", "signum", _
        "segno", "segno", _
        "coda", "coda")


    Dim i As Integer
    Dim stichwort As String
    stichwort = oReplacer.value
    For i = LBound(keysAndReplacers) To UBound(keysAndReplacers) Step 2

        If UCase(oReplacer.value) = UCase(keysAndReplacers(i)) Then
            ' Dann ersetze das Stichwot (oReplacer.value) durch das Ersatz-Stichwort:
            oReplacer.value = keysAndReplacers(i + 1)
            oReplacer.isImage = True
        
            ' modifyReplacerIfSonderzeichen = True
            ' Merker, dass es einen Taktartwechsel gegeben hat.
            Exit Sub
            ' Exit Function
            
        End If
            

    Next i

' End Function
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



<a name="modifyReplacerIfTaktwechsel"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Sub modifyReplacerIfTaktwechsel```</span><small>(Zeile 578)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Prueft, ob das Stichwort einem Taktartwechsel entspricht.
 Je nach Taktart wird das uebergebene Replacer-Objekt inplace modifiziert und zurueckgegeben (mal wird ein Bild gewuenscht, mal nciht...)
 ALLE MOEGLICHEN Taktarten muessen EXPLIZIT IM ARRAY definiert sein (keysAndReplacers), damit sie als solche erkannt wird!







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```getReplacerPre1```](#getReplacerPre1) : <small>  [Zeile 422] : ```    Call modifyReplacerIfTaktwechsel(oErsatz)``` </small>





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
Private Sub modifyReplacerIfTaktwechsel(oReplacer As Replacer)
    ' Prueft, ob das Stichwort einem Taktartwechsel entspricht.
    ' Je nach Taktart wird das uebergebene Replacer-Objekt inplace modifiziert und zurueckgegeben (mal wird ein Bild gewuenscht, mal nciht...)
    ' ALLE MOEGLICHEN Taktarten muessen EXPLIZIT IM ARRAY definiert sein (keysAndReplacers), damit sie als solche erkannt wird!
    

    Dim keysAndReplacers() As Variant
    ' Aufbau des Arrays: ("Key1", "IMG:ersatzString1, "key2", "IMG:ersatz2", ...) wobei immer wenn ein Bild eingefuegt werden soll ein "IMG:" davor steht, sonst nicht, dann wird ein Apostroph vorangestellt
    keysAndReplacers = Array("TC|", "IMG:allabreve", _
    "T3/4", "'3/4", _
        "T4/4", "'4/4", _
        "T6/8", "'6/8", _
        "T3/8", "'3/8", _
    "TC", "IMG:TC")

    Dim i As Integer
    Dim stichwort As String
    stichwort = oReplacer.value
    For i = LBound(keysAndReplacers) To UBound(keysAndReplacers) Step 2

        If stichwort = keysAndReplacers(i) Then
            ' Dann ersetze das Stichwot durch das Ersatz-Stichwort:
            stichwort = keysAndReplacers(i + 1)
            ' pruefe, ob bei diesem Ersatz-Stichwort ein Bild verwendet werden soll oder nicht:
            If Left(stichwort, 4) = "IMG:" Then
                oReplacer.isImage = True
                oReplacer.value = Right(stichwort, Len(stichwort) - 4)
            Else
                oReplacer.isImage = False
                oReplacer.value = stichwort
            End If
        
            ' modifyReplacerIfTaktwechsel = True
            ' Merker, dass es einen Taktartwechsel gegeben hat.
            Exit Sub
        
        End If
    
    Next i
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



<a name="notengriffeErzeugen"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Sub notengriffeErzeugen```</span><small>(Zeile 1169)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

> *No information availible. For more information expand source code.*






<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (0)</summary>

<div style="padding-left:1em;">



Kein Aufruf gefunden.







</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (6)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```init```](#init) : <small>  [Zeile 1172] : ```    Call init``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```addSeitenformat```](#addSeitenformat) : <small>  [Zeile 1183] : ```        oSeite = addSeitenformat(startzeileSteuertabelle)``` </small>


  - [```newSeitenformat```](#newSeitenformat) : <small>  [Zeile 970] : ```    oFormat = newSeitenformat()``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


  - [```identifyLastRow```](#identifyLastRow) : <small>  [Zeile 980] : ```    letzteZeileSteuertabelle = identifyLastRow(startzeileSteuertabelle, oFormat.maxAnzahlEintraege)``` </small>


    - [```newSeitenformat```](#newSeitenformat) : <small>  [Zeile 876] : ```    tempObj = newSeitenformat(readParameters:=False)``` </small>



      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

    - [```areaEmpty```](#areaEmpty) : <small>  [Zeile 882] : ```        If Not areaEmpty(wsTodo, zeile, tempObj.excelStartSpaltenNr, zeile, tempObj.excelStartSpaltenNr + tempObj.anzahlSubSpaltenProSpalte - 1) Then``` </small>



      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```erstelleKonvertierteSeite```](#erstelleKonvertierteSeite) : <small>  [Zeile 1191] : ```        Call erstelleKonvertierteSeite(startzeileSteuertabelle, oSeite)``` </small>


  - [```insertSheet```](#insertSheet) : <small>  [Zeile 662] : ```        Call insertSheet(oSeite.blattname & pageNum)``` </small>


    - [```existsSheetname```](#existsSheetname) : <small>  [Zeile 283] : ```        If existsSheetname(sheetname) = False Then``` </small>



      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


  - [```konvertiereZeile```](#konvertiereZeile) : <small>  [Zeile 677] : ```            If konvertiereZeile(zeileSteuertabelle, oSeite) = False Then``` </small>


    - [```getReplacerPre1```](#getReplacerPre1) : <small>  [Zeile 723] : ```                    oReplacer = getReplacerPre1(stichwort)``` </small>


        - [```modifyReplacerIfSonderzeichen```](#modifyReplacerIfSonderzeichen) : <small>  [Zeile 419] : ```    Call modifyReplacerIfSonderzeichen(oErsatz)``` </small>



          - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

        - [```modifyReplacerIfTaktwechsel```](#modifyReplacerIfTaktwechsel) : <small>  [Zeile 422] : ```    Call modifyReplacerIfTaktwechsel(oErsatz)``` </small>



          - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


    - [```getReplacerPre2```](#getReplacerPre2) : <small>  [Zeile 725] : ```                    oReplacer = getReplacerPre2(stichwort)``` </small>


        - [```getReplacerPre1```](#getReplacerPre1) : <small>  [Zeile 445] : ```    getReplacerPre2 = getReplacerPre1(stichwort)``` </small>


          - [```modifyReplacerIfSonderzeichen```](#modifyReplacerIfSonderzeichen) : <small>  [Zeile 419] : ```    Call modifyReplacerIfSonderzeichen(oErsatz)``` </small>



            - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

          - [```modifyReplacerIfTaktwechsel```](#modifyReplacerIfTaktwechsel) : <small>  [Zeile 422] : ```    Call modifyReplacerIfTaktwechsel(oErsatz)``` </small>



            - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


          - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


    - [```getReplacerNote```](#getReplacerNote) : <small>  [Zeile 727] : ```                    oReplacer = getReplacerNote(stichwort)``` </small>



        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


    - [```getReplacerCount```](#getReplacerCount) : <small>  [Zeile 729] : ```                    oReplacer = getReplacerCount(stichwort)``` </small>



        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


    - [```getReplacerPost```](#getReplacerPost) : <small>  [Zeile 731] : ```                    oReplacer = getReplacerPost(stichwort)``` </small>



        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


    - [```getReplacerEmpty```](#getReplacerEmpty) : <small>  [Zeile 733] : ```                    oReplacer = getReplacerEmpty()``` </small>



        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


    - [```getFilePath```](#getFilePath) : <small>  [Zeile 742] : ```                pathImage = getFilePath(oReplacer.value)``` </small>


        - [```getFilePath```](#getFilePath) : <small>  [Zeile 366] : ```    getFilePath = getFilePath(ERROR_FILENAME)``` </small>


          - <small> *... recursivly calls itself under certain conditions ...* </small> 



        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


    - [```InsertImageFromFileToCell```](#InsertImageFromFileToCell) : <small>  [Zeile 744] : ```                Call InsertImageFromFileToCell(oSeite.excelZeilenNr, zielspalte, pathImage, oReplacer.cellFitTo)``` </small>



        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

    - [```getTypeOfFollowingSeperatorLine```](#getTypeOfFollowingSeperatorLine) : <small>  [Zeile 757] : ```    followingLineStyle = getTypeOfFollowingSeperatorLine(zeileSteuertabelle, oSeite)``` </small>



        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

    - [```insertFrameLineBelow```](#insertFrameLineBelow) : <small>  [Zeile 762] : ```        Call insertFrameLineBelow(zeileSteuertabelle, oSeite, CBool(followingLineStyle - 1))``` </small>



        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

    - [```SeitenformatNextPosition```](#SeitenformatNextPosition) : <small>  [Zeile 767] : ```    Call SeitenformatNextPosition(oSeite)``` </small>



        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```applyLayoutPrintArea```](#applyLayoutPrintArea) : <small>  [Zeile 1192] : ```        Call applyLayoutPrintArea(oSeite)``` </small>


  - [```getCountsOfColumnsToPrint```](#getCountsOfColumnsToPrint) : <small>  [Zeile 254] : ```    anzahlColumnsToPrint = getCountsOfColumnsToPrint(oSeite)``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```applyLayoutToAllPages```](#applyLayoutToAllPages) : <small>  [Zeile 1201] : ```    Call applyLayoutToAllPages(oSeite)``` </small>


  - [```applyLayoutToSinglePage```](#applyLayoutToSinglePage) : <small>  [Zeile 150] : ```        Call applyLayoutToSinglePage(ws, page, oSeite)``` </small>


    - [```getCountsOfColumnsToPrint```](#getCountsOfColumnsToPrint) : <small>  [Zeile 179] : ```    anzahlColumnsToMerge = getCountsOfColumnsToPrint(oSeite)``` </small>



        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

    - [```mergeCells```](#mergeCells) : <small>  [Zeile 195] : ```    Call mergeCells(wsResult, titelZeile, titelSpalte, titelZeile, anzahlColumnsToMerge)``` </small>



        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

    - [```mergeCells```](#mergeCells) : <small>  [Zeile 211] : ```    Call mergeCells(wsResult, seitenangabeZeile, seitenangabeSpalte, seitenangabeZeile, anzahlColumnsToMerge)``` </small>



        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```printLastSheetsAsPdf```](#printLastSheetsAsPdf) : <small>  [Zeile 1210] : ```        Call printLastSheetsAsPdf(pageNum)``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
Public Sub notengriffeErzeugen()

    ' Initialisieren der modulweiten Variablen (Worksheets usw...)
    Call init

    Dim oSeite As Seitenformat
    Dim startzeileSteuertabelle As Integer
    startzeileSteuertabelle = 30


    Do
        ' Endlosschleife, bis eine Seite generiert werden wuerde, auf der kein Eintrag zu finden ist.

        ' Erstelle eine neues Seitenformat-Objekts:
        oSeite = addSeitenformat(startzeileSteuertabelle)


        If oSeite.anzahlEintraege <= 0 Then
            Exit Do ' Ende der STB erreicht
        End If


        Call erstelleKonvertierteSeite(startzeileSteuertabelle, oSeite)
        Call applyLayoutPrintArea(oSeite)

        startzeileSteuertabelle = startzeileSteuertabelle + oSeite.maxAnzahlEintraege + oSeite.countOfSkippedLinesTodo
        ' BUGFIX ab Version 1.10: Die countOfSkippedLinesTodo wird nie zurueckgesetzt. Aber die startzeileSteuertabelle ist immer relativ PRO KONVERTIERTE ERGEBNISSEITE. Also muss die Variable immer beim Neuerschaffung einer neuen Seite zurueckgesetzt werden. Alternative: Speichern in oSeite-Objekt

    Loop

    ' Damit es nicht in mehreren Durchlaeufen passieren wird, wird der gesamte Text am Ende - nach Erstellen aller Seiten ergaenzt auf jede Seite. Damit ist die Gesamtanzahl der erstellten Seiten bereits bekannt und es muss nichts mehrfach definiert werden!
    ' call addPageNumbers()
    Call applyLayoutToAllPages(oSeite)

    wsResult.Activate

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Abfrage zur PDF-Export-Funktion:
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If wsTodo.Cells(1, 26) = True Then
        ' In dieser ZElle ist die Zellverknuepfung mit der Checkbox fuer die Export-Funktion.
        Call printLastSheetsAsPdf(pageNum)
    End If


    MsgBox "Herzlichen Glueckwunsch (nachtraeglich) zum Geburtstag ;-) <3." & vbCrLf & vbCrLf & "Umwandlung abgeschlossen." & vbCrLf & vbCrLf & "Sofern das Ergebnis (nochmals) als PDF exportiert werden sollte, muss die Checkbox in Zeile 1 der Steuertabelle aktiviert werden und die Umwandlung erneut durchgefuehrt werden. (Aktuell gibt es kein Automatismusm um bereits erstellte Tabellenblaetter automatisiert zu exportieren. Manuelle Version: Markieren aller relevanten Tabellenblaetter (mittels Strg + Anklicken), danach auf Drucken, unter den Einstellungen 'Aktive Blaetter drucken' auswaehlen und die Skalierung so stellen, dass alle SPALTEN auf einer Seite dargestellt werden...) --> oder eben abweichend nach Bedarf ;-)", vbOK, getProjectTitle()


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



<a name="printLastSheetsAsPdf"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Sub printLastSheetsAsPdf```</span><small>(Zeile 1081)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Speichert die hintersten (ERgebnis-Tabellenblaetter) als PDF ab, inkl. Skalierung auf eine Seite der Spalten.







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```notengriffeErzeugen```](#notengriffeErzeugen) : <small>  [Zeile 1210] : ```        Call printLastSheetsAsPdf(pageNum)``` </small>





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
Sub printLastSheetsAsPdf(anzahlSheets As Integer)
' Speichert die hintersten (ERgebnis-Tabellenblaetter) als PDF ab, inkl. Skalierung auf eine Seite der Spalten.

    Dim sheetnames() As Variant
    Dim i As Integer


    ReDim sheetnames(1 To anzahlSheets)

    For i = 1 To anzahlSheets
    
        sheetnames(i) = Sheets(Sheets.Count - anzahlSheets + i).name

    Next i
    
    Sheets(sheetnames).Select
    Sheets(Sheets.Count).Activate

    Application.PrintCommunication = False

    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""

        ' # HACK:  Sofer ein anderes Layout gewaehlt werden soll, muss es  hier umgestellt werden (oder mittels MAcroRekorder!)
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.787401575)
        .BottomMargin = Application.InchesToPoints(0.787401575)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)


        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        ' .Orientation = xlLandscape ' auskommentiert, da vorher bereits gesettet, und hier wuerde es ueberschrieben werden
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1 ' Alternativwert = 0. Aber hier 1, da alle Spalten auf ein Seite angepasst werden sollen
        .FitToPagesTall = 1 ' Alternativwert = 0. Aber hier 1, da alle Zeilen auf ein Seite angepasst werden sollen
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""

    End With

    Application.PrintCommunication = True
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False

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



<a name="resetWorkbook"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Sub resetWorkbook```</span><small>(Zeile 87)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

> *No information availible. For more information expand source code.*






<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (0)</summary>

<div style="padding-left:1em;">



Kein Aufruf gefunden.







</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (1)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```getProjectTitle```](#getProjectTitle) : <small>  [Zeile 111] : ```    MsgBox "Alle generierten Tabellenblaetter wurden geloescht. Dieser Vorgang kann nicht rueckgaengig gemacht werden! (--> falls Falsch ginge ggf. schliessen ohne speichern, sofern alle relevanten Modifikationen vorab gespeichert wurden...)", vbOK, getProjectTitle()``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
Sub resetWorkbook()

    If InputBox("Sollen alle generierten Tabellenblaetter UNWIDERUFLICH ENTFERNT werden? (dann eingabe von: 'y', sonst was beliebiges anderes)", getProjectTitle()) <> "y" Then
        Exit Sub
    End If
    Application.DisplayAlerts = False
    
      
    Dim i As Integer
    
    Do Until Sheets.Count = 3
        ' Auffuehrung der Namen aller Tabellenbletter, die beim Reset NICHT entfernt werden sollen.
        If Sheets(Sheets.Count).name <> "TEMPLATE_LAYOUT" And _
           Sheets(Sheets.Count).name <> "steuertabelle" And _
           Sheets(Sheets.Count).name <> "INFOS und Legende" Then

               Sheets(Sheets.Count).Delete
               
        End If
        
    Loop
    
    Application.DisplayAlerts = True
    
    MsgBox "Alle generierten Tabellenblaetter wurden geloescht. Dieser Vorgang kann nicht rueckgaengig gemacht werden! (--> falls Falsch ginge ggf. schliessen ohne speichern, sofern alle relevanten Modifikationen vorab gespeichert wurden...)", vbOK, getProjectTitle()
    
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



<a name="addSeitenformat"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function addSeitenformat```</span><small>(Zeile 964)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Erstellt ein Objekt des Typs Seitenformat und initialisiert die Konstanten Werte sowie zusaetzlich berechnete.
 Berechnet alle Attribute, das Obj wird zurück gegeben, auf Grundlage von diesem kann ein neues Blatt genereiert werden.







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```notengriffeErzeugen```](#notengriffeErzeugen) : <small>  [Zeile 1183] : ```        oSeite = addSeitenformat(startzeileSteuertabelle)``` </small>





</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (2)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```newSeitenformat```](#newSeitenformat) : <small>  [Zeile 970] : ```    oFormat = newSeitenformat()``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```identifyLastRow```](#identifyLastRow) : <small>  [Zeile 980] : ```    letzteZeileSteuertabelle = identifyLastRow(startzeileSteuertabelle, oFormat.maxAnzahlEintraege)``` </small>


  - [```newSeitenformat```](#newSeitenformat) : <small>  [Zeile 876] : ```    tempObj = newSeitenformat(readParameters:=False)``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

  - [```areaEmpty```](#areaEmpty) : <small>  [Zeile 882] : ```        If Not areaEmpty(wsTodo, zeile, tempObj.excelStartSpaltenNr, zeile, tempObj.excelStartSpaltenNr + tempObj.anzahlSubSpaltenProSpalte - 1) Then``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
Private Function addSeitenformat(startzeileSteuertabelle As Integer) As Seitenformat
' Erstellt ein Objekt des Typs Seitenformat und initialisiert die Konstanten Werte sowie zusaetzlich berechnete.
' Berechnet alle Attribute, das Obj wird zurück gegeben, auf Grundlage von diesem kann ein neues Blatt genereiert werden.

    ' Iniitalisiere ein neues Seitenformat-Objekt:
    Dim oFormat As Seitenformat
    oFormat = newSeitenformat()





    ' Berechnen der tataechlich moeglichen Anzahl an Eintragungen auf die Seite - in Abhaengigkeit der ausstehenden Anzahl der in Steuertabelle verfuehgbaren Inputs:
    oFormat.maxAnzahlEintraege = oFormat.maxAnzahlZeilen * oFormat.maxAnzahlSpalten

    Dim letzteZeileSteuertabelle As Integer
    letzteZeileSteuertabelle = identifyLastRow(startzeileSteuertabelle, oFormat.maxAnzahlEintraege)


    ' Berechnnung der Zeilen pro verfuegbaren Blattbereich:
    oFormat.anzahlEintraege = letzteZeileSteuertabelle - startzeileSteuertabelle + 1

    ' Aufteilen der Eintraege auf die verfuegbaren Blattbereiche (Spalten:) - immer AUFGERUNDET auf naechste Ganzzahl (sodass die letzte Spalte nie mehr Zeilen als die erste Spalte hat)
    Dim tempAnzahlZeilen As Double
    
    tempAnzahlZeilen = oFormat.anzahlEintraege / oFormat.anzahlSpalten
    'oFormat.anzahlZeilen = oFormat.anzahlEintraege / oFormat.anzahlSpalten
    If tempAnzahlZeilen <> Int(tempAnzahlZeilen) Then
        tempAnzahlZeilen = Int(tempAnzahlZeilen) + 1
    End If
    
    If tempAnzahlZeilen > oFormat.maxAnzahlZeilen Then
        tempAnzahlZeilen = tempAnzahlZeilen - 1
    End If
    
    oFormat.anzahlZeilen = tempAnzahlZeilen
    
    addSeitenformat = oFormat


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



<a name="areaEmpty"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function areaEmpty```</span><small>(Zeile 800)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Prueft einen Bereich auf leere Zellen.
 Rueckgabewerte:
 True falls ALLE Zellen im Bereich leer sind
 False falls NICHT ALLE ZEllen leer sind







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```identifyLastRow```](#identifyLastRow) : <small>  [Zeile 882] : ```        If Not areaEmpty(wsTodo, zeile, tempObj.excelStartSpaltenNr, zeile, tempObj.excelStartSpaltenNr + tempObj.anzahlSubSpaltenProSpalte - 1) Then``` </small>





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
Private Function areaEmpty(ws As Worksheet, lineFrom As Integer, columnFrom As Integer, lineTo As Integer, columnTo As Integer)
    ' Prueft einen Bereich auf leere Zellen.
    ' Rueckgabewerte:
        ' True falls ALLE Zellen im Bereich leer sind
        ' False falls NICHT ALLE ZEllen leer sind

        Dim line As Integer
        Dim column As Integer
        
        For line = lineFrom To lineTo

            For column = columnFrom To columnTo
                
                If ws.Cells(line, column) <> "" Then

                    ' Funktion beenden mit False:
                    areaEmpty = False
                    Exit Function

                End If

            Next column

        Next line

        areaEmpty = True
        
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



<a name="existsSheetname"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function existsSheetname```</span><small>(Zeile 300)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

> *No information availible. For more information expand source code.*






<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```insertSheet```](#insertSheet) : <small>  [Zeile 283] : ```        If existsSheetname(sheetname) = False Then``` </small>





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
Private Function existsSheetname(name As String) As Boolean

    Dim ws As Worksheet
    For Each ws In Sheets
        If ws.name = name Then
        
            existsSheetname = True
            Exit Function
        
        End If
            
    Next ws
    
    existsSheetname = False

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



<a name="getCountsOfColumnsToPrint"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Function getCountsOfColumnsToPrint```</span><small>(Zeile 235)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Berechnet die tatsaechliche Anzahl der zu druckenden Spalten
 Dafuer wird die letzte Spalte abgezogen, weil wg. Blattrand keine zusaetzliche Luft erforderlich ist:







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (2)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```applyLayoutToSinglePage```](#applyLayoutToSinglePage) : <small>  [Zeile 179] : ```    anzahlColumnsToMerge = getCountsOfColumnsToPrint(oSeite)``` </small>

- [```applyLayoutPrintArea```](#applyLayoutPrintArea) : <small>  [Zeile 254] : ```    anzahlColumnsToPrint = getCountsOfColumnsToPrint(oSeite)``` </small>





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
Function getCountsOfColumnsToPrint(oSeite As Seitenformat) As Integer
    ' Berechnet die tatsaechliche Anzahl der zu druckenden Spalten
    ' Dafuer wird die letzte Spalte abgezogen, weil wg. Blattrand keine zusaetzliche Luft erforderlich ist:

    getCountsOfColumnsToPrint = oSeite.anzahlSpalten * oSeite.anzahlSubSpaltenProSpalte - 1

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



<a name="getFilePath"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function getFilePath```</span><small>(Zeile 338)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Generiert einen Path zur Bilddatei basierend auf einem Stichwort.
 Sofern es keine passende Datei gibt, wird ein Platzhalter-Bild referenziert, durch das ein FEHLER verdeutlcihet wird.







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (2)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```getFilePath```](#getFilePath) : <small>  [Zeile 366] : ```    getFilePath = getFilePath(ERROR_FILENAME)``` </small>

- [```konvertiereZeile```](#konvertiereZeile) : <small>  [Zeile 742] : ```                pathImage = getFilePath(oReplacer.value)``` </small>





</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (1)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```getFilePath```](#getFilePath) : <small>  [Zeile 366] : ```    getFilePath = getFilePath(ERROR_FILENAME)``` </small>


  - <small> *... recursivly calls itself under certain conditions ...* </small> 



- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
Private Function getFilePath(stichwort As String)
    ' Generiert einen Path zur Bilddatei basierend auf einem Stichwort.
    ' Sofern es keine passende Datei gibt, wird ein Platzhalter-Bild referenziert, durch das ein FEHLER verdeutlcihet wird.

    Dim folderPath As String
    Dim imagePath As String
    Dim ERROR_FILENAME As String
    

    ' Einlesen des PFades aus der TodoTanbellenblattes:
    folderPath = wsTodo.Cells(2, 3)

    ' DEFAULT-Bildname:
    ERROR_FILENAME = "ERROR"
    
    ' Zusammenbau des Datepfades (als png-Dateien, da svg bei Franziska nicht funktionierte - bei mir schon):
    imagePath = folderPath & stichwort & ".png"

    ' Fehlervermeidung: Pruefe ob Datei existiert:
    If Dir(imagePath) <> "" Then
        
        ' Datei existiert.
        getFilePath = imagePath
        Exit Function
    End If
    

    ' Alternativ: Fehler-Bild default nehmen (rekursiv):
    getFilePath = getFilePath(ERROR_FILENAME)

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



<a name="getModifiedKeyIfTaktwechsel"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function getModifiedKeyIfTaktwechsel```</span><small>(Zeile 622)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Prueft, ob das Stichwort einem Taktartwechsel entspricht.
 Dazu werden konkrete Stichworte geprueftund durch ein alternativ-stichwort fuer einen Dateinamen ersetzt zurueckgegeben.







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
Private Function getModifiedKeyIfTaktwechsel(stichwort As String) As String
    ' Prueft, ob das Stichwort einem Taktartwechsel entspricht.
    ' Dazu werden konkrete Stichworte geprueftund durch ein alternativ-stichwort fuer einen Dateinamen ersetzt zurueckgegeben.

    getModifiedKeyIfTaktwechsel = "!" 'default initial:

    Dim keysAndReplacers() As Variant
    ' Aufbau des Arrays: ("Key1", "ersatzString1, "key2", "ersatz2", ...)
    keysAndReplacers = Array("TC|", "allabreve", _
    "T3/4", "T3-4", _
        "T4/4", "T4-4", _
        "T6/8", "T6-8", _
        "T3/8", "T3-8", _
    "TC", "TC")

    Dim i As Integer
    For i = LBound(keysAndReplacers) To UBound(keysAndReplacers)
        '# TODO: Eig. muesste hier step 2 noch rein! (pruefe nur die WErte mit ungeradem Index...)
        If stichwort = keysAndReplacers(i) Then
            ' Gebe den ERsatzwert fuer diesen KEy zurueck:
            getModifiedKeyIfTaktwechsel = keysAndReplacers(i + 1)
            Exit Function
        End If

    Next i

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



<a name="getProjectTitle"></a>
<span style="background-color: lightgrey; padding: 2px;">```Public Function getProjectTitle```</span><small>(Zeile 79)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

> *No information availible. For more information expand source code.*






<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```resetWorkbook```](#resetWorkbook) : <small>  [Zeile 111] : ```    MsgBox "Alle generierten Tabellenblaetter wurden geloescht. Dieser Vorgang kann nicht rueckgaengig gemacht werden! (--> falls Falsch ginge ggf. schliessen ohne speichern, sofern alle relevanten Modifikationen vorab gespeichert wurden...)", vbOK, getProjectTitle()``` </small>





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
Function getProjectTitle() As String

    getProjectTitle = "Notengriff-Converter"

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



<a name="getReplacerCount"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function getReplacerCount```</span><small>(Zeile 482)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Prueft das Stichwort fuer diese spezielle Spalte/ das Attribut und gibt darauf basierend eine ZEichenkette als Replacer zurueck. Diese Zeichenkette entspricht entweder dem Stichwort (Text uebernehmen) oder dem Dateipfad zu einem Bilddatei.







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```konvertiereZeile```](#konvertiereZeile) : <small>  [Zeile 729] : ```                    oReplacer = getReplacerCount(stichwort)``` </small>





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
Private Function getReplacerCount(stichwort As String) As Replacer
    ' Prueft das Stichwort fuer diese spezielle Spalte/ das Attribut und gibt darauf basierend eine ZEichenkette als Replacer zurueck. Diese Zeichenkette entspricht entweder dem Stichwort (Text uebernehmen) oder dem Dateipfad zu einem Bilddatei.

    Dim oErsatz  As Replacer

    ' Einfaches durchschleusen fuer diese Spalte:

    oErsatz.isImage = False
    oErsatz.value = stichwort
    oErsatz.cellfitTo = "Height"


    getReplacerCount = oErsatz

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



<a name="getReplacerEmpty"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function getReplacerEmpty```</span><small>(Zeile 516)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Gibt leeren Replacer zurueck







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```konvertiereZeile```](#konvertiereZeile) : <small>  [Zeile 733] : ```                    oReplacer = getReplacerEmpty()``` </small>





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
Private Function getReplacerEmpty() As Replacer
    ' Gibt leeren Replacer zurueck

    Dim oErsatz  As Replacer
    
    oErsatz.isImage = False
    oErsatz.value = ""
    oErsatz.cellfitTo = "Height"

    getReplacerEmpty = oErsatz

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



<a name="getReplacerNote"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function getReplacerNote```</span><small>(Zeile 451)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Prueft das Stichwort fuer diese spezielle Spalte/ das Attribut und gibt darauf basierend eine ZEichenkette als Replacer zurueck. Diese Zeichenkette entspricht entweder dem Stichwort (Text uebernehmen) oder dem Dateipfad zu einem Bilddatei.







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```konvertiereZeile```](#konvertiereZeile) : <small>  [Zeile 727] : ```                    oReplacer = getReplacerNote(stichwort)``` </small>





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
Private Function getReplacerNote(stichwort As String) As Replacer
    ' Prueft das Stichwort fuer diese spezielle Spalte/ das Attribut und gibt darauf basierend eine ZEichenkette als Replacer zurueck. Diese Zeichenkette entspricht entweder dem Stichwort (Text uebernehmen) oder dem Dateipfad zu einem Bilddatei.

    Dim oErsatz  As Replacer
    'Dim extensionWithDot As String

    Dim conditionStartsWithUpperCase As Boolean
    Dim conditionEndsWith0 As Boolean

    ' # TODO:  Abhaengig von Gross/Kleinschreibweise wird das stichwort so umgeformt, dass das entsprechende Bild hierzu spaeter heraus gesucht wird
    ' Zuordnung: / fuer einfaches und schnelles Eintragen in Excel werden die Oktaven ergaenzt:
    ' Soll-zuordnung = {"ais" : "ais.png", "Ais" : "Ais1.png", "A0" : "A0.png")
    ' Alg: Sofern Anfangsbuchstabe GROSS ist UND am Ende keine 0 steht handelt es  sich um Toene der Oktav 1, bei der die 1 ergaenzt wird. Andernfalls immer standardmaessiges Durchschleusen des Keys mit Verbindung mit der Dateiendung

    conditionStartsWithUpperCase = Left(stichwort, 1) = UCase(Left(stichwort, 1))
    conditionEndsWith0 = Right(stichwort, 1) = 0

    If conditionStartsWithUpperCase And Not conditionEndsWith0 Then
        stichwort = stichwort & "1"
    End If

    oErsatz.isImage = True
    oErsatz.value = stichwort ' Die Dateierweiterung wird in der MEthode getFilePath angestellt.
    oErsatz.cellfitTo = "Width"

    getReplacerNote = oErsatz

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



<a name="getReplacerPost"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function getReplacerPost```</span><small>(Zeile 500)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Prueft das Stichwort fuer diese spezielle Spalte/ das Attribut und gibt darauf basierend eine ZEichenkette als Replacer zurueck. Diese Zeichenkette entspricht entweder dem Stichwort (Text uebernehmen) oder dem Dateipfad zu einem Bilddatei.







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```konvertiereZeile```](#konvertiereZeile) : <small>  [Zeile 731] : ```                    oReplacer = getReplacerPost(stichwort)``` </small>





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
Private Function getReplacerPost(stichwort As String) As Replacer
    ' Prueft das Stichwort fuer diese spezielle Spalte/ das Attribut und gibt darauf basierend eine ZEichenkette als Replacer zurueck. Diese Zeichenkette entspricht entweder dem Stichwort (Text uebernehmen) oder dem Dateipfad zu einem Bilddatei.

    Dim oErsatz  As Replacer

    ' Einfaches durchschleusen fuer diese Spalte:
    oErsatz.isImage = False
    oErsatz.value = stichwort

    getReplacerPost = oErsatz
    oErsatz.cellfitTo = "Height"

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



<a name="getReplacerPre1"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function getReplacerPre1```</span><small>(Zeile 378)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Prueft das Stichwort fuer diese spezielle Spalte/ das Attribut und gibt darauf basierend eine ZEichenkette als Replacer zurueck. Diese Zeichenkette entspricht entweder dem Stichwort (Text uebernehmen) oder dem Dateipfad zu einem Bilddatei.







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (2)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```getReplacerPre2```](#getReplacerPre2) : <small>  [Zeile 445] : ```    getReplacerPre2 = getReplacerPre1(stichwort)``` </small>

- [```konvertiereZeile```](#konvertiereZeile) : <small>  [Zeile 723] : ```                    oReplacer = getReplacerPre1(stichwort)``` </small>





</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (2)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```modifyReplacerIfSonderzeichen```](#modifyReplacerIfSonderzeichen) : <small>  [Zeile 419] : ```    Call modifyReplacerIfSonderzeichen(oErsatz)``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

- [```modifyReplacerIfTaktwechsel```](#modifyReplacerIfTaktwechsel) : <small>  [Zeile 422] : ```    Call modifyReplacerIfTaktwechsel(oErsatz)``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
Private Function getReplacerPre1(stichwort As String) As Replacer
    ' Prueft das Stichwort fuer diese spezielle Spalte/ das Attribut und gibt darauf basierend eine ZEichenkette als Replacer zurueck. Diese Zeichenkette entspricht entweder dem Stichwort (Text uebernehmen) oder dem Dateipfad zu einem Bilddatei.

    ' Moegliche Optionen:
    ' - Durchschleusen des Strings
    ' - Wiederholungszeichen (Begin) --> str in TimesNewRoman ODER IMG
    ' - Klammer-Zahl (Wiederholungsklammer) --> str
    ' - Taktart-Wechsel --> IMG
    ' - Paukenschlag --> IMG
    ' [Trennerlinien werden unabhängig davon im Anschluss fuer jede Zeile gesondert geprueft!]

    Dim oErsatz  As Replacer
    ' dim found as Boolean



    ' Default-Init:
    oErsatz.isImage = False
    oErsatz.value = stichwort
    oErsatz.cellfitTo = "Height"
    ' found = False



    ' ' # OBSOLET / ALT: Bis V.1.16
    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ' Abfragen fuer einzelne Zeichen/Symbole:
    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ' Pruefe Option: Paukenschlag
    ' If UCase(stichwort) = "X" Then
    '     oErsatz.value = "paukenschlag"
    '     oErsatz.isImage = True



    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Abfragen fuer Gruppen / Kategorien:
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ' found = modifyReplacerIfSonderzeichen(oErsatz)
    Call modifyReplacerIfSonderzeichen(oErsatz)
    'Suche Sonderzeichen wie Paukenschlaege, Coda, Signum, ....

    Call modifyReplacerIfTaktwechsel(oErsatz)
    ' if not found then

        ' found = modifyReplacerIfTaktwechsel(oErsatz)
            ' Modifikation des Replacers erfolgte bereits innerhabl der Methode.

    '     ' ' # ALT: Bis V.1.16:
    '     ' ElseIf getModifiedKeyIfTaktwechsel(stichwort) <> "!" Then
    '     '     oErsatz.isImage = True
    '     '     oErsatz.value = getModifiedKeyIfTaktwechsel(stichwort)

    ' End If
    
    getReplacerPre1 = oErsatz

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



<a name="getReplacerPre2"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function getReplacerPre2```</span><small>(Zeile 441)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Prueft das Stichwort fuer diese spezielle Spalte/ das Attribut und gibt darauf basierend eine ZEichenkette als Replacer zurueck. Diese Zeichenkette entspricht entweder dem Stichwort (Text uebernehmen) oder dem Dateipfad zu einem Bilddatei.







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```konvertiereZeile```](#konvertiereZeile) : <small>  [Zeile 725] : ```                    oReplacer = getReplacerPre2(stichwort)``` </small>





</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (1)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```getReplacerPre1```](#getReplacerPre1) : <small>  [Zeile 445] : ```    getReplacerPre2 = getReplacerPre1(stichwort)``` </small>


  - [```modifyReplacerIfSonderzeichen```](#modifyReplacerIfSonderzeichen) : <small>  [Zeile 419] : ```    Call modifyReplacerIfSonderzeichen(oErsatz)``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

  - [```modifyReplacerIfTaktwechsel```](#modifyReplacerIfTaktwechsel) : <small>  [Zeile 422] : ```    Call modifyReplacerIfTaktwechsel(oErsatz)``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
Private Function getReplacerPre2(stichwort As String) As Replacer
    ' Prueft das Stichwort fuer diese spezielle Spalte/ das Attribut und gibt darauf basierend eine ZEichenkette als Replacer zurueck. Diese Zeichenkette entspricht entweder dem Stichwort (Text uebernehmen) oder dem Dateipfad zu einem Bilddatei.

    ' AKTUELL: DIE SELBEN ABFRAGEN WIE BEI PRE1-SPALTE
    getReplacerPre2 = getReplacerPre1(stichwort)

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



<a name="getTypeOfFollowingSeperatorLine"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function getTypeOfFollowingSeperatorLine```</span><small>(Zeile 1044)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  returns 0 falls kein eLinie folgt
 returns 1 falls durchgezogene Linie folgt
 returns -1 falls gestrichelte Linie folgt.







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```konvertiereZeile```](#konvertiereZeile) : <small>  [Zeile 757] : ```    followingLineStyle = getTypeOfFollowingSeperatorLine(zeileSteuertabelle, oSeite)``` </small>





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
Private Function getTypeOfFollowingSeperatorLine(lastRowNumber As Integer, oSeite As Seitenformat) As Integer
' returns 0 falls kein eLinie folgt
' returns 1 falls durchgezogene Linie folgt
' returns -1 falls gestrichelte Linie folgt.

    Dim spalteSteuertabelle As Integer
    Dim value As Variant
    
    getTypeOfFollowingSeperatorLine = 0 ' Default

    For spalteSteuertabelle = oSeite.excelStartSpaltenNr To oSeite.excelStartSpaltenNr + oSeite.anzahlSubSpaltenProSpalte - 1
        
        value = UCase(wsTodo.Cells(lastRowNumber + 1, spalteSteuertabelle))
        If value = "LINE" Then

            getTypeOfFollowingSeperatorLine = 1
            Exit Function

        ElseIf value = "LINE--" Then

            getTypeOfFollowingSeperatorLine = -1
            Exit Function

        End If

    Next spalteSteuertabelle


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



<a name="identifyLastRow"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function identifyLastRow```</span><small>(Zeile 867)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  Liesst die STB ausgehend ab startzeile aus bis dass die Anzahl maxZeilen abgearbeitet wurde. Pro Zeile wird sich die Zeilennummer gemerkt, sofern Inhalt vorhanden ist.
 Ziel davon ist, dass man so erkennt, wenn die letzten n zeilen nur freizeilen sind (um manuell einen seitenumbruch zu generieren), oder ob nur etwas abstand dadurch generiert werden sollte...







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```addSeitenformat```](#addSeitenformat) : <small>  [Zeile 980] : ```    letzteZeileSteuertabelle = identifyLastRow(startzeileSteuertabelle, oFormat.maxAnzahlEintraege)``` </small>





</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (2)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```newSeitenformat```](#newSeitenformat) : <small>  [Zeile 876] : ```    tempObj = newSeitenformat(readParameters:=False)``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

- [```areaEmpty```](#areaEmpty) : <small>  [Zeile 882] : ```        If Not areaEmpty(wsTodo, zeile, tempObj.excelStartSpaltenNr, zeile, tempObj.excelStartSpaltenNr + tempObj.anzahlSubSpaltenProSpalte - 1) Then``` </small>



  - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
Private Function identifyLastRow(startzeile As Integer, maxZeilen As Integer) As Integer
    ' Liesst die STB ausgehend ab startzeile aus bis dass die Anzahl maxZeilen abgearbeitet wurde. Pro Zeile wird sich die Zeilennummer gemerkt, sofern Inhalt vorhanden ist.
    ' Ziel davon ist, dass man so erkennt, wenn die letzten n zeilen nur freizeilen sind (um manuell einen seitenumbruch zu generieren), oder ob nur etwas abstand dadurch generiert werden sollte...

    identifyLastRow = -1 ' default-initial

    ' Nur als temporaeres Hilfsobjekt, dmait alle Parameter nicht wiederholt werde muessen:
    Dim zeile As Integer
    Dim tempObj As Seitenformat
    tempObj = newSeitenformat(readParameters:=False)


    For zeile = startzeile To startzeile + maxZeilen


        If Not areaEmpty(wsTodo, zeile, tempObj.excelStartSpaltenNr, zeile, tempObj.excelStartSpaltenNr + tempObj.anzahlSubSpaltenProSpalte - 1) Then

            identifyLastRow = zeile
        
        End If

    Next zeile

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



<a name="konvertiereZeile"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function konvertiereZeile```</span><small>(Zeile 696)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

>  # TODOC: ALT:
 Konvertiert eine Zeile des Todo-Tabellenblattes in eine Zeile des Ergebnis-Tabellenblattes.
 Geht alle Spalten der STB durch und ruft fuer diese Spalte den jeweiligen Getter zum Holen des ReplacerObjektes auf. Die Logik bzgl. der Ersetzung erfolgt in den einzelnen Gettern. Rueckgabewert ist jeweils ein Replacer-Objekt mit den Attributen isImage:bool und value:string
 Die Zielpositionen fuer die Replacer werden aus dem uebergebenen Objekt oSeite gezogen.
 Danach wird dann die eigentliche Ersetzung durchgefuehrt.
 Rueckgabewert: True fuer normale Zeile ,False falls eine zusaetzliche ZEile der STB ueberspringen werden soll (auslassungsstichwort!)







<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (1)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```erstelleKonvertierteSeite```](#erstelleKonvertierteSeite) : <small>  [Zeile 677] : ```            If konvertiereZeile(zeileSteuertabelle, oSeite) = False Then``` </small>





</details

</div>








<!--  CALL SEQUENCE Abruffolge: -->


<details>
    <summary>      Interne Aufrufabfolge (11)</summary>

---


Innehalb der Prozedur werden die folgenden, untergeordneten Prozeduren aufgerufen:





- [```getReplacerPre1```](#getReplacerPre1) : <small>  [Zeile 723] : ```                    oReplacer = getReplacerPre1(stichwort)``` </small>


    - [```modifyReplacerIfSonderzeichen```](#modifyReplacerIfSonderzeichen) : <small>  [Zeile 419] : ```    Call modifyReplacerIfSonderzeichen(oErsatz)``` </small>



      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

    - [```modifyReplacerIfTaktwechsel```](#modifyReplacerIfTaktwechsel) : <small>  [Zeile 422] : ```    Call modifyReplacerIfTaktwechsel(oErsatz)``` </small>



      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```getReplacerPre2```](#getReplacerPre2) : <small>  [Zeile 725] : ```                    oReplacer = getReplacerPre2(stichwort)``` </small>


    - [```getReplacerPre1```](#getReplacerPre1) : <small>  [Zeile 445] : ```    getReplacerPre2 = getReplacerPre1(stichwort)``` </small>


      - [```modifyReplacerIfSonderzeichen```](#modifyReplacerIfSonderzeichen) : <small>  [Zeile 419] : ```    Call modifyReplacerIfSonderzeichen(oErsatz)``` </small>



        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

      - [```modifyReplacerIfTaktwechsel```](#modifyReplacerIfTaktwechsel) : <small>  [Zeile 422] : ```    Call modifyReplacerIfTaktwechsel(oErsatz)``` </small>



        - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


      - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```getReplacerNote```](#getReplacerNote) : <small>  [Zeile 727] : ```                    oReplacer = getReplacerNote(stichwort)``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```getReplacerCount```](#getReplacerCount) : <small>  [Zeile 729] : ```                    oReplacer = getReplacerCount(stichwort)``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```getReplacerPost```](#getReplacerPost) : <small>  [Zeile 731] : ```                    oReplacer = getReplacerPost(stichwort)``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```getReplacerEmpty```](#getReplacerEmpty) : <small>  [Zeile 733] : ```                    oReplacer = getReplacerEmpty()``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```getFilePath```](#getFilePath) : <small>  [Zeile 742] : ```                pathImage = getFilePath(oReplacer.value)``` </small>


    - [```getFilePath```](#getFilePath) : <small>  [Zeile 366] : ```    getFilePath = getFilePath(ERROR_FILENAME)``` </small>


      - <small> *... recursivly calls itself under certain conditions ...* </small> 



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- [```InsertImageFromFileToCell```](#InsertImageFromFileToCell) : <small>  [Zeile 744] : ```                Call InsertImageFromFileToCell(oSeite.excelZeilenNr, zielspalte, pathImage, oReplacer.cellFitTo)``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

- [```getTypeOfFollowingSeperatorLine```](#getTypeOfFollowingSeperatorLine) : <small>  [Zeile 757] : ```    followingLineStyle = getTypeOfFollowingSeperatorLine(zeileSteuertabelle, oSeite)``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

- [```insertFrameLineBelow```](#insertFrameLineBelow) : <small>  [Zeile 762] : ```        Call insertFrameLineBelow(zeileSteuertabelle, oSeite, CBool(followingLineStyle - 1))``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>

- [```SeitenformatNextPosition```](#SeitenformatNextPosition) : <small>  [Zeile 767] : ```    Call SeitenformatNextPosition(oSeite)``` </small>



    - <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>


- <small>*Keine weiteren Aufrufe zu anderen, hier dokumentierten Prozeduren.*</small>









</details>





<!--  Source Code: -->



<details>
    <summary>      Source Code</summary>

---

```
Private Function konvertiereZeile(zeileSteuertabelle As Integer, oSeite As Seitenformat) As Boolean
    ' # TODOC: ALT:
    ' Konvertiert eine Zeile des Todo-Tabellenblattes in eine Zeile des Ergebnis-Tabellenblattes.
    ' Geht alle Spalten der STB durch und ruft fuer diese Spalte den jeweiligen Getter zum Holen des ReplacerObjektes auf. Die Logik bzgl. der Ersetzung erfolgt in den einzelnen Gettern. Rueckgabewert ist jeweils ein Replacer-Objekt mit den Attributen isImage:bool und value:string
    ' Die Zielpositionen fuer die Replacer werden aus dem uebergebenen Objekt oSeite gezogen.
    ' Danach wird dann die eigentliche Ersetzung durchgefuehrt.
    ' Rueckgabewert: True fuer normale Zeile ,False falls eine zusaetzliche ZEile der STB ueberspringen werden soll (auslassungsstichwort!)

    Dim spalteSteuertabelle As Integer
    Dim stichwort As String
    Dim pathImage As String
    Dim oReplacer As Replacer
    Dim zielspalte As Integer
    konvertiereZeile = True


    For spalteSteuertabelle = oSeite.excelStartSpaltenNr To oSeite.excelStartSpaltenNr + oSeite.anzahlSubSpaltenProSpalte - 1
    
        zielspalte = spalteSteuertabelle + oSeite.anzahlSubSpaltenProSpalte * (oSeite.spaltengruppe - 1)
        stichwort = wsTodo.Cells(zeileSteuertabelle, spalteSteuertabelle)


        If stichwort <> "" Then

            Select Case spalteSteuertabelle

                Case 1
                    oReplacer = getReplacerPre1(stichwort)
                Case 2
                    oReplacer = getReplacerPre2(stichwort)
                Case 3
                    oReplacer = getReplacerNote(stichwort)
                Case 4
                    oReplacer = getReplacerCount(stichwort)
                Case 5
                    oReplacer = getReplacerPost(stichwort)
                Case Else ' eigentlich obsolet
                    oReplacer = getReplacerEmpty()

            End Select
            '''''''''''''''''''''''''''''''''''
            ' Eigentliches Ersetzen:
            '''''''''''''''''''''''''''''''''''
            If oReplacer.isImage = True Then
                ' Bild einsetzen:

                pathImage = getFilePath(oReplacer.value)

                Call InsertImageFromFileToCell(oSeite.excelZeilenNr, zielspalte, pathImage, oReplacer.cellFitTo)
                    
            Else
                ' Text uebernehmen:
                wsResult.Cells(oSeite.excelZeilenNr, zielspalte) = oReplacer.value

            End If
        End If

    Next spalteSteuertabelle

    ' Pruefe, ob durch das Stichwort in der folgenden Zeile der STB ein Trennlinie unterhalb der gerade beschriebenen Zeilen eingefuegt wwerden soll:
    Dim followingLineStyle As Integer
    followingLineStyle = getTypeOfFollowingSeperatorLine(zeileSteuertabelle, oSeite)

    If followingLineStyle <> 0 Then

        ' Linienart modifizieren via Uebergabeparameter
        Call insertFrameLineBelow(zeileSteuertabelle, oSeite, CBool(followingLineStyle - 1))
        konvertiereZeile = False
        oSeite.countOfSkippedLinesTodo = oSeite.countOfSkippedLinesTodo + 1
    End If

    Call SeitenformatNextPosition(oSeite)

    ' Anscheinend muss ich oSeite ähnlich wie bei python nicht zurueckuebergeben.... (modulweit / "global"?)

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



<a name="newSeitenformat"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function newSeitenformat```</span><small>(Zeile 896)</small>






<!--  DocString der Prozedur: -->




<div style="padding-left:2em;">

> 
 Standard-Parameter aus TODO-Steuertabelle auslesen und Standardwerte parametrisieren.
 # ACHTUNG: DEPENDENCIES zur Excel-Worksheet!








<!--  References der Procedure: -->

<details>

<summary> Referenzierungen dieser Prozedur (2)</summary>

<div style="padding-left:1em;">



Die Prozedur wird in den folgenden, uebergeordneten Prozeduren aufgerufen:



- [```identifyLastRow```](#identifyLastRow) : <small>  [Zeile 876] : ```    tempObj = newSeitenformat(readParameters:=False)``` </small>

- [```addSeitenformat```](#addSeitenformat) : <small>  [Zeile 970] : ```    oFormat = newSeitenformat()``` </small>





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
Private Function newSeitenformat(Optional readParameters As Boolean = True) As Seitenformat
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Standard-Parameter aus TODO-Steuertabelle auslesen und Standardwerte parametrisieren.
    ' # ACHTUNG: DEPENDENCIES zur Excel-Worksheet!
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim oSeitenformat As Seitenformat

    Dim paramZeilenOffsetAusrichtung As Integer

    oSeitenformat.countOfSkippedLinesTodo = 0 ' Anzahl der uebersprungenen Zeilen innerhalb der Steuertabelle im bereich der fuer diese Seite umgewandelten Zeilen

    oSeitenformat.anzahlSubSpaltenProSpalte = 6 ' Azahl der Subspalten pro Spaltengruppe
    
    
    oSeitenformat.zeilengruppe = 1
    oSeitenformat.spaltengruppe = 1
    
    oSeitenformat.excelStartZeilenNr = 5 ' Excel-Zeilennummer, auf der die Seite den ersten Eintrag haben soll
    oSeitenformat.excelStartSpaltenNr = 1 ' Excel-Spaltennummer, auf der die Seite den ersten Eintrag haben soll
    
    
    oSeitenformat.excelZeilenNr = oSeitenformat.excelStartZeilenNr ' Tatsaechliche Zielzeilennummer
    oSeitenformat.excelSpaltenNr = oSeitenformat.excelStartSpaltenNr ' Tatsaechliche Zielspaltennummer
    
    oSeitenformat.finished = False

    If readParameters = False Then
        ' Vorzeitiges Verlassen der Funktion. Damit habe ich trotzdem duplizierten Code vermieden (hardcodierte parameter)
        newSeitenformat = oSeitenformat
        Exit Function
    
    End If


    '''''''''''''''''''''''''''''''''''''''''''''''
    ' AUSLESEN DER PARAMETER AUS DER STEUERTABELLE:
    '''''''''''''''''''''''''''''''''''''''''''''''
    Dim paramSpalte As Integer
    paramSpalte = 6 ' Spaltennummer der Steuertabelle, in der die Parameterwerte stehen
    
    
    oSeitenformat.titel = wsTodo.Cells(6, paramSpalte)
    oSeitenformat.zusatztext = wsTodo.Cells(7, paramSpalte)
    oSeitenformat.blattname = wsTodo.Cells(8, paramSpalte)
    oSeitenformat.ausrichtung = wsTodo.Cells(9, paramSpalte)

    ' Max. Zeilen- und Spaltenanzahl auslesen:
    paramZeilenOffsetAusrichtung = 0
    If oSeitenformat.ausrichtung = "Querformat" Then
        paramZeilenOffsetAusrichtung = 3
    End If

    oSeitenformat.maxAnzahlZeilen = wsTodo.Cells(12 + paramZeilenOffsetAusrichtung, paramSpalte)
    oSeitenformat.maxAnzahlSpalten = wsTodo.Cells(13 + paramZeilenOffsetAusrichtung, paramSpalte)

    ' Limitierend ist immer vorrangig die Spaltenanzahl - d.h. es werden IMMER alle Spalten benutzt, sofern es insgesamt mind. die Anzahl an Eintraegen wie Spalten gibt.
    oSeitenformat.anzahlSpalten = oSeitenformat.maxAnzahlSpalten

    newSeitenformat = oSeitenformat

End Function

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

<small>Dokumentation generiert am 2024-02-07 21:06:14 durch das  automatisierte Code-Dokumentationstool von Matthias Kader (Commit vom 2024-01-21 23:36:42: '157a6014c474dcc6ccfbd9ce96e6ea264ce29dc2')</small> 
