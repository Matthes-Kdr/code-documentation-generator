Attribute VB_Name = "Notenumwandlung"
Option Explicit


' @author: Matthias Kader


Type Replacer

    isImage As Boolean
    value As String
    cellFitTo as String

End Type




Type Seitenformat

    maxAnzahlZeilen As Integer
    maxAnzahlSpalten As Integer
    
    anzahlZeilen As Integer
    anzahlSpalten As Integer
    
    anzahlSubSpaltenProSpalte As Integer
    ausrichtung As String

    ' Ist letztendlich erstmal egal und wird nur einmalig genutzt, aber egal - bliebt jetzt hier!
    titel As String
    zusatztext As String
    blattname As String
    ' pageNum as integer

    anzahlEintraege As Integer
    maxAnzahlEintraege As Integer
    




    zeilengruppe As Integer ' Startet immer von 1, ist lalso unabhaengig von der tatsaechliche Zielzeile - es geht um die Platzhalter fuer die gruppierten Grafiken. Aktuell ist es zwar keine Gruppe (nur eine Zeile), aber um einheitlichkeit der Namensgebung mit Spaltengruppen zu verdeutlichen, so gewaehlt. Ausserdem: Sofern Leerzeilen eingefuegt werden koennen / sollen, ist es doch wieder eine  Zeilengruppe!
    spaltengruppe As Integer ' Startet immer von 1, ist also unabhaengig von der tatsaechlichen Zielspalte. Umfasst ausserdem eine gruppierte Einheit fuer die Eintragung einer Note inkl. aller Pre + Post-Zeichen
    
    
    

    excelStartZeilenNr As Integer
    excelStartSpaltenNr As Integer
    
    excelZeilenNr As Integer
    excelSpaltenNr As Integer

    finished As Boolean

    countOfSkippedLinesTodo As Integer

    

End Type




''' Abkuerzung: STB: Steuertabelle


Dim wsTemplate As Worksheet
Dim wsTodo As Worksheet
Dim wsResult As Worksheet
Dim pageNum As Integer
'Dim countOfSkippedLinesTodo As Integer





Function getProjectTitle() As String

    getProjectTitle = "Notengriff-Converter"

End Function



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





Private Sub init()
    ' Parameter fuer modulweite Variablen laden

    ' Vorlage fuer Format:
    Set wsTemplate = Worksheets("TEMPLATE_LAYOUT")


    ' TODO-Tabellenblatt:
    Set wsTodo = Worksheets("STEUERTABELLE")
    
    ' Seitenzahl:
    pageNum = 0

End Sub



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







Function getCountsOfColumnsToPrint(oSeite As Seitenformat) As Integer
    ' Berechnet die tatsaechliche Anzahl der zu druckenden Spalten
    ' Dafuer wird die letzte Spalte abgezogen, weil wg. Blattrand keine zusaetzliche Luft erforderlich ist:

    getCountsOfColumnsToPrint = oSeite.anzahlSpalten * oSeite.anzahlSubSpaltenProSpalte - 1

End Function





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




Private Sub mergeCells(ws As Worksheet, lineFrom As Integer, columnFrom As Integer, lineTo As Integer, columnTo As Integer)
    ' Verbindet den uebergebenen Zellbereich

    Dim cellTopLeft As Range
    Dim cellBottomRight As Range

    Set cellTopLeft = ws.Cells(lineFrom, columnFrom)
    Set cellBottomRight = ws.Cells(lineTo, columnTo)
   
    ' Verbinden der Zellen:
    Range(cellTopLeft, cellBottomRight).Merge

End Sub





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



Private Function getReplacerPre2(stichwort As String) As Replacer
    ' Prueft das Stichwort fuer diese spezielle Spalte/ das Attribut und gibt darauf basierend eine ZEichenkette als Replacer zurueck. Diese Zeichenkette entspricht entweder dem Stichwort (Text uebernehmen) oder dem Dateipfad zu einem Bilddatei.

    ' AKTUELL: DIE SELBEN ABFRAGEN WIE BEI PRE1-SPALTE
    getReplacerPre2 = getReplacerPre1(stichwort)

End Function



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



Private Function getReplacerCount(stichwort As String) As Replacer
    ' Prueft das Stichwort fuer diese spezielle Spalte/ das Attribut und gibt darauf basierend eine ZEichenkette als Replacer zurueck. Diese Zeichenkette entspricht entweder dem Stichwort (Text uebernehmen) oder dem Dateipfad zu einem Bilddatei.

    Dim oErsatz  As Replacer

    ' Einfaches durchschleusen fuer diese Spalte:

    oErsatz.isImage = False
    oErsatz.value = stichwort
    oErsatz.cellfitTo = "Height"


    getReplacerCount = oErsatz

End Function



Private Function getReplacerPost(stichwort As String) As Replacer
    ' Prueft das Stichwort fuer diese spezielle Spalte/ das Attribut und gibt darauf basierend eine ZEichenkette als Replacer zurueck. Diese Zeichenkette entspricht entweder dem Stichwort (Text uebernehmen) oder dem Dateipfad zu einem Bilddatei.

    Dim oErsatz  As Replacer

    ' Einfaches durchschleusen fuer diese Spalte:
    oErsatz.isImage = False
    oErsatz.value = stichwort

    getReplacerPost = oErsatz
    oErsatz.cellfitTo = "Height"

End Function



Private Function getReplacerEmpty() As Replacer
    ' Gibt leeren Replacer zurueck

    Dim oErsatz  As Replacer
    
    oErsatz.isImage = False
    oErsatz.value = ""
    oErsatz.cellfitTo = "Height"

    getReplacerEmpty = oErsatz

End Function






' Private Function modifyReplacerIfSonderzeichen(oReplacer As Replacer) As Boolean
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



' Private Function modifyReplacerIfTaktwechsel(oReplacer As Replacer) As Boolean
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


Private Sub THIS_VERSION__1_17()
' Dummy-Methode, um Versionsnummer in MacroList anzeigen zu lassen...

End Sub








