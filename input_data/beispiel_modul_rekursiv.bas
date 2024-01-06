


'''  
' Beispiel Modul zum Testen der Dokumentation der Abruffolge.
'
' ## !!! # ACHTUNG BUGS: !!!
' > Sub 'liebherr' wird insgesamt 6x referenziert, nicht 5x, wie aktuell dokumentiert!

'''





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





Private Function rekursiv(tx as string) as string
' wird nirgendwo aufgerufen.
' Anzahl weiterer internen Aufrufe : 3 (1 + rekursiv sich selbst + 1)
'
' ruft addieren auf (innerhalb string-func)
' ruft unter bestimmten Bedingungen wiederum rekursiv die selbe Funktion rekursiv auf
' ruft liebherr auf
'
'


    ' Addieren:
    tx = tx + string(addieren(1,2))

    if len(tx) < 10 then

        rekursiv(tx)

    end if


    call liebherr

    call liebherr("nochmal")




End Function




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


''' Hier steht Text, der nirgendwo in derDoku auftauchen soll, denn es ist nicht innerhalb einre Prozedur und kein Modulkopf,



Private Sub main()
' Anzahl der Referenzierungen im Modul: 0
' Anzahl weiterer internen Aufrufe : 5
'
''' Ruft die MEthode 'addieren' auf
'
''' Ruft die MEthode 'subtrahieren' auf (und darin dann wieder addieren)
'
''' Ruft die MEthode 'liebherr' auf
'
''' Ruft die MEthode 'rekursiv' auf
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


    call liebherr("vor rekursivem Aufruf")

    call rekursiv

    call liebherr("NACH rekursivem Aufruf")


End Sub






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





   Sub liebherr()
   ' Anzahl der Referenzierungen im Modul: 6
    ' Anzahl weiterer internen Aufrufe : 0
    '
    ''' ' Ruft keine weitere Prozedur auf.


    MsgBox("Dies ist ein implizit als public gekennzeichnetes Sub.")


End Sub






   Sub casio()
    ' Anzahl der Referenzierungen im Modul: 0
    ' Anzahl weiterer internen Aufrufe : 0
    '
    ''' ' Ruft keine weitere Prozedur auf.

    MsgBox("Dies ist ein implizit als public gekennzeichnetes Sub.")


End Sub


' Sub schrott() ' Auskommentiert
''' Deklaration erfolgte via sub. - Auskommentiert!

    ' MsgBox("Dies ist ein implizit als public gekennzeichnetes Sub.")


' End Sub

