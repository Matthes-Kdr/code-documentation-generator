


'''  
' Beispiel Modul zum Testen der Dokumentation der Abruffolge.
'''





Private Function addieren(a as integer, b as integer) as integer
''' Diese Funktion addiert beide Zahlen miteinander und übergibt das Ergebnis zurück.
'
' Ruft keine weitere Prozedur auf.


    ' Addieren:
    addieren = a + b

End Function


Private Function subtrahieren(a as integer, b as integer) as integer
''' Diese Funktion subtrahiert b von a und übergibt das Ergebnis zurück.
'
' Ruft die  Prozedur 'adddieren' auf.

    ' Benutze die addieren Funktion:
    subtrahieren = addieren(a, -b) ' Parameter b wird mit -1 multipliziert übergeben


end Function


''' Hier steht Text, der nirgendwo in derDoku auftauchen soll, denn es ist nicht innerhalb einre Prozedur und kein Modulkopf,



Private Sub main()
''' Ruft die MEthode 'addieren' auf
'
''' Ruft die MEthode 'subtrahieren' auf

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






Public Sub bauer()
''' Ruft MEHRFACH die Prozedur 'liebherr' auf (insgesamt 5 mal nacheinander)

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
    ''' ' Ruft keine weitere Prozedur auf.


    MsgBox("Dies ist ein implizit als public gekennzeichnetes Sub.")


End Sub






   Sub casio()
    ''' ' Ruft keine weitere Prozedur auf.

    MsgBox("Dies ist ein implizit als public gekennzeichnetes Sub.")


End Sub


' Sub schrott() ' Auskommentiert
''' Deklaration erfolgte via sub. - Auskommentiert!

    ' MsgBox("Dies ist ein implizit als public gekennzeichnetes Sub.")


' End Sub

