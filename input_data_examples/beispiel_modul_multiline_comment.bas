


'''  
' Beispiel Modul zum Testen der Dokumentation der Abruffolge.

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
' Anzahl weiterer internen Aufrufe : 3
'
''' Ruft die MEthode 'addieren' auf
'
''' Ruft die MEthode 'subtrahieren' auf (und darin dann wieder addieren)
'
''' Ruft die MEthode 'liebherr' auf


' HACK: Durch 5-maliges Wiederholen von "#" wird simuliert, als ob es einen Syntax Multiline-Comment gäbe
#####
Dies hier soll ein Block kommentar-Demo sein, den es in VBA gar nicht gibt, aber egal
Hier ist noch eine Zeile davon!
wert = addieren(i, i) ' Dies sollte NICHT dokumentiert werden, sofern das Multiline-Feature implementiert ist!
#####



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

