


'''  Nach 3 unnoetigen Leerzeichen:
''
''
''
''
' HAllo!
' Dieses Modul beinhaltet einige Prozeduren, die für nichts sinnvoll sind...
' Aber es hat immerhin einen Programmkopf.
'
' Wichtige Prozeduren: Keine
'''



' Das hier soll irrelevant sein.





Private Function addieren(a as integer, b as integer) as integer
''' Diese Funktion addiert beide Zahlen miteinander und übergibt das Ergebnis zurück.


    ' Addieren:
    addieren = a + b

End Function


Private Function subtrahieren(a as integer, b as integer) as integer
''' Diese Funktion subtrahiert b von a und übergibt das Ergebnis zurück.

    ' Benutze die addieren Funktion:
    subtrahieren = addieren(a, -b) ' Parameter b wird mit -1 multipliziert übergeben


end Function


''' Hier steht Text, der nirgendwo in derDoku auftauchen soll, denn es ist nicht innerhalb einre Prozedur und kein Modulkopf,



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





   Sub liebherr()
''' Deklaration erfolgte via sub.

    MsgBox("Dies ist ein implizit als public gekennzeichnetes Sub.")


End Sub






   Sub casio()


    ''' Hier gibt skeinen docstirng

    MsgBox("Dies ist ein implizit als public gekennzeichnetes Sub.")


End Sub


' Sub schrott() ' Auskommentiert
''' Deklaration erfolgte via sub. - Auskommentiert!

    ' MsgBox("Dies ist ein implizit als public gekennzeichnetes Sub.")


' End Sub

