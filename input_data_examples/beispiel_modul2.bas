'''  
' Demo Modul zum Testen der Aufrufebenen / Aufruf-Abfolge / Ablauf
'
' Das Modul besteht aus insgesamt 7 Subs in folgender Aufrufhierarchie:
' ' 1 x main
' ' 3 x hauptfunc
' ' 2 x unterfuncktion
' ' 1 x wiederholungsfunktion
'
'''


Sub wiederholungsfunktion()
    ' Func für standard-aufgaben
    msgbox("Standardfunktion")
end sub


Sub unterfunktionA()
    ' weiterer Aufruf einer wiederholungsfunktion:
    call wiederholungsfunktion()
end sub


Sub unterfunktionB()
    ' kein externer aufruf
    x = 42
end sub



Sub hauptfunc1()
    ' weiterere 2 Aufrufe in dieser Funktion
    call unterfunktionA()
    call unterfunktionB()
end sub



Sub hauptfunc2()
    ' weiterer Aufruf von unterfunktionA in dieser Funktion
    call unterfunktionA()
    
    msgbox("standalone")
end sub



Sub hauptfunc3()
    ' kein weiterer Aufruf in dieser Funktion
    msgbox("standalone")
end sub





    

Sub main()
    ' Sequentieller Aufruf aller 3 Hauptfunktionen
    call hauptfunc1
    call hauptfunc2
    call hauptfunc3
    call wiederholungsfunktion
end sub

