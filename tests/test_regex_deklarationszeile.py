



import re



def test1():
    pattern = re.compile(r'''
        ^\s*                  # Optionale Leerzeichen am Anfang
        (?:                   # Nicht erforderliche Gruppe für den Modifier
            (?:public|private) # Möglicher Modifier (public oder private)
            \s+                # Mindestens ein Leerzeichen nach dem Modifier
        )?                    # Ende der optionalen Modifier-Gruppe
        Sub                   # Keyword 'Sub'
        \s+                   # Mindestens ein Leerzeichen nach 'Sub'
        (\w+)                 # Gruppierung für den Sub-Routinen-Namen (\w+ steht für einen Wort-Zeichenkette)
        \s*                   # Optionale Leerzeichen
        (?:[^'\n]|'(?! Sub ))* # Optionaler Code ohne Kommentarzeichen oder Kommentare, die nicht ' Sub ' enthalten
        $                     # Ende der Zeile
    ''', re.IGNORECASE | re.VERBOSE)







    # Beispielanwendung
    vba_code = '''
        ' Kommentierte Zeile
        Public Sub Example_Sub()
        Private Sub Another_Sub()
        Sub Sub_Without_Modifier()
        Sub Sub_With_Comment ' Kommentar am Ende
        ' Private Sub blabla()
        '    Sub bla()
    '''




    matches = pattern.findall(vba_code)
    for match in matches:
        print(match)



    muster = re.compile(r"""
        ^\s*        # Optionale Leerzeichen am Anfangen
        \d{2}       # 2 Ziffern
        \.           # Punkt
        \w+                
                        """, re.VERBOSE)



    raw = """

    22.April
    11.November
    44.03.Noi
    """
    matches = muster.findall(raw)
    print(matches)
    for match in matches:
        print(match)




def test3():

    vba_code = """
 Sub ExampleProcedure()
    ' Code der Prozedur
End Sub

Private Function CalculateSomething()
    ' Code der Funktion
End Function

Sub AnotherProcedure()
    ' Code einer weiteren Prozedur
End Sub
    """

    pattern = re.compile(r'\b(?:Sub|Function)\s+([a-zA-Z_]\w*)\s*\([^)]*\)\s*\n(?:.*\n)*?End\s+(?:Sub|Function)\b', re.IGNORECASE | re.DOTALL)

    matches = pattern.findall(vba_code)


    print(matches)
    for match in matches:
        print(match)        



# test3()        


def test4():

    vba_code = """
 Sub ExampleProcedure()
    ' Code der Prozedur
End Sub

Private Function CalculateSomething()
    ' Code der Funktion
End Function

Sub AnotherProcedure()
    ' Code einer weiteren Prozedur
End Sub
    """

    pattern = re.compile(r'\s*(?:Sub|Function)\s+([a-zA-Z_]\w*)\s*\([^)]*\)\s*\n(?:.*\n)*?End\s+(?:Sub|Function)\b', re.IGNORECASE | re.DOTALL)

    matches = pattern.findall(vba_code)


    print(matches)
    for match in matches:
        print(match)        



# test4()        

def test5():

    vba_code = """

   Private Sub main() ' 1
Private Sub main() ' 2
Public Sub oeffentlichesSub() 3
 Public Sub oeffentlichesSub()4
Sub    implizietOeffentlichesSub()5
Sub implizietOeffentlichesSub()6
' Sub implizietOeffentlichesSub()7
'Sub implizietOeffentlichesSub()8





    """
    print(vba_code)
    pattern = re.compile(r"[^']\s*(?:Sub)\s+\w*\(", re.IGNORECASE)

    matches = pattern.findall(vba_code)


    print(matches, len(matches))
    for match in matches:
        print(match)        



test5()        



# (?:Sub)\s+\w*\(