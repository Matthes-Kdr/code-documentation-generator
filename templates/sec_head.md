# Code-Dokumentation: Modul '@PLACEHOLDER_INPUT_FILE@'



**Letzte Änderung** der Quelldatei '@PLACEHOLDER_INPUT_FILE@' vor der Generierung dieser automatischen Dokumentation: **@PLACEHOLDER_TIMESTAMP_SOURCEFILE@**


Generierungsdatum dieser Dokumentation: **@PLACEHOLDER_TIMESTAMP_NOW@**









<!-- TODO: nur temporrary!  -->
# ZWISCHENGELAGERT ALS ZIEL-VORGABE FÜR ABRUFSEQUENZ:


**Aktuelle Bugs:**

- Probleme mit inkorrekter Einrückungen nach diversen Aufrufebenen scheint behoben zu sein 2024-01-07 - 22:47:41 (Teste nochmal!)
  
- Es werden nicht alle Aufrufe erkannt (ODER??!)
    
  - siehe beispiel_modul.bas --> liebherr : sollte 6 referenzierungen haben, es werden nur 5 dokumentiert... Zeile 106 fehlt:  ```var = liebherr```

  - wäre nicht tragisch, weil liebherr in diesem Syntax bei VBA nur eine Funktion mit Rückgabewert sein kann, und dann könnte man auch liebherr() schreiben, das würde erkannt werden. Allerdings funktioniert auch die Schreibweise ohne KLammern, weshalb sie auch erkannt werden sollte (auch wenn ich es nie so schreiben wollen würde...)


