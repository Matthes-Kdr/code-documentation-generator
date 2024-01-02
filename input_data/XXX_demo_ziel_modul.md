# Code-Dokumentation: Modul 'demo_ziel_modul.bas'

Organisatorische Hinweise zur Verwendeten bzw. dokumentierten Datei.




## Index

Verlinkte Auflistung aller Subs und Functions, bestenfalls alphabetisch sortiert:

* [**Modulinformationen / Modulkopf**](#sec_modulinfos)
* [**Subs**](#sec_subs) (1)
  * [```main```](#main)

* [**Functions**](#sec_functions) (2)
  * [```addieren```](#addieren)
  * [```subtrahieren```](#subtrahieren)
















































<a name="sec_modulinfos"></a>
## Modulinfos und Kopf



Dieses Modul beinhaltet einige Prozeduren, die für nichts sinnvoll sind...
Aber es hat immerhin einen Programmkopf.

Wichtige Prozeduren: Keine




<a name="sec_subs"></a>
## Subs



<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->





<a name="main"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Sub main```</span><small>(Zeile 224)</small>

<div style="padding-left:2em;">


> Hier soll das HAuptprogramm stehen.
Alles was als KOMMENTAR hier unter der Definitionszeile einer Funktion steht, BEVOR EINE LEERZEILE folgt, soll später als Zusammenfassung angezeigt werden in der Code-Dokumentation - also ähnlich wie im docstring bei python.



<details>

<summary> Referenzierungen dieser Prozedur (0)</summary>

<div style="padding-left:1em;">

 
Kein Aufruf gefunden.

</details

</div>




<details>
    <summary>      Expand Source Code</summary>

---

```
Private Sub main()
''' Hier soll das HAuptprogramm stehen.
''' Alles was als KOMMENTAR hier unter der Definitionszeile einer Funktion steht, BEVOR EINE LEERZEILE folgt, soll später als Zusammenfassung angezeigt werden in der Code-Dokumentation - also ähnlich wie im docstring bei python.


    dim i as integer

    i = 10

    for i = 0 to 10
        msgbox(i)
        wert = addieren(i, i)
    next i

End Sub
```

</details>


</div>


---



































<a name="sec_functions"></a>
## Functions






<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->

<a name="addieren"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function addieren(a as integer, b as integer) as integer```</span><small>(Zeile 24)</small>

<div style="padding-left:2em;">


> Diese Funktion addiert beide Zahlen miteinander und übergibt das Ergebnis zurück.

<details>

<summary> Referenzierungen dieser Prozedur (3)</summary>

<div style="padding-left:1em;">

Die Prozedur wird in den folgenden, übergeordneten Prozeduren aufgerufen:
 
* ```Private Sub main``` : Zeile 20 : ```MsgBox(addieren(4, 3))```
* ```Private Sub main``` : Zeile 140 : ```MsgBox(addieren(variable_a, variable_b))```
* ```Private Sub XX_unbenutzt``` : Zeile 22 : ```MsgBox(addieren(42, 33))```

</details

</div>




<details>
    <summary>      Expand Source Code</summary>

---

```


Private Function addieren(a as integer, b as integer) as integer
''' Diese Funktion addiert beide Zahlen miteinander und übergibt das Ergebnis zurück.

    addieren = a + bei

end Function
```

</details>


</div>


---
























<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->
<!-- NEUE PROZEDUR-DOKUMENTATION -->

<a name="subtrahieren"></a>
<span style="background-color: lightgrey; padding: 2px;">```Private Function subtrahieren(a as integer, b as integer) as integer```</span><small>(Zeile 51)</small>

<div style="padding-left:2em;">


> Diese Funktion subtrahiert die Zahl b von Zahl a und übergibt das Ergebnis zurück.

<details>

<summary> Referenzierungen dieser Prozedur (2)</summary>

<div style="padding-left:1em;">

Die Prozedur wird in den folgenden, übergeordneten Prozeduren aufgerufen:
 
* ```Private Sub main``` : Zeile 20 : ```MsgBox(subtrahieren(4, 3))```
* ```Private Sub XX_unbenutzt``` : Zeile 22 : ```MsgBox(subtrahieren(42, 33))```

</details

</div>




<details>
    <summary>      Expand Source Code</summary>

---

```


Private Function subtrahieren(a as integer, b as integer) as integer
''' Diese Funktion subtrahiert die Zahl b von Zahl a und übergibt das Ergebnis zurück.

    subtrahieren = a + b

end Function
```

</details>


</div>


---













<small> Erstellt am (Datum) durch das  automatisierte Code-Dokumentationstool von .... in der Version ....</small> 
