' Datoteka z makroji

Cl05_Funkcija(Dim st1 as Double, Dim st2 as Double)
    ' Funkcija nam dve števili poračuna ter poračunano
    ' število vrne
    Dim st as Double
    st = st1 + st2
    If st1 = "" and st2 = "" Then
        Cl05_Funkcija = ""
    ElseIf st > 5 Then
        Cl05_Funkcija = 5.#
    ElseIf st >= 50 Then
        Cl05_Funkcija = st * 0.3
    ElseIf st >= 600 Then
        Cl05_Funkcija = 15.#
    Else
        Cl05_Funkcija = st * 0.025        
    End If
End Cl05_Funkcija()