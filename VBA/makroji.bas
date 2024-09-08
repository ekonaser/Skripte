' Datoteka z makroji

Function Cl05_Funkcija(st1, st2)
    ' Funkcija nam dve števili poračuna ter poračunano
    ' število vrne
    Dim st as Double
    If st1 = "" or st1 = "None" Then
        st1 = 0
    If st2 = "" or st2 = "None" Then
        st2 = 0
    st = st1 + st2
    If st = 0 Then
        ' vrnemo prazen niz v tem primeru
        Cl05_Funkcija = ""
    ElseIf st < 5 Then
        Cl05_Funkcija = 5.#
    ElseIf st <= 50 Then
        Cl05_Funkcija = st * 0.3
    ElseIf st <= 600 Then
        Cl05_Funkcija = 15.#
    Else
        Cl05_Funkcija = st * 0.025        
    End If
End Function