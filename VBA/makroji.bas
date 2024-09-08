' Datoteka z makroji

Function Cl05_Funkcija(st1 As Variant, st2 As Variant) As Variant
    ' Funkcija nam dve števili poracuna ter poracunano
    ' število vrne
    Dim st As Double
    If CStr(st1) = "" Or CStr(st1) = "None" Then
        st1 = 0
    End If
    If CStr(st2) = "" Or CStr(st2) = "None" Then
        st2 = 0
    End If
    st = CDbl(st1) + CDbl(st2)
    If st = 0 Then
        ' vrnemo prazen niz v tem primeru
        Cl05_Funkcija = ""
    ElseIf st < 5 Then
        Cl05_Funkcija = 5#
    ElseIf st <= 50 Then
        Cl05_Funkcija = st * 0.3
    ElseIf st <= 600 Then
        Cl05_Funkcija = 15#
    Else
        Cl05_Funkcija = st * 0.025
    End If
End Function