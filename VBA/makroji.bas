' Datoteka z makroji

Function Cl5_Funkcija(st1 As Variant, st2 As Variant) As Variant
    ' Funkcija nam dve števili poracuna ter poracunano
    ' število vrne
    Dim st As Double
    Dim poracunana As Double
    If CStr(st1) = "" Or CStr(st1) = "None" Then
        st1 = 0
    End If
    If CStr(st2) = "" Or CStr(st2) = "None" Then
        st2 = 0
    End If
    st = CDbl(st1) + CDbl(st2)
    If st < 50 Then
        poracunana = st * 0.3
        If poracunana = 0 Then
            Cl5_Funkcija = ""
        ElseIf poracunana <= 5 Then
            Cl5_Funkcija = 5#
        Else
            Cl5_Funkcija = poracunana
        End If
    ElseIf st <= 600 Then
        Cl5_Funkcija = 15#
    Else
        Cl5_Funkcija = st * 0.025
    End If
End Function

Function Cl0_Funkcija(st1 As Variant) As Variant
    ' Funkcija nam poracuna postavke
    If CStr(st1) = "" Then
        st1 = 0
    End If
    st = CDbl(st1)
    If st > 5 Then
        Cl0_Funkcija = Format((st - 5) * 8, "#0.00")
    Else
        Cl0_Funkcija = ""
    End If
End Function

Function vNizu(a As String, b As String) As Boolean
    ' Funkcija preveri ce je niz a v nizu b
    ' Funkcija vrne 'bool' vrednost
    If InStr(b, a) > 0 Then
        vNizu = True
    Else
        vNizu = False
    End If
End Function
