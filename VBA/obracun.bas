Dim steviloStrank As Integer
Public slovar As Object
Dim datum As Date

Sub UvozniMakro()
    UstvariOBRACUN
    UvoziTrinet
    UvoziCarmSkeniranje
End Sub

Sub GlavniMakro()
    Dim ws As Worksheet
    PrestejAWB
    DodajCarM
    DodajSlovar
    ZapolniVrstice
    Fizikalci
    InputDatum
    CustomerReference
    Davek
    CL0
    CL5
    OpravilnaStevilka
    CL4CL6CL8
    IzbrisiVrstice
    Set ws = ThisWorkbook.Sheets("OBRACUN")
    ws.Columns.AutoFit
End Sub

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
        ElseIf poracunana <= 8 Then
            Cl5_Funkcija = 8
        Else
            Cl5_Funkcija = poracunana
        End If
    ElseIf st <= 600 Then
        Cl5_Funkcija = 15
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
        Cl0_Funkcija = (st - 5) * 8
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

Private Sub PrestejAWB()
    Dim ws As Worksheet
    Dim rowNumber As Integer
    Dim count As Integer
    Dim cellValue As String

    ' Set the row number you want to count the strings in
    rowNumber = 2
    
    Set ws = ThisWorkbook.Sheets("carm_scan")

    ' Initialize the count
    steviloStrank = 0

    ' Loop through the cells in the row
    col = 1
    Do
        cellValue = ws.Cells(rowNumber, 1).Value
        If cellValue = "" Then
            Exit Do
        End If
        steviloStrank = steviloStrank + 1
        rowNumber = rowNumber + 1
    Loop
    steviloStrank = steviloStrank + 1
End Sub

Private Sub DodajCarM()
    Dim cellValue1 As String
    Dim cellValue2 As String
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    
    Set ws1 = ThisWorkbook.Sheets("carm_scan")
    Set ws2 = ThisWorkbook.Sheets("OBRACUN")
    
    rowNumber = 2
        
    Do
        cellValue1 = Int(ws1.Cells(rowNumber, 1).Value)
        cellValue2 = ws1.Cells(rowNumber, 2).Value
        
        ws2.Cells(rowNumber, 1).Value = cellValue1
        ws2.Cells(rowNumber, 3).Value = cellValue2
        
        ws2.Cells(rowNumber, 1).NumberFormat = "0"
        
        If rowNumber = steviloStrank Then
            Exit Do
        End If
        rowNumber = rowNumber + 1
    Loop
End Sub

Private Sub DodajSlovar()
    Dim key As String
    Dim cellValue As String
    Dim rowNumber As Integer
    Dim infoArray() As Variant
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets("trinet")
    Set slovar = CreateObject("Scripting.Dictionary")

    ' Initialize rowNumber
    rowNumber = 2

    ' Loop to populate dictionary
    Do
        cellValue = ws.Cells(rowNumber, 1).Value
        If cellValue = "" Then
            Exit Do
        End If
        
        If ws.Cells(rowNumber, 2) <> "" Then
            key = ws.Cells(rowNumber, 2).Value
            
            If ws.Cells(rowNumber, 3).Value <> "" And IsDate(ws.Cells(rowNumber, 3).Value) Then
                infoArray = Array(ws.Cells(rowNumber, 1).Value, DateValue(ws.Cells(rowNumber, 3).Value), ws.Cells(rowNumber, 4).Value, ws.Cells(rowNumber, 5).Value, _
                                  ws.Cells(rowNumber, 6).Value, ws.Cells(rowNumber, 7).Value, ws.Cells(rowNumber, 8).Value, ws.Cells(rowNumber, 9).Value, _
                                  ws.Cells(rowNumber, 10).Value, ws.Cells(rowNumber, 11).Value, ws.Cells(rowNumber, 12).Value, ws.Cells(rowNumber, 2).Value, _
                                  ws.Cells(rowNumber, 13).Value, ws.Cells(rowNumber, 14).Value, ws.Cells(rowNumber, 15).Value, ws.Cells(rowNumber, 16).Value)
            Else
                infoArray = Array(ws.Cells(rowNumber, 1).Value, ws.Cells(rowNumber, 3).Value, ws.Cells(rowNumber, 4).Value, ws.Cells(rowNumber, 5).Value, _
                                  ws.Cells(rowNumber, 6).Value, ws.Cells(rowNumber, 7).Value, ws.Cells(rowNumber, 8).Value, ws.Cells(rowNumber, 9).Value, _
                                  ws.Cells(rowNumber, 10).Value, ws.Cells(rowNumber, 11).Value, ws.Cells(rowNumber, 12).Value, ws.Cells(rowNumber, 2).Value, _
                                  ws.Cells(rowNumber, 13).Value, ws.Cells(rowNumber, 14).Value, ws.Cells(rowNumber, 15).Value, ws.Cells(rowNumber, 16).Value)
            End If
            slovar.Add key, infoArray
        End If
        rowNumber = rowNumber + 1
    Loop
End Sub

Private Sub ZapolniVrstice()
    Dim ws As Worksheet
    Dim kljuc As String
    Set ws = ThisWorkbook.Sheets("OBRACUN")
    Dim rowNumber As Integer
    
    rowNumber = 2
    Do
        If rowNumber = (steviloStrank + 1) Then
            Exit Do
        End If
        kljuc = ws.Cells(rowNumber, 3).Value
        If slovar.Exists(kljuc) Then
            ws.Cells(rowNumber, 4).Value = slovar(kljuc)(1)
            ws.Cells(rowNumber, 5).Value = slovar(kljuc)(0)
            ws.Cells(rowNumber, 7).Value = slovar(kljuc)(9)
            ws.Cells(rowNumber, 8).Value = slovar(kljuc)(10)
            ws.Cells(rowNumber, 9).Value = slovar(kljuc)(7)
            ws.Cells(rowNumber, 10).Value = slovar(kljuc)(6)
            ws.Cells(rowNumber, 11).Value = slovar(kljuc)(4)
            ws.Cells(rowNumber, 12).Value = slovar(kljuc)(5)
            ws.Cells(rowNumber, 15).Value = slovar(kljuc)(12)
            ws.Cells(rowNumber, 16).Value = slovar(kljuc)(13)
            ws.Cells(rowNumber, 17).Value = slovar(kljuc)(14)
            ws.Cells(rowNumber, 18).Value = slovar(kljuc)(15)
            ws.Cells(rowNumber, 19).Value = slovar(kljuc)(2)
            ws.Cells(rowNumber, 20).Value = slovar(kljuc)(3)
            ws.Cells(rowNumber, 21).Value = slovar(kljuc)(8)
        End If
        rowNumber = rowNumber + 1
    Loop
End Sub

Private Sub CustomerReference()
    Dim ws As Worksheet
    Dim niz As String
    Set ws = ThisWorkbook.Sheets("OBRACUN")
    Dim rowNumber As Integer
    Dim nizDatum As String

    rowNumber = 2

    Do
        If rowNumber = (steviloStrank + 1) Then
            Exit Do
        End If

        nizDatum = ws.Cells(rowNumber, 4)
        If nizDatum <> "" Then
            niz = "MRN " & Right(ws.Cells(rowNumber, 3), 6) & " DNE " & nizDatum
        Else
            niz = "MRN " & Right(ws.Cells(rowNumber, 3), 6) & " DNE " & Format(datum, "DD.MM.YYYY")
            ws.Cells(rowNumber, 4).Value = datum
        End If
        ws.Cells(rowNumber, 2).Value = niz
        rowNumber = rowNumber + 1
    Loop
End Sub

Private Sub InputDatum()
    Dim niz As String
    Dim dayPart As Integer
    Dim monthPart As Integer
    Dim yearPart As Integer
    niz = InputBox("Vnesi datum zadnjega delovnega dne [DD.MM.YYYY]:", "Datum")
    
    dayPart = CInt(Mid(niz, 1, 2))
    monthPart = CInt(Mid(niz, 4, 2))
    yearPart = CInt(Mid(niz, 7, 4))
    datum = DateSerial(yearPart, monthPart, dayPart)
End Sub

Private Sub Fizikalci()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("OBRACUN")
    Dim rowNumber As Integer
    
    rowNumber = 2
    
    Do
        If rowNumber = (steviloStrank + 1) Then
            Exit Do
        End If
        If ws.Cells(rowNumber, 8) <> "" And ws.Cells(rowNumber, 7) = "" Then
            ws.Cells(rowNumber, 7).Value = "SIA5555555"
        End If
        rowNumber = rowNumber + 1
    Loop
End Sub

Private Sub Davek()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("OBRACUN")
    Dim rowNumber As Integer
    
    rowNumber = 2
    
    Do
        If rowNumber = (steviloStrank + 1) Then
            Exit Do
        End If
        If ws.Cells(rowNumber, 11) = "Da" Then
            ws.Cells(rowNumber, 9).Value = ""
        End If
        ws.Cells(rowNumber, 11) = ""
        rowNumber = rowNumber + 1
    Loop
End Sub

Private Sub OpravilnaStevilka()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("OBRACUN")
    Dim rowNumber As Integer
    Dim opravilna As String
    
    rowNumber = 2
    
    Do
        If rowNumber = steviloStrank + 1 Then
            Exit Do
        End If
        opravilna = UCase(ws.Cells(rowNumber, 5).Value)
        If vNizu("ROD", opravilna) Then
            ws.Cells(rowNumber, 3).Value = ws.Cells(rowNumber, 3) & " ROD"
            ws.Cells(rowNumber, 5).Value = "CPT"
        ElseIf vNizu("DDP", opravilna) Then
            ws.Cells(rowNumber, 5).Value = "DDP"
        ElseIf vNizu("VRA", opravilna) Then
            ws.Cells(rowNumber, 5).Value = "VRACILO 42"
        ElseIf vNizu("ZA", opravilna) Then
            ws.Cells(rowNumber, 5).Value = "ZACASNI UVOZ 42"
        ElseIf vNizu("SANI", opravilna) Then
            ws.Cells(rowNumber, 5).Value = "SANITARC/DRUG VLADNI ORGAN"
        Else
            ws.Cells(rowNumber, 5).Value = "CPT"
        End If
        rowNumber = rowNumber + 1
    Loop
End Sub

Private Sub CL4CL6CL8()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("OBRACUN")
    Dim rowNumber As Integer
    
    rowNumber = 2
    Do
        If rowNumber = steviloStrank + 1 Then
            Exit Do
        End If
        If ws.Cells(rowNumber, 19) = "H3" Or ws.Cells(rowNumber, 19) = "H4" Then
            ws.Cells(rowNumber, 5).Value = "ZACASNI UVOZ 42"
        End If
        If ws.Cells(rowNumber, 20) = "61" Then
            ws.Cells(rowNumber, 5).Value = "VRACILO 42"
        End If
        ws.Cells(rowNumber, 19).Value = ""
        ws.Cells(rowNumber, 20).Value = ""
        ws.Cells(rowNumber, 21).Value = ""
        rowNumber = rowNumber + 1
    Loop
End Sub

Private Sub IzbrisiVrstice()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("OBRACUN")
    Dim rowNumber As Integer
    
    rowNumber = 2
    
    While ws.Cells(rowNumber, 1) <> ""
        If ws.Cells(rowNumber, 7) <> "" And ws.Cells(rowNumber, 8) Like ws.Cells(rowNumber, 18) And ws.Cells(rowNumber, 12) = "" And ws.Cells(rowNumber, 5) <> "VRACILO 42" And ws.Cells(rowNumber, 5) <> "ZACASNI UVOZ 42" And ws.Cells(rowNumber, 5) <> "SANITARC/DRUG VLADNI ORGAN" Then
            ws.Rows(rowNumber).Delete
        ElseIf ws.Cells(rowNumber, 7) <> "" And ws.Cells(rowNumber, 9) = "" And ws.Cells(rowNumber, 10) = "" And ws.Cells(rowNumber, 11) = "" And ws.Cells(rowNumber, 12) = "" And ws.Cells(rowNumber, 13) = "" And ws.Cells(rowNumber, 14) = "" And ws.Cells(rowNumber, 5) <> "VRACILO 42" And ws.Cells(rowNumber, 5) <> "ZACASNI UVOZ 42" And ws.Cells(rowNumber, 5) <> "SANITARC/DRUG VLADNI ORGAN" Then
            ws.Rows(rowNumber).Delete
        Else
            rowNumber = rowNumber + 1
        End If
    Wend
End Sub

Private Sub CL0()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("OBRACUN")
    Dim rowNumber As Integer
    
    rowNumber = 2
    
    Do
        If rowNumber = (steviloStrank + 1) Then
            Exit Do
        End If
        ws.Cells(rowNumber, 12).Value = Cl0_Funkcija(ws.Cells(rowNumber, 12))
        ws.Cells(rowNumber, 12).NumberFormat = "0.00"
        rowNumber = rowNumber + 1
    Loop
End Sub

Private Sub CL5()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("OBRACUN")
    Dim rowNumber As Integer
    
    rowNumber = 2
    
    Do
        If rowNumber = (steviloStrank + 1) Then
            Exit Do
        End If
        ws.Cells(rowNumber, 14).Value = Cl5_Funkcija(ws.Cells(rowNumber, 9), ws.Cells(rowNumber, 10))
        ws.Cells(rowNumber, 14).NumberFormat = "0.00"
        rowNumber = rowNumber + 1
    Loop
End Sub

Private Sub UvoziTrinet()
    Dim stevecVrstic As Integer
    Dim rowNumber As Integer
    Dim hs As Worksheet
    Set hs = ThisWorkbook.Sheets.Add
    hs.Name = "trinet" ' Ime lista
    Dim ws As Worksheet
    Dim filePath As Variant
    
    Set ws = ThisWorkbook.Sheets("trinet")
    filePath = Application.GetOpenFilename("CSV Datoteke (*.csv), *.csv", , "Izberi Trinet CSV datoteko")
    
    If filePath <> False Then
        With ws.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=ws.Range("A1"))
            .TextFilePlatform = 65001
            .TextFileOtherDelimiter = "|"
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = True
            .TextFileCommaDelimiter = False
            .TextFileColumnDataTypes = Array(1)
            .TextFileDecimalSeparator = "."
            .TextFileThousandsSeparator = ","
            .TextFileTrailingMinusNumbers = True
            .TextFilePlatform = xlWindows
            .Refresh
        End With
    Else
        MsgBox "Niste izbrali datoteke."
    End If
    rowNumber = 2
    stevecVrstic = 2
    Do
        If ws.Cells(rowNumber, 1) = "" Then
            Exit Do
        End If
        rowNumber = rowNumber + 1
        stevecVrstic = stevecVrstic + 1
    Loop
    rowNumber = 2
    Do
        If rowNumber = stevecVrstic Then
            Exit Do
        End If
        If ws.Cells(rowNumber, 8) <> "" Then
            ws.Cells(rowNumber, 8).Value = CDbl(ws.Cells(rowNumber, 8))
            ws.Cells(rowNumber, 8).NumberFormat = "0.00"
        End If
        If ws.Cells(rowNumber, 9) <> "" Then
            ws.Cells(rowNumber, 9).Value = CDbl(ws.Cells(rowNumber, 9))
            ws.Cells(rowNumber, 9).NumberFormat = "0.00"
        End If
        rowNumber = rowNumber + 1
    Loop
End Sub

Private Sub UvoziCarmSkeniranje()
    Dim stevecVrstic As Integer
    Dim rowNumber As Integer
    Dim hs As Worksheet
    Set hs = ThisWorkbook.Sheets.Add
    hs.Name = "carm_scan" ' Ime lista
    Dim ws As Worksheet
    Dim filePath As Variant
    
    Set ws = ThisWorkbook.Sheets("carm_scan")
    filePath = Application.GetOpenFilename("CSV Datoteke (*.csv), *.csv", , "Izberi CarM skeniranje")
    
    If filePath <> False Then
        With ws.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=ws.Range("A1"))
            .TextFilePlatform = 65001
            .TextFileOtherDelimiter = "|"
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = True
            .TextFileCommaDelimiter = False
            .TextFileColumnDataTypes = Array(1)
            .TextFileDecimalSeparator = "."
            .TextFileThousandsSeparator = ","
            .TextFileTrailingMinusNumbers = True
            .TextFilePlatform = xlWindows
            .Refresh
        End With
    Else
        MsgBox "Niste izbrali datoteke."
    End If
    rowNumber = 2
    stevecVrstic = 2
    Do
        If ws.Cells(rowNumber, 1) = "" Then
            Exit Do
        End If
        rowNumber = rowNumber + 1
        stevecVrstic = stevecVrstic + 1
    Loop
    rowNumber = 2
    Do
        If rowNumber = stevecVrstic Then
            Exit Do
        End If
        If ws.Cells(rowNumber, 1) <> "" Then
            ws.Cells(rowNumber, 1).NumberFormat = "0"
        End If
        rowNumber = rowNumber + 1
    Loop
End Sub

Private Sub UstvariOBRACUN()
    Dim filterRange As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "OBRACUN"
    ws.Cells(1, 1).Value = "Con Note"
    ws.Cells(1, 2).Value = "Customer reference"
    ws.Cells(1, 3).Value = "MRN"
    ws.Cells(1, 4).Value = "Date"
    ws.Cells(1, 5).Value = "Opravilna stevilka"
    ws.Cells(1, 6).Value = "Customer name"
    ws.Cells(1, 7).Value = "VAT ID"
    ws.Cells(1, 8).Value = "customer_name"
    ws.Cells(1, 9).Value = "VT-B00"
    ws.Cells(1, 10).Value = "DT-A00"
    ws.Cells(1, 11).Value = "HF"
    ws.Cells(1, 12).Value = "CL0"
    ws.Cells(1, 13).Value = "CL3"
    ws.Cells(1, 14).Value = "CL5"
    ws.Cells(1, 15).Value = "AD0"
    ws.Cells(1, 16).Value = "AF2"
    ws.Cells(1, 17).Value = "CC1"
    ws.Cells(1, 18).Value = "CG1"
    ws.Cells(1, 19).Value = "CL4"
    ws.Cells(1, 20).Value = "CL6"
    ws.Cells(1, 21).Value = "CL8"
    
    Set filterRange = ws.Range("A1:U1") ' Nastavimo zeljeni obseg
    filterRange.AutoFilter
End Sub
