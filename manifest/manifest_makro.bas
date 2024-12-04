Dim stPrazne As Long
Dim stT1 As Integer

Private Sub PrestejPrazne()
    Dim ws As Worksheet
    
    Set ws = Worksheets("Transit Manifest")
    stPrazne = 1
    Do While ws.Cells(stPrazne, 1) = ""
        stPrazne = stPrazne + 1
    Loop
End Sub

Private Sub PrestejT1()
    Dim ws As Worksheet
    
    Set ws = Worksheets("Transit Manifest")
    stT1 = 1
    Do While ws.Cells(stT1, 14) <> "T1"
        stT1 = stT1 + 1
    Loop
End Sub

Sub Manifest()
    Dim i As Long
    Dim ws As Worksheet
    Dim st As Long
    
    Set ws = Worksheets("Transit Manifest")
    i = 1
    
    Do While i < 10000
        If ws.Cells(i, 5).Value <> "" And Not Application.WorksheetFunction.IsText(ws.Cells(i, 5).Value) _
        And IsNumeric(ws.Cells(i, 5).Value) And Int(Val(ws.Cells(i, 5).Value)) > 1 And ws.Cells(i, 14).Value = "T1" Then
            st = 1
            ws.Cells(i, 6).Value = Round(CDbl(ws.Cells(i, 6).Value) / ws.Cells(i, 5), 2)
            ws.Cells(i, 12).Value = Round(CDbl(ws.Cells(i, 12).Value) / ws.Cells(i, 5), 2)
            Do While ws.Cells(i + st, 6).Value = ""
                ws.Cells(i + st, 6).Value = ws.Cells(i, 6).Value
                ws.Cells(i + st, 5).Value = ws.Cells(i, 5).Value
                ws.Cells(i + st, 4).Value = ws.Cells(i, 4).Value
                ws.Cells(i + st, 7).Value = ws.Cells(i, 7).Value
                ws.Cells(i + st, 8).Value = ws.Cells(i, 8).Value
                ws.Cells(i + st, 9).Value = ws.Cells(i, 9).Value
                ws.Cells(i + st, 10).Value = ws.Cells(i, 10).Value
                ws.Cells(i + st, 11).Value = ws.Cells(i, 11).Value
                ws.Cells(i + st, 12).Value = ws.Cells(i, 12).Value
                ws.Cells(i + st, 13).Value = ws.Cells(i, 13).Value
                
                st = st + 1
            Loop
            
            i = i + st
            
        Else
            i = i + 1
        End If
    Loop
End Sub

Sub UstvariSLO()
    PrestejPrazne
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim i As Integer
    
    Set ws1 = ThisWorkbook.Worksheets.Add
    ws1.Name = "SLO"
    Set ws2 = Worksheets("Transit Manifest")
    
    i = 1
    
    Do While i <> stPrazne + 1
        ws1.Cells(i, 1).Value = ws2.Cells(i, 1)
        ws1.Cells(i, 2).Value = ws2.Cells(i, 2)
        ws1.Cells(i, 3).Value = ws2.Cells(i, 3)
        ws1.Cells(i, 4).Value = ws2.Cells(i, 4)
        ws1.Cells(i, 5).Value = ws2.Cells(i, 5)
        ws1.Cells(i, 6).Value = ws2.Cells(i, 6)
        ws1.Cells(i, 7).Value = ws2.Cells(i, 7)
        ws1.Cells(i, 8).Value = ws2.Cells(i, 8)
        ws1.Cells(i, 9).Value = ws2.Cells(i, 9)
        ws1.Cells(i, 10).Value = ws2.Cells(i, 10)
        ws1.Cells(i, 11).Value = ws2.Cells(i, 11)
        ws1.Cells(i, 12).Value = ws2.Cells(i, 12)
        ws1.Cells(i, 13).Value = ws2.Cells(i, 13)
        ws1.Cells(i, 14).Value = ws2.Cells(i, 14)
        i = i + 1
    Loop
    
    PrestejT1
    
    i = stT1
    
    Do While ws2.Cells(stT1, 14).Value = "T1"
        If ws2.Cells(stT1, 10).Value = "SI" Then
            ws1.Cells(i, 1).Value = ws2.Cells(stT1, 1)
            ws1.Cells(i, 2).Value = ws2.Cells(stT1, 2)
            ws1.Cells(i, 3).Value = ws2.Cells(stT1, 3)
            ws1.Cells(i, 4).Value = ws2.Cells(stT1, 4)
            ws1.Cells(i, 5).Value = ws2.Cells(stT1, 5)
            ws1.Cells(i, 6).Value = ws2.Cells(stT1, 6)
            ws1.Cells(i, 7).Value = ws2.Cells(stT1, 7)
            ws1.Cells(i, 8).Value = ws2.Cells(stT1, 8)
            ws1.Cells(i, 9).Value = ws2.Cells(stT1, 9)
            ws1.Cells(i, 10).Value = ws2.Cells(stT1, 10)
            ws1.Cells(i, 11).Value = ws2.Cells(stT1, 11)
            ws1.Cells(i, 12).Value = ws2.Cells(stT1, 12)
            ws1.Cells(i, 13).Value = ws2.Cells(stT1, 13)
            ws1.Cells(i, 14).Value = ws2.Cells(stT1, 14)
            
            ws1.Cells(i, 2).NumberFormat = "#0"
            ws1.Cells(i, 3).NumberFormat = "#0"
            i = i + 1
        End If
        
        stT1 = stT1 + 1
    Loop
End Sub

Sub UstvariHR()
    PrestejPrazne
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim i As Integer
    
    Set ws1 = ThisWorkbook.Worksheets.Add
    ws1.Name = "HR"
    Set ws2 = Worksheets("Transit Manifest")
    
    i = 1
    
    Do While i <> stPrazne + 1
        ws1.Cells(i, 1).Value = ws2.Cells(i, 1)
        ws1.Cells(i, 2).Value = ws2.Cells(i, 2)
        ws1.Cells(i, 3).Value = ws2.Cells(i, 3)
        ws1.Cells(i, 4).Value = ws2.Cells(i, 4)
        ws1.Cells(i, 5).Value = ws2.Cells(i, 5)
        ws1.Cells(i, 6).Value = ws2.Cells(i, 6)
        ws1.Cells(i, 7).Value = ws2.Cells(i, 7)
        ws1.Cells(i, 8).Value = ws2.Cells(i, 8)
        ws1.Cells(i, 9).Value = ws2.Cells(i, 9)
        ws1.Cells(i, 10).Value = ws2.Cells(i, 10)
        ws1.Cells(i, 11).Value = ws2.Cells(i, 11)
        ws1.Cells(i, 12).Value = ws2.Cells(i, 12)
        ws1.Cells(i, 13).Value = ws2.Cells(i, 13)
        ws1.Cells(i, 14).Value = ws2.Cells(i, 14)
        i = i + 1
    Loop
    
    PrestejT1
    
    i = stT1
    
    Do While ws2.Cells(stT1, 14).Value = "T1"
        If ws2.Cells(stT1, 10).Value = "HR" Then
            ws1.Cells(i, 1).Value = ws2.Cells(stT1, 1)
            ws1.Cells(i, 2).Value = ws2.Cells(stT1, 2)
            ws1.Cells(i, 3).Value = ws2.Cells(stT1, 3)
            ws1.Cells(i, 4).Value = ws2.Cells(stT1, 4)
            ws1.Cells(i, 5).Value = ws2.Cells(stT1, 5)
            ws1.Cells(i, 6).Value = ws2.Cells(stT1, 6)
            ws1.Cells(i, 7).Value = ws2.Cells(stT1, 7)
            ws1.Cells(i, 8).Value = ws2.Cells(stT1, 8)
            ws1.Cells(i, 9).Value = ws2.Cells(stT1, 9)
            ws1.Cells(i, 10).Value = ws2.Cells(stT1, 10)
            ws1.Cells(i, 11).Value = ws2.Cells(stT1, 11)
            ws1.Cells(i, 12).Value = ws2.Cells(stT1, 12)
            ws1.Cells(i, 13).Value = ws2.Cells(stT1, 13)
            ws1.Cells(i, 14).Value = ws2.Cells(stT1, 14)
            
            ws1.Cells(i, 2).NumberFormat = "#0"
            ws1.Cells(i, 3).NumberFormat = "#0"
            i = i + 1
        End If
        
        stT1 = stT1 + 1
    Loop
End Sub
