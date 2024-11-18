Sub izbrisiVrstice()
  ' makro nam izbrise vrstice v obsegu od 1 do 3000
    Dim ws As Worksheet
    Dim i As Integer
    
    Set ws = Worksheets("FDX")
    
    i = 1
    
    Do While i <> 3000
        If ws.Cells(i, 1) = "" Then
            ws.Rows(i).Delete
        End If
        i = i + 1
    Loop
End Sub
