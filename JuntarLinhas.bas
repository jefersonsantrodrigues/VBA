Sub JuntarLinhasProduto()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1) ' ajustar se for outra planilha

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = lastRow To 2 Step -1 ' de baixo pra cima
        If Trim(ws.Cells(i, 1).Value) = "" Then
        
            ' concatena valores das colunas B até última coluna usada na linha i na linha acima
            Dim col As Long
            Dim lastCol As Long
            lastCol = ws.Cells(i, ws.Columns.Count).End(xlToLeft).Column
            
            Dim j As Long
            For j = 1 To lastCol
                If Trim(ws.Cells(i, j).Value) <> "" Then
                    ws.Cells(i - 1, j).Value = ws.Cells(i - 1, j).Value & " " & ws.Cells(i, j).Value
                End If
            Next j
            
            ' Apaga a linha concatenada
            ws.Rows(i).Delete
        End If
    Next i
End Sub
