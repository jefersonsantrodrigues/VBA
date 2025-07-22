Sub RemoverLinhasVazias()
    Dim ultimaLinha As Long
    Dim i As Long
    Dim linhaVazia As Boolean

    'Identifica a última linha usada na planilha
    ultimaLinha = Cells(Rows.Count, 1).End(xlUp).Row

    ' Percorre de baixo para cima para evitar pular linhas após exclusão
    For i = ultimaLinha To 1 Step -1
        linhaVazia = WorksheetFunction.CountA(Rows(i)) = 0
        If linhaVazia Then
            Rows(i).Delete
        End If
    Next i

    MsgBox "Linhas vazias removidas com sucesso!", vbInformation
End Sub
