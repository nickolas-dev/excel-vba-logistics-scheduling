Attribute VB_Name = "modUX"
Public Sub AjustarTabelaAtivacao(ws As Worksheet)

    Dim tbl As ListObject
    Dim UltimaLinha As Long
    Dim i As Long

    Set tbl = ws.ListObjects("tbl_Ativacao")

    ' Descobrir a última linha com dados em A ou B
    UltimaLinha = tbl.HeaderRowRange.Row

    For i = ws.Rows.Count To tbl.HeaderRowRange.Row + 1 Step -1
        If Trim(ws.Cells(i, "A").Value) <> "" _
           Or Trim(ws.Cells(i, "B").Value) <> "" Then
            UltimaLinha = i
            Exit For
        End If
    Next i

    ' Redimensionar a tabela somente até a última linha com dados
    tbl.Resize ws.Range(tbl.HeaderRowRange.Cells(1, 1), ws.Cells(UltimaLinha, tbl.Range.Columns.Count))

    ' UX: posicionar o cursor
    ws.Range("A2").Activate

End Sub

