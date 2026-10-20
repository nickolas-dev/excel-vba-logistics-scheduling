Attribute VB_Name = "modFormatacao"
Public Sub NormalizarFormatoAtivacao(rng As Range)

    With rng
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = False
        .Font.Italic = False
        .Font.Subscript = False
        .Font.Superscript = False

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Interior.Pattern = xlNone
    End With

    ' Garantir altura da linha (aplicar para TODAS as linhas afetadas)
    Dim linha As Range
    For Each linha In rng.Rows
        linha.RowHeight = 14.5
    Next linha

End Sub
