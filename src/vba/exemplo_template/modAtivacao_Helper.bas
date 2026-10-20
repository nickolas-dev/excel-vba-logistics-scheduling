Attribute VB_Name = "modAtivacao_Helper"
Public Sub SincronizarDocumentoAtivacao(ws As Worksheet, Target As Range)

    ' Não roda no modelo
    If LCase(ThisWorkbook.Name) = "template_modelo.xlsm" Then Exit Sub

    Dim tbl As ListObject
    Dim rngValido As Range
    Dim r As Range
    Dim linha As Long
    Dim valorA As String, valorB As String
    Dim cd As String, DataFat As String

    Set tbl = ws.ListObjects("tbl_Ativacao")

    ' ? Ignora qualquer alteração fora do corpo da tabela
    Set rngValido = Intersect(Target, tbl.DataBodyRange)
    If rngValido Is Nothing Then Exit Sub

    Application.EnableEvents = False

    ' Percorre TODAS as linhas afetadas (digitar ou colar)
    For Each r In rngValido.Rows

        linha = r.Row
        valorA = Trim(ws.Cells(linha, "A").Value)
        valorB = Trim(ws.Cells(linha, "B").Value)
        
        'cd = ObterCDDoArquivo
    
        cd = ObterCDPorRota(valorB)
        DataFat = ObterDataFaturamentoDoArquivo

        ' =====================================================
        ' REGRA FORTE: apagou A ou B ? limpa A:E
        ' =====================================================
        If (Not Intersect(Target, ws.Cells(linha, "A")) Is Nothing And valorA = "") _
        Or (Not Intersect(Target, ws.Cells(linha, "B")) Is Nothing And valorB = "") Then

            ws.Range("A" & linha & ":E" & linha).ClearContents
            GoTo ProximaLinha
        End If

        ' =====================================================
        ' A ? B dinâmicos
        ' =====================================================
        If Not Intersect(Target, ws.Cells(linha, "A")) Is Nothing And valorA <> "" Then
            ws.Cells(linha, "B").Value = Mid(valorA, 3, 6)
        End If

        If Not Intersect(Target, ws.Cells(linha, "B")) Is Nothing And valorB <> "" Then
            ws.Cells(linha, "A").Value = cd & valorB & DataFat
        End If

        ' =====================================================
        ' ? MERGE / SOBRESCRITA DE DUPLICADOS
        ' =====================================================
        Call UnificarLinhaAtivacao(ws, linha)

ProximaLinha:
    Next r

    Application.EnableEvents = True

End Sub

Public Function ObterCDDoArquivo() As String
    ' Ex: Template FT - 28.03.26.xlsm
    ObterCDDoArquivo = Mid(ThisWorkbook.Name, 10, 2)
End Function

Public Function ObterDataFaturamentoDoArquivo() As String

    Dim parteData As String
    Dim dataConvertida As Date

    ' Pega "28.03.26"
    parteData = Mid(ThisWorkbook.Name, InStr(ThisWorkbook.Name, "-") + 2, 8)

    ' Converte para Date
    dataConvertida = DateSerial( _
        2000 + Right(parteData, 2), _
        Mid(parteData, 4, 2), _
        Left(parteData, 2))

    ObterDataFaturamentoDoArquivo = Format(dataConvertida, "dd/mm/yyyy")

End Function

Public Sub UnificarLinhaAtivacao(ws As Worksheet, linhaAtual As Long)

    Dim tbl As ListObject
    Dim dictA As Object, dictB As Object
    Dim r As Long, linhaReal As Long
    Dim chaveA As String, chaveB As String
    Dim linhaExistente As Long

    Set tbl = ws.ListObjects("tbl_Ativacao")
    Set dictA = CreateObject("Scripting.Dictionary")
    Set dictB = CreateObject("Scripting.Dictionary")

    ' Indexa a tabela
    For r = 1 To tbl.DataBodyRange.Rows.Count

        linhaReal = tbl.DataBodyRange.Rows(r).Row

        If linhaReal <> linhaAtual Then

            If Trim(ws.Cells(linhaReal, "A").Value) <> "" Then
                dictA(ws.Cells(linhaReal, "A").Value) = linhaReal
            End If

            If Trim(ws.Cells(linhaReal, "B").Value) <> "" Then
                dictB(ws.Cells(linhaReal, "B").Value) = linhaReal
            End If

        End If
    Next r

    chaveA = Trim(ws.Cells(linhaAtual, "A").Value)
    chaveB = Trim(ws.Cells(linhaAtual, "B").Value)
    linhaExistente = 0

    If chaveA <> "" And dictA.Exists(chaveA) Then
        linhaExistente = dictA(chaveA)
    ElseIf chaveB <> "" And dictB.Exists(chaveB) Then
        linhaExistente = dictB(chaveB)
    End If

    ' Encontrou duplicado ? MERGE
    If linhaExistente > 0 Then

        Call MesclarLinhas(ws, linhaExistente, linhaAtual)

        ' Remove a linha recém-colada (evita duplicata)
        ws.Range("A" & linhaAtual & ":E" & linhaAtual).ClearContents

    End If

End Sub

Private Sub MesclarLinhas(ws As Worksheet, linhaBase As Long, linhaNova As Long)

    Dim col As Long

    ' C:E apenas se vierem com valor
    For col = 3 To 5
        If Trim(ws.Cells(linhaNova, col).Value) <> "" Then
            ws.Cells(linhaBase, col).Value = ws.Cells(linhaNova, col).Value
        End If
    Next col

End Sub

Public Function ObterCDPorRota(ByVal Rota As String) As String

    Dim wsFormulas As Worksheet
    Dim UltimaLinha As Long
    Dim i As Long

    Dim Embarque As String
    Dim DataFat As String
    Dim RotaExtraida As String

    Set wsFormulas = ThisWorkbook.Sheets("Formulas")

    UltimaLinha = _
        wsFormulas.Cells(wsFormulas.Rows.Count, "D").End(xlUp).Row

    For i = 2 To UltimaLinha

        Embarque = Trim(wsFormulas.Cells(i, "D").Value)

        If Embarque <> "" Then

            DataFat = Right(Embarque, 10)

            RotaExtraida = Mid( _
                Embarque, _
                3, _
                Len(Embarque) - 12)

            If Trim(RotaExtraida) = Trim(Rota) Then

                ObterCDPorRota = Left(Embarque, 2)

                Exit Function

            End If

        End If

    Next i

    'Fallback para distribuição
    ObterCDPorRota = ObterCDDoArquivo

End Function

Public Sub AtualizarTodaAtivacao()

    Dim ws As Worksheet
    Dim UltimaLinha As Long

    Set ws = ThisWorkbook.Sheets("Ativação")

    UltimaLinha = _
        ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    If UltimaLinha < 2 Then Exit Sub

    Call SincronizarDocumentoAtivacao( _
        ws, _
        ws.Range("B2:B" & UltimaLinha))

    Call Atualizar_Ativacao( _
        ThisWorkbook.Sheets("Formulas"), _
        ThisWorkbook.Sheets("Ativação"))

    'Call AjustarTabelaAtivacao(ws)

End Sub
