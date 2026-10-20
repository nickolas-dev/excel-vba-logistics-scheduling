Attribute VB_Name = "modpontoFocal"
Public Function ExtrairNomeBase(nomeArquivo As String) As String

    Dim nome As String
    Dim pos As Long

    nome = Replace(nomeArquivo, ".xlsx", "")
    nome = Replace(nome, ".xlsm", "")

    pos = InStr(1, nome, "Tbl_", vbTextCompare)

    If pos > 0 Then
        ExtrairNomeBase = Mid(nome, pos + 4)
    Else
        ExtrairNomeBase = nome
    End If

End Function

Public Sub CriarCabecalhosPontoFocal(ws As Worksheet)

    ws.Cells(1, 1).Value = "CD FATURAMENTO"
    ws.Cells(1, 2).Value = "DATA FATURAMENTO"
    ws.Cells(1, 3).Value = "ROTA"
    ws.Cells(1, 4).Value = "TURNO 1"
    ws.Cells(1, 5).Value = "TURNO 2"
    ws.Cells(1, 6).Value = "TURNO 3"
    ws.Cells(1, 7).Value = "ATRIBUIR"
    ws.Cells(1, 8).Value = "OPERACAO"

End Sub

Public Function CopiarDadosBasePorData( _
    wsOrigem As Worksheet, _
    wsDestino As Worksheet, _
    DataFat As Date, _
    LinhaDestino As Long, _
    dictIndice As Object) As Variant

    Dim UltLinha As Long
    Dim i As Long
    Dim linhaAtual As Long

    Dim Chave As String
    Dim arrChaves() As String
    Dim qtdNovas As Long

    UltLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "D").End(xlUp).Row
    linhaAtual = LinhaDestino
    qtdNovas = 0

    For i = 2 To UltLinha

        If IsDate(wsOrigem.Cells(i, "D").Value) Then
            If wsOrigem.Cells(i, "D").Value = DataFat Then

                Chave = _
                    wsOrigem.Cells(i, "B").Value & _
                    wsOrigem.Cells(i, "E").Value & _
                    Format(DataFat, "dd/mm/yyyy")

                ' Só copia se a chave NÃO existir no índice
                If Not dictIndice.Exists(Chave) Then

                    wsDestino.Cells(linhaAtual, "A").Value = wsOrigem.Cells(i, "B").Value
                    wsDestino.Cells(linhaAtual, "B").Value = wsOrigem.Cells(i, "D").Value
                    wsDestino.Cells(linhaAtual, "C").Value = wsOrigem.Cells(i, "E").Value
                    
                    wsDestino.Cells(linhaAtual, "H").Value = _
                        ObterOperacaoPF(wsOrigem.Cells(i, "A").Value)

                    linhaAtual = linhaAtual + 1

                    ' Guarda a chave nova
                    ReDim Preserve arrChaves(0 To qtdNovas)
                    arrChaves(qtdNovas) = Chave
                    qtdNovas = qtdNovas + 1

                End If

            End If
        End If

    Next i

    ' Retorno
    If qtdNovas > 0 Then
        CopiarDadosBasePorData = arrChaves
    Else
        CopiarDadosBasePorData = Empty
    End If

End Function

Public Sub RemoverDuplicatas(ws As Worksheet)

    Dim UltLinha As Long
    Dim i As Long
    Dim dict As Object
    Dim embarque As String

    Set dict = CreateObject("Scripting.Dictionary")

    UltLinha = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    For i = UltLinha To 2 Step -1

        embarque = _
            CStr(ws.Cells(i, "A").Value) & _
            CStr(ws.Cells(i, "C").Value) & _
            Format(ws.Cells(i, "B").Value, "dd/mm/yyyy")

        If dict.Exists(embarque) Then
            ws.Rows(i).Delete
        Else
            dict.Add embarque, True
        End If

    Next i

End Sub

Public Sub ProtegerPlanilha(ws As Worksheet)

    ws.Protect Password:="teste", _
               AllowFiltering:=True, _
               AllowSorting:=True
End Sub

Public Sub ValidacaoPontoFocal(ws As Worksheet)

    Dim UltimaLinha As Long

    Dim ListaTurno0 As String
    Dim ListaTurno1 As String
    Dim ListaTurno2 As String
    Dim ListaTurno3 As String
    Dim ListaGeral As String

    ws.Activate

    UltimaLinha = _
        ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    ws.Range("D2:G" & UltimaLinha).Validation.Delete

    ListaTurno0 = ObterListaPFPorTurno("0")
    ListaTurno1 = ObterListaPFPorTurno("1")
    ListaTurno2 = ObterListaPFPorTurno("2")
    ListaTurno3 = ObterListaPFPorTurno("3")

    ListaGeral = _
        ListaTurno0 & "," & _
        ListaTurno1 & "," & _
        ListaTurno2 & "," & _
        ListaTurno3

    '=========================
    ' TURNO 1
    '=========================

    With ws.Range("D2:D" & UltimaLinha).Validation

        .Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            formula1:=ListaTurno1

        .IgnoreBlank = True
        .InCellDropdown = True

    End With

    '=========================
    ' TURNO 2
    '=========================

    With ws.Range("E2:E" & UltimaLinha).Validation

        .Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            formula1:=ListaTurno2

        .IgnoreBlank = True
        .InCellDropdown = True

    End With

    '=========================
    ' TURNO 3
    '=========================

    With ws.Range("F2:F" & UltimaLinha).Validation

        .Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            formula1:=ListaTurno3

        .IgnoreBlank = True
        .InCellDropdown = True

    End With

    '=========================
    ' GERAL
    '=========================

    With ws.Range("G2:G" & UltimaLinha).Validation

        .Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            formula1:=ListaGeral

        .IgnoreBlank = True
        .InCellDropdown = True

    End With

End Sub

Public Sub GerarPontosFocaisSeVazio(ws1 As Worksheet)

    Dim dict As Object
    Dim dictPF As Object

    Dim UltLinha As Long
    Dim i As Long

    Dim Rota As String
    Dim Operacao As String

    Dim pontos As Variant

    Dim idx As Long

    Dim ultimoChar As String
    Dim regraEspecial As Boolean

    Dim Turno1 As Variant
    Dim Turno2 As Variant
    Dim Turno3 As Variant

    Set dict = CreateObject("Scripting.Dictionary")

    Set dictPF = CarregarPFCache()

    Turno1 = CarregarPFComPeso("1")
    Turno2 = CarregarPFComPeso("2")
    Turno3 = CarregarPFComPeso("3")

    UltLinha = ws1.Cells(ws1.Rows.Count, "C").End(xlUp).Row

    Randomize

    idx = Int(Rnd * 1000)

    For i = 2 To UltLinha

        Rota = ws1.Cells(i, "C").Value
        Operacao = UCase(Trim(ws1.Cells(i, "H").Value))

        ultimoChar = Right(ws1.Cells(i, "C").Value, 1)

        regraEspecial = False

        If ultimoChar = "M" Then

            If ws1.Cells(i, "A").Value = "CDFT" _
               Or ws1.Cells(i, "A").Value = "CDRJ" Then

                regraEspecial = True

            End If

        End If

        If Rota <> "" Then

            If ws1.Cells(i, "D").Value = "" Then

                If Not dict.Exists(Rota) Then

                    pontos = Array( _
                        Turno1(idx Mod (UBound(Turno1) + 1)), _
                        Turno2(idx Mod (UBound(Turno2) + 1)), _
                        Turno3(idx Mod (UBound(Turno3) + 1)) _
                    )

                    dict.Add Rota, pontos

                    idx = idx + 1

                End If

                '=====================================
                ' ATRIBUIÇÃO PADRÃO
                '=====================================

                ws1.Cells(i, "D").Value = dict(Rota)(0)
                ws1.Cells(i, "E").Value = dict(Rota)(1)
                ws1.Cells(i, "F").Value = dict(Rota)(2)

                '=====================================
                ' REGRAS ESPECIAIS POR OPERAÇÃO
                '=====================================

                Select Case Operacao

                    Case "PANIFRESH", "CROSS DOCKING", "BATATAS"

                        ws1.Cells(i, "D").Value = "EDUARDO"
                        ws1.Cells(i, "E").Value = "LEONARDO"
                        ws1.Cells(i, "F").Value = "MARIO"

                    Case "OUTBACK"

                        ws1.Cells(i, "D").Value = "VICTOR"
                        ws1.Cells(i, "E").Value = "LEONARDO"
                        ws1.Cells(i, "F").Value = "MARIO"

                End Select

            End If

        End If

    Next i

End Sub

Public Sub AjustarEstiloTblPontoFocal(ws As Worksheet)

    Dim tbl As ListObject

    ' Ativa a planilha (OBRIGATÓRIO para filtro e freeze)
    ws.Activate

    ' Ocultar linhas de grade
    ActiveWindow.DisplayGridlines = False

    ' Formatação geral
    With ws.Cells
        .Font.Name = "Aptos Narrow"
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Altura das linhas
    ws.Rows.RowHeight = 14.5

    '=================================================
    ' CABEÇALHO PRINCIPAL (A:G)
    '=================================================

    With ws.Range("A1:G1")

        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(25, 107, 36) ' #196B24

    End With

    '=================================================
    ' OPERACAO (H)
    '=================================================

    With ws.Range("H1")

        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(12, 118, 158) ' #0C769E

    End With

    ' Largura das colunas
    ws.Columns("A").ColumnWidth = 15.73
    ws.Columns("B").ColumnWidth = 17.18
    ws.Columns("C").ColumnWidth = 7.82
    ws.Columns("D").ColumnWidth = 9.82
    ws.Columns("E").ColumnWidth = 9.82
    ws.Columns("F").ColumnWidth = 9.82
    ws.Columns("G").ColumnWidth = 10.095
    ws.Columns("H").ColumnWidth = 11.45

    '=================================================
    ' GARANTIR FILTROS (SETINHAS)
    '=================================================
    
        If ws.FilterMode Then
        ws.ShowAllData
    End If
    
    ws.AutoFilterMode = False

    ws.Range("A1:H" & _
    ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter
    
    '=================================================
    ' CONGELAR CABEÇALHO
    '=================================================

    ActiveWindow.FreezePanes = False

    ws.Range("A2").Select

    ActiveWindow.FreezePanes = True

End Sub

Public Sub CriarTabelaPontoFocal(ws As Worksheet)
    Dim UltLinha As Long
    Dim rng As Range
    Dim tbl As ListObject

    UltLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set rng = ws.Range("A1:G" & UltLinha) ' Remove tabela existente,se houver
    If ws.ListObjects.Count > 0 Then
        ws.ListObjects(1).Unlist
    End If

    ' Cria nova tabela
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = "Tbl_PontoFocal"
    tbl.TableStyle = "" ' remove estilo visual

End Sub

Public Sub CriarTblPontoFocalPorData(wbImportado As Workbook)
    Dim wsOrigem As Worksheet
    Dim dictDatas As Object
    Dim UltLinha As Long
    Dim i As Long
    Dim DataFat As Date
    Dim vData As Variant   ' <-variável correta para For Each
    
    Dim dictIndice As Object
    Set dictIndice = CarregarIndicePontoFocal()

    Set wsOrigem = wbImportado.Sheets(1)
    Set dictDatas = CreateObject("Scripting.Dictionary")
    UltLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "D").End(xlUp).Row

    ' Coleta datas únicas
    For i = 2 To UltLinha
        If IsDate(wsOrigem.Cells(i, "D").Value) Then
            DataFat = wsOrigem.Cells(i, "D").Value
            If Not dictDatas.Exists(CLng(DataFat)) Then
                dictDatas.Add CLng(DataFat), DataFat
            End If
        End If
    Next i

    ' Processa cada data separadamente
    For Each vData In dictDatas.Items
    
        If ExisteChaveNovaParaData(wsOrigem, CDate(vData), dictIndice) Then
            Call CriarTblPontoFocal_UmaData(wbImportado, CDate(vData))
        End If
    
    Next vData

End Sub

Public Sub CriarTblPontoFocal_UmaData(wbImportado As Workbook, DataFat As Date)

    Const MODO_PRODUCAO As Boolean = False ' True = Produção | False = Teste

    Dim wsOrigem As Worksheet
    Dim wbNovo As Workbook
    Dim wsNovo As Worksheet

    Dim caminhoBase As String
    Dim caminhoAno As String
    Dim caminhoSemana As String
    Dim caminhoCompleto As String
    Dim nomeArquivo As String

    Dim AnoFat As String
    Dim semanaFat As String
    Dim ProximaLinha As Long
    Dim arquivoExiste As Boolean

    Dim arrNovasChaves As Variant   ' <<< NOVO
     
    ' Planilha de origem
    Set wsOrigem = wbImportado.Sheets(1)

    ' Ano e semana da data de faturamento
    AnoFat = CStr(Year(DataFat))
    semanaFat = "Sem" & Format(DatePart("ww", DataFat, vbMonday, vbFirstFourDays), "00")

    ' Caminho base
    If MODO_PRODUCAO Then
        caminhoBase = PortfolioPath("Tbl_PontoFocal\")
    Else
        caminhoBase = PortfolioPath("Tbl_PontoFocal\")
    End If

    caminhoAno = caminhoBase & AnoFat & "\"
    caminhoSemana = caminhoAno & semanaFat & "\"

    ' Garante pastas
    If Dir(caminhoBase, vbDirectory) = "" Then MkDir caminhoBase
    If Dir(caminhoAno, vbDirectory) = "" Then MkDir caminhoAno
    If Dir(caminhoSemana, vbDirectory) = "" Then MkDir caminhoSemana

    ' Nome e caminho completo do arquivo
    nomeArquivo = "Tbl_PontoFocal_Fat" & Format(DataFat, "dd_mm_yyyy") & ".xlsx"
    caminhoCompleto = caminhoSemana & nomeArquivo

    ' Verifica se o arquivo já existe
    arquivoExiste = (Dir(caminhoCompleto) <> "")

    ' Abre ou cria arquivo
    If arquivoExiste Then

        Set wbNovo = Workbooks.Open(caminhoCompleto)
        Set wsNovo = wbNovo.Sheets("Sheet1")

        ' Desprotege para permitir escrita
        wsNovo.Unprotect Password:="teste"

        ProximaLinha = wsNovo.Cells(wsNovo.Rows.Count, "A").End(xlUp).Row + 1

    Else

        Set wbNovo = Workbooks.Add(xlWBATWorksheet)
        Set wsNovo = wbNovo.Sheets(1)
        wsNovo.Name = "Sheet1"

        Call CriarCabecalhosPontoFocal(wsNovo)
        ProximaLinha = 2

    End If

    ' === COPIA DADOS E CAPTURA CHAVES NOVAS ===
    arrNovasChaves = CopiarDadosBasePorData( _
                        wsOrigem, _
                        wsNovo, _
                        DataFat, _
                        ProximaLinha, _
                        CarregarIndicePontoFocal _
                     )

    ' === ATUALIZA ÍNDICE SOMENTE SE HOUVER NOVAS CHAVES ===
    If Not IsEmpty(arrNovasChaves) Then
        Call AdicionarChavesNoIndice(arrNovasChaves, DataFat)
    End If

    ' Ajustes finais
    Call RemoverDuplicatas(wsNovo)
    Call ValidacaoPontoFocal(wsNovo)
    Call GerarPontosFocaisSeVazio(wsNovo)
    Call AjustarEstiloTblPontoFocal(wsNovo)

    ' ===== FIX DEFINITIVO DAS SETINHAS =====
    wsNovo.Activate

    If wsNovo.AutoFilterMode Then
        wsNovo.AutoFilterMode = False
    End If

    wsNovo.Range("A1:H1").AutoFilter

    Call ProtegerPlanilha(wsNovo)

    ' Salva e fecha
    Application.DisplayAlerts = False
    wbNovo.SaveAs caminhoCompleto, xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wbNovo.Close False

End Sub

Public Function ExisteChaveNovaParaData( _
    wsOrigem As Worksheet, _
    DataFat As Date, _
    dictIndice As Object) As Boolean

    Dim UltLinha As Long
    Dim i As Long
    Dim Chave As String

    UltLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "D").End(xlUp).Row

    For i = 2 To UltLinha

        ' ? DATA FATURAMENTO ESTÁ NA COLUNA D
        If DateValue(wsOrigem.Cells(i, "D").Value) = DateValue(DataFat) Then

            ' ? CHAVE = CD FAT (B) + ROTA (E) + DATA
            Chave = _
                wsOrigem.Cells(i, "B").Value & _
                wsOrigem.Cells(i, "E").Value & _
                Format(DataFat, "dd/mm/yyyy")

            If Not dictIndice.Exists(Chave) Then
                ExisteChaveNovaParaData = True
                Exit Function
            End If

        End If

    Next i

    ExisteChaveNovaParaData = False

End Function

Public Sub AtualizarCachePF()

    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet
    Dim tblOrigem As ListObject

    Dim wsDestino As Worksheet
    Dim tblDestino As ListObject

    Dim AtualizacaoOK As Boolean

    On Error GoTo TrataErro

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set wsDestino = _
        ThisWorkbook.Sheets("PARAMETROS_PF_CACHE")

    Set tblDestino = _
        wsDestino.ListObjects("tblPFCache")

    'Limpar cache atual
    If Not tblDestino.DataBodyRange Is Nothing Then

        tblDestino.DataBodyRange.Delete

    End If

    Set wbOrigem = Workbooks.Open( _
        PortfolioPath("..\..\Redistribuir_Ponto_Focal.xlsm"), _
        ReadOnly:=True)

    Set wsOrigem = _
        wbOrigem.Sheets("PARAMETROS_PF")

    Set tblOrigem = _
        wsOrigem.ListObjects("tblPF")

    If Not tblOrigem.DataBodyRange Is Nothing Then

        tblOrigem.DataBodyRange.Copy

        wsDestino.Range("A2").PasteSpecial xlPasteValues

        Application.CutCopyMode = False

    End If

    wbOrigem.Close SaveChanges:=False

    ThisWorkbook.Save

    AtualizacaoOK = True

Finalizar:

    On Error Resume Next

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.CutCopyMode = False

    If AtualizacaoOK Then

        MsgBox _
            "Cache de Pontos Focais atualizado com sucesso!", _
            vbInformation, _
            "Control Tower MB"

    End If

    Exit Sub

TrataErro:

    If Not wbOrigem Is Nothing Then

        wbOrigem.Close SaveChanges:=False

    End If

    MsgBox _
        "Erro ao atualizar o cache de Pontos Focais." & _
        vbCrLf & vbCrLf & _
        "Descrição: " & Err.Description, _
        vbCritical, _
        "Control Tower MB"

    Resume Finalizar

End Sub

Public Function CarregarPFCache() As Object

    Dim dict As Object

    Dim ws As Worksheet
    Dim tbl As ListObject

    Dim Linha As ListRow

    Dim NomePF As String
    Dim Turno As String

    Set dict = CreateObject("Scripting.Dictionary")

    dict.Add "0", New Collection
    dict.Add "1", New Collection
    dict.Add "2", New Collection
    dict.Add "3", New Collection

    Set ws = _
        ThisWorkbook.Sheets("PARAMETROS_PF_CACHE")

    Set tbl = _
        ws.ListObjects("tblPFCache")

    For Each Linha In tbl.ListRows

        NomePF = _
            UCase(Trim(Linha.Range.Cells(1, 1).Value))

        Turno = _
            Trim(Linha.Range.Cells(1, 2).Value)

        If dict.Exists(Turno) Then

            dict(Turno).Add NomePF

        End If

    Next Linha

    Set CarregarPFCache = dict

End Function

Public Function ObterListaPFPorTurno( _
    TurnoDesejado As String) As String

    Dim ws As Worksheet
    Dim tbl As ListObject

    Dim Linha As ListRow

    Dim Lista As String

    Set ws = _
        ThisWorkbook.Sheets("PARAMETROS_PF_CACHE")

    Set tbl = _
        ws.ListObjects("tblPFCache")

    For Each Linha In tbl.ListRows

        If Trim(Linha.Range.Cells(1, 2).Value) = _
           Trim(TurnoDesejado) Then

            Lista = Lista & _
                Linha.Range.Cells(1, 1).Value & ","

        End If

    Next Linha

    If Right(Lista, 1) = "," Then

        Lista = Left(Lista, Len(Lista) - 1)

    End If

    ObterListaPFPorTurno = Lista

End Function

Public Function CollectionParaArray(col As Collection) As Variant

    Dim arr() As String
    Dim i As Long

    ReDim arr(0 To col.Count - 1)

    For i = 1 To col.Count

        arr(i - 1) = col(i)

    Next i

    CollectionParaArray = arr

End Function

Public Function CarregarPFComPeso( _
    TurnoDesejado As String) As Variant

    Dim ws As Worksheet
    Dim tbl As ListObject

    Dim Linha As ListRow

    Dim NomePF As String
    Dim Peso As Double

    Dim arrTemp() As String
    Dim TotalItens As Long
    Dim Repeticoes As Long

    Dim i As Long

    Set ws = _
        ThisWorkbook.Sheets("PARAMETROS_PF_CACHE")

    Set tbl = _
        ws.ListObjects("tblPFCache")

    TotalItens = 0

    For Each Linha In tbl.ListRows

        If Trim(Linha.Range.Cells(1, 2).Value) = _
           Trim(TurnoDesejado) Then

            Peso = CDbl(Linha.Range.Cells(1, 3).Value)

            Repeticoes = Round(Peso * 10, 0)

            TotalItens = _
                TotalItens + Repeticoes

        End If

    Next Linha

    ReDim arrTemp(0 To TotalItens - 1)

    TotalItens = 0

    For Each Linha In tbl.ListRows

        If Trim(Linha.Range.Cells(1, 2).Value) = _
           Trim(TurnoDesejado) Then

            NomePF = _
                UCase(Trim( _
                Linha.Range.Cells(1, 1).Value))

            Peso = _
                CDbl(Linha.Range.Cells(1, 3).Value)

            Repeticoes = _
                Round(Peso * 10, 0)

            For i = 1 To Repeticoes

                arrTemp(TotalItens) = NomePF

                TotalItens = TotalItens + 1

            Next i

        End If

    Next Linha

    CarregarPFComPeso = arrTemp

End Function

Public Sub AtualizarPontosFocaisPendentes()

    Dim Caminho As String
    Dim Arquivo As String

    Dim wb As Workbook

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Caminho = _
        PortfolioPath("Tbl_Prog\")

    Arquivo = Dir(Caminho & "*.xlsx")

    Do While Arquivo <> ""

        Set wb = Workbooks.Open( _
            Caminho & Arquivo)

        Call CriarTblPontoFocalPorData(wb)

        wb.Close SaveChanges:=False

        Arquivo = Dir

    Loop

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox _
        "Verificação de Pontos Focais concluída!", _
        vbInformation, _
        "Control Tower MB"

End Sub

Public Function ObterOperacaoPF(OperacaoOrigem As String) As String

    OperacaoOrigem = UCase(Trim(OperacaoOrigem))

    Select Case True

        Case OperacaoOrigem = "DISTR OUTBACK"

            ObterOperacaoPF = "OUTBACK"

        Case Left(OperacaoOrigem, 5) = "DISTR"

            ObterOperacaoPF = "DISTRIBUICAO"

        Case OperacaoOrigem = "COLETA PANIFRESH"

            ObterOperacaoPF = "PANIFRESH"

        Case Left(OperacaoOrigem, 6) = "COLETA"

            ObterOperacaoPF = "COLETA"

        Case OperacaoOrigem = "TRANSFERENCIA"

            ObterOperacaoPF = "TRANSFERENCIA"

        Case OperacaoOrigem = "TRANSF BATATAS"

            ObterOperacaoPF = "BATATAS"

        Case Left(OperacaoOrigem, 5) = "CROSS"

            ObterOperacaoPF = "CROSS DOCKING"

        Case Else

            ObterOperacaoPF = OperacaoOrigem

    End Select

End Function

