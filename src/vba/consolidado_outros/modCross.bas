Attribute VB_Name = "modCross"
Sub Tbl_CrossDocking()

    On Error GoTo CleanExit

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet

    Dim wbDestino As Workbook
    Dim wsDestino As Worksheet

    Dim dictArquivos As Object

    Dim linha As Long
    Dim ultLinhaOrigem As Long
    Dim ultimaLinhaDestino As Long

    Dim nomeArquivo As String
    Dim caminhoDestino As String

    Dim dataTexto As String
    Dim dataConvertida As Date
    Dim mesExtenso As String

    Dim embarque As String
    Dim cdFat As String

    Set wbOrigem = ThisWorkbook
    Set wsOrigem = wbOrigem.Sheets(1)

    Set dictArquivos = CreateObject("Scripting.Dictionary")

    'Produção
    'caminhoDestino = _
        PortfolioPath("Tbl_Prog\")
        
    'Teste
    caminhoDestino = _
        PortfolioPath("Tbl_Prog_Testes\")

    dataTexto = wsOrigem.Range("B1").Value

    dataConvertida = DateValue(dataTexto)

    mesExtenso = StrConv(Format(dataConvertida, "mmmm"), vbProperCase)

    ultLinhaOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, "P").End(xlUp).Row

    For linha = 2 To ultLinhaOrigem

        embarque = Trim(wsOrigem.Cells(linha, "P").Value)

        cdFat = Trim(wsOrigem.Cells(linha, "E").Value)

        nomeArquivo = ""

        '=============================
        ' DEFINIÇÃO DO DESTINO
        '=============================

        If InStr(1, embarque, "TRJD", vbTextCompare) > 0 Then
        
            nomeArquivo = _
                "Tbl_Transfer_Cross_JD_" & _
                mesExtenso & ".xlsx"
        
        ElseIf cdFat = "GR" Then
        
            nomeArquivo = _
                "Tbl_Transfer_Cross_GR_" & _
                mesExtenso & ".xlsx"
        
        ElseIf cdFat = "JC" _
            Or cdFat = "NE" _
            Or cdFat = "PR" Then
        
            nomeArquivo = _
                "Tbl_Transfer_Cross_JC_" & _
                mesExtenso & ".xlsx"
        
        End If

        If nomeArquivo = "" Then GoTo ProximaLinha

        If Not dictArquivos.Exists(nomeArquivo) Then

            Set wbDestino = Nothing
            
            If Dir(caminhoDestino & nomeArquivo) <> "" Then
            
                On Error Resume Next
            
                Set wbDestino = Workbooks.Open( _
                    caminhoDestino & nomeArquivo, _
                    UpdateLinks:=0)
            
                On Error GoTo 0
            
            End If


            If wbDestino Is Nothing Then

                Call CriarArquivoDoZero(caminhoDestino, nomeArquivo)

                Set wbDestino = _
                    Workbooks.Open(caminhoDestino & nomeArquivo)

            End If

            Set wsDestino = wbDestino.Sheets(1)

            dictArquivos.Add nomeArquivo, wbDestino

        Else

            Set wbDestino = dictArquivos(nomeArquivo)
            Set wsDestino = wbDestino.Sheets(1)

        End If

        EnsureHeaderFilters wsDestino, "A1:AI1"

        ultimaLinhaDestino = _
            wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row + 1

        '=============================
        ' VERIFICA DUPLICIDADE
        ' EMBARQUE1 + LOJA
        '=============================

        Dim EmbarqueCross As String
        Dim Pos As Variant

        EmbarqueCross = _
            Trim$(wsOrigem.Cells(linha, "P").Value) & _
            Trim$(wsOrigem.Cells(linha, "I").Value)

        If Len(EmbarqueCross) > 0 Then

            Pos = Application.Match( _
                EmbarqueCross, _
                wsDestino.Columns("P"), _
                0)

            If Not IsError(Pos) Then

                GoTo ProximaLinha

            End If

        End If


        '=============================
        ' MAPEAMENTO
        '=============================

        wsDestino.Cells(ultimaLinhaDestino, "A").Value = wsOrigem.Cells(linha, "D").Value
        wsDestino.Cells(ultimaLinhaDestino, "B").Value = wsOrigem.Cells(linha, "E").Value
        wsDestino.Cells(ultimaLinhaDestino, "C").Value = wsOrigem.Cells(linha, "F").Value

        wsDestino.Cells(ultimaLinhaDestino, "D").Value = _
            Format(wsOrigem.Cells(linha, "G").Value, "mm/dd/yy")

        wsDestino.Cells(ultimaLinhaDestino, "E").Value = wsOrigem.Cells(linha, "H").Value
        wsDestino.Cells(ultimaLinhaDestino, "F").Value = wsOrigem.Cells(linha, "I").Value

        wsDestino.Cells(ultimaLinhaDestino, "G").Value = _
            Format(wsOrigem.Cells(linha, "J").Value, "mm/dd/yy")

        wsDestino.Cells(ultimaLinhaDestino, "H").Value = wsOrigem.Cells(linha, "K").Value

        ' I = G + H
        wsDestino.Cells(ultimaLinhaDestino, "I").Formula = _
            "=G" & ultimaLinhaDestino & _
            "+H" & ultimaLinhaDestino

        wsDestino.Cells(ultimaLinhaDestino, "J").Value = _
            Format(wsOrigem.Cells(linha, "M").Value, "mm/dd/yy")

        wsDestino.Cells(ultimaLinhaDestino, "K").Value = _
            wsOrigem.Cells(linha, "N").Value

        ' M = J + K
        wsDestino.Cells(ultimaLinhaDestino, "M").Formula = _
            "=J" & ultimaLinhaDestino & _
            "+K" & ultimaLinhaDestino
        
        ' N = M + L
        wsDestino.Cells(ultimaLinhaDestino, "N").Formula = _
            "=M" & ultimaLinhaDestino & _
            "+L" & ultimaLinhaDestino

        ' O = B + E + DATA
        wsDestino.Cells(ultimaLinhaDestino, "O").Formula = _
            "=B" & ultimaLinhaDestino & _
            "&E" & ultimaLinhaDestino & _
            "&TEXT(D" & ultimaLinhaDestino & ",""dd/mm/aaaa"")"
        
        ' P = O + F
        wsDestino.Cells(ultimaLinhaDestino, "P").Formula = _
            "=O" & ultimaLinhaDestino & _
            "&F" & ultimaLinhaDestino
            
        'Q:R
        Dim col As Variant

        For Each col In Array("Q", "R")
            wsDestino.Cells(ultimaLinhaDestino, col).Value = "-"
        Next col

        '=============================
        ' S
        '=============================

        If Trim(wsOrigem.Cells(linha, "T").Value) <> "" Then

            wsDestino.Cells(ultimaLinhaDestino, "S").Value = _
                wsOrigem.Cells(linha, "T").Value

        Else

            wsDestino.Cells(ultimaLinhaDestino, "S").Value = "-"

        End If

        '=============================
        ' T
        '=============================

        If Trim(wsOrigem.Cells(linha, "S").Value) <> "" Then

            wsDestino.Cells(ultimaLinhaDestino, "T").Value = _
                wsOrigem.Cells(linha, "S").Value

        Else

            wsDestino.Cells(ultimaLinhaDestino, "T").Value = "-"

        End If

        '=============================
        ' U
        '=============================

        wsDestino.Cells(ultimaLinhaDestino, "U").Value = _
            wsOrigem.Cells(linha, "U").Value

        '=============================
        ' Z
        '=============================

        wsDestino.Cells(ultimaLinhaDestino, "Z").Value = "NÃO"

        '=============================
        ' AA
        '=============================

        If cdFat = "GR" Then

            wsDestino.Cells(ultimaLinhaDestino, "AA").Value = "Fs"

        Else

            wsDestino.Cells(ultimaLinhaDestino, "AA").Value = "McD"

        End If

        '=============================
        ' AB
        '=============================

        wsDestino.Cells(ultimaLinhaDestino, "AB").Value = _
            wsOrigem.Cells(linha, "Q").Value

        '=============================
        ' AC
        '=============================

        wsDestino.Cells(ultimaLinhaDestino, "AC").Value = "NÃO"
        
        '=============================
        ' AD
        '=============================

        wsDestino.Cells(ultimaLinhaDestino, "AD").Value = ""

        '=============================
        ' AE
        '=============================

        wsDestino.Cells(ultimaLinhaDestino, "AE").Value = "-"

        '=============================
        ' AF
        '=============================

        wsDestino.Cells(ultimaLinhaDestino, "AF").Value = ""

        '=============================
        ' AG
        '=============================

        wsDestino.Cells(ultimaLinhaDestino, "AG").Formula = _
            "=I" & ultimaLinhaDestino & _
            "+AF" & ultimaLinhaDestino

        '=============================
        ' AH
        '=============================

        wsDestino.Cells(ultimaLinhaDestino, "AH").Value = "NÃO"

        '=============================
        ' AI
        '=============================

        wsDestino.Cells(ultimaLinhaDestino, "AI").Value = 0

ProximaLinha:

    Next linha

    Dim chaveArquivo As Variant

    '=========================
    ' SALVAR ARQUIVOS
    '=========================
    
    For Each chaveArquivo In dictArquivos.Keys

        Set wbDestino = dictArquivos(chaveArquivo)

        EnsureHeaderFilters wbDestino.Sheets(1), "A1:AI1"

        With wbDestino.Sheets(1)

            .Activate
        
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
        
            .Range("A1").Select

        End With

        wbDestino.Save

    Next chaveArquivo
    
    '=====================================
    ' PONTO FOCAL
    '=====================================

    Dim wbPontoFocalEngine As Workbook
    Dim wbAtual As Workbook

    Application.ScreenUpdating = False

    'PRODUÇÃO

    'Set wbPontoFocalEngine = Workbooks.Open( _
    '    PortfolioPath("criacao_prog_v5_6_testes.xlsm"), _
    '    UpdateLinks:=False)

    'TESTE

    Set wbPontoFocalEngine = Workbooks.Open( _
        PortfolioPath("criacao_prog_v5_6_testes.xlsm"), _
        UpdateLinks:=False)

    For Each wbAtual In Application.Workbooks

        If wbAtual.Name <> wbPontoFocalEngine.Name _
           And wbAtual.Name <> ThisWorkbook.Name Then

            If LCase(Right(wbAtual.Name, 5)) = ".xlsx" Then

                On Error Resume Next

                If UCase(Trim(wbAtual.Sheets(1).Range("A1").Value)) = _
                   "TIPO OPERACAO" Then

                    On Error GoTo 0

                    Application.Run _
                        "'" & wbPontoFocalEngine.Name & _
                        "'!CriarTblPontoFocalPorData", _
                        wbAtual

                Else

                    On Error GoTo 0

                End If

            End If

        End If

    Next wbAtual

    wbPontoFocalEngine.Close SaveChanges:=False

CleanExit:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub TemplateInfologOUTROS()

    Dim wbOrigem As Workbook
    Dim wbModelo As Workbook
    Dim wbDestino As Workbook
    Dim wsOrigem As Worksheet
    Dim wsFormulas As Worksheet
    Dim wsAtivacao As Worksheet

    Dim caminhoModelo As String
    Dim pastaTemplates As String
    Dim caminhoDestino As String
    Dim nomeBase As String
    Dim nomeTemplateBase As String
    Dim nomeTemplateFinal As String

    Dim contador As Long
    Dim ultimaLinha As Long
    Dim ultimaLinhaFormulas As Long
    Dim linhaDestino As Long
    Dim i As Long

    Dim NomeArquivoOrigem As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbOrigem = ThisWorkbook
    Set wsOrigem = wbOrigem.Sheets(1)

    ' === CAMINHOS ===

    caminhoModelo = _
        PortfolioPath("modelo_template.xlsm")

    'Produção
    'pastaTemplates = PortfolioPath("Templates\")

    'Teste
    pastaTemplates = _
        PortfolioPath("Templates\")

    nomeBase = wbOrigem.Name
    nomeTemplateBase = Replace(nomeBase, "Base", "Template", , 1, vbTextCompare)

    ' === VERIFICAR SE JÁ EXISTE ===

    contador = 0
    nomeTemplateFinal = nomeTemplateBase
    caminhoDestino = pastaTemplates & nomeTemplateFinal

    Do While Dir(caminhoDestino) <> ""

        contador = contador + 1

        nomeTemplateFinal = _
            Replace(nomeTemplateBase, _
            ".xlsm", _
            " (" & contador & ").xlsm")

        caminhoDestino = pastaTemplates & nomeTemplateFinal

    Loop

    ' === CRIAR CÓPIA DO MODELO ===

    Set wbModelo = Workbooks.Open(caminhoModelo, ReadOnly:=True)

    wbModelo.SaveCopyAs caminhoDestino

    wbModelo.Close False

    ' === ABRIR TEMPLATE ===

    Set wbDestino = Workbooks.Open(caminhoDestino)

    Set wbTemplateCriado = wbDestino

    Application.Run _
        "'" & wbDestino.Name & "'!SetTemplateEmCarga", True

    Set wsFormulas = wbDestino.Sheets("Formulas")
    Set wsAtivacao = wbDestino.Sheets("Ativação")

    wsFormulas.Range("2:" & wsFormulas.Rows.Count).ClearContents

    ' === FORMATAÇÃO ===

    With wsFormulas.Range("A2:U" & wsFormulas.Rows.Count)

        .Font.Name = "Calibri"
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    ultimaLinha = _
        wsOrigem.Cells(wsOrigem.Rows.Count, "I").End(xlUp).Row

    linhaDestino = 2

    NomeArquivoOrigem = UCase(wbOrigem.Name)

    For i = 2 To ultimaLinha

        '=========================
        ' MAPEAMENTO
        '=========================

        wsFormulas.Cells(linhaDestino, "R").Value = _
            wsOrigem.Cells(i, "I").Value

        wsFormulas.Cells(linhaDestino, "Q").Value = _
            wsOrigem.Cells(i, "F").Value

        wsFormulas.Cells(linhaDestino, "P").Value = _
            wsOrigem.Cells(i, "U").Value

        wsFormulas.Cells(linhaDestino, "N").Value = _
            wsOrigem.Cells(i, "O").Value

        wsFormulas.Cells(linhaDestino, "L").Value = _
            wsOrigem.Cells(i, "L").Value

        wsFormulas.Cells(linhaDestino, "D").Value = _
            wsOrigem.Cells(i, "P").Value

        '=========================
        ' V - CAIXAS
        '=========================

        If wsFormulas.Cells(linhaDestino, "M").Value <> "-" Then

            wsFormulas.Cells(linhaDestino, "V").Value = _
                wsOrigem.Cells(i, "Q").Value

        Else

            wsFormulas.Cells(linhaDestino, "V").Value = _
                TimeSerial(1, 0, 0)

        End If

        '=========================
        ' FIXOS
        '=========================

        wsFormulas.Cells(linhaDestino, "J").Value = 1260
        wsFormulas.Cells(linhaDestino, "I").Value = 1257
        wsFormulas.Cells(linhaDestino, "H").Value = 1237

        wsFormulas.Cells(linhaDestino, "B").Value = _
            "RASTREADO"

        '=========================
        ' COLUNA C
        '=========================
        
        If InStr(UCase$(NomeArquivoOrigem), _
                 "CROSS DOCKING") > 0 Then
        
            wsFormulas.Cells(linhaDestino, "C").Value = _
                "CROSS DOCKING"
        
        ElseIf InStr(UCase$(NomeArquivoOrigem), _
                     "PANI") > 0 _
            Or InStr(UCase$(NomeArquivoOrigem), _
                     "COLETAS") > 0 Then
        
            wsFormulas.Cells(linhaDestino, "C").Value = _
                "COLETA"
        
        ElseIf InStr(UCase$(NomeArquivoOrigem), _
                     "BATATAS") > 0 Then
        
            wsFormulas.Cells(linhaDestino, "C").Value = _
                "NOVOS NEGOCIOS"
        
        Else
        
            wsFormulas.Cells(linhaDestino, "C").Value = ""
        
        End If

        '=========================
        ' ATIVAÇÃO
        '=========================

        wsAtivacao.Cells(linhaDestino, "B").Value = _
            wsOrigem.Cells(i, "H").Value

        If Trim(wsOrigem.Cells(i, "S").Value) <> "" Then

            wsAtivacao.Cells(linhaDestino, "C").Value = _
                wsOrigem.Cells(i, "S").Value

        End If

        If Trim(wsOrigem.Cells(i, "T").Value) <> "" Then

            wsAtivacao.Cells(linhaDestino, "E").Value = _
                wsOrigem.Cells(i, "T").Value

        End If

        linhaDestino = linhaDestino + 1

    Next i

    ' === VALIDAÇÃO COLUNA M ===

    Dim celula As Range
    Dim erroEncontrado As Boolean

    ultimaLinhaFormulas = _
        wsFormulas.Cells(wsFormulas.Rows.Count, "R").End(xlUp).Row

    For Each celula In wsFormulas.Range("M2:M" & ultimaLinhaFormulas)

        If Trim(celula.Value) = "-" Then

            erroEncontrado = True
            Exit For

        End If

    Next celula

    If erroEncontrado Then

        MsgBox _
            "Por favor verificar registros com loja não cadastrada.", _
            vbCritical, _
            "Erro de validação"

        wsFormulas.Range("A1:U" & ultimaLinhaFormulas).AutoFilter

        wsFormulas.Range("A1:U" & ultimaLinhaFormulas) _
            .AutoFilter Field:=13, Criteria1:="-"

        FiltroLojaNaoCadastrada = True

    End If

    Application.Run _
        "'" & wbDestino.Name & _
        "'!Atualizar_Formulas_Externas", _
        wsFormulas, _
        True

    Application.Run _
        "'" & wbDestino.Name & _
        "'!Atualizar_Ativacao", _
        wsFormulas, _
        wbDestino.Sheets("Ativação")

    Call PreencherTurnosPontoFocal(wbDestino)
    
    On Error Resume Next
    
    Application.Run _
        "'" & wbDestino.Name & _
        "'!AtualizarTodaAtivacao"
    
    On Error GoTo 0

    Application.Run _
        "'" & wbDestino.Name & _
        "'!SetTemplateEmCarga", False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub
