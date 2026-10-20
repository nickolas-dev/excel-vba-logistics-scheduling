Attribute VB_Name = "modPaniColetas"
Public Sub Tbl_PaniColetas()

    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet

    Dim wbDestino As Workbook
    Dim wsDestino As Worksheet

    Dim dictArquivos As Object

    Dim linha As Long
    Dim ultLinhaOrigem As Long
    Dim ultimaLinhaDestino As Long

    Dim caminhoDestino As String

    Dim dataTexto As String
    Dim dataConvertida As Date
    Dim mesExtenso As String

    Dim embarque As String
    Dim nomeArquivo As String

    Dim chaveArquivo As Variant

    Set wbOrigem = ThisWorkbook
    Set wsOrigem = wbOrigem.Sheets(1)

    Set dictArquivos = CreateObject("Scripting.Dictionary")

    'TESTE
    caminhoDestino = _
        PortfolioPath("Tbl_Prog_Testes\")
    
    'PRODUÇÃO
    'caminhoDestino = _
        PortfolioPath("Tbl_Prog\")

    dataTexto = wsOrigem.Range("B1").Value

    dataConvertida = DateValue(dataTexto)

    mesExtenso = _
        StrConv(Format(dataConvertida, "mmmm"), vbProperCase)

    ultLinhaOrigem = _
        wsOrigem.Cells(wsOrigem.Rows.Count, "P").End(xlUp).Row

    For linha = 2 To ultLinhaOrigem

        embarque = _
            Trim(wsOrigem.Cells(linha, "P").Value)

        nomeArquivo = _
            ObterArquivoTbl(embarque, mesExtenso)

        If nomeArquivo <> "" Then

            '=========================
            ' ABRIR / CRIAR ARQUIVO
            '=========================

            If Not dictArquivos.Exists(nomeArquivo) Then

                Set wbDestino = Nothing

                If Dir(caminhoDestino & nomeArquivo) <> "" Then

                    Set wbDestino = _
                        Workbooks.Open( _
                        caminhoDestino & nomeArquivo)

                Else

                    Call CriarArquivoDoZero( _
                        caminhoDestino, _
                        nomeArquivo)

                    Set wbDestino = _
                        Workbooks.Open( _
                        caminhoDestino & nomeArquivo)

                End If

                dictArquivos.Add _
                    nomeArquivo, _
                    wbDestino

            Else

                Set wbDestino = _
                    dictArquivos(nomeArquivo)

            End If

            '=========================
            ' GRAVAR DADOS
            '=========================

            Set wsDestino = wbDestino.Sheets(1)

            ultimaLinhaDestino = _
                wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row + 1
            
            '=============================
            ' VERIFICA DUPLICIDADE
            ' EMBARQUE1 + LOJA
            '=============================
            
            Dim EmbarquePani As String
            Dim Pos As Variant
            
            EmbarquePani = _
                Trim$(wsOrigem.Cells(linha, "P").Value) & _
                Trim$(wsOrigem.Cells(linha, "I").Value)
            
            If Len(EmbarquePani) > 0 Then
            
                Pos = Application.Match( _
                    EmbarquePani, _
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
            
            wsDestino.Cells(ultimaLinhaDestino, "H").Value = _
                wsOrigem.Cells(linha, "K").Value
            
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
            
            If wsOrigem.Cells(linha, "E").Value = "GR" Then
            
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
            

        End If

ProximaLinha:

    Next linha

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
        
        dictArquivos(chaveArquivo).Save

    Next chaveArquivo
    
    '=====================================
    ' PONTO FOCAL
    '=====================================
    
    Dim wbPontoFocalEngine As Workbook
    Dim wbAtual As Workbook
    
    Application.ScreenUpdating = False
    
    'TESTE
    
    Set wbPontoFocalEngine = Workbooks.Open( _
        PortfolioPath("criacao_prog_v5_6_testes.xlsm"), _
        UpdateLinks:=False)
        
    'PRODUÇÃO
    
    'Set wbPontoFocalEngine = Workbooks.Open( _
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

End Sub
