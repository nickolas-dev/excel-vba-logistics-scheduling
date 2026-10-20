Attribute VB_Name = "Módulo1"
Public ultimaLinhaFormulas As Long
Public verificaRetorno As Long

Private Sub CommandButton5_Click() ' BOTÃO EXPORTAR
    
    'Verificando se há inconsistências
    ult_linha_analise = Range("I1").End(xlDown).Row
    
    For linha_analise = 2 To ult_linha_analise
    
        If Cells(linha_analise, "L").Value > Cells(linha_analise, "O").Value Then
        
            inconsistencia = "DTH CARRG PROG é maior do que DTH ENTR PROG"
            
            MsgBox "Não é possível exportar os dados. Verifique as inconsistências: " & vbNewLine & vbNewLine & inconsistencia, vbCritical, "Control Tower MB"
            
            Exit Sub
        
        Else
        
            If Cells(linha_analise, "X").Value = "Duplicada" Then
            
                inconsistencia = "Lojas duplicadas"
            
                MsgBox "Não é possível exportar os dados. Verifique as inconsistências: " & vbNewLine & vbNewLine & inconsistencia, vbCritical, "Control Tower MB"
                
            Else
            
                If Cells(linha_analise, "X").Value = "Alerta: Carregamento programado para 24 horas antes da primeira entrega" Then
            
                    inconsistencia = "Há rotas com carregamento programado para 24 horas antes da primeira entrega." & vbNewLine & vbNewLine & "Deseja prosseguir com a exportação dos dados?"
                
                   confirmaExportar = MsgBox(inconsistencia, vbYesNo + vbQuestion, "Control Tower MB")
                   
                    If confirmaExportar = 7 Then
                    
                         Exit Sub
                         
                    Else
                    
                         UserForm1.Show
                         
                         Exit Sub
                    
                    End If
                
                End If
            
            End If
        
        End If
        
    Next

    UserForm1.Show

End Sub

Sub Tbl_Distribuicao()

    On Error GoTo CleanExit

    Application.ScreenUpdating = False
    Application.Calculation = False
    
    Application.Visible = False
    Application.StatusBar = "Processando arquivos..."
    
    
    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet
    Dim wbDestino As Workbook
    Dim wsDestino As Worksheet
    Dim Linha As Long, ultimaLinhaDestino As Long
    Dim cd_fat2 As String, uf As String, nomeArquivo As String
    Dim caminhoDestino As String
    Dim dataTexto As String, dataConvertida As Date, mesExtenso As String
    Dim semana As String
    Dim dictArquivos As Object

    Set wbOrigem = ThisWorkbook
    Set wsOrigem = wbOrigem.Sheets(1) ' Ajuste se necessário
    Set dictArquivos = CreateObject("Scripting.Dictionary")

    'caminhoDestino = PortfolioPath("Tbl_Prog\")
    caminhoDestino = PortfolioPath("Tbl_Prog_Testes\")
    dataTexto = wsOrigem.Range("B1").Value
    semana = wsOrigem.Range("C1").Value

    ' Converte a data para obter o mês por extenso
    dataConvertida = DateValue(dataTexto)
    mesExtenso = StrConv(Format(dataConvertida, "mmmm"), vbProperCase)

    ' Percorre as linhas da planilha de origem

' >>> DECLARE AQUI (ANTES DO LOOP) <<<
Dim nomeArquivoAnterior As String
nomeArquivoAnterior = ""

ult_linha_wsOrigem = Range("I1").End(xlDown).Row

For Linha = 2 To ult_linha_wsOrigem
    cd_fat2 = Trim(wsOrigem.Cells(Linha, "E").Value)
    uf = Left(Trim(wsOrigem.Cells(Linha, "H").Value), 2)
    Dim redeLoja As String
    redeLoja = Right(Trim(wsOrigem.Cells(Linha, "H").Value), 1)
    Dim numCaractLoja As Long
    numCaractLoja = Len(wsOrigem.Cells(Linha, "I").Value)
    nomeArquivo = ""
    
    Dim tipoOper As String
    Dim isDireta As Boolean
    
    tipoOper = Trim$(UCase$(wsOrigem.Cells(Linha, "D").Value))
    ' Ex.: "DISTR CDJC - DIR" ou "DISTR CDGR - DIR"
    isDireta = (tipoOper = ("DISTR CD" & UCase$(cd_fat2) & " - DIR"))
    
    'DISTR OUTBACK
    If tipoOper = "DISTR OUTBACK" Then
    
        nomeArquivo = "Tbl_DistrOUTBACK_" & mesExtenso & ".xlsx"
    
    Else
    
        Select Case cd_fat2
            Case "NE"
                If numCaractLoja = 4 Then
                    ' Sem fallback: sempre repete o nome da linha anterior
                    nomeArquivo = nomeArquivoAnterior
                ElseIf redeLoja = "M" Then
                    Select Case numCaractLoja
                        Case 3
                            nomeArquivo = "Tbl_DistrMCD_" & cd_fat2 & "_Sem" & semana & ".xlsx"
                        Case 7
                            nomeArquivo = "Tbl_DistrFS_" & cd_fat2 & "_Sem" & semana & ".xlsx"
                    End Select
                Else
                    ' redeLoja <> "M" (ex.: "F")
                    Select Case numCaractLoja
                        Case 7
                            nomeArquivo = "Tbl_DistrFS_" & cd_fat2 & "_Sem" & semana & ".xlsx"
                        Case 3
                            nomeArquivo = "Tbl_DistrMCD_" & cd_fat2 & "_Sem" & semana & ".xlsx"
                    End Select
                End If
                    
            Case "FT"
                If uf = "ON" Then
                    nomeArquivo = "Tbl_DistrMCD_" & cd_fat2 & "_Sem" & semana & "_OnePallet" & ".xlsx"
                Else
                    nomeArquivo = "Tbl_DistrMCD_" & cd_fat2 & "_Sem" & semana & ".xlsx"
                End If
    
            Case "JC"
                If isDireta Then
                    ' DIRETAS por TIPO OPERACAO (coluna D)
                    nomeArquivo = "Tbl_DistrMCD_DIRETAS_" & mesExtenso & ".xlsx"
                Else
                    ' Regras existentes por UF
                    If (uf = "MS" Or uf = "RO" Or uf = "SM" Or uf = "TM") Then
                        nomeArquivo = "Tbl_DistrMCD_DIRETAS_" & mesExtenso & ".xlsx"
                    ElseIf uf = "BA" Or uf = "BS" Or uf = "CS" Or uf = "MG" Or uf = "IG" Or uf = "MC" Then
                        nomeArquivo = "Tbl_DistrMCD_BA_BS_MG_Sem" & semana & ".xlsx"
                    ElseIf uf = "DF" Or uf = "GO" Or uf = "MT" Or uf = "TO" Or uf = "MD" Or uf = "ML" Then
                        nomeArquivo = "Tbl_DistrMCD_DF_GO_MT_TO_Sem" & semana & ".xlsx"
                    ElseIf uf = "ES" Or uf = "RJ" Or uf = "NF" Or uf = "RL" Or uf = "SF" Or uf = "SR" Or uf = "JF" Then
                        nomeArquivo = "Tbl_DistrMCD_ES_RJ_Sem" & semana & ".xlsx"
                    'ElseIf uf = "ES" Then
                        'nomeArquivo = "Tbl_DistrMCD_ES_" & mesExtenso & ".xlsx"
                    Else
                        nomeArquivo = "Tbl_DistrMCD_JC_Sem" & semana & ".xlsx"
                    End If
                End If
    
            Case "PR"
                ' PR: distingue RS vs OUTROS + redeLoja e numCaractLoja
                If numCaractLoja = 4 Then
                    ' Sem fallback: sempre repete o nome da linha anterior
                    nomeArquivo = nomeArquivoAnterior
                ElseIf uf = "RS" Then
                    If redeLoja = "M" Then
                        Select Case numCaractLoja
                            Case 3
                                nomeArquivo = "Tbl_DistrMCD_RS_Sem" & semana & ".xlsx"
                            Case 7
                                nomeArquivo = "Tbl_DistrFS_RS_Sem" & semana & ".xlsx"
                        End Select
                    Else
                        Select Case numCaractLoja
                            Case 7
                                nomeArquivo = "Tbl_DistrFS_RS_Sem" & semana & ".xlsx"
                            Case 3
                                nomeArquivo = "Tbl_DistrMCD_RS_Sem" & semana & ".xlsx"
                        End Select
                    End If
                Else
                    ' UF <> RS
                    If redeLoja = "M" Then
                        Select Case numCaractLoja
                            Case 3
                                nomeArquivo = "Tbl_DistrMCD_PR_Sem" & semana & ".xlsx"
                            Case 7
                                nomeArquivo = "Tbl_DistrFS_PR_Sem" & semana & ".xlsx"
                        End Select
                    Else
                        Select Case numCaractLoja
                            Case 7
                                nomeArquivo = "Tbl_DistrFS_PR_Sem" & semana & ".xlsx"
                            Case 3
                                nomeArquivo = "Tbl_DistrMCD_PR_Sem" & semana & ".xlsx"
                        End Select
                    End If
                End If
                    
            Case "GR"
                If isDireta Then
                    ' DIRETAS por TIPO OPERACAO (coluna D)
                    nomeArquivo = "Tbl_DistrFS_DIRETAS_" & mesExtenso & ".xlsx"
                Else
                    ' Regras existentes por UF
                    If (uf = "MS" Or uf = "RO" Or uf = "SM" Or uf = "TM") Then
                        nomeArquivo = "Tbl_DistrFS_DIRETAS_" & mesExtenso & ".xlsx"
                    ElseIf uf = "BA" Or uf = "BS" Or uf = "CS" Or uf = "MG" Or uf = "IG" Or uf = "MC" Then
                        nomeArquivo = "Tbl_DistrFS_BA_BS_MG_Sem" & semana & ".xlsx"
                    ElseIf uf = "DF" Or uf = "GO" Or uf = "MT" Or uf = "TO" Or uf = "MD" Or uf = "ML" Then
                        nomeArquivo = "Tbl_DistrFS_DF_GO_MT_TO_Sem" & semana & ".xlsx"
                    ElseIf uf = "RJ" Or uf = "NF" Or uf = "RL" Or uf = "SF" Or uf = "SR" Or uf = "JF" Then
                        nomeArquivo = "Tbl_DistrFS_RJ_Sem" & semana & ".xlsx"
                    ElseIf uf = "ES" Then
                        nomeArquivo = "Tbl_DistrFS_ES_" & mesExtenso & ".xlsx"
                    Else
                        nomeArquivo = "Tbl_DistrFS_GR_Sem" & semana & ".xlsx"
                    End If
                End If
    
        End Select
        
    End If

        ' Se não foi definido nome, pula a linha
        If nomeArquivo = "" Then GoTo ProximaLinha
        
        ' Resetar wbDestino para garantir que não carregue valor anterior
        Set wbDestino = Nothing

        If Not dictArquivos.Exists(nomeArquivo) Then
            On Error Resume Next
            Application.DisplayAlerts = False
            Set wbDestino = Workbooks.Open(caminhoDestino & nomeArquivo)
            Application.DisplayAlerts = True
            On Error GoTo 0
        
            If wbDestino Is Nothing Then
                Call CriarArquivoDoZero(caminhoDestino, nomeArquivo)
                Set wbDestino = Workbooks(nomeArquivo)
                DoEvents
            End If
        
            Set wsDestino = wbDestino.Sheets(1)
            dictArquivos.Add nomeArquivo, wbDestino
        Else
            Set wbDestino = dictArquivos(nomeArquivo)
            Set wsDestino = wbDestino.Sheets(1)
        End If
        
        
        ' Garante setinhas ativas e sem filtros no destino
        EnsureHeaderFilters wsDestino, "A1:AI1"

        'Copiando informações:
        ultimaLinhaDestino = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row + 1
        
        'Verificando duplicidade em EMBARQUE3 (Não sobe para a Tbl_Prog)
        wbOrigem.Activate
        
        Dim embarque3 As String
        
        embarque3 = Cells(Linha, "P").Value & Cells(Linha, "I").Value
        
        Dim pos As Variant
        Dim valorBusca As String
        
        valorBusca = Trim$(CStr(embarque3))
        
        If Len(valorBusca) > 0 Then
            
            ' 0 = correspondência exata
            pos = Application.Match(valorBusca, wsDestino.Columns("P"), 0)
        
            If Not IsError(pos) Then
                GoTo ProximaLinha
            End If
        
        End If
                
        ' 1. TIPO OPERACAO (Coluna D origem ? Coluna A destino)
        wsDestino.Cells(ultimaLinhaDestino, "A").Value = wsOrigem.Cells(Linha, "D").Value
        
        ' 2. CD FATURAMENTO (Coluna E origem ? Coluna B destino)
        wsDestino.Cells(ultimaLinhaDestino, "B").Value = wsOrigem.Cells(Linha, "E").Value
        
        ' 3. ORIGEM (Coluna F origem ? Coluna C destino)
        wsDestino.Cells(ultimaLinhaDestino, "C").Value = wsOrigem.Cells(Linha, "F").Value
        
        ' 4. DATA FATURAMENTO (Coluna G origem ? Coluna D destino)
        wsDestino.Cells(ultimaLinhaDestino, "D").Value = Format(wsOrigem.Cells(Linha, "G").Value, "mm/dd/yy")

        ' 5. ROTA (Coluna H origem ? Coluna E destino)
        wsDestino.Cells(ultimaLinhaDestino, "E").Value = wsOrigem.Cells(Linha, "H").Value
        
        ' 6. LOJA (Coluna I origem ? Coluna F destino)
        wsDestino.Cells(ultimaLinhaDestino, "F").Value = wsOrigem.Cells(Linha, "I").Value
        
        ' 7. DATA CARREG (Coluna J origem ? Coluna G destino)
        wsDestino.Cells(ultimaLinhaDestino, "G").Value = Format(wsOrigem.Cells(Linha, "J").Value, "mm/dd/yy")

        ' 8. HORA CARREG (Coluna K origem ? Coluna H destino)
        wsDestino.Cells(ultimaLinhaDestino, "H").Value = wsOrigem.Cells(Linha, "K").Value
        
        ' 9. Fórmula na coluna I: =Glinha + Hlinha
        wsDestino.Cells(ultimaLinhaDestino, "I").Formula = "=G" & ultimaLinhaDestino & "+H" & ultimaLinhaDestino
        
        ' 10. DATA ENTREGA LOJA (Coluna M origem ? Coluna J destino)
        wsDestino.Cells(ultimaLinhaDestino, "J").Value = Format(wsOrigem.Cells(Linha, "M").Value, "mm/dd/yy")

        ' 11. HORA ENTREGA LOJA (Coluna N origem ? Coluna K destino)
        wsDestino.Cells(ultimaLinhaDestino, "K").Value = wsOrigem.Cells(Linha, "N").Value
        
        ' 12. Fórmula na coluna M: =Jlinha + Klinha
        wsDestino.Cells(ultimaLinhaDestino, "M").Formula = "=J" & ultimaLinhaDestino & "+K" & ultimaLinhaDestino
        
        ' 13. Fórmula na coluna N: =Mlinha + Llinha
        wsDestino.Cells(ultimaLinhaDestino, "N").Formula = "=M" & ultimaLinhaDestino & "+L" & ultimaLinhaDestino
        
        ' 14. Fórmula na coluna O: =Blinha & Elinha & TEXTO(Dlinha;"dd/mm/aaaa")
        wsDestino.Cells(ultimaLinhaDestino, "O").Formula = "=B" & ultimaLinhaDestino & "&E" & ultimaLinhaDestino & "&TEXT(D" & ultimaLinhaDestino & ",""dd/mm/aaaa"")"
        
        ' 15. Fórmula na coluna P: =Olinha & Flinha
        wsDestino.Cells(ultimaLinhaDestino, "P").Formula = "=O" & ultimaLinhaDestino & "&F" & ultimaLinhaDestino
        
        ' 16. CAIXAS (Coluna Q origem ? Coluna AB destino)
        wsDestino.Cells(ultimaLinhaDestino, "AB").Value = wsOrigem.Cells(Linha, "Q").Value
        
        ' 17. Preencher colunas R, S, T, U, AA, AC, AE com "-"
        Dim col As Variant
        For Each col In Array("Q", "R", "S", "T", "U", "AA", "AC", "AE")
            wsDestino.Cells(ultimaLinhaDestino, col).Value = "-"
        Next col
        
        ' 18. GRUPO ESTRELA (Coluna Z destino)
        wsDestino.Cells(ultimaLinhaDestino, "Z").Value = "NÃO"
        
        ' 19. Fórmula na coluna AG: =Ilinha + AFlinha
        wsDestino.Cells(ultimaLinhaDestino, "AG").Formula = "=I" & ultimaLinhaDestino & "+AF" & ultimaLinhaDestino
        
        ' 20. LOJA IMPACTADA ATRASO (Coluna AH)
        wsDestino.Cells(ultimaLinhaDestino, "AH").Value = "NÃO"
        
        ' 21. QTD LOJA IMPACTADA ATRASO (Coluna AI)
        wsDestino.Cells(ultimaLinhaDestino, "AI").Value = 0
        
ProximaLinha:
    ' >>> ATUALIZA O "ANTERIOR" PARA A PRÓXIMA ITERAÇÃO <<<
    If Len(nomeArquivo) > 0 Then
        nomeArquivoAnterior = nomeArquivo
    End If
    
Next Linha
    
    If wsDestino Is Nothing Then
    
    Exit Sub
    
    Else
    
    End If
    
    ' Reaplicar os filtros na linha 1 (sem filtrar nada)
    EnsureHeaderFilters wsDestino, "A1:AI1"

    ' Salva e fecha todos os arquivos de destino
    Dim chaveArquivo As Variant

    For Each chaveArquivo In dictArquivos.Keys
        Set wbDestino = dictArquivos(chaveArquivo)
        wbDestino.Save
        DoEvents
    Next chaveArquivo
    
    ' ===== PONTO FOCAL =====
    Dim wbPontoFocalEngine As Workbook
    
    Application.ScreenUpdating = False
    
    ' Abre o engine (UMA ÚNICA VEZ)
    
    'PRODUÇÃO
        
        'On Error Resume Next
        
        'If wbPontoFocalEngine Is Nothing Then
            'Err.Clear
            
            'Set wbPontoFocalEngine = Workbooks.Open( _
                PortfolioPath("criacao_prog_v5_6_testes.xlsm"), _
                UpdateLinks:=False)
        'End If
        
        'On Error GoTo 0
        
        'If wbPontoFocalEngine Is Nothing Then
            'MsgBox "Não foi possível localizar o arquivo Macro.xlsm nem o Beta.xlsm.", vbCritical
            'Exit Sub
        'End If
    
    'TESTE
    
    Dim eventosAntesDoEngine As Boolean
    eventosAntesDoEngine = Application.EnableEvents
    Application.EnableEvents = False
    Set wbPontoFocalEngine = Workbooks.Open( _
        PortfolioPath("criacao_prog_v5_6_testes.xlsm"), _
        UpdateLinks:=False _
    )
    Application.EnableEvents = eventosAntesDoEngine
    
'Aqui vai esse código que vamos criar juntos
Dim wbAtual As Workbook

For Each wbAtual In Application.Workbooks
    
    ' Ignora o próprio engine e o workbook da macro
    If wbAtual.Name <> wbPontoFocalEngine.Name _
       And wbAtual.Name <> ThisWorkbook.Name Then
        
        ' Apenas arquivos .xlsx (sem macro)
        If LCase(Right(wbAtual.Name, 5)) = ".xlsx" Then
            
            ' Confere padrão mínimo do layout
            On Error Resume Next
            If UCase(Trim(wbAtual.Sheets(1).Range("A1").Value)) = "TIPO OPERACAO" Then
                On Error GoTo 0
                
                ' ? MESMA CHAMADA DO USERFORM, AGORA EM LOTE
                Application.Run _
                    "'" & wbPontoFocalEngine.Name & "'!CriarTblPontoFocalPorData", _
                    wbAtual
            Else
                On Error GoTo 0
            End If
            
        End If
    End If
    
Next wbAtual

' Fecha o engine
wbPontoFocalEngine.Close SaveChanges:=False

CleanExit:
    Application.StatusBar = False
    Application.Visible = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Sub CriarArquivoDoZero(caminhoDestino As String, nomeArquivo As String)

    Dim wbNovo As Workbook
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim i As Long

    ' 1. Criar novo arquivo
    Set wbNovo = Workbooks.Add(xlWBATWorksheet)
    DoEvents
    Set ws1 = wbNovo.Sheets(1)
    ws1.Name = "Sheet1"

    ' 2. Ocultar linhas de grade e congelar linha superior
    ActiveWindow.DisplayGridlines = False
    
    wbNovo.Sheets(1).Activate
    wbNovo.Sheets(1).Range("A2").Select
    ActiveWindow.FreezePanes = True

    ' 3. Largura das colunas A:AI
    Dim larguras As Variant
    larguras = Array(14.55, 15.73, 9.45, 17.18, 7.82, 7.45, 18.27, 18.64, 15.36, 16.82, 17.18, 20.36, 14.18, 19.73, 12.91, 18.18, 23.45, 18.73, 12.09, 11.36, 14.36, 14.27, 14.18, 15.36, 15.27, 14.45, 10.27, 9.09, 16.91, 21.27, 12.73, 25.82, 21.73, 22.18, 25.36)
    For i = 1 To 35
        ws1.Columns(i).ColumnWidth = larguras(i - 1)
    Next i

    ' 4. Cabeçalhos e Linhas até 2000
    Dim headers As Variant
    headers = Array("TIPO OPERACAO", "CD FATURAMENTO", "ORIGEM", "DATA FATURAMENTO", "ROTA", "LOJA", "DATA CARREGAMENTO", "HORA CARREGAMENTO", "DTH CARRG PROG", "DATA ENTREGA LOJA", "HORA ENTREGA LOJA", "HORA ADICIONAL ATRASO", "DTH ENTR PROG", "DTH ENTR PROG_ATRASO", "EMBARQUE1", "EMBARQUE3", "OCORRENCIA CARREGAMENTO", "OCORRENCIA ENTREGA", "OBS", "PLACA CAV", "PLACA CARRETA", "DATA CH CD aux", "DATA SD CD aux", "DATA CH LOJA aux", "DATA SD LOJA aux", "GRUPO ESTRELA", "ROTA_AM", "CAIXAS", "EMBARQUE DIRETAS", "DATA INICIO DESC LOJA aux", "OBS INI DESC.", "HORA ADICIONAL ATRASO CARREG", "DTH CARREG PROG_ATRASO", "LOJA IMPACTADA ATRASO", "QTD LOJA IMPACTADA ATRASO")
    
    For i = 0 To UBound(headers)
        With ws1.Cells(1, i + 1)
            .Value = headers(i)
            .Font.Name = "Aptos Narrow"
            .Font.Size = 8
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next i
    
    'Formula na coluna AG
    Range("AG2").Formula2 = "=I2+AF2"
    'Range("AG2").Copy
    'Range("AG3:AG2000").PasteSpecial xlPasteFormulas
    
    'Coluna AH
    'Range("AH2:AH2000").Value = "NÃO"
    
    'Coluna AI
    'Range("AI2:AI2000").Value = 0
    
    'Formatação das linhas
    With ws1.Range("A2:AI2000")
        .Font.Name = "Aptos Narrow"
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    'Formato hh:mm:ss nas colunas H, K e L
    With ws1
        .Range("H2:H2000").NumberFormat = "hh:mm:ss"
        .Range("K2:K2000").NumberFormat = "hh:mm:ss"
        .Range("L2:L2000").NumberFormat = "hh:mm:ss"
        .Range("AF2:AF2000").NumberFormat = "hh:mm:ss"
    End With
    
    'Formato dd/mm/aaaa hh:mm nas colunas I, M, N, V, W, X, Y, AD
    With ws1
        .Range("I2:I2000").NumberFormat = "dd/mm/yyyy hh:mm"
        .Range("M2:M2000").NumberFormat = "dd/mm/yyyy hh:mm"
        .Range("N2:N2000").NumberFormat = "dd/mm/yyyy hh:mm"
        .Range("V2:V2000").NumberFormat = "dd/mm/yyyy hh:mm"
        .Range("W2:W2000").NumberFormat = "dd/mm/yyyy hh:mm"
        .Range("X2:X2000").NumberFormat = "dd/mm/yyyy hh:mm"
        .Range("Y2:Y2000").NumberFormat = "dd/mm/yyyy hh:mm"
        .Range("AD2:AD2000").NumberFormat = "dd/mm/yyyy hh:mm"
        .Range("AG2:AG2000").NumberFormat = "dd/mm/yyyy hh:mm"
    End With

    ' 5. Cores de preenchimento
    Dim azul As Long: azul = RGB(11, 29, 47)
    Dim marrom As Long: marrom = RGB(126, 53, 14)
    Dim verde As Long: verde = RGB(18, 80, 26)

    For i = 1 To 35
        Select Case i
            Case 1 To 8, 10, 11, 17 To 21, 26 To 28
                ws1.Cells(1, i).Interior.Color = azul
            Case 9, 12 To 16, 22 To 25, 29, 32 To 35
                ws1.Cells(1, i).Interior.Color = marrom
            Case 30, 31
                ws1.Cells(1, i).Interior.Color = verde
        End Select
    Next i

    ' 6. Filtro
    ws1.Range("A1:AI1").AutoFilter

    ' 7. Criar Planilha2
    Set ws2 = wbNovo.Sheets.Add(After:=ws1)
    ws2.Name = "Planilha2"

    ' Coluna A
    Dim listaA As Variant
    listaA = Array("Justificativa Inicio Descarga Loja", "-", "AG RETORNO MOTORISTA", "AG TEMPERATURA - ABAIXO DA CONDICIONAL", "DESC INICIADA - AG ATUALIZAÇÃO RCS", "MOTORISTA CHEGOU ADIANTADO AO LOCAL", "MOTORISTA CONFIRMOU Ñ INICIO DE DESC, AG MAIS 10 MINUTOS", "MOTORISTA SEM MOBILE", "SEM RETORNO MOTORISTA", "SEM RETORNO MOTORISTA - SOLICITADO APOIO TD", "SOLICITAÇÃO DE RETIRADA SEM SUCESSO", "SOLICITADO RETIRADA DO VEICULO")
    For i = 0 To UBound(listaA)
        ws2.Cells(i + 1, 1).Value = listaA(i)
    Next i
    
    Call ConfigurarValidacoesWBNovo(wbNovo)

    ' Coluna D
    'Dim listaD As Variant
    'listaD = Array( _
        "JUSTIFICATIVAS - RCS ATRASO", "-", _
        "Atraso Ajudante", "Atraso Bimbo", "Atraso Cia Aérea", "Atraso do Hortifrúti", _
        "Atraso do Transportador Terceiro", "Atraso Motorista", "Atraso na emissão de NF", _
        "Atraso na liberação das NF's", "Atraso na saída da balsa", "Atraso no Carregamento BACKHAUL", _
        "Atraso no Carregamento CROSS", "Atraso no Carregamento WH", "Atraso Panifresh", _
        "Condições climáticas na região", "Encaixe Urgente", "Entrega simultânea", _
        "Erro de Planejamento Transp.", "Erro do Suporte Operacional", "Erro na janela", _
        "Erro na Preliminar", "Erro no planejamento (OffLoand)", "Erro Transportes", _
        "Falta de Carga de Bateria Celular", "Falta de Energia", "Falta do Motorista Titular", _
        "Fim da Jornada do motorista", "Horário de Almoço", "Indisponibilidade de Veículo", _
        "Inversão de Entrega Autorizada", "Inversão Não Autorizada", "Loja fechada", "Loja sem senha", _
        "Manifestações / Eventos / Acidente na Rodovia", "Manutenção de Carreta", "Manutenção de Cavalo", _
        "Manutenção no armazém", "Manutenção Predial", "Manutenção TK", "Morosidade Loja", _
        "Ordem de Carregamento Loja", "Paletização invertida", "Posto Fiscal", "Problema Mecânico", _
        "Problemas de saúde do motorista", "Restrição de Acesso Loja", "Retenção de Balança", _
        "Sinistro", "Sistema / Mobile / BI", "Transbordo de carga", "Transferência/Transbordo", _
        "Trânsito", "Trânsito na Transferência" _
    )
   
'For i = 0 To UBound(listaD)
 '   ws2.Cells(i + 1, 4).Value = listaD(i)
'Next i

    ' Coluna F
    ws2.Range("F1").Value = "GRUPO ESTRELA"
    ws2.Range("F2").Value = "NÃO"
    ws2.Range("F3").Value = "SIM"
    
    With wbNovo.Sheets("Planilha2")
        ' Ocultar linhas de grade
        .Activate
        ActiveWindow.DisplayGridlines = False
    
        ' Formatação geral no intervalo A1:G2000
        With .Range("A1:G2000")
            .Font.Name = "Aptos Narrow"
            .Font.Size = 8
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    
        ' Cabeçalhos em A1, D1 e F1
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(255, 255, 255)
        .Range("A1").Interior.Color = RGB(&HB, &H30, &H40)
    
        .Range("D1").Font.Bold = True
        .Range("D1").Font.Color = RGB(255, 255, 255)
        .Range("D1").Interior.Color = RGB(&HB, &H30, &H40)
    
        .Range("F1").Font.Bold = True
        .Range("F1").Font.Color = RGB(255, 255, 255)
        .Range("F1").Interior.Color = RGB(&HB, &H30, &H40)
    
        ' Largura das colunas
        .Columns("A").ColumnWidth = 39.82
        .Columns("D").ColumnWidth = 27.64
        .Columns("F").ColumnWidth = 9.91
    End With

    ' Ocultar Planilha2
    ws2.Visible = xlSheetHidden

    ' 8. Validação de data nas colunas I, M, N, V, W, X, Y, AD, AG
    Dim colunasData As Variant
    Dim col As Variant
    
    colunasData = Array("I", "M", "N", "V", "W", "X", "Y", "AD", "AG")
    For Each col In colunasData
        With ws1.Range(col & "2:" & col & "2000").Validation
            .Delete
            .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            formula1:=DateSerial(1900, 1, 1), Formula2:=DateSerial(9999, 12, 31)
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = "Data inválida"
            .InputMessage = ""
            .ErrorMessage = "Insira uma data válida no formato dd/mm/aaaa hh:mm"
            .ShowInput = True
            .ShowError = True
        End With
    Next col

    ' 10. Validação de lista na coluna R
    'With ws1.Range("R2:R2000").Validation
     '   .Delete
      '  .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Planilha2!$D$2:$D$54"
       ' .IgnoreBlank = True
        '.InCellDropdown = True
    'End With

    ' 11. Validação de lista na coluna Z
    With ws1.Range("Z2:Z2000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, formula1:="=Planilha2!$F$2:$F$3"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ' 12. Validação de lista na coluna AE
    With ws1.Range("AE2:AE2000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, formula1:="=Planilha2!$A$2:$A$12"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    ' 13. Salvando o arquivo
    Application.DisplayAlerts = False
    wbNovo.SaveAs Filename:=caminhoDestino & nomeArquivo, FileFormat:=xlOpenXMLWorkbook
    DoEvents
    
    CompatibilizarParaSharedWorkbook wbNovo

    ' 14. Tornando o arquivo compartilhado
    Application.DisplayAlerts = False
    wbNovo.Save ' Demonstracao: sem compartilhamento legado
    Application.DisplayAlerts = True

End Sub

Sub TemplateInfolog()
    Dim wbOrigem As Workbook
    Dim wbDestino As Workbook
    Dim wsOrigem As Worksheet
    Dim wsPlanilha1 As Worksheet
    Dim wsFormulas As Worksheet
    Dim caminhoDestino As String
    Dim UltimaLinha As Long
    Dim i As Long
    Dim celulaValor As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbOrigem = ThisWorkbook
    Set wsOrigem = wbOrigem.Sheets(1)

    ' Caminho e abertura do arquivo de destino
    'caminhoDestino = PortfolioPath("legacy\Template Transportes Infolog 2.0.xlsx")
    caminhoDestino = PortfolioPath("legacy\Template Transportes Infolog 2.0.xlsx")
    Set wbDestino = Workbooks.Open(caminhoDestino, UpdateLinks:=0)
    DoEvents

    ' Definir planilhas principais
    Set wsPlanilha1 = wbDestino.Sheets("Planilha1")
    Set wsFormulas = wbDestino.Sheets("Formulas")

    ' Ocultar todas as outras planilhas
    Dim ws As Worksheet
    For Each ws In wbDestino.Sheets
        If ws.Name <> "Planilha1" And ws.Name <> "Formulas" Then
            ws.Visible = xlSheetHidden
        End If
    Next ws

    ' Limpar conteúdo das planilhas (exceto cabeçalho)
    wsPlanilha1.Range("2:" & wsPlanilha1.Rows.Count).ClearContents
    wsFormulas.Range("2:" & wsFormulas.Rows.Count).ClearContents

    ' Formatação geral
    With wsPlanilha1.Range("A2:Z" & wsPlanilha1.Rows.Count)
        .Font.Name = "Calibri"
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With wsFormulas.Range("A2:Z" & wsFormulas.Rows.Count)
        .Font.Name = "Calibri"
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Descobrir última linha com dados na coluna I de wbOrigem
    UltimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "I").End(xlUp).Row
    
    'Verificando se origem é CDRJ e não inserindo no Template, se não for copia as colunas especificas
    Dim linhaDestino As Long
    linhaDestino = 2 ' Começa na linha 2 da wsFormulas
    
    For i = 2 To UltimaLinha
        'If wsOrigem.Cells(i, "F").Value <> "CDRJ" Then
            wsFormulas.Cells(linhaDestino, "R").Value = wsOrigem.Cells(i, "I").Value
            wsFormulas.Cells(linhaDestino, "Q").Value = wsOrigem.Cells(i, "F").Value
            wsFormulas.Cells(linhaDestino, "P").Value = wsOrigem.Cells(i, "W").Value
            wsFormulas.Cells(linhaDestino, "N").Value = wsOrigem.Cells(i, "O").Value
            wsFormulas.Cells(linhaDestino, "L").Value = wsOrigem.Cells(i, "L").Value
            wsFormulas.Cells(linhaDestino, "D").Value = wsOrigem.Cells(i, "P").Value
    
            wsFormulas.Cells(linhaDestino, "J").Value = 1260
            wsFormulas.Cells(linhaDestino, "I").Value = 1257
            wsFormulas.Cells(linhaDestino, "H").Value = 1237
            wsFormulas.Cells(linhaDestino, "B").Value = "RASTREADO"
    
            'celulaValor = Trim(wsOrigem.Cells(i, "H").Value)
            'If Len(celulaValor) > 0 Then
                'Select Case Right(celulaValor, 1)
                    'Case "F"
                        'wsFormulas.Cells(linhaDestino, "C").Value = "DISTRIBUICAO FS"
                    'Case "M"
                        'wsFormulas.Cells(linhaDestino, "C").Value = "DISTRIBUICAO MCD"
                    'Case "X"
                        'wsFormulas.Cells(linhaDestino, "C").Value = "DISTRIBUICAO MCD"
                    'Case Else
                        'wsFormulas.Cells(linhaDestino, "C").Value = ""
                'End Select
            'End If
            
            ' --- Início do trecho ajustado (agrupado por embarque - col P) ---
            Static ultimoEmbarque As String
            'Dim celulaValor As String
            Dim tam As Long
            Dim embarqueAtual As String
            
            ' embarque da linha atual (coluna P da wsOrigem)
            embarqueAtual = Trim$(CStr(wsOrigem.Cells(i, "P").Text))
            
            If linhaDestino > 2 And embarqueAtual = ultimoEmbarque Then
                ' Mesmo embarque do registro anterior: replica valor da coluna C da linha anterior em wsFormulas
                wsFormulas.Cells(linhaDestino, "C").Value = wsFormulas.Cells(linhaDestino - 1, "C").Value
            Else
                ' Novo embarque (ou primeira linhaDestino=2): aplica a regra pelos caracteres da coluna I
                celulaValor = Trim$(CStr(wsOrigem.Cells(i, "I").Text))
                tam = Len(celulaValor)
            
                If tam > 0 Then
                    Select Case tam
                        Case 7
                            wsFormulas.Cells(linhaDestino, "C").Value = "DISTRIBUICAO FS"
                        Case 3
                            wsFormulas.Cells(linhaDestino, "C").Value = "DISTRIBUICAO MCD"
                        Case 4
                            ' herda o valor anterior em C (se for a primeira linhaDestino, define um fallback)
                            If linhaDestino > 1 Then
                                wsFormulas.Cells(linhaDestino, "C").Value = wsFormulas.Cells(linhaDestino - 1, "C").Value
                            Else
                                wsFormulas.Cells(linhaDestino, "C").Value = "DISTRIBUICAO MCD" ' fallback opcional
                            End If
                        Case Else
                            wsFormulas.Cells(linhaDestino, "C").Value = ""
                    End Select
                Else
                    wsFormulas.Cells(linhaDestino, "C").Value = ""
                End If
            End If
            
            ' Atualiza o controle do último embarque processado
            ultimoEmbarque = embarqueAtual
            ' --- Fim do trecho ajustado ---
    
            linhaDestino = linhaDestino + 1 ' Só incrementa se a linha for copiada
        'End If
    Next i

    
    ' Reexibir planilhas auxiliares temporariamente
    wbDestino.Sheets("Dados_Transportador").Visible = xlSheetVisible
    wbDestino.Sheets("Dados_Lojas").Visible = xlSheetVisible
    
    ' Inserir fórmulas nas colunas A, K e M
    With wsFormulas
        'Dim ultimaLinhaFormulas As Long
        ultimaLinhaFormulas = .Cells(.Rows.Count, "R").End(xlUp).Row
    
        .Range("A2:A" & ultimaLinhaFormulas).FormulaLocal = "=PROCX(P2;Dados_Transportador!A:A;Dados_Transportador!B:B;""-"")"
        .Range("K2:K" & ultimaLinhaFormulas).FormulaLocal = "=PROCX(Q2;Dados_Lojas!A:A;Dados_Lojas!B:B;""-"")"
        .Range("M2:M" & ultimaLinhaFormulas).FormulaLocal = "=PROCX(R2;Dados_Lojas!A:A;Dados_Lojas!B:B;""-"")"
    End With
    
    ' Ocultar novamente as planilhas auxiliares
    wbDestino.Sheets("Dados_Transportador").Visible = xlSheetHidden
    wbDestino.Sheets("Dados_Lojas").Visible = xlSheetHidden

    'Analisando se já tem retorno para o CD (vindo da Tbl_Prog)
    wsFormulas.Activate
    
    verificaRetorno = 2
    
    Do While verificaRetorno <= ultimaLinhaFormulas + 1
    
        'Concatenando embarque com a loja
        embarqueLoja = Cells(verificaRetorno, "D").Value & Cells(verificaRetorno, "R").Value
    
        If Len(embarqueLoja) = 21 Or Len(embarqueLoja) = 25 Then
    
            Call RetornoParaOCD(wsFormulas)
        
        End If
        
        verificaRetorno = verificaRetorno + 1
    
    Loop
    
    wsFormulas.Activate
    Range("A1").Activate
    
    Dim Celula As Range
    Dim erroEncontrado As Boolean
    
    ' Verificar se existe "-" na coluna M
    ultimaLinhaFormulas = wsFormulas.Cells(wsFormulas.Rows.Count, "M").End(xlUp).Row
    erroEncontrado = False
    
    For Each Celula In wsFormulas.Range("M2:M" & ultimaLinhaFormulas)
        If Trim(Celula.Value) = "-" And Len(Celula.Value) = 1 Then
            erroEncontrado = True
            Exit For
        End If
    Next Celula
    
    If erroEncontrado Then
        MsgBox "Por favor verificar as lojas não cadastradas - cópia dos valores para a Planilha1 não realizada", vbCritical, "Erro de validação"
        Sheets("Dados_Lojas").Visible = True
        ' Filtrar na wsFormulas as linhas com "-" (exatamente) na coluna M
        wsFormulas.Range("A1:R" & ultimaLinhaFormulas).AutoFilter Field:=13, Criteria1:="="
        wsFormulas.Range("A1:R" & ultimaLinhaFormulas).AutoFilter Field:=13, Criteria1:="-"
    Else
        ' Copiar A:R de wsFormulas e colar como valores em wsPlanilha1
        wsFormulas.Range("A2:R" & ultimaLinhaFormulas).Copy
        wsPlanilha1.Range("A2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If
        
    wsPlanilha1.Activate
    Range("A1").Activate
    
    wsFormulas.Activate
    Range("A1").Activate

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    'MsgBox "Dados copiados para o template com sucesso!", vbInformation, "Infolog"
End Sub

Sub RetornoParaOCD(wsFormulas As Worksheet)
    
    Dim i As Long
    'Dim ultimaLinhaFormulas As Long
    Dim valorDAtual As Variant, valorDAnterior As Variant
    Dim valorPAnterior As Variant
    Dim valorRAnterior As Variant

    'ultimaLinhaFormulas = wsFormulas.Cells(wsFormulas.Rows.Count, "D").End(xlUp).Row + 1
    'i = 3 ' Começa da linha 3
    i = verificaRetorno

    'Do While i <= ultimaLinhaFormulas
        valorDAtual = wsFormulas.Cells(i, "D").Value
        valorDAnterior = wsFormulas.Cells(i - 1, "D").Value
        valorPAnterior = wsFormulas.Cells(i - 1, "P").Value
        valorRAnterior = wsFormulas.Cells(i - 1, "R").Value

        If valorDAtual <> valorDAnterior And UCase(valorPAnterior) = "MARTIN BROWER" And Len(valorRAnterior) <> 4 Then
            ' Inserir nova linha
            wsFormulas.Rows(i).Insert Shift:=xlDown

            ' Copiar A:Q da linha anterior (valores e fórmulas)
            wsFormulas.Range("A" & i - 1 & ":Q" & i - 1).Copy
            wsFormulas.Range("A" & i & ":Q" & i).PasteSpecial Paste:=xlPasteFormulasAndNumberFormats

            ' Colocar valor da coluna Q na coluna R da nova linha
            wsFormulas.Cells(i, "R").Value = wsFormulas.Cells(i - 1, "Q").Value

            ' Atualizar última linha (já que inserimos uma nova)
            ultimaLinhaFormulas = ultimaLinhaFormulas + 1
            
            i = i + 1

        End If

        i = i + 1
        
    'Loop
    
    If i = ultimaLinhaFormulas Then
    
        ' Verifica se a última linha precisa de retorno
        valorDAtual = wsFormulas.Cells(ultimaLinhaFormulas, "D").Value
        valorDAnterior = wsFormulas.Cells(ultimaLinhaFormulas - 1, "D").Value
        valorPAnterior = wsFormulas.Cells(ultimaLinhaFormulas - 1, "P").Value
        
        If valorDAtual <> valorDAnterior And UCase(valorPAnterior) = "MARTIN BROWER" Then
            wsFormulas.Rows(ultimaLinhaFormulas + 1).Insert Shift:=xlDown
            wsFormulas.Range("A" & ultimaLinhaFormulas & ":Q" & ultimaLinhaFormulas).Copy
            wsFormulas.Range("A" & ultimaLinhaFormulas + 1 & ":Q" & ultimaLinhaFormulas + 1).PasteSpecial Paste:=xlPasteFormulasAndNumberFormats
            wsFormulas.Cells(ultimaLinhaFormulas + 1, "R").Value = wsFormulas.Cells(ultimaLinhaFormulas, "Q").Value
        End If
    
    Else
    
    End If

Application.CutCopyMode = False

End Sub

' Chame este procedimento passando o Workbook recém-criado (wbNovo)
Sub ConfigurarValidacoesWBNovo(ByVal wbNovo As Workbook)
    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = wbNovo.Worksheets("Sheet1")
    Set ws2 = wbNovo.Worksheets("Planilha2")

    ' ========= 1) (Opcional) Preencher a lista original em D com sua array =========
    ' Se você já preenche em outro trecho, pode ignorar esta parte.
    Dim listaD As Variant
    listaD = Array( _
        "JUSTIFICATIVAS - RCS ATRASO", "-", _
        "Analisando com o Transportador", "Analisando com o Transportes", "Atraso Ajudante", "Atraso Bimbo", "Atraso Cia Aérea", "Atraso do Hortifrúti", _
        "Atraso do Transportador Terceiro", "Atraso de Transportes", "Atraso Motorista", "Atraso na emissão de NF", _
        "Atraso na liberação das NF's", "Atraso na saída da balsa", "Atraso no Carregamento BACKHAUL", _
        "Atraso no Carregamento CROSS", "Atraso no Carregamento WH", "Atraso Panifresh", _
        "Condições climáticas na região", "Encaixe Urgente", "Entrega simultânea", _
        "Erro de Planejamento Transp.", "Erro do Suporte Operacional", "Erro na janela", _
        "Erro na Preliminar", "Erro no planejamento (OffLoand)", "Erro Transportes", _
        "Falta de Carga de Bateria Celular", "Falta de Energia", "Falta do Motorista Titular", _
        "Fim da Jornada do motorista", "Horário de Almoço", "Indisponibilidade de Veículo", _
        "Inversão de Entrega Autorizada", "Inversão Não Autorizada", "Loja fechada", "Loja sem senha", _
        "Manifestações / Eventos / Acidente na Rodovia", "Manutenção de Carreta", "Manutenção de Cavalo", _
        "Manutenção no armazém", "Manutenção Predial", "Manutenção TK", "Morosidade Loja", _
        "Ordem de Carregamento Loja", "Paletização invertida", "Posto Fiscal", "Problema Mecânico", _
        "Problemas de saúde do motorista", "Restrição de Acesso Loja", "Retenção de Balança", _
        "Sinistro", "Sistema / Mobile / BI", "Transbordo de carga", "Transferência/Transbordo", _
        "Trânsito", "Trânsito na Transferência" _
    )
    Dim i As Long
    For i = 0 To UBound(listaD)
        ws2.Cells(i + 1, 4).Value = listaD(i)  ' D1 = título; itens começam em D2
    Next i

    ' ========= 2) Tornar a coluna D uma Tabela (expande sozinha quando você adicionar itens) =========
    Dim rngD As Range, loD As ListObject
    Set rngD = ws2.Range("D1", ws2.Cells(ws2.Rows.Count, "D").End(xlUp))
    On Error Resume Next
    Set loD = ws2.ListObjects("tblMotivos")
    On Error GoTo 0
    If loD Is Nothing Then
        Set loD = ws2.ListObjects.Add(xlSrcRange, rngD, , xlYes)
        loD.Name = "tblMotivos"
        loD.ListColumns(1).Name = "Motivo"
    Else
        loD.Resize rngD
    End If

    ' ========= 3) Construir a lista sem "-" em H:H =========
    ' Tenta com Excel 365 (FILTRAR). Se não existir, cai no plano B (preenche H via código).
    Dim temFiltro As Boolean
    temFiltro = False
    On Error Resume Next
    ws2.Range("H1").Value = "Motivo (sem '-')"
    ws2.Range("H2").Formula2 = "=FILTER(tblMotivos[Motivo], tblMotivos[Motivo]<>""-"")"
    If Err.Number = 0 Then temFiltro = True
    Err.Clear
    On Error GoTo 0

    If temFiltro Then
        ' Usaremos o intervalo "H2#" (spill) como a lista sem "-".
        ' Define nomes (Named Ranges) para usar na Validação de Dados
        On Error Resume Next
        wbNovo.Names("Lista_R_Com_Hifen").Delete
        wbNovo.Names("Lista_R_Sem_Hifen").Delete
        On Error GoTo 0

        wbNovo.Names.Add Name:="Lista_R_Com_Hifen", RefersTo:="=Planilha2!tblMotivos[Motivo]"
        wbNovo.Names.Add Name:="Lista_R_Sem_Hifen", RefersTo:="=Planilha2!$H$2#"

    Else
        ' -------- Plano B (sem FILTRAR): copia D->H ignorando "-" agora --------
        Dim lastD As Long, r As Long, h As Long
        lastD = ws2.Cells(ws2.Rows.Count, "D").End(xlUp).Row
        h = 2
        For r = 2 To lastD
            If ws2.Cells(r, "D").Value <> "-" Then
                ws2.Cells(h, "H").Value = ws2.Cells(r, "D").Value
                h = h + 1
            End If
        Next r

        ' Define nomes dinâmicos até a última célula não vazia (evita pegar brancos)
        On Error Resume Next
        wbNovo.Names("Lista_R_Com_Hifen").Delete
        wbNovo.Names("Lista_R_Sem_Hifen").Delete
        On Error GoTo 0

        wbNovo.Names.Add Name:="Lista_R_Com_Hifen", _
            RefersTo:="=Planilha2!$D$2:INDEX(Planilha2!$D:$D, MATCH(""zzz"", Planilha2!$D:$D))"
        wbNovo.Names.Add Name:="Lista_R_Sem_Hifen", _
            RefersTo:="=Planilha2!$H$2:INDEX(Planilha2!$H:$H, MATCH(""zzz"", Planilha2!$H:$H))"
        ' Obs.: MATCH("zzz",Coluna) devolve a última célula com texto.
        ' Se algum item for numérico, troque por MATCH(1E+99,Coluna) (último número) ou use COUNTA com ajuste.
    End If

    ' ========= 4) Validação de Dados na coluna R (lista alterna conforme L) =========
    With ws1.Range("R2:R2000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
              formula1:="=INDIRECT(IF($L2="""",""Lista_R_Com_Hifen"",""Lista_R_Sem_Hifen""))"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "Ocorrência inválida"
        .ErrorMessage = "A entrega está com HORA ADICIONAL ATRASO, então precisa ter uma justificativa."
    End With

    ' ========= 5) Validação de Dados na coluna L (hora só quando R <> "-") =========
    With ws1.Range("L2:L2000").Validation
        .Delete
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
              formula1:="=OR($L2="""", AND($R2<>""-"", ISNUMBER($L2), $L2>=0, $L2<1))"
        .IgnoreBlank = True
        .ErrorTitle = "HORA ADIC SEM JUSTIFICATIVA"
        .ErrorMessage = "Não é possível inserir HORA ADICIONAL ATRASO sem uma justificativa na ocorrência de entrega."
    End With

    ' (Opcional) Formatação da coluna L
    ws1.Range("L2:L2000").NumberFormat = "hh:mm:ss"
    
    
    ' ========= Validação de Dados na coluna Q (lista alterna conforme AF) =========
    With ws1.Range("Q2:Q2000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
              formula1:="=INDIRECT(IF($AF2="""",""Lista_R_Com_Hifen"",""Lista_R_Sem_Hifen""))"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "Ocorrência inválida"
        .ErrorMessage = "O carregamento está com HORA ADICIONAL ATRASO, então precisa ter uma justificativa."
    End With
    
    ' ========= Validação de Dados na coluna AF (hora só quando Q <> "-") =========
    With ws1.Range("AF2:AF2000").Validation
        .Delete
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
              formula1:="=OR($AF2="""", AND($Q2<>""-"", ISNUMBER($AF2), $AF2>=0, $AF2<1))"
        .IgnoreBlank = True
        .ErrorTitle = "HORA ADIC SEM JUSTIFICATIVA"
        .ErrorMessage = "Não é possível inserir HORA ADICIONAL ATRASO sem uma justificativa na ocorrência de carregamento."
    End With
    
    ' (Opcional) Formatação da coluna AF como hora
    ws1.Range("AF2:AF2000").NumberFormat = "hh:mm:ss"

End Sub

Public Sub CompatibilizarParaSharedWorkbook(ByVal wbNovo As Workbook)
    On Error GoTo TrataErro

    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = wbNovo.Worksheets("Sheet1")
    Set ws2 = wbNovo.Worksheets("Planilha2")

    ' 1) Shared Workbook legado não suporta Tabelas
    Dim lo As ListObject
    For Each lo In ws2.ListObjects
        lo.Unlist
    Next lo

    ' 2) Recriar H sem "-" (sem fórmulas dinâmicas/derrame)
    ws2.Range("H2", ws2.Cells(ws2.Rows.Count, "H")).ClearContents

    Dim lastD As Long, lastH As Long, r As Long, h As Long, valD As String
    lastD = ws2.Cells(ws2.Rows.Count, "D").End(xlUp).Row

    h = 2
    For r = 2 To lastD
        valD = Trim$(CStr(ws2.Cells(r, "D").Value))
        If valD <> "" And valD <> "-" Then
            ws2.Cells(h, "H").Value = valD
            h = h + 1
        End If
    Next r
    lastH = ws2.Cells(ws2.Rows.Count, "H").End(xlUp).Row

    ' 2a) Evitar listas vazias (fonte nula causa 1004)
    If lastD < 2 Then ws2.Range("D2").Value = "-": lastD = 2
    If lastH < 2 Then ws2.Range("H2").Value = "(sem itens)": lastH = 2

    ' 3) Nomes estáticos que apontam para Planilha2
    On Error Resume Next
    wbNovo.Names("Lista_R_Com_Hifen").Delete
    wbNovo.Names("Lista_R_Sem_Hifen").Delete
    On Error GoTo 0

    wbNovo.Names.Add Name:="Lista_R_Com_Hifen", RefersTo:="=" & ws2.Name & "!$D$2:$D$" & lastD
    wbNovo.Names.Add Name:="Lista_R_Sem_Hifen", RefersTo:="=" & ws2.Name & "!$H$2:$H$" & lastH

    ' 4) Desproteger Sheet1 para aplicar validações
    If ws1.ProtectContents Then ws1.Unprotect

    ' 5) Validações LISTA (R/Q) com ESCOLHER/CHOOSE (sem INDIRETO)
    Dim fListaR_PTBR As String, fListaQ_PTBR As String
    Dim fListaR_EN As String, fListaQ_EN As String

    fListaR_PTBR = "=ESCOLHER(SE($L2="""";1;2);Lista_R_Com_Hifen;Lista_R_Sem_Hifen)"
    fListaR_EN = "=CHOOSE(IF($L2="""",1,2),Lista_R_Com_Hifen,Lista_R_Sem_Hifen)"

    fListaQ_PTBR = "=ESCOLHER(SE($AF2="""";1;2);Lista_R_Com_Hifen;Lista_R_Sem_Hifen)"
    fListaQ_EN = "=CHOOSE(IF($AF2="""",1,2),Lista_R_Com_Hifen,Lista_R_Sem_Hifen)"

    Call ApplyValidationListSmart(ws1, ws1.Range("R2:R2000"), fListaR_EN, fListaR_PTBR)
    Call ApplyValidationListSmart(ws1, ws1.Range("Q2:Q2000"), fListaQ_EN, fListaQ_PTBR)

    ' 6) Validações CUSTOM (hora + dependência), SEM ISNUMBER/ÉNÚM
    ' Coerção a número via IFERROR/SEERRO ? testa 0<=x<1
    Dim sep As String: sep = Application.International(xlListSeparator)

    ' EN primeiro (geralmente mais estável via VBA)
    Dim fL_en As String, fAF_en As String
    fL_en = "=OR($L2=""""" & sep & "AND($R2<>""-""" & sep & "IFERROR($L2*1" & sep & "-1)>=0" & sep & "IFERROR($L2*1" & sep & "-1)<1))"
    fAF_en = "=OR($AF2=""""" & sep & "AND($Q2<>""-""" & sep & "IFERROR($AF2*1" & sep & "-1)>=0" & sep & "IFERROR($AF2*1" & sep & "-1)<1))"

    ' PT-BR (;) – mesma lógica
    Dim fL_ptbr As String, fAF_ptbr As String
    fL_ptbr = "=OU($L2="""";E($R2<>""-"";SEERRO($L2*1;-1)>=0;SEERRO($L2*1;-1)<1))"
    fAF_ptbr = "=OU($AF2="""";E($Q2<>""-"";SEERRO($AF2*1;-1)>=0;SEERRO($AF2*1;-1)<1))"

    Call ApplyValidationCustomSmart(ws1, ws1.Range("L2:L2000"), fL_en, fL_ptbr, _
                                    "Hora adicional sem justificativa", _
                                    "Não é possível inserir HORA ADICIONAL ATRASO sem uma justificativa na ocorrência de entrega.")
    ws1.Range("L2:L2000").NumberFormat = "hh:mm:ss"

    Call ApplyValidationCustomSmart(ws1, ws1.Range("AF2:AF2000"), fAF_en, fAF_ptbr, _
                                    "Hora adicional sem justificativa", _
                                    "Não é possível inserir HORA ADICIONAL ATRASO sem uma justificativa na ocorrência de carregamento.")
    ws1.Range("AF2:AF2000").NumberFormat = "hh:mm:ss"

    ' 7) Sanity check seguro (não estoura 1004 ao ler Formula1)
    Debug.Print "R2:", TipoValidacao(ws1.Range("R2")), SafeFormula1(ws1.Range("R2"))
    Debug.Print "Q2:", TipoValidacao(ws1.Range("Q2")), SafeFormula1(ws1.Range("Q2"))
    Debug.Print "L2:", TipoValidacao(ws1.Range("L2")), SafeFormula1(ws1.Range("L2"))
    Debug.Print "AF2:", TipoValidacao(ws1.Range("AF2")), SafeFormula1(ws1.Range("AF2"))

    ' 8) Fallback mínimo: se L/AF ficarem sem validação, aplica regra básica (dependência)
    Call EnsureMinimalValidation(ws1)

    Exit Sub

TrataErro:
    MsgBox "Falha na compatibilização (Shared Workbook): " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' =================== HELPERS ===================

' LISTA com fallback: EN ? PT-BR + Selection (workaround 1004)
Private Sub ApplyValidationListSmart(ByVal ws As Worksheet, ByVal rng As Range, _
                                     ByVal f_en As String, ByVal f_ptbr As String)
    Application.DisplayAlerts = False

    If TryAddValidationList(rng, f_en) Then GoTo Fim
    If TryAddValidationList(rng, f_ptbr) Then GoTo Fim

    ' Selection workaround
    ws.Activate: rng.Select
    If TryAddValidationList(Selection, f_en) Then GoTo Fim
    If TryAddValidationList(Selection, f_ptbr) Then GoTo Fim

Fim:
    Application.DisplayAlerts = True
End Sub

' CUSTOM com fallback: EN ? PT-BR + Selection + linha-a-linha (somente L/AF)
Private Sub ApplyValidationCustomSmart(ByVal ws As Worksheet, ByVal rng As Range, _
                                       ByVal f_en As String, ByVal f_ptbr As String, _
                                       ByVal errTitle As String, ByVal errMsg As String)
    Application.DisplayAlerts = False

    If TryAddValidationCustom(rng, f_en, errTitle, errMsg) Then GoTo Fim
    If TryAddValidationCustom(rng, f_ptbr, errTitle, errMsg) Then GoTo Fim

    ' Selection workaround
    ws.Activate: rng.Select
    If TryAddValidationCustom(Selection, f_en, errTitle, errMsg) Then GoTo Fim
    If TryAddValidationCustom(Selection, f_ptbr, errTitle, errMsg) Then GoTo Fim

    ' Último recurso: linha a linha (ajustando a linha das referências)
    Dim c As Range
    For Each c In rng.Cells
        Dim f_en_row As String, f_ptbr_row As String
        If c.Column = ws.Range("L1").Column Then
            f_en_row = "=OR($L" & c.Row & "="""", AND($R" & c.Row & "<>""-"" , IFERROR($L" & c.Row & "*1,-1)>=0, IFERROR($L" & c.Row & "*1,-1)<1))"
            f_ptbr_row = "=OU($L" & c.Row & "=""""; E($R" & c.Row & "<>""-""; SEERRO($L" & c.Row & "*1;-1)>=0; SEERRO($L" & c.Row & "*1;-1)<1))"
        ElseIf c.Column = ws.Range("AF1").Column Then
            f_en_row = "=OR($AF" & c.Row & "="""", AND($Q" & c.Row & "<>""-"" , IFERROR($AF" & c.Row & "*1,-1)>=0, IFERROR($AF" & c.Row & "*1,-1)<1))"
            f_ptbr_row = "=OU($AF" & c.Row & "=""""; E($Q" & c.Row & "<>""-""; SEERRO($AF" & c.Row & "*1;-1)>=0; SEERRO($AF" & c.Row & "*1;-1)<1))"
        Else
            ' Não é L nem AF: não aplica per-row
            GoTo Fim
        End If

        If Not TryAddValidationCustom(c, f_en_row, errTitle, errMsg) Then
            Call TryAddValidationCustom(c, f_ptbr_row, errTitle, errMsg)
        End If
    Next c

Fim:
    Application.DisplayAlerts = True
End Sub

' Tenta aplicar validação de LISTA; retorna True/False, sem interromper em 1004
Private Function TryAddValidationList(ByVal tgt As Variant, ByVal formulaList As String) As Boolean
    On Error Resume Next
    With GetRange(tgt).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, formula1:=formulaList
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    TryAddValidationList = (Err.Number = 0)
    Err.Clear
End Function

' Tenta aplicar validação CUSTOM; retorna True/False, sem interromper em 1004
Private Function TryAddValidationCustom(ByVal tgt As Variant, ByVal formula1 As String, _
                                        ByVal errTitle As String, ByVal errMsg As String) As Boolean
    On Error Resume Next
    With GetRange(tgt).Validation
        .Delete
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, formula1:=formula1
        .IgnoreBlank = True
        .ErrorTitle = errTitle
        .ErrorMessage = errMsg
    End With
    TryAddValidationCustom = (Err.Number = 0)
    Err.Clear
End Function

' Converte Variant (Range/Selection) em Range
Private Function GetRange(ByVal tgt As Variant) As Range
    If TypeName(tgt) = "Range" Then
        Set GetRange = tgt
    Else
        Set GetRange = Selection
    End If
End Function

' Tipo de validação (depuração) - não dispara 1004
Private Function TipoValidacao(ByVal cel As Range) As String
    On Error Resume Next
    Select Case cel.Validation.Type
        Case xlValidateList:       TipoValidacao = "Lista"
        Case xlValidateCustom:     TipoValidacao = "Custom"
        Case xlValidateDate:       TipoValidacao = "Data"
        Case xlValidateTime:       TipoValidacao = "Hora"
        Case xlValidateInputOnly:  TipoValidacao = "Sem validação"
        Case Else:                 TipoValidacao = "Outro"
    End Select
    On Error GoTo 0
End Function

' Lê Formula1 com segurança (não estoura 1004 se não houver validação)
Private Function SafeFormula1(ByVal cel As Range) As String
    On Error Resume Next
    Dim v As Validation
    Set v = cel.Validation
    If v.Type = xlValidateInputOnly Then
        SafeFormula1 = "(sem validação)"
    Else
        SafeFormula1 = v.formula1
    End If
    If Err.Number <> 0 Then
        SafeFormula1 = "(sem fórmula / erro ao ler)"
        Err.Clear
    End If
End Function

' Garante um fallback mínimo em L/AF (dependência de R/Q) se a Custom não fixar
Private Sub EnsureMinimalValidation(ByVal ws1 As Worksheet)
    Dim sep As String: sep = Application.International(xlListSeparator)

    If ws1.Range("L2").Validation.Type = xlValidateInputOnly Then
        Dim fMinL_en As String, fMinL_ptbr As String
        fMinL_en = "=OR($L2=""""" & sep & "$R2<>""-"")"
        fMinL_ptbr = "=OU($L2="""";$R2<>""-"")"
        If Not TryAddValidationCustom(ws1.Range("L2:L2000"), fMinL_en, _
                                      "Hora adicional sem justificativa", _
                                      "Não é possível inserir HORA ADICIONAL ATRASO sem uma justificativa na ocorrência de entrega.") Then
            Call TryAddValidationCustom(ws1.Range("L2:L2000"), fMinL_ptbr, _
                                        "Hora adicional sem justificativa", _
                                        "Não é possível inserir HORA ADICIONAL ATRASO sem uma justificativa na ocorrência de entrega.")
        End If
    End If

    If ws1.Range("AF2").Validation.Type = xlValidateInputOnly Then
        Dim fMinAF_en As String, fMinAF_ptbr As String
        fMinAF_en = "=OR($AF2=""""" & sep & "$Q2<>""-"")"
        fMinAF_ptbr = "=OU($AF2="""";$Q2<>""-"")"
        If Not TryAddValidationCustom(ws1.Range("AF2:AF2000"), fMinAF_en, _
                                      "Hora adicional sem justificativa", _
                                      "Não é possível inserir HORA ADICIONAL ATRASO sem uma justificativa na ocorrência de carregamento.") Then
            Call TryAddValidationCustom(ws1.Range("AF2:AF2000"), fMinAF_ptbr, _
                                        "Hora adicional sem justificativa", _
                                        "Não é possível inserir HORA ADICIONAL ATRASO sem uma justificativa na ocorrência de carregamento.")
        End If
    End If
End Sub

Private Sub EnsureHeaderFilters(ByVal ws As Worksheet, ByVal headerRangeAddress As String)
    On Error Resume Next
    
    ' Se houver filtro aplicado (linhas escondidas por filtro), mostra todos os dados
    If ws.FilterMode Then ws.ShowAllData
    
    ' Garante que o AutoFilter (setinhas) esteja ativo sobre o cabeçalho
    If Not ws.AutoFilterMode Then
        ws.Range(headerRangeAddress).AutoFilter
    Else
        ' Algumas situações mantêm AutoFilterMode True mas removem as setas do cabeçalho desejado.
        ' Reaplica no cabeçalho alvo para garantir as setinhas lá.
        ws.AutoFilterMode = False
        ws.Range(headerRangeAddress).AutoFilter
    End If
    
    On Error GoTo 0
End Sub
