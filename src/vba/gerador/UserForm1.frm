Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{DAB090CD-FC2C-4803-BCA8-940B9FA950BC}{47D82C32-FECA-4050-A50F-F6F08A46A624}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub cmdGerarPontoFocal_Click()

    Dim Resposta As VbMsgBoxResult

    Resposta = MsgBox( _
        "Selecione SIM para carregar um arquivo e atribuir Pontos Focais." & _
        vbCrLf & vbCrLf & _
        "Selecione NÃO para procurar rotas pendentes em Tbl_Prog.", _
        vbYesNoCancel + vbQuestion, _
        "Control Tower MB")

    Select Case Resposta

        Case vbYes
        
            Dim CaminhoArquivo As Variant
            Dim wbSelecionado As Workbook
        
            ChDrive Left$(PortfolioPath("Tbl_Prog\"), 1)
            ChDir PortfolioPath("Tbl_Prog\")
        
            CaminhoArquivo = _
                Application.GetOpenFilename( _
                "Arquivos do Excel (*.xlsx), *.xlsx", _
                , _
                "Selecione um arquivo")
        
            If CaminhoArquivo = False Then
        
                Exit Sub
        
            End If
        
            Application.ScreenUpdating = False
            Application.EnableEvents = False
        
            Set wbSelecionado = _
                Workbooks.Open(CaminhoArquivo)
        
            Call CriarTblPontoFocalPorData(wbSelecionado)
        
            wbSelecionado.Close SaveChanges:=False
        
            Application.EnableEvents = True
            Application.ScreenUpdating = True
        
            MsgBox _
                "Pontos Focais processados com sucesso!", _
                vbInformation, _
                "Control Tower MB"

        Case vbNo

            Call AtualizarPontosFocaisPendentes

        Case vbCancel

            Exit Sub

    End Select

End Sub

Private Sub CommandButton1_Click()
    Dim CaminhoArquivo As Variant
    Dim wb As Workbook

    ' Abre a caixa de diálogo para selecionar o arquivo
    ChDrive Left$(PortfolioPath("Tbl_Prog\"), 1)
    ChDir PortfolioPath("Tbl_Prog\")
    CaminhoArquivo = Application.GetOpenFilename("Arquivos do Excel (*.xls; *.xlsx), *.xls; *.xlsx", , "Selecione um arquivo")

    ' Verifica se o usuário clicou em Cancelar
    If CaminhoArquivo = False Then
        'MsgBox "Nenhum arquivo foi selecionado. O formulário será fechado.", vbInformation
        Unload Me
        Exit Sub
    End If
    
    'Fecha o UserForm
    Unload Me
    
   Application.ScreenUpdating = False
   Application.EnableEvents = False
   Application.DisplayAlerts = False
   Application.Visible = False ' Oculta o Excel
    
   UserForm2.Show vbModeless
   UserForm2.UpdateProgressBar 0

    ' Abre o arquivo e armazena a referência
    Set wb = Workbooks.Open(CaminhoArquivo)

        'Voltar para esse workbook
        ThisWorkbook.Activate
        
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
        
        UserForm2.UpdateProgressBar 25
    
        ' 3. Largura das colunas A:AI
        Dim larguras As Variant
        larguras = Array(14.55, 15.73, 9.45, 17.18, 7.82, 7.45, 18.27, 18.64, 15.36, 16.82, 17.18, 20.36, 14.18, 19.73, 12.91, 18.18, 23.45, 18.73, 12.09, 11.36, 14.36, 14.27, 14.18, 15.36, 15.27, 14.45, 10.27, 9.09, 16.91, 21.27, 12.73, 25.82, 21.73, 22.18, 25.36)
        For i = 1 To 35 ' A:AI
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
        
        'Formatação das linhas
        With ws1.Range("A2:AI2000")
            .Font.Name = "Aptos Narrow"
            .Font.Size = 8
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        'Formato hh:mm:ss nas colunas H, K, L e AF
        With ws1
            .Range("H2:H2000").NumberFormat = "hh:mm:ss"
            .Range("K2:K2000").NumberFormat = "hh:mm:ss"
            .Range("L2:L2000").NumberFormat = "hh:mm:ss"
            .Range("AF2:AF2000").NumberFormat = "hh:mm:ss"
        End With
        
        'Formato dd/mm/aaaa hh:mm nas colunas I, M, N, V, W, X, Y, AD, AG
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
        
        'Formato dd/mm/aaaa nas colunas D, G e J
        With ws1
            .Range("D2:D2000").NumberFormat = "dd/mm/yyyy"
            .Range("G2:G2000").NumberFormat = "dd/mm/yyyy"
            .Range("J2:J2000").NumberFormat = "dd/mm/yyyy"
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
        
        UserForm2.UpdateProgressBar 50
    
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
        
        ' Chama o método do próprio UserForm
        ConfigurarValidacoesWBNovo wbNovo
    
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

   ' For i = 0 To UBound(listaD)
    '    ws2.Cells(i + 1, 4).Value = listaD(i)
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
        
        UserForm2.UpdateProgressBar 75
    
        ' 8. Validação de data nas colunas I, M, N, V, W, X, Y, AD e AG
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
       ' With ws1.Range("R2:R2000").Validation
        '    .Delete
         '   .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Planilha2!$D$2:$D$54"
          '  .IgnoreBlank = True
           ' .InCellDropdown = True
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
        
        '13. Importando dados de outra Tbl_Prog
        wb.Activate
        
        '13. Importando dados de outra Tbl_Prog
        'wb.Activate
        'ult_linha = Range("A1").End(xlDown).Row
        
        wb.Activate
        
        Set ws = wb.Worksheets("Sheet1")
        
        With ws
            'Se existir filtro aplicado (mesmo sem dados filtrados)
            If .AutoFilterMode Then
                'Se houver linhas efetivamente filtradas, evita erro 1004
                If .FilterMode Then
                    .ShowAllData
                End If
                'Remove o AutoFilter da linha de cabeçalho
                .AutoFilterMode = False
            End If
        
            ult_linha = .Range("A1").End(xlDown).Row
        End With
        
            'Copiando colunas que possuem somente dados - A:H
            wb.Sheets(1).Range("A2:H" & ult_linha).Copy
            
            'Colar em wbNovo - A:H
            ws1.Range("A2").PasteSpecial Paste:=xlPasteValues
            
            'Copiando colunas que possuem somente dados - J:L
            wb.Sheets(1).Range("J2:L" & ult_linha).Copy
            
            'Colar em wbNovo - J:L
            ws1.Range("J2").PasteSpecial Paste:=xlPasteValues
            
            'Copiando colunas que possuem somente dados - Q:AE
            wb.Sheets(1).Range("Q2:AE" & ult_linha).Copy
            
            'Colar em wbNovo - Q:AE
            ws1.Range("Q2").PasteSpecial Paste:=xlPasteValues
            
            'Copiando colunas que possuem fórmulas - I
            wb.Sheets(1).Range("I2:I" & ult_linha).Copy
            
            'Colar em wbNovo - I
            ws1.Range("I2").PasteSpecial Paste:=xlPasteFormulas
    
            'Copiando colunas que possuem fórmulas - L:P
            wb.Sheets(1).Range("L2:P" & ult_linha).Copy
            
            'Colar em wbNovo - L:P
            ws1.Range("L2").PasteSpecial Paste:=xlPasteFormulas
            
            'Copiando colunas que possuem fórmulas - AG
            wb.Sheets(1).Range("AG2:AG" & ult_linha).Copy
            
            'Colar em wbNovo - AG
            ws1.Range("AG2").PasteSpecial Paste:=xlPasteFormulas
            
            'Copiando colunas que possuem somente dados - AH:AI
            wb.Sheets(1).Range("AH2:AI" & ult_linha).Copy
            
            'Colar em wbNovo - AH
            ws1.Range("AH2").PasteSpecial Paste:=xlPasteValues
            
            Application.CutCopyMode = False
            
        UserForm2.UpdateProgressBar 90
        
        Dim wbImportado As Workbook
        Set wbImportado = Workbooks.Open(CaminhoArquivo)
        
        Call CriarTblPontoFocalPorData(wbImportado)
        
        'Preenche Ponto Focal
        'ws1.Activate
        'Call ValidacaoPontoFocal
        'Call GerarPontosFocaisSeVazio(ws1)

        ' 14. Salvar o novo arquivo com o mesmo nome e no mesmo local
        
        Dim caminhoDestino As String
        Dim nomeArquivo As String
        Dim caminhoCompletoNovo As String
        
        ' Obter o caminho e nome do arquivo original (wb)
        caminhoDestino = wb.Path & Application.PathSeparator
        nomeArquivo = wb.Name
        caminhoCompletoNovo = caminhoDestino & nomeArquivo
        
        'Excluir o arquivo original
        wb.Close SaveChanges:=False
        Set wb = Nothing
        Kill caminhoCompletoNovo
        
        Application.DisplayAlerts = False
        wbNovo.SaveAs FileName:=caminhoCompletoNovo, FileFormat:=xlOpenXMLWorkbook
        DoEvents
        
        CompatibilizarParaSharedWorkbook wbNovo
        
        ' 15. Tornar o arquivo compartilhado
        Application.DisplayAlerts = False
        wbNovo.SaveAs FileName:=caminhoCompletoNovo, AccessMode:=xlShared
        DoEvents
        Application.DisplayAlerts = True
        
        UserForm2.UpdateProgressBar 100
        
        Unload UserForm2
        Application.Visible = True
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        
        Range("A1").Activate

        MsgBox "O novo arquivo Tbl_Prog foi criado com sucesso!", vbInformation, "Control Tower MB"
        
        ThisWorkbook.Close SaveChanges:=False
        DoEvents
                
End Sub

Private Sub CommandButton2_Click()

Application.ScreenUpdating = False

    'Fecha o UserForm
    Unload Me
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Visible = False ' Oculta o Excel
    
    UserForm2.Show vbModeless
    UserForm2.UpdateProgressBar 0

    
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
        
        UserForm2.UpdateProgressBar 25
    
        ' 3. Largura das colunas A:AI
        Dim larguras As Variant
        larguras = Array(14.55, 15.73, 9.45, 17.18, 7.82, 7.45, 18.27, 18.64, 15.36, 16.82, 17.18, 20.36, 14.18, 19.73, 12.91, 18.18, 23.45, 18.73, 12.09, 11.36, 14.36, 14.27, 14.18, 15.36, 15.27, 14.45, 10.27, 9.09, 16.91, 21.27, 12.73, 25.82, 21.73, 22.18, 25.36)
        For i = 1 To 35 ' A:AI
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
        
        'Formatação das linhas
        With ws1.Range("A2:AI2000")
            .Font.Name = "Aptos Narrow"
            .Font.Size = 8
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        'Formato hh:mm:ss nas colunas H, K, L e AF
        With ws1
            .Range("H2:H2000").NumberFormat = "hh:mm:ss"
            .Range("K2:K2000").NumberFormat = "hh:mm:ss"
            .Range("L2:L2000").NumberFormat = "hh:mm:ss"
            .Range("AF2:AF2000").NumberFormat = "hh:mm:ss"
        End With
        
        'Formato dd/mm/aaaa hh:mm nas colunas I, M, N, V, W, X, Y, AD e AG
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
        
        'Formato dd/mm/aaaa nas colunas D, G e J
        With ws1
            .Range("D2:D2000").NumberFormat = "dd/mm/yyyy"
            .Range("G2:G2000").NumberFormat = "dd/mm/yyyy"
            .Range("J2:J2000").NumberFormat = "dd/mm/yyyy"
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
        
        UserForm2.UpdateProgressBar 50
    
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
        
        ' Chama o método do próprio UserForm
        ConfigurarValidacoesWBNovo wbNovo
    
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
        
        UserForm2.UpdateProgressBar 75
    
        ' 8. Validação de data nas colunas I, M, N, V, W, X, Y, AD e AG
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
       ' With ws1.Range("R2:R2000").Validation
        '    .Delete
         '   .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Planilha2!$D$2:$D$54"
          '  .IgnoreBlank = True
           ' .InCellDropdown = True
'        End With
    
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
        
        '13. Apenas na linha 2 como exemplo
        With ws1
            .Range("I2").Formula = "=G2+H2"
            .Range("M2").Formula = "=J2+K2"
            .Range("N2").Formula = "=M2+L2"
            .Range("O2").Formula2 = "=B2&E2&TEXT(D2,""dd/mm/aaaa"")"
            .Range("P2").Formula2 = "=O2&F2"
            .Range("Z2:Z2000").Value = "NÃO"
            .Range("AG2").Formula2 = "=I2+AF2"
            '.Range("AH2:AH2000").value = "NÃO"
            '.Range("AI2:AI2000").value = 0
        End With
        
        '14. Preenchendo "-" nas colunas
        Dim colunasPreencher As Variant
        Dim col2 As Variant
        
        colunasPreencher = Array("Q", "R", "S", "T", "U", "AA", "AC", "AE")
        
        For Each col2 In colunasPreencher
            ws1.Range(col2 & "2:" & col2 & "2000").Value = "-"
        Next col2
        
        UserForm2.UpdateProgressBar 90

        '15. Salvando o arquivo
        Dim pastaDestino As String
        Dim nomeBase As String
        Dim nomeFinal As String
        Dim contador As Integer
        
        pastaDestino = ThisWorkbook.Path & Application.PathSeparator
        nomeBase = "Tbl_Prog_Modelo"
        nomeFinal = nomeBase & ".xlsx"
        
        ' Verifica se já existe e incrementa
        Do While Dir(pastaDestino & nomeFinal) <> ""
            contador = contador + 1
            nomeFinal = nomeBase & "(" & contador & ").xlsx"
        Loop
        
        Application.DisplayAlerts = False
        wbNovo.SaveAs FileName:=pastaDestino & nomeFinal, FileFormat:=xlOpenXMLWorkbook
        DoEvents
        
        CompatibilizarParaSharedWorkbook wbNovo
        
        'Proteger planilha Sheet1
        'ProtegerPontoFocal

        '16. Tornar o arquivo compartilhado
        Application.DisplayAlerts = False
        wbNovo.SaveAs FileName:=pastaDestino & nomeFinal, AccessMode:=xlShared
        DoEvents
        Application.DisplayAlerts = True
        
        UserForm2.UpdateProgressBar 100
        
        Unload UserForm2
        Application.Visible = True
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        
        Range("A1").Activate

        MsgBox "O novo arquivo " & nomeFinal & " foi criado com sucesso!", vbInformation, "Control Tower MB"
 
        ThisWorkbook.Close SaveChanges:=False
        DoEvents
        
    Application.ScreenUpdating = True

End Sub

Private Sub UserForm_Click()

End Sub
