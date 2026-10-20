Attribute VB_Name = "modCrossDocking"
Public wb3 As Workbook
Public dictRotasJC As Object
Public dictRotasGR As Object

Public Sub Etapa1CrossDocking()

    Dim Resposta As VbMsgBoxResult

    Dim caminho_modelo_consolidado_OUTROS As String

    Dim wb2 As Workbook
    'Dim wb3 As Workbook

    Dim nome_arquivo As String
    'Dim novoNome As String

    Dim contador As Long

    Resposta = MsgBox( _
    "Deseja prosseguir com a criação da programação Cross Docking?" & _
    vbCrLf & vbCrLf & _
    "Fat.: " & frmCross.txtDataFat.Value, _
    vbQuestion + vbYesNo, _
    "Confirmação")

    If Resposta = vbNo Then Exit Sub

    'PRODUÇÃO
    'caminho_modelo_consolidado_OUTROS = _
        PortfolioPath("modelo_consolidado_OUTROS.xlsm")

    'TESTE
    caminho_modelo_consolidado_OUTROS = _
        PortfolioPath("backup modelo_consolidado\modelo_consolidado_OUTROS.xlsm")

    Call AtualizandoCalendarioModeloConsolidadoOUTROS(caminho_modelo_consolidado_OUTROS)

    Set wb2 = Workbooks.Open(caminho_modelo_consolidado_OUTROS)

    nome_arquivo = _
        "Base CROSS DOCKING - " & _
        Format(CDate(frmCross.txtDataFat.Value), "dd.mm.yy")

    novoNome = PortfolioPath(nome_arquivo) & ".xlsm"

    contador = 1

    Do While Dir(novoNome) <> ""

        novoNome = ThisWorkbook.Path & "\" & _
                   nome_arquivo & _
                   " (" & contador & ").xlsm"

        contador = contador + 1

    Loop

    wb2.SaveAs _
        FileName:=novoNome, _
        FileFormat:=xlOpenXMLWorkbookMacroEnabled, _
        CreateBackup:=False

    wb2.Close SaveChanges:=False

    Set wb3 = Workbooks.Open(novoNome)

    'MsgBox _
        "Arquivo criado com sucesso:" & vbCrLf & vbCrLf & _
        wb3.Name, _
        vbInformation, _
        "Cross Docking"

End Sub

Public Sub Etapa2CrossDocking()

    With wb3.Sheets("Planilha1")

        'Data de Faturamento como texto
        .Range("B1").NumberFormat = "@"
        .Range("B1").Value = Format( _
            CDate(frmCross.txtDataFat.Value), _
            "dd/mm/yy")

    End With

    'Preenchendo Semana de Faturamento
    Call PreencherC1ComValor

    'MsgBox _
        "Etapa 2 concluída com sucesso!", _
        vbInformation, _
        "Cross Docking"

End Sub

Public Sub Etapa3CrossDocking()

    Dim ws As Worksheet
    Dim UltimaLinha As Long
    Dim Prefixo As String
    Dim DataFat As String

    Dim DiaSemana As Long
    Dim i As Long

    Set ws = wb3.Sheets("Planilha1")

    'Forçar DATA FAT como texto
    ws.Columns("G").NumberFormat = "@"

    DataFat = Format(CDate(frmCross.txtDataFat.Value), "dd/mm/yy")

    DiaSemana = Weekday(CDate(frmCross.txtDataFat.Value), vbSunday)

    Select Case DiaSemana

        Case 1: Prefixo = "0"
        Case 2: Prefixo = "2"
        Case 3: Prefixo = "3"
        Case 4: Prefixo = "4"
        Case 5: Prefixo = "5"
        Case 6: Prefixo = "6"
        Case 7: Prefixo = "7"

    End Select

    UltimaLinha = 2

    '=====================================
    ' CROSS JD - FT
    '=====================================

    For i = 1 To CLng(frmCross.txtFTxJD.Value)

        ws.Cells(UltimaLinha, "D").Value = "CROSS CDJD"
        ws.Cells(UltimaLinha, "E").Value = "FT"
        ws.Cells(UltimaLinha, "F").Value = "CDFT"
        ws.Cells(UltimaLinha, "G").Value = DataFat
        ws.Cells(UltimaLinha, "H").Value = "TRJD" & Prefixo & Format(i, "0")
        ws.Cells(UltimaLinha, "I").Value = "CDJD"

        UltimaLinha = UltimaLinha + 1

    Next i

    '=====================================
    ' CROSS JD - GR
    '=====================================

    For i = 1 To CLng(frmCross.txtGRxJD.Value)

        ws.Cells(UltimaLinha, "D").Value = "CROSS CDJD"
        ws.Cells(UltimaLinha, "E").Value = "GR"
        ws.Cells(UltimaLinha, "F").Value = "CDGR"
        ws.Cells(UltimaLinha, "G").Value = DataFat
        ws.Cells(UltimaLinha, "H").Value = "TRJD" & Prefixo & Format(i, "0")
        ws.Cells(UltimaLinha, "I").Value = "CDJD"

        UltimaLinha = UltimaLinha + 1

    Next i

    '=====================================
    ' CROSS NE - CE
    '=====================================

    For i = 1 To CLng(frmCross.txtNExCE.Value)

        ws.Cells(UltimaLinha, "D").Value = "CROSS CDCE"
        ws.Cells(UltimaLinha, "E").Value = "NE"
        ws.Cells(UltimaLinha, "F").Value = "CDNE"
        ws.Cells(UltimaLinha, "G").Value = DataFat
        ws.Cells(UltimaLinha, "H").Value = "TRCE" & Prefixo & Format(i, "0")
        ws.Cells(UltimaLinha, "I").Value = "CDCE"

        UltimaLinha = UltimaLinha + 1

    Next i

    '=====================================
    ' CROSS NE - PI
    '=====================================

    For i = 1 To CLng(frmCross.txtNExPI.Value)

        ws.Cells(UltimaLinha, "D").Value = "CROSS CDPI"
        ws.Cells(UltimaLinha, "E").Value = "NE"
        ws.Cells(UltimaLinha, "F").Value = "CDNE"
        ws.Cells(UltimaLinha, "G").Value = DataFat
        ws.Cells(UltimaLinha, "H").Value = "TRPI" & Prefixo & Format(i, "0")
        ws.Cells(UltimaLinha, "I").Value = "CDPI"

        UltimaLinha = UltimaLinha + 1

    Next i

    '=====================================
    ' CROSS PR
    '=====================================

    For i = 1 To CLng(frmCross.txtPRxRS.Value)

        ws.Cells(UltimaLinha, "D").Value = "CROSS CDRS"
        ws.Cells(UltimaLinha, "E").Value = "PR"
        ws.Cells(UltimaLinha, "F").Value = "CDPR"
        ws.Cells(UltimaLinha, "G").Value = DataFat
        ws.Cells(UltimaLinha, "H").Value = "TRRS" & Prefixo & Format(i, "0")
        ws.Cells(UltimaLinha, "I").Value = "CDRS"

        UltimaLinha = UltimaLinha + 1

    Next i

    'MsgBox _
        "Etapa 3 concluída com sucesso!", _
        vbInformation, _
        "Cross Docking"

End Sub

Public Sub Etapa4CrossDocking()

    Dim ws As Worksheet
    Dim wsCrossDocking As Worksheet

    Dim UltimaLinha As Long
    Dim Linha As Long
    Dim LinhaCross As Long
    Dim LinhaEncontrada As Long

    Dim DataFat As Date
    Dim DataCarreg As Date
    Dim DataEntrega As Date

    Dim Origem As String
    Dim Loja As String

    Dim DiasCarreg As Long
    Dim DiasEntrega As Long

    Dim EncontrouEmColunaA As Boolean

    Set ws = wb3.Sheets("Planilha1")
    Set wsCrossDocking = ThisWorkbook.Sheets("CROSS DOCKING")

    UltimaLinha = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row

    DataFat = CDate(frmCross.txtDataFat.Value)

    For Linha = 2 To UltimaLinha

        Origem = Trim(ws.Cells(Linha, "F").Value)
        Loja = Trim(ws.Cells(Linha, "I").Value)

        LinhaEncontrada = 0
        EncontrouEmColunaA = False

        '=================================================
        ' PROCURA ORIGEM + DESTINO
        '=================================================

        For LinhaCross = 2 To wsCrossDocking.Cells(wsCrossDocking.Rows.Count, "A").End(xlUp).Row

            If wsCrossDocking.Cells(LinhaCross, "A").Value = Origem _
               And wsCrossDocking.Cells(LinhaCross, "C").Value = Loja Then

                LinhaEncontrada = LinhaCross
                EncontrouEmColunaA = True
                Exit For

            End If

            If wsCrossDocking.Cells(LinhaCross, "B").Value = Origem _
               And wsCrossDocking.Cells(LinhaCross, "C").Value = Loja Then

                LinhaEncontrada = LinhaCross
                EncontrouEmColunaA = False
                Exit For

            End If

        Next LinhaCross

        If LinhaEncontrada > 0 Then

            '=====================================
            ' DATA CARREG
            '=====================================

            If Weekday(DataFat, vbSunday) = 7 Then

                DiasCarreg = wsCrossDocking.Cells(LinhaEncontrada, "E").Value

            Else

                DiasCarreg = wsCrossDocking.Cells(LinhaEncontrada, "D").Value

            End If

            DataCarreg = DataFat + DiasCarreg

            ws.Cells(Linha, "J").NumberFormat = "@"
            ws.Cells(Linha, "J").Value = Format(DataCarreg, "dd/mm/yy")

            '=====================================
            ' HORA CARREG
            '=====================================

            ws.Cells(Linha, "K").Value = _
                wsCrossDocking.Cells(LinhaEncontrada, "F").Value

            ws.Cells(Linha, "K").NumberFormat = "hh:mm:ss"

            '=====================================
            ' TRANSPORTADORA
            '=====================================

            If EncontrouEmColunaA Then

                ws.Cells(Linha, "U").Value = _
                    wsCrossDocking.Cells(LinhaEncontrada, "I").Value

            Else

                ws.Cells(Linha, "U").Value = _
                    wsCrossDocking.Cells(LinhaEncontrada, "J").Value

            End If

            '=====================================
            ' DATA ENTREGA
            '=====================================

            DiasEntrega = wsCrossDocking.Cells(LinhaEncontrada, "G").Value

            DataEntrega = DataCarreg + DiasEntrega

            '=====================================
            ' EXCEÇÃO OPERACIONAL
            ' CDGR -> CDJD
            '=====================================

            If Origem = "CDGR" And Loja = "CDJD" Then

                DataEntrega = DataEntrega + 1

                ws.Cells(Linha, "K").Value = TimeValue("23:00:00")
                ws.Cells(Linha, "N").Value = TimeValue("03:00:00")

            Else

                ws.Cells(Linha, "N").Value = _
                    wsCrossDocking.Cells(LinhaEncontrada, "H").Value

            End If

            ws.Cells(Linha, "M").NumberFormat = "@"
            ws.Cells(Linha, "M").Value = Format(DataEntrega, "dd/mm/yy")

            ws.Cells(Linha, "N").NumberFormat = "hh:mm:ss"

        End If

        '=================================================
        ' DTH CARRG PROG
        '=================================================

        ws.Cells(Linha, "L").Formula = _
            "=J" & Linha & "+K" & Linha

        ws.Cells(Linha, "L").NumberFormat = _
            "dd/mm/yyyy hh:mm"

        '=================================================
        ' DTH ENTR PROG
        '=================================================

        ws.Cells(Linha, "O").Formula = _
            "=M" & Linha & "+N" & Linha

        ws.Cells(Linha, "O").NumberFormat = _
            "dd/mm/yyyy hh:mm"

    Next Linha

    'MsgBox _
        "Etapa 4 concluída com sucesso!", _
        vbInformation, _
        "Cross Docking"

End Sub

Public Sub Etapa5CrossDocking()

    Dim ws As Worksheet
    Dim UltimaLinha As Long
    Dim Linha As Long

    Dim CdFat As String
    Dim Rota As String
    Dim DataFat As String

    Set ws = wb3.Sheets("Planilha1")

    UltimaLinha = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row

    For Linha = 2 To UltimaLinha

        CdFat = Trim(ws.Cells(Linha, "E").Value)
        Rota = Trim(ws.Cells(Linha, "H").Value)

        DataFat = Format( _
            CDate(frmCross.txtDataFat.Value), _
            "dd/mm/yyyy")

        '=====================================
        ' EMBARQUE1 (P)
        '=====================================

        ws.Cells(Linha, "P").FormulaLocal = _
            "=E" & Linha & _
            "&H" & Linha & _
            "&TEXTO(G" & Linha & ";" & Chr(34) & _
            "dd/mm/aaaa" & Chr(34) & ")"

        '=====================================
        ' CAIXAS (Q)
        '=====================================

        ws.Cells(Linha, "Q").Value = 1000

        '=====================================
        ' ROTAS ÚNICAS (R)
        '=====================================

        ws.Cells(Linha, "R").Value = _
            ws.Cells(Linha, "H").Value

    Next Linha

    'MsgBox _
        "Etapa 5 concluída com sucesso!", _
        vbInformation, _
        "Cross Docking"

End Sub

Public Sub Etapa6ACrossDocking()

    Set dictRotasJC = CreateObject("Scripting.Dictionary")
    Set dictRotasGR = CreateObject("Scripting.Dictionary")

    If frmCross.optNao.Value Then Exit Sub

    If frmCross.optJC.Value Then

        Call ProcessarArquivoMontaPallet("JC", dictRotasJC)

    ElseIf frmCross.optGR.Value Then

        Call ProcessarArquivoMontaPallet("GR", dictRotasGR)

    ElseIf frmCross.optAmbos.Value Then

        Call ProcessarArquivoMontaPallet("JC", dictRotasJC)
        Call ProcessarArquivoMontaPallet("GR", dictRotasGR)

    End If
    
    'MsgBox _
        "Etapa 6A concluída com sucesso!", _
        vbInformation, _
        "Cross Docking"

End Sub

Private Sub ProcessarArquivoMontaPallet( _
    ByVal TipoCD As String, _
    ByRef dictRotas As Object)

    Dim AnoFat As String
    Dim MesFat As String
    Dim DiaFat As String

    Dim PastaBase As String
    Dim PastaMes As String

    Dim CaminhoArquivo As String

    Dim wbMP As Workbook
    Dim wsMP As Worksheet

    Dim UltimaLinha As Long
    Dim Linha As Long

    Dim Rota As String
    Dim LojaDestino As String

    Dim ArquivoSelecionado As Variant

    AnoFat = Format(CDate(frmCross.txtDataFat.Value), "yyyy")
    MesFat = Format(CDate(frmCross.txtDataFat.Value), "mm")
    DiaFat = Format(CDate(frmCross.txtDataFat.Value), "dd")

    If TipoCD = "JC" Then

        PastaBase = _
            PortfolioPath("Planeamento\CDJC\") & AnoFat & _
            "\07. PCO - OPERATION PLAN AND CONTROL\1 - PLANEJAMENTO\Monta Palete\"

    Else

        PastaBase = _
            PortfolioPath("Planeamento\CDGR\PLANEJADO CDGR SEMANAL\ESQUEMA DE CARREGAMENTO\1 - PLANEJAMENTO\") & _
            AnoFat & "\Monta Palete\"

    End If

    PastaMes = Dir(PastaBase & MesFat & "*", vbDirectory)

    If PastaMes <> "" Then

        CaminhoArquivo = _
            PastaBase & PastaMes & "\" & _
            "Base MPallet " & TipoCD & " " & _
            Format(CLng(DiaFat), "00") & ".xlsx"

    End If

    If Dir(CaminhoArquivo) = "" Then

        ChDir PastaBase

        ArquivoSelecionado = Application.GetOpenFilename( _
            "Arquivos Excel (*.xlsx), *.xlsx", _
            , _
            "Selecione o arquivo Monta Pallet")

        If ArquivoSelecionado = False Then Exit Sub

        CaminhoArquivo = ArquivoSelecionado

    End If

    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    
    Set wbMP = Workbooks.Open( _
        FileName:=CaminhoArquivo, _
        UpdateLinks:=0, _
        ReadOnly:=True)
    
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True

    On Error Resume Next

    Set wsMP = wbMP.Sheets("EC D")

    On Error GoTo 0

    If wsMP Is Nothing Then

        wbMP.Close False
        Exit Sub

    End If

    If wsMP.FilterMode Then wsMP.ShowAllData

    If wsMP.AutoFilterMode Then

        wsMP.AutoFilterMode = False

    End If

    wsMP.Rows(1).AutoFilter

    UltimaLinha = wsMP.Cells(wsMP.Rows.Count, "J").End(xlUp).Row

    For Linha = 2 To UltimaLinha

        Rota = Trim(wsMP.Cells(Linha, "J").Value)

        If Left(Rota, 2) = "TR" Then

            LojaDestino = "CD" & Mid(Rota, 3, 2)

            Select Case LojaDestino

                Case "CDJD", "CDCE", "CDPI", "CDRS"

                    'Ignora

                Case Else

                    If DestinoExisteCrossDocking(LojaDestino) Then

                        If Not dictRotas.Exists(Rota) Then
                        
                            dictRotas.Add Rota, _
                                wsMP.Cells(Linha, "G").Value & "|" & _
                                wsMP.Cells(Linha, "H").Value & "|" & _
                                wsMP.Cells(Linha, "I").Value
                        
                        End If

                    End If

            End Select

        End If

    Next Linha

    wbMP.Close SaveChanges:=False

End Sub

Private Function DestinoExisteCrossDocking( _
    ByVal LojaDestino As String) As Boolean

    Dim wsCrossDocking As Worksheet
    Dim UltimaLinha As Long
    Dim Linha As Long

    Set wsCrossDocking = ThisWorkbook.Sheets("CROSS DOCKING")

    UltimaLinha = _
        wsCrossDocking.Cells( _
        wsCrossDocking.Rows.Count, "C").End(xlUp).Row

    For Linha = 2 To UltimaLinha

        If Trim(wsCrossDocking.Cells(Linha, "C").Value) = LojaDestino Then

            DestinoExisteCrossDocking = True
            Exit Function

        End If

    Next Linha

End Function

Public Sub Etapa6BCrossDocking()

    If frmCross.optNao.Value Then Exit Sub

    If frmCross.optJC.Value Then

        Call GerarMontaPallet("JC", dictRotasJC)

    ElseIf frmCross.optGR.Value Then

        Call GerarMontaPallet("GR", dictRotasGR)

    ElseIf frmCross.optAmbos.Value Then

        Call GerarMontaPallet("JC", dictRotasJC)
        Call GerarMontaPallet("GR", dictRotasGR)

    End If

    'MsgBox _
        "Etapa 6B concluída com sucesso!", _
        vbInformation, _
        "Cross Docking"

End Sub

Private Sub GerarMontaPallet( _
    ByVal CdFat As String, _
    ByVal dictRotas As Object)

    Dim ws As Worksheet
    Dim wsCrossDocking As Worksheet

    Dim UltimaLinha As Long

    Dim Chave As Variant

    Dim Origem As String
    Dim Loja As String
    Dim Rota As String

    Dim Dados As Variant

    Dim DataCarreg As Date
    Dim HoraCarreg As Variant

    Dim DataEntrega As Date
    Dim DiasEntrega As Long
    Dim HoraEntrega As Variant

    Dim CelDestino As Range

    Set ws = wb3.Sheets("Planilha1")
    Set wsCrossDocking = ThisWorkbook.Sheets("CROSS DOCKING")

    For Each Chave In dictRotas.Keys

        UltimaLinha = _
            ws.Cells(ws.Rows.Count, "H").End(xlUp).Row + 1

        Rota = Chave

        Dados = Split(dictRotas(Chave), "|")

        Loja = "CD" & Mid(Rota, 3, 2)

        Origem = "CD" & CdFat

        '=====================================
        ' D
        '=====================================

        ws.Cells(UltimaLinha, "D").Value = _
            "CROSS " & Loja

        '=====================================
        ' E
        '=====================================

        ws.Cells(UltimaLinha, "E").Value = CdFat

        '=====================================
        ' F
        '=====================================

        ws.Cells(UltimaLinha, "F").Value = Origem

        '=====================================
        ' G
        '=====================================

        ws.Cells(UltimaLinha, "G").NumberFormat = "@"

        ws.Cells(UltimaLinha, "G").Value = _
            Format(CDate(frmCross.txtDataFat.Value), _
            "dd/mm/yy")

        '=====================================
        ' H
        '=====================================

        ws.Cells(UltimaLinha, "H").Value = Rota

        '=====================================
        ' I
        '=====================================

        ws.Cells(UltimaLinha, "I").Value = Loja

        '=====================================
        ' J
        '=====================================

        DataCarreg = CDate(Dados(0))

        ws.Cells(UltimaLinha, "J").NumberFormat = "@"

        ws.Cells(UltimaLinha, "J").Value = _
            Format(DataCarreg, "dd/mm/yy")

        
        '========================================
        ' K
        '========================================
        
        If Trim(Dados(1)) <> "" Then
        
            If IsNumeric(Dados(1)) Then
        
                ws.Cells(UltimaLinha, "K").Value = CDbl(Dados(1))
                ws.Cells(UltimaLinha, "K").NumberFormat = "hh:mm:ss"
        
            Else
        
                ws.Cells(UltimaLinha, "K").Value = Dados(1)
        
            End If
        
        End If
        
        '=====================================
        ' U
        '=====================================

        ws.Cells(UltimaLinha, "U").Value = Dados(2)
        
        '=====================================
        ' FALLBACK TRANSPORTADORA
        '=====================================
        
        If Trim(ws.Cells(UltimaLinha, "U").Value) = "" Then
        
            Dim LinhaCross As Long
            Dim LinhaEncontrada As Long
            Dim EncontrouEmColunaA As Boolean
        
            LinhaEncontrada = 0
            EncontrouEmColunaA = False
        
            For LinhaCross = 2 To wsCrossDocking.Cells(wsCrossDocking.Rows.Count, "A").End(xlUp).Row
        
                If wsCrossDocking.Cells(LinhaCross, "A").Value = Origem _
                   And wsCrossDocking.Cells(LinhaCross, "C").Value = Loja Then
        
                    LinhaEncontrada = LinhaCross
                    EncontrouEmColunaA = True
                    Exit For
        
                End If
        
                If wsCrossDocking.Cells(LinhaCross, "B").Value = Origem _
                   And wsCrossDocking.Cells(LinhaCross, "C").Value = Loja Then
        
                    LinhaEncontrada = LinhaCross
                    EncontrouEmColunaA = False
                    Exit For
        
                End If
        
            Next LinhaCross
        
            If LinhaEncontrada > 0 Then
        
                If EncontrouEmColunaA Then
        
                    ws.Cells(UltimaLinha, "U").Value = _
                        wsCrossDocking.Cells(LinhaEncontrada, "I").Value
        
                Else
        
                    ws.Cells(UltimaLinha, "U").Value = _
                        wsCrossDocking.Cells(LinhaEncontrada, "J").Value
        
                End If
        
            End If
        
        End If

        '=====================================
        ' DESTINO
        '=====================================

        Set CelDestino = wsCrossDocking.Columns("C").Find( _
            What:=Loja, _
            LookAt:=xlWhole)

        If Not CelDestino Is Nothing Then

            DiasEntrega = _
                wsCrossDocking.Cells(CelDestino.Row, "G").Value

            DataEntrega = _
                DataCarreg + DiasEntrega

            ws.Cells(UltimaLinha, "M").NumberFormat = "@"

            ws.Cells(UltimaLinha, "M").Value = _
                Format(DataEntrega, "dd/mm/yy")

            HoraEntrega = _
                wsCrossDocking.Cells(CelDestino.Row, "H").Value

            ws.Cells(UltimaLinha, "N").Value = _
                HoraEntrega

            ws.Cells(UltimaLinha, "N").NumberFormat = _
                "hh:mm:ss"

        End If

        '=====================================
        ' L
        '=====================================

        ws.Cells(UltimaLinha, "L").Formula = _
            "=J" & UltimaLinha & _
            "+K" & UltimaLinha

        ws.Cells(UltimaLinha, "L").NumberFormat = _
            "dd/mm/yyyy hh:mm"

        '=====================================
        ' O
        '=====================================

        ws.Cells(UltimaLinha, "O").Formula = _
            "=M" & UltimaLinha & _
            "+N" & UltimaLinha

        ws.Cells(UltimaLinha, "O").NumberFormat = _
            "dd/mm/yyyy hh:mm"

        '=====================================
        ' P
        '=====================================

        ws.Cells(UltimaLinha, "P").Value = _
            CdFat & _
            Rota & _
            Format( _
            CDate(frmCross.txtDataFat.Value), _
            "dd/mm/yyyy")

        '=====================================
        ' Q
        '=====================================

        ws.Cells(UltimaLinha, "Q").Value = 1000

        '=====================================
        ' R
        '=====================================

        ws.Cells(UltimaLinha, "R").Value = Rota

    Next Chave

End Sub

Public Sub Etapa7CrossDocking()

    Dim ws As Worksheet

    Dim UltimaLinha As Long
    Dim tbl As ListObject

    Set ws = wb3.Sheets("Planilha1")

    UltimaLinha = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

    'Remover tabela existente (se houver)
    On Error Resume Next

    For Each tbl In ws.ListObjects
        tbl.Unlist
    Next tbl

    On Error GoTo 0

    'Criar nova tabela
    Set tbl = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=ws.Range("D1:V" & UltimaLinha), _
        XlListObjectHasHeaders:=xlYes)

    tbl.Name = "tbCrossDocking"

    'Tabela Clara 21
    tbl.TableStyle = "TableStyleLight21"

    'MsgBox _
        "Etapa 7 concluída com sucesso!", _
        vbInformation, _
        "Cross Docking"

End Sub

Public Sub Etapa8CrossDocking()

    Dim ws As Worksheet
    Dim UltimaLinha As Long

    Dim TotalRotas As Long
    Dim TotalLojas As Long

    Set ws = wb3.Sheets("Planilha1")

    UltimaLinha = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row

    TotalRotas = Application.WorksheetFunction.CountA(ws.Range("R2:R" & UltimaLinha))
    TotalLojas = Application.WorksheetFunction.CountA(ws.Range("I2:I" & UltimaLinha))

    wb3.Save

    MsgBox _
        "Programação Cross Docking criada com sucesso!" & vbCrLf & vbCrLf & _
        "Total de rotas: " & Format(TotalRotas, "#,##0") & vbCrLf & _
        "Total de lojas: " & Format(TotalLojas, "#,##0"), _
        vbInformation, _
        "Control Tower MB"

    ThisWorkbook.Saved = False
    ThisWorkbook.Close SaveChanges:=False

End Sub
