Attribute VB_Name = "modBatatas"
Public wb3 As Workbook

Public Sub AbrirFrmBatatas()

    If frmCross.Visible Then

        MsgBox _
            "Feche o formulário Cross antes de abrir Panifresh e Coletas.", _
            vbExclamation
            
        Exit Sub
            
    ElseIf frmPaniColetas.Visible Then

        MsgBox _
            "Feche o formulário Pani e Coletas antes de abrir Batatas.", _
            vbExclamation

        Exit Sub

    End If

    Unload frmBatatas

    frmBatatas.Show vbModeless

End Sub

Public Sub Etapa1Batatas()

    Dim Resposta As VbMsgBoxResult

    Dim caminho_modelo_consolidado_OUTROS As String

    Dim wb2 As Workbook

    Dim nome_arquivo As String
    'Dim novoNome As String

    Dim contador As Long

    Dim TemBatatas As Boolean

    TemBatatas = _
        CLng(frmBatatas.txtBatatasxFT.Value) > 0 _
        Or CLng(frmBatatas.txtBatatasxJC.Value) > 0 _
        Or CLng(frmBatatas.txtBatatasxRJ.Value) > 0

    If Not TemBatatas Then

        MsgBox _
            "Nenhuma programação foi informada.", _
            vbExclamation, _
            "Batatas"

        Exit Sub

    End If

    Resposta = MsgBox( _
        "Deseja prosseguir com a criação da programação Batatas?" & _
        vbCrLf & vbCrLf & _
        "Fat.: " & frmBatatas.txtDataFat.Value, _
        vbQuestion + vbYesNo, _
        "Confirmação")

    If Resposta = vbNo Then Exit Sub

    '=========================
    ' MODELO
    '=========================

    'TESTE
    caminho_modelo_consolidado_OUTROS = _
        PortfolioPath("backup modelo_consolidado\modelo_consolidado_OUTROS.xlsm")

    'PRODUÇÃO
    'caminho_modelo_consolidado_OUTROS = _
    '    PortfolioPath("modelo_consolidado_OUTROS.xlsm")

    Call AtualizandoCalendarioModeloConsolidadoOUTROS( _
        caminho_modelo_consolidado_OUTROS)

    Set wb2 = Workbooks.Open( _
        caminho_modelo_consolidado_OUTROS)

    '=========================
    ' NOME DO ARQUIVO
    '=========================

    nome_arquivo = _
        "Base BATATAS - " & _
        Format(CDate(frmBatatas.txtDataFat.Value), "dd.mm.yy")

    novoNome = _
        ThisWorkbook.Path & "\" & _
        nome_arquivo & ".xlsm"

    contador = 1

    Do While Dir(novoNome) <> ""

        novoNome = _
            ThisWorkbook.Path & "\" & _
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
        "Etapa 1 concluída com sucesso!" & _
        vbCrLf & vbCrLf & _
        wb3.Name, _
        vbInformation, _
        "Batatas"

End Sub

Public Sub Etapa2Batatas()

    With wb3.Sheets("Planilha1")

        'Data de Faturamento como texto
        .Range("B1").NumberFormat = "@"

        .Range("B1").Value = _
            Format( _
                CDate(frmBatatas.txtDataFat.Value), _
                "dd/mm/yy")

    End With

    'Preenchendo Semana de Faturamento
    Call PreencherC1ComValor

    'MsgBox _
        "Etapa 2 concluída com sucesso!", _
        vbInformation, _
        "Batatas"

End Sub

Public Sub Etapa3Batatas()

    Dim ws As Worksheet

    Dim UltimaLinha As Long
    Dim SufixoDia As String
    Dim DataFat As String

    Dim i As Long

    Set ws = wb3.Sheets("Planilha1")

    ws.Columns("G").NumberFormat = "@"

    DataFat = _
        Format( _
            CDate(frmBatatas.txtDataFat.Value), _
            "dd/mm/yy")

    SufixoDia = frmBatatas.PrefixoDiaSemana

    UltimaLinha = 2

    '=========================
    ' BATATA FT
    '=========================

    For i = 1 To CLng(frmBatatas.txtBatatasxFT.Value)

        ws.Cells(UltimaLinha, "D").Value = _
            "TRANSF BATATA"

        ws.Cells(UltimaLinha, "E").Value = _
            "FT"

        ws.Cells(UltimaLinha, "F").Value = _
            "CDSC"

        ws.Cells(UltimaLinha, "G").Value = _
            DataFat

        ws.Cells(UltimaLinha, "H").Value = _
            "BATATAFT" & SufixoDia & i

        ws.Cells(UltimaLinha, "I").Value = _
            "CDFT"

        UltimaLinha = UltimaLinha + 1

    Next i

    '=========================
    ' BATATA JC
    '=========================

    For i = 1 To CLng(frmBatatas.txtBatatasxJC.Value)

        ws.Cells(UltimaLinha, "D").Value = _
            "TRANSF BATATA"

        ws.Cells(UltimaLinha, "E").Value = _
            "JC"

        ws.Cells(UltimaLinha, "F").Value = _
            "CDSC"

        ws.Cells(UltimaLinha, "G").Value = _
            DataFat

        ws.Cells(UltimaLinha, "H").Value = _
            "BATATAJC" & SufixoDia & i

        ws.Cells(UltimaLinha, "I").Value = _
            "CDJC"

        UltimaLinha = UltimaLinha + 1

    Next i

    '=========================
    ' BATATA RJ
    '=========================

    For i = 1 To CLng(frmBatatas.txtBatatasxRJ.Value)

        ws.Cells(UltimaLinha, "D").Value = _
            "TRANSF BATATA"

        ws.Cells(UltimaLinha, "E").Value = _
            "JC"

        ws.Cells(UltimaLinha, "F").Value = _
            "CDSC"

        ws.Cells(UltimaLinha, "G").Value = _
            DataFat

        ws.Cells(UltimaLinha, "H").Value = _
            "BATATARJ" & SufixoDia & i

        ws.Cells(UltimaLinha, "I").Value = _
            "CDRJ"

        UltimaLinha = UltimaLinha + 1

    Next i

    'MsgBox _
        "Etapa 3 concluída com sucesso!", _
        vbInformation, _
        "Batatas"

End Sub

Public Sub Etapa4Batatas()

    Dim ws As Worksheet

    Dim UltimaLinha As Long
    Dim Linha As Long

    Dim Rota As String
    Dim DataFat As Date

    Set ws = wb3.Sheets("Planilha1")

    UltimaLinha = _
        ws.Cells(ws.Rows.Count, "H").End(xlUp).Row

    DataFat = _
        CDate(frmBatatas.txtDataFat.Value)

    For Linha = 2 To UltimaLinha

        Rota = _
            UCase(Trim(ws.Cells(Linha, "H").Value))

        '=========================
        ' J - DATA CARREGAMENTO
        '=========================

        ws.Cells(Linha, "J").Value = DataFat

        '=========================
        ' K - HORA CARREGAMENTO
        '=========================

        ws.Cells(Linha, "K").Value = _
            TimeValue("23:00:00")

        '=========================
        ' M - DATA ENTREGA
        '=========================

        If Left(Rota, 8) = "BATATARJ" Then

            ws.Cells(Linha, "M").Value = _
                DataFat + 2

        Else

            ws.Cells(Linha, "M").Value = _
                DataFat + 1

        End If

        '=========================
        ' N - HORA ENTREGA
        '=========================

        ws.Cells(Linha, "N").Value = _
            TimeValue("20:00:00")

    Next Linha

    ws.Columns("J").NumberFormat = "dd/mm/yyyy"
    ws.Columns("K").NumberFormat = "hh:mm:ss"

    ws.Columns("M").NumberFormat = "dd/mm/yyyy"
    ws.Columns("N").NumberFormat = "hh:mm:ss"

    'MsgBox _
        "Etapa 4 concluída com sucesso!", _
        vbInformation, _
        "Batatas"

End Sub

Public Sub Etapa5Batatas()

    Dim ws As Worksheet

    Dim UltimaLinha As Long
    Dim Linha As Long

    Set ws = wb3.Sheets("Planilha1")

    UltimaLinha = _
        ws.Cells(ws.Rows.Count, "H").End(xlUp).Row

    For Linha = 2 To UltimaLinha

        '=========================
        ' L = DTH CARREG PROG
        '=========================

        ws.Cells(Linha, "L").Formula = _
            "=J" & Linha & "+K" & Linha

        '=========================
        ' O = DTH ENTREGA LOJA
        '=========================

        ws.Cells(Linha, "O").Formula = _
            "=M" & Linha & "+N" & Linha

        '=========================
        ' P = EMBARQUE
        '=========================

        ws.Cells(Linha, "P").FormulaLocal = _
            "=E" & Linha & _
            "&H" & Linha & _
            "&TEXTO(G" & Linha & ";" & Chr(34) & _
            "dd/mm/aaaa" & Chr(34) & ")"

        '=========================
        ' Q = CAIXAS
        '=========================

        ws.Cells(Linha, "Q").Value = 1000

        '=========================
        ' R = ROTA ÚNICA
        '=========================

        ws.Cells(Linha, "R").Value = _
            ws.Cells(Linha, "H").Value

    Next Linha

    ws.Columns("L").NumberFormat = _
        "dd/mm/yyyy hh:mm"

    ws.Columns("O").NumberFormat = _
        "dd/mm/yyyy hh:mm"

    'MsgBox _
        "Etapa 5 concluída com sucesso!", _
        vbInformation, _
        "Batatas"

End Sub

Public Sub Etapa6Batatas()

    Dim ws As Worksheet

    Dim UltimaLinha As Long
    Dim tbl As ListObject

    Set ws = wb3.Sheets("Planilha1")

    UltimaLinha = _
        ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

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

    tbl.Name = "tbBatatas"

    'Tabela Clara 21
    tbl.TableStyle = "TableStyleLight21"

    'MsgBox _
        "Etapa 6 concluída com sucesso!", _
        vbInformation, _
        "Batatas"

End Sub

Public Sub Etapa7Batatas()

    Dim ws As Worksheet
    Dim UltimaLinha As Long

    Dim TotalRotas As Long
    Dim TotalLojas As Long

    Set ws = wb3.Sheets("Planilha1")

    UltimaLinha = _
        ws.Cells(ws.Rows.Count, "H").End(xlUp).Row

    TotalRotas = _
        Application.WorksheetFunction.CountA( _
        ws.Range("R2:R" & UltimaLinha))

    TotalLojas = _
        Application.WorksheetFunction.CountA( _
        ws.Range("I2:I" & UltimaLinha))

    wb3.Save

    MsgBox _
        "Programação Batatas criada com sucesso!" & _
        vbCrLf & vbCrLf & _
        "Total de rotas: " & _
        Format(TotalRotas, "#,##0") & vbCrLf & _
        "Total de lojas: " & _
        Format(TotalLojas, "#,##0"), _
        vbInformation, _
        "Control Tower MB"

    ThisWorkbook.Saved = False

    ThisWorkbook.Close SaveChanges:=False

End Sub
