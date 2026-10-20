Attribute VB_Name = "modPaniColetas"
Public wb3 As Workbook

Public Sub AbrirFrmPaniColetas()

    If frmCross.Visible Then

        MsgBox _
            "Feche o formulário Cross antes de abrir Panifresh e Coletas.", _
            vbExclamation

        Exit Sub

    End If

    Unload frmPaniColetas

    frmPaniColetas.Show vbModeless

End Sub

Public Sub Etapa1PaniColetas()

    Dim Resposta As VbMsgBoxResult

    Dim caminho_modelo_consolidado_OUTROS As String

    Dim wb2 As Workbook

    Dim nome_arquivo As String
    'Dim novoNome As String

    Dim contador As Long

    Dim TemPani As Boolean
    Dim TemColetas As Boolean

    Dim NomeProgramacao As String

    '=========================
    ' IDENTIFICAR PROGRAMAÇÃO
    '=========================

    TemPani = _
        CLng(frmPaniColetas.txtPanixMG.Value) > 0 _
        Or CLng(frmPaniColetas.txtPanixRJ.Value) > 0 _
        Or CLng(frmPaniColetas.txtPanixSP.Value) > 0

    TemColetas = _
        CLng(frmPaniColetas.txtColBomSabor.Value) > 0 _
        Or CLng(frmPaniColetas.txtEntBomSabor.Value) > 0 _
        Or CLng(frmPaniColetas.txtColCellier.Value) > 0 _
        Or CLng(frmPaniColetas.txtEntCellier.Value) > 0 _
        Or CLng(frmPaniColetas.txtColDohler.Value) > 0 _
        Or CLng(frmPaniColetas.txtEntDohler.Value) > 0 _
        Or CLng(frmPaniColetas.txtColKerry.Value) > 0 _
        Or CLng(frmPaniColetas.txtEntKerry.Value) > 0 _
        Or CLng(frmPaniColetas.txtColOmamori.Value) > 0 _
        Or CLng(frmPaniColetas.txtEntOmamori.Value) > 0 _
        Or CLng(frmPaniColetas.txtColPoty.Value) > 0 _
        Or CLng(frmPaniColetas.txtEntPoty.Value) > 0 _
        Or CLng(frmPaniColetas.txtTransfJC.Value) > 0

    If TemPani And TemColetas Then

        NomeProgramacao = "Panifresh e Coletas"

    ElseIf TemPani Then

        NomeProgramacao = "Panifresh"

    ElseIf TemColetas Then

        NomeProgramacao = "Coletas"

    Else

        MsgBox _
            "Nenhuma programação foi informada.", _
            vbExclamation, _
            "Panifresh e Coletas"

        Exit Sub

    End If

    '=========================
    ' CONFIRMAÇÃO
    '=========================

    Resposta = MsgBox( _
        "Deseja prosseguir com a criação da programação " & _
        NomeProgramacao & "?" & _
        vbCrLf & vbCrLf & _
        "Fat.: " & frmPaniColetas.txtDataFat.Value, _
        vbQuestion + vbYesNo, _
        "Confirmação")

    If Resposta = vbNo Then Exit Sub

    '=========================
    ' MODELO
    '=========================

    'Teste
    caminho_modelo_consolidado_OUTROS = _
        PortfolioPath("backup modelo_consolidado\modelo_consolidado_OUTROS.xlsm")

    'Produção
    'caminho_modelo_consolidado_OUTROS = _
        PortfolioPath("modelo_consolidado_OUTROS.xlsm")
    
    Call AtualizandoCalendarioModeloConsolidadoOUTROS( _
        caminho_modelo_consolidado_OUTROS)

    Set wb2 = Workbooks.Open( _
        caminho_modelo_consolidado_OUTROS)

    '=========================
    ' NOME DO ARQUIVO
    '=========================

    If TemPani And TemColetas Then

        nome_arquivo = _
            "Base PANI e Coletas - " & _
            Format(CDate(frmPaniColetas.txtDataFat.Value), "dd.mm.yy")

    ElseIf TemPani Then

        nome_arquivo = _
            "Base PANIFRESH - " & _
            Format(CDate(frmPaniColetas.txtDataFat.Value), "dd.mm.yy")

    Else

        nome_arquivo = _
            "Base Coletas - " & _
            Format(CDate(frmPaniColetas.txtDataFat.Value), "dd.mm.yy")

    End If

    novoNome = ThisWorkbook.Path & "\" & _
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
        "Panifresh e Coletas"

End Sub

Public Sub Etapa2PaniColetas()

    With wb3.Sheets("Planilha1")

        'Data de Faturamento como texto
        .Range("B1").NumberFormat = "@"

        .Range("B1").Value = _
            Format( _
                CDate(frmPaniColetas.txtDataFat.Value), _
                "dd/mm/yy")

    End With

    'Preenchendo Semana de Faturamento
    Call PreencherC1ComValor

    'MsgBox _
        "Etapa 2 concluída com sucesso!", _
        vbInformation, _
        "Panifresh e Coletas"

End Sub

Public Sub Etapa3APaniColetas()

    Dim ws As Worksheet

    Dim UltimaLinha As Long
    Dim SufixoDia As String
    Dim DataFat As String

    Dim i As Long

    Set ws = wb3.Sheets("Planilha1")

    ws.Columns("G").NumberFormat = "@"

    DataFat = _
        Format( _
            CDate(frmPaniColetas.txtDataFat.Value), _
            "dd/mm/yy")

    SufixoDia = frmPaniColetas.PrefixoDiaSemana

    UltimaLinha = 2

    '=========================
    ' PANI MG
    '=========================

    For i = 1 To CLng(frmPaniColetas.txtPanixMG.Value)

        ws.Cells(UltimaLinha, "D").Value = _
            "COLETA PANIFRESH"

        ws.Cells(UltimaLinha, "E").Value = _
            "JC"

        ws.Cells(UltimaLinha, "F").Value = _
            "PANIFRESH"

        ws.Cells(UltimaLinha, "G").Value = _
            DataFat

        ws.Cells(UltimaLinha, "H").Value = _
            "PANIMG" & SufixoDia & i

        ws.Cells(UltimaLinha, "I").Value = _
            "CDMG"

        UltimaLinha = UltimaLinha + 1

    Next i

    '=========================
    ' PANI RJ
    '=========================

    For i = 1 To CLng(frmPaniColetas.txtPanixRJ.Value)

        ws.Cells(UltimaLinha, "D").Value = _
            "COLETA PANIFRESH"

        ws.Cells(UltimaLinha, "E").Value = _
            "JC"

        ws.Cells(UltimaLinha, "F").Value = _
            "PANIFRESH"

        ws.Cells(UltimaLinha, "G").Value = _
            DataFat

        ws.Cells(UltimaLinha, "H").Value = _
            "PANIRJ" & SufixoDia & i

        ws.Cells(UltimaLinha, "I").Value = _
            "CDRJ"

        UltimaLinha = UltimaLinha + 1

    Next i

    '=========================
    ' PANI SP
    '=========================
    
    Select Case CLng(frmPaniColetas.txtPanixSP.Value)
    
        Case 1
    
            ws.Cells(UltimaLinha, "D").Value = "COLETA PANIFRESH"
            ws.Cells(UltimaLinha, "E").Value = "JC"
            ws.Cells(UltimaLinha, "F").Value = "PANIFRESH"
            ws.Cells(UltimaLinha, "G").Value = DataFat
            ws.Cells(UltimaLinha, "H").Value = "PANIJC" & SufixoDia & "1"
            ws.Cells(UltimaLinha, "I").Value = "CDJC"
    
            UltimaLinha = UltimaLinha + 1
    
        Case Is >= 2
    
            ' PANIJC 1
            ws.Cells(UltimaLinha, "D").Value = "COLETA PANIFRESH"
            ws.Cells(UltimaLinha, "E").Value = "JC"
            ws.Cells(UltimaLinha, "F").Value = "PANIFRESH"
            ws.Cells(UltimaLinha, "G").Value = DataFat
            ws.Cells(UltimaLinha, "H").Value = "PANIJC" & SufixoDia & "1"
            ws.Cells(UltimaLinha, "I").Value = "CDJC"
    
            UltimaLinha = UltimaLinha + 1
    
            ' PANIJC 2
            ws.Cells(UltimaLinha, "D").Value = "COLETA PANIFRESH"
            ws.Cells(UltimaLinha, "E").Value = "JC"
            ws.Cells(UltimaLinha, "F").Value = "PANIFRESH"
            ws.Cells(UltimaLinha, "G").Value = DataFat
            ws.Cells(UltimaLinha, "H").Value = "PANIJC" & SufixoDia & "2"
            ws.Cells(UltimaLinha, "I").Value = "CDJC"
    
            UltimaLinha = UltimaLinha + 1
    
            ' PANIPRO a partir da 3ª
            For i = 1 To CLng(frmPaniColetas.txtPanixSP.Value) - 2
    
                ws.Cells(UltimaLinha, "D").Value = "COLETA PANIFRESH"
                ws.Cells(UltimaLinha, "E").Value = "JC"
                ws.Cells(UltimaLinha, "F").Value = "PANIFRESH"
                ws.Cells(UltimaLinha, "G").Value = DataFat
                ws.Cells(UltimaLinha, "H").Value = "PANIPRO" & SufixoDia & i
                ws.Cells(UltimaLinha, "I").Value = "CDJC"
    
                UltimaLinha = UltimaLinha + 1
    
            Next i
    
    End Select

    'MsgBox _
        "Etapa 3A concluída com sucesso!", _
        vbInformation, _
        "Panifresh e Coletas"

End Sub

Public Sub Etapa3BPaniColetas()

    Dim ws As Worksheet

    Dim UltimaLinha As Long
    Dim SufixoDia As String
    Dim DataFat As String

    Dim i As Long

    Set ws = wb3.Sheets("Planilha1")

    DataFat = _
        Format( _
            CDate(frmPaniColetas.txtDataFat.Value), _
            "dd/mm/yy")

    SufixoDia = frmPaniColetas.PrefixoDiaSemana

    UltimaLinha = _
        ws.Cells(ws.Rows.Count, "D").End(xlUp).Row + 1

    '=========================
    ' BOM SABOR
    '=========================

    For i = 1 To CLng(frmPaniColetas.txtColBomSabor.Value)

        ws.Cells(UltimaLinha, "D").Value = "COLETA BOM SABOR"
        ws.Cells(UltimaLinha, "E").Value = "FT"
        ws.Cells(UltimaLinha, "F").Value = "CDFT"
        ws.Cells(UltimaLinha, "G").Value = DataFat
        ws.Cells(UltimaLinha, "H").Value = "COLBOMSABOR" & SufixoDia & i
        ws.Cells(UltimaLinha, "I").Value = "BOM SABOR"

        UltimaLinha = UltimaLinha + 1

    Next i

    For i = 1 To CLng(frmPaniColetas.txtEntBomSabor.Value)

        ws.Cells(UltimaLinha, "D").Value = "ENTREGA BOM SABOR"
        ws.Cells(UltimaLinha, "E").Value = "FT"
        ws.Cells(UltimaLinha, "F").Value = "CDFT"
        ws.Cells(UltimaLinha, "G").Value = DataFat
        ws.Cells(UltimaLinha, "H").Value = "ENTBOMSABOR" & SufixoDia & i
        ws.Cells(UltimaLinha, "I").Value = ""

        UltimaLinha = UltimaLinha + 1

    Next i

    '=========================
    ' CELLIER
    '=========================

    For i = 1 To CLng(frmPaniColetas.txtColCellier.Value)

        ws.Cells(UltimaLinha, "D").Value = "COLETA CELLIER"
        ws.Cells(UltimaLinha, "E").Value = "FT"
        ws.Cells(UltimaLinha, "F").Value = "CDFT"
        ws.Cells(UltimaLinha, "G").Value = DataFat
        ws.Cells(UltimaLinha, "H").Value = "COLCELLIER" & SufixoDia & i
        ws.Cells(UltimaLinha, "I").Value = "CELLIER"

        UltimaLinha = UltimaLinha + 1

    Next i

    For i = 1 To CLng(frmPaniColetas.txtEntCellier.Value)

        ws.Cells(UltimaLinha, "D").Value = "ENTREGA CELLIER"
        ws.Cells(UltimaLinha, "E").Value = "FT"
        ws.Cells(UltimaLinha, "F").Value = "CDFT"
        ws.Cells(UltimaLinha, "G").Value = DataFat
        ws.Cells(UltimaLinha, "H").Value = "ENTCELLIER" & SufixoDia & i
        ws.Cells(UltimaLinha, "I").Value = ""

        UltimaLinha = UltimaLinha + 1

    Next i

    '=========================
    ' DOHLER
    '=========================

    For i = 1 To CLng(frmPaniColetas.txtColDohler.Value)

        ws.Cells(UltimaLinha, "D").Value = "COLETA DOHLER"
        ws.Cells(UltimaLinha, "E").Value = "FT"
        ws.Cells(UltimaLinha, "F").Value = "CDFT"
        ws.Cells(UltimaLinha, "G").Value = DataFat
        ws.Cells(UltimaLinha, "H").Value = "COLDOHLER" & SufixoDia & i
        ws.Cells(UltimaLinha, "I").Value = "DOHLER"

        UltimaLinha = UltimaLinha + 1

    Next i

    For i = 1 To CLng(frmPaniColetas.txtEntDohler.Value)

        ws.Cells(UltimaLinha, "D").Value = "ENTREGA DOHLER"
        ws.Cells(UltimaLinha, "E").Value = "FT"
        ws.Cells(UltimaLinha, "F").Value = "CDFT"
        ws.Cells(UltimaLinha, "G").Value = DataFat
        ws.Cells(UltimaLinha, "H").Value = "ENTDOHLER" & SufixoDia & i
        ws.Cells(UltimaLinha, "I").Value = ""

        UltimaLinha = UltimaLinha + 1

    Next i

    '=========================
    ' KERRY
    '=========================

    For i = 1 To CLng(frmPaniColetas.txtColKerry.Value)

        ws.Cells(UltimaLinha, "D").Value = "COLETA KERRY"
        ws.Cells(UltimaLinha, "E").Value = "FT"
        ws.Cells(UltimaLinha, "F").Value = "CDFT"
        ws.Cells(UltimaLinha, "G").Value = DataFat
        ws.Cells(UltimaLinha, "H").Value = "COLKERRY" & SufixoDia & i
        ws.Cells(UltimaLinha, "I").Value = "KERRY"

        UltimaLinha = UltimaLinha + 1

    Next i

    For i = 1 To CLng(frmPaniColetas.txtEntKerry.Value)

        ws.Cells(UltimaLinha, "D").Value = "ENTREGA KERRY"
        ws.Cells(UltimaLinha, "E").Value = "FT"
        ws.Cells(UltimaLinha, "F").Value = "CDFT"
        ws.Cells(UltimaLinha, "G").Value = DataFat
        ws.Cells(UltimaLinha, "H").Value = "ENTKERRY" & SufixoDia & i
        ws.Cells(UltimaLinha, "I").Value = ""

        UltimaLinha = UltimaLinha + 1

    Next i

    '=========================
    ' OMAMORI
    '=========================

    For i = 1 To CLng(frmPaniColetas.txtColOmamori.Value)

        ws.Cells(UltimaLinha, "D").Value = "COLETA OMAMORI"
        ws.Cells(UltimaLinha, "E").Value = "FT"
        ws.Cells(UltimaLinha, "F").Value = "CDFT"
        ws.Cells(UltimaLinha, "G").Value = DataFat
        ws.Cells(UltimaLinha, "H").Value = "COLOMAMORI" & SufixoDia & i
        ws.Cells(UltimaLinha, "I").Value = "OMAMORI"

        UltimaLinha = UltimaLinha + 1

    Next i

    For i = 1 To CLng(frmPaniColetas.txtEntOmamori.Value)

        ws.Cells(UltimaLinha, "D").Value = "ENTREGA OMAMORI"
        ws.Cells(UltimaLinha, "E").Value = "FT"
        ws.Cells(UltimaLinha, "F").Value = "CDFT"
        ws.Cells(UltimaLinha, "G").Value = DataFat
        ws.Cells(UltimaLinha, "H").Value = "ENTOMAMORI" & SufixoDia & i
        ws.Cells(UltimaLinha, "I").Value = ""

        UltimaLinha = UltimaLinha + 1

    Next i

    '=========================
    ' POTY
    '=========================

    For i = 1 To CLng(frmPaniColetas.txtColPoty.Value)

        ws.Cells(UltimaLinha, "D").Value = "COLETA POTY"
        ws.Cells(UltimaLinha, "E").Value = "FT"
        ws.Cells(UltimaLinha, "F").Value = "CDFT"
        ws.Cells(UltimaLinha, "G").Value = DataFat
        ws.Cells(UltimaLinha, "H").Value = "COLPOTY" & SufixoDia & i
        ws.Cells(UltimaLinha, "I").Value = "POTY"

        UltimaLinha = UltimaLinha + 1

    Next i

    For i = 1 To CLng(frmPaniColetas.txtEntPoty.Value)

        ws.Cells(UltimaLinha, "D").Value = "ENTREGA POTY"
        ws.Cells(UltimaLinha, "E").Value = "FT"
        ws.Cells(UltimaLinha, "F").Value = "CDFT"
        ws.Cells(UltimaLinha, "G").Value = DataFat
        ws.Cells(UltimaLinha, "H").Value = "ENTPOTY" & SufixoDia & i
        ws.Cells(UltimaLinha, "I").Value = ""

        UltimaLinha = UltimaLinha + 1

    Next i

    '=========================
    ' TRANSFERÊNCIA JC
    '=========================

    For i = 1 To CLng(frmPaniColetas.txtTransfJC.Value)

        ws.Cells(UltimaLinha, "D").Value = "TRANSFERENCIA"
        ws.Cells(UltimaLinha, "E").Value = "FT"
        ws.Cells(UltimaLinha, "F").Value = "CDFT"
        ws.Cells(UltimaLinha, "G").Value = DataFat
        ws.Cells(UltimaLinha, "H").Value = _
            "JC" & SufixoDia & Format(i, "000")
        ws.Cells(UltimaLinha, "I").Value = "COLJC"

        UltimaLinha = UltimaLinha + 1

    Next i

    'MsgBox _
        "Etapa 3B concluída com sucesso!", _
        vbInformation, _
        "Panifresh e Coletas"

End Sub

Public Sub Etapa3CPaniColetas()

    Dim ws As Worksheet

    Dim UltimaLinha As Long
    Dim Linha As Long

    Dim LojaAtual As String
    Dim RotaAtual As String

    Set ws = wb3.Sheets("Planilha1")

    UltimaLinha = _
        ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

    For Linha = UltimaLinha To 2 Step -1

        RotaAtual = Trim(ws.Cells(Linha, "H").Value)

        If RotaAtual <> "" Then

            Select Case True

                '=========================
                ' PANIFT
                '=========================

                'Case Left(RotaAtual, 6) = "PANIFT"

                    'ws.Rows(Linha + 1).Insert

                    'ws.Range("D" & Linha & ":H" & Linha).Copy _
                        ws.Range("D" & Linha + 1)

                    'ws.Cells(Linha + 1, "I").Value = "CDFT"

                '=========================
                ' PANIJC
                '=========================

                'Case Left(RotaAtual, 6) = "PANIJC"

                    'ws.Rows(Linha + 1).Insert

                    'ws.Range("D" & Linha & ":H" & Linha).Copy _
                        ws.Range("D" & Linha + 1)

                    'ws.Cells(Linha + 1, "I").Value = "CDJC"

                '=========================
                ' PANIPRO
                '=========================

                'Case Left(RotaAtual, 7) = "PANIPRO"

                    'ws.Rows(Linha + 1).Insert

                    'ws.Range("D" & Linha & ":H" & Linha).Copy _
                        ws.Range("D" & Linha + 1)

                    'ws.Cells(Linha + 1, "I").Value = "CDJC"

                '=========================
                ' COLETAS
                '=========================

                Case Left(RotaAtual, 3) = "COL"

                    ws.Rows(Linha + 1).Insert

                    ws.Range("D" & Linha & ":H" & Linha).Copy _
                        ws.Range("D" & Linha + 1)

                    ws.Cells(Linha + 1, "I").Value = "CDFT"

                '=========================
                ' ENTREGAS
                '=========================

                Case Left(RotaAtual, 3) = "ENT"

                    ws.Rows(Linha + 1).Insert

                    ws.Range("D" & Linha & ":H" & Linha).Copy _
                        ws.Range("D" & Linha + 1)

                    ws.Cells(Linha + 1, "I").Value = "CDFT"

                '=========================
                ' TRANSFERÊNCIA
                '=========================

                Case Left(RotaAtual, 2) = "JC"

                    ws.Rows(Linha + 1).Insert

                    ws.Range("D" & Linha & ":H" & Linha).Copy _
                        ws.Range("D" & Linha + 1)

                    ws.Cells(Linha + 1, "I").Value = "CDFT"

            End Select

        End If

    Next Linha

    'MsgBox _
        "Etapa 3C concluída com sucesso!", _
        vbInformation, _
        "Panifresh e Coletas"

End Sub

Public Sub Etapa4APaniColetas()

    Dim wsProg As Worksheet
    Dim wsBase As Worksheet

    Dim UltimaLinha As Long
    Dim Linha As Long
    Dim LinhaBase As Long

    Dim ModeloRota As String
    Dim Rota As String

    Dim DataFat As Date
    Dim DataCarreg As Date
    Dim DataEntrega As Date

    Dim DiasCarreg As Long
    Dim DiasViagem As Long

    Dim HoraCarreg As Date
    Dim HoraEntrega As Date

    Set wsProg = wb3.Sheets("Planilha1")
    Set wsBase = ThisWorkbook.Sheets("PANI e COLETAS")

    UltimaLinha = _
        wsProg.Cells(wsProg.Rows.Count, "H").End(xlUp).Row

    DataFat = CDate(frmPaniColetas.txtDataFat.Value)

    For Linha = 2 To UltimaLinha

        Rota = Trim(wsProg.Cells(Linha, "H").Value)

        If Rota <> "" Then

            '=========================
            ' IDENTIFICAR MODELO
            '=========================

            Select Case True

                Case Left(Rota, 6) = "PANIMG"
                    ModeloRota = "PANIMG"

                Case Left(Rota, 6) = "PANIRJ"
                    ModeloRota = "PANIRJ"

                Case Left(Rota, 6) = "PANIFT"
                    ModeloRota = "PANIFT"

                Case Left(Rota, 6) = "PANIJC"
                    ModeloRota = "PANIJC"

                Case Left(Rota, 7) = "PANIPRO"
                    ModeloRota = "PANIPRO"

                Case Left(Rota, 11) = "COLBOMSABOR"
                    ModeloRota = "COLBOMSABOR"

                Case Left(Rota, 11) = "ENTBOMSABOR"
                    ModeloRota = "ENTBOMSABOR"

                Case Left(Rota, 10) = "COLCELLIER"
                    ModeloRota = "COLCELLIER"

                Case Left(Rota, 10) = "ENTCELLIER"
                    ModeloRota = "ENTCELLIER"

                Case Left(Rota, 9) = "COLDOHLER"
                    ModeloRota = "COLDOHLER"

                Case Left(Rota, 9) = "ENTDOHLER"
                    ModeloRota = "ENTDOHLER"

                Case Left(Rota, 8) = "COLKERRY"
                    ModeloRota = "COLKERRY"

                Case Left(Rota, 8) = "ENTKERRY"
                    ModeloRota = "ENTKERRY"

                Case Left(Rota, 10) = "COLOMAMORI"
                    ModeloRota = "COLOMAMORI"

                Case Left(Rota, 10) = "ENTOMAMORI"
                    ModeloRota = "ENTOMAMORI"

                Case Left(Rota, 7) = "COLPOTY"
                    ModeloRota = "COLPOTY"

                Case Left(Rota, 7) = "ENTPOTY"
                    ModeloRota = "ENTPOTY"

                Case Left(Rota, 2) = "JC"
                    ModeloRota = "TRANSFERENCIA"

            End Select

            '=========================
            ' PROCURAR NA TABELA
            '=========================

            LinhaBase = 0

            For LinhaBase = 2 To _
                wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row

                If Trim(wsBase.Cells(LinhaBase, "A").Value) = ModeloRota Then
                    Exit For
                End If

            Next LinhaBase

            '=========================
            ' CARREGAMENTO
            '=========================

            If Weekday(DataFat, vbSunday) = 7 Then

                DiasCarreg = Val(wsBase.Cells(LinhaBase, "F").Value)

            Else

                DiasCarreg = Val(wsBase.Cells(LinhaBase, "E").Value)

            End If

            HoraCarreg = wsBase.Cells(LinhaBase, "G").Value

            DiasViagem = Val(wsBase.Cells(LinhaBase, "H").Value)

            HoraEntrega = wsBase.Cells(LinhaBase, "I").Value

            '=========================
            ' REGRA ESPECIAL
            ' TRANSFERENCIA + SABADO
            '=========================

            If ModeloRota = "TRANSFERENCIA" _
            And Weekday(DataFat, vbSunday) = 7 Then

                HoraCarreg = TimeValue("23:00:00")

                DiasViagem = 1

                HoraEntrega = TimeValue("02:00:00")

            End If

            DataCarreg = DataFat + DiasCarreg

            DataEntrega = DataCarreg + DiasViagem

            '=========================
            ' PREENCHIMENTO
            '=========================

            wsProg.Cells(Linha, "J").Value = DataCarreg
            wsProg.Cells(Linha, "K").Value = HoraCarreg

            wsProg.Cells(Linha, "M").Value = DataEntrega
            wsProg.Cells(Linha, "N").Value = HoraEntrega

            wsProg.Cells(Linha, "U").Value = _
                wsBase.Cells(LinhaBase, "J").Value

        End If

    Next Linha

    wsProg.Columns("J").NumberFormat = "dd/mm/yyyy"
    wsProg.Columns("K").NumberFormat = "hh:mm:ss"

    wsProg.Columns("M").NumberFormat = "dd/mm/yyyy"
    wsProg.Columns("N").NumberFormat = "hh:mm:ss"

    'MsgBox _
        "Etapa 4A concluída com sucesso!", _
        vbInformation, _
        "Panifresh e Coletas"

End Sub

Public Sub Etapa4BPaniColetas()

    Dim ws As Worksheet

    Dim UltimaLinha As Long
    Dim Linha As Long

    Dim DataHoraEntrega As Date
    Dim DataHoraRetorno As Date

    Set ws = wb3.Sheets("Planilha1")

    UltimaLinha = _
        ws.Cells(ws.Rows.Count, "H").End(xlUp).Row

    For Linha = 3 To UltimaLinha

        If ws.Cells(Linha, "H").Value = _
           ws.Cells(Linha - 1, "H").Value Then

            '=========================
            ' J / K
            '=========================

            ws.Cells(Linha, "J").Value = _
                ws.Cells(Linha - 1, "J").Value

            ws.Cells(Linha, "K").Value = _
                ws.Cells(Linha - 1, "K").Value

            '=========================
            ' DATA/HORA ENTREGA
            '=========================

            DataHoraEntrega = _
                ws.Cells(Linha - 1, "M").Value + _
                ws.Cells(Linha - 1, "N").Value

            DataHoraRetorno = _
                DataHoraEntrega + _
                TimeValue("05:00:00")

            ws.Cells(Linha, "M").Value = _
                Int(DataHoraRetorno)

            ws.Cells(Linha, "N").Value = _
                DataHoraRetorno - Int(DataHoraRetorno)

            '=========================
            ' TRANSPORTADORA
            '=========================

            ws.Cells(Linha, "U").Value = _
                ws.Cells(Linha - 1, "U").Value

        End If

    Next Linha

    ws.Columns("J").NumberFormat = "dd/mm/yyyy"
    ws.Columns("K").NumberFormat = "hh:mm:ss"

    ws.Columns("M").NumberFormat = "dd/mm/yyyy"
    ws.Columns("N").NumberFormat = "hh:mm:ss"

    'MsgBox _
        "Etapa 4B concluída com sucesso!", _
        vbInformation, _
        "Panifresh e Coletas"

End Sub

Public Sub Etapa5APaniColetas()

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
        ' P = EMBARQUE1
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

    Next Linha

    ws.Columns("L").NumberFormat = _
        "dd/mm/yyyy hh:mm"

    ws.Columns("O").NumberFormat = _
        "dd/mm/yyyy hh:mm"

    'MsgBox _
        "Etapa 5A concluída com sucesso!", _
        vbInformation, _
        "Panifresh e Coletas"

End Sub

Public Sub Etapa5BPaniColetas()

    Dim ws As Worksheet

    Dim UltimaLinha As Long
    Dim Linha As Long
    Dim LinhaDestino As Long

    Set ws = wb3.Sheets("Planilha1")

    UltimaLinha = _
        ws.Cells(ws.Rows.Count, "H").End(xlUp).Row

    LinhaDestino = 2

    For Linha = 2 To UltimaLinha

        If Linha = 2 Then

            ws.Cells(LinhaDestino, "R").Value = _
                ws.Cells(Linha, "H").Value

            LinhaDestino = LinhaDestino + 1

        Else

            If ws.Cells(Linha, "H").Value <> _
               ws.Cells(Linha - 1, "H").Value Then

                ws.Cells(LinhaDestino, "R").Value = _
                    ws.Cells(Linha, "H").Value

                LinhaDestino = LinhaDestino + 1

            End If

        End If

    Next Linha

    'MsgBox _
        "Etapa 5B concluída com sucesso!", _
        vbInformation, _
        "Panifresh e Coletas"

End Sub

Public Sub Etapa5CPaniColetas()

    Dim ws As Worksheet

    Dim UltimaLinha As Long
    Dim Linha As Long

    Dim Lista As String
    Dim Rota As String

    Set ws = wb3.Sheets("Planilha1")

    UltimaLinha = _
        ws.Cells(ws.Rows.Count, "H").End(xlUp).Row

    For Linha = 2 To UltimaLinha

        Rota = _
            UCase(Trim(ws.Cells(Linha, "H").Value))

        Lista = ""

        Select Case True

            Case Left(Rota, 11) = "ENTBOMSABOR"

                Lista = _
                    "COLGR,COLJC,BOM SABOR"

            Case Left(Rota, 10) = "ENTCELLIER"

                Lista = _
                    "COLGR,COLJC,CELLIER"

            Case Left(Rota, 9) = "ENTDOHLER"

                Lista = _
                    "COLGR,COLJC,DOHLER"

            Case Left(Rota, 8) = "ENTKERRY"

                Lista = _
                    "COLGR,COLJC,KERRY"

            Case Left(Rota, 10) = "ENTOMAMORI"

                Lista = _
                    "COLGR,COLJC,OMAMORI"

            Case Left(Rota, 7) = "ENTPOTY"

                Lista = _
                    "COLGR,COLJC,POTY"

        End Select

        If Lista <> "" Then

            With ws.Cells(Linha, "I").Validation

                .Delete

                .Add _
                    Type:=xlValidateList, _
                    AlertStyle:=xlValidAlertStop, _
                    Operator:=xlBetween, _
                    formula1:=Lista

                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True

            End With

        End If

    Next Linha

    'MsgBox _
        "Etapa 5C concluída com sucesso!", _
        vbInformation, _
        "Panifresh e Coletas"

End Sub

Public Sub Etapa6PaniColetas()

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

    tbl.Name = "tbPaniColetas"

    'Tabela Clara 21
    tbl.TableStyle = "TableStyleLight21"

    'MsgBox _
        "Etapa 6 concluída com sucesso!", _
        vbInformation, _
        "Panifresh e Coletas"

End Sub

Public Sub Etapa7PaniColetas()

    Dim ws As Worksheet
    Dim UltimaLinha As Long

    Dim TotalRotas As Long
    Dim TotalLojas As Long

    Dim TemPani As Boolean
    Dim TemColetas As Boolean

    Dim NomeProgramacao As String

    Set ws = wb3.Sheets("Planilha1")

    UltimaLinha = _
        ws.Cells(ws.Rows.Count, "H").End(xlUp).Row

    TotalRotas = _
        Application.WorksheetFunction.CountA( _
        ws.Range("R2:R" & UltimaLinha))

    TotalLojas = _
        Application.WorksheetFunction.CountA( _
        ws.Range("I2:I" & UltimaLinha))

    TemPani = _
        CLng(frmPaniColetas.txtPanixMG.Value) > 0 _
        Or CLng(frmPaniColetas.txtPanixRJ.Value) > 0 _
        Or CLng(frmPaniColetas.txtPanixSP.Value) > 0

    TemColetas = _
        CLng(frmPaniColetas.txtColBomSabor.Value) > 0 _
        Or CLng(frmPaniColetas.txtEntBomSabor.Value) > 0 _
        Or CLng(frmPaniColetas.txtColCellier.Value) > 0 _
        Or CLng(frmPaniColetas.txtEntCellier.Value) > 0 _
        Or CLng(frmPaniColetas.txtColDohler.Value) > 0 _
        Or CLng(frmPaniColetas.txtEntDohler.Value) > 0 _
        Or CLng(frmPaniColetas.txtColKerry.Value) > 0 _
        Or CLng(frmPaniColetas.txtEntKerry.Value) > 0 _
        Or CLng(frmPaniColetas.txtColOmamori.Value) > 0 _
        Or CLng(frmPaniColetas.txtEntOmamori.Value) > 0 _
        Or CLng(frmPaniColetas.txtColPoty.Value) > 0 _
        Or CLng(frmPaniColetas.txtEntPoty.Value) > 0 _
        Or CLng(frmPaniColetas.txtTransfJC.Value) > 0

    If TemPani And TemColetas Then

        NomeProgramacao = _
            "Panifresh e Coletas"

    ElseIf TemPani Then

        NomeProgramacao = _
            "Panifresh"

    Else

        NomeProgramacao = _
            "Coletas"

    End If

    wb3.Save

    MsgBox _
        "Programação " & NomeProgramacao & _
        " criada com sucesso!" & _
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

