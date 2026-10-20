Attribute VB_Name = "Módulo4"
Public caminho_modelo_consolidado As String
Public novoNome As String
Public wb3 As Workbook
Public ult_linha As Integer
Public cd_fat As String
Public cd_fat2 As String
Public data_fat As String
Public data_fat_NE As String
Public sem_fat As Integer
Public linha_origem As Long
Public wbTransitTime As Workbook
Public linha_FS As Long

Public Sub Etapa1()

Application.DisplayAlerts = False

'cd_fat = ComboBox1.Value

'sem_fat = Sheets("Planilha2").Cells(2, 12).Value

'cd_fat2 = Right(cd_fat, 2)

'caminho_modelo_consolidado = PortfolioPath("modelo_consolidado.xlsm")
caminho_modelo_consolidado = PortfolioPath("backup modelo_consolidado\modelo_consolidado.xlsm")

Call AtualizandoCalendarioModeloConsolidado(caminho_modelo_consolidado)

'Set wb1 = ThisWorkbook
'Set ws2 = wb1.Sheets("Planilha2")

Set wb2 = Workbooks.Open(caminho_modelo_consolidado)

' Substituindo "/" por "." nas datas
nova_data_fat = Replace(data_fat, "/", ".")
nova_data_fat_NE = Replace(data_fat_NE, "/", ".")

' Definindo o nome do arquivo com base em cd_fat2
Dim total_caract_OUTBACK As Long
total_caract_OUTBACK = 7
Dim num_caract_OUTBACK As Long

Dim planilhaDados As Worksheet
Set planilhaDados = ThisWorkbook.Sheets("Planilha2")
Dim ult_linhaOUTBACK As Long
Dim linha_OUTBACK As Long

ult_linhaOUTBACK = planilhaDados.Range("D1048576").End(xlUp).Row

For linha_OUTBACK = 2 To ult_linhaOUTBACK

    num_caract_OUTBACK = Len(planilhaDados.Cells(linha_OUTBACK, "D").Value)
    
    If num_caract_OUTBACK <> total_caract_OUTBACK Then
    
        Exit For
    
    Else
    
    End If

Next

If cd_fat2 = "JC" And num_caract_OUTBACK = 7 Then

    nome_arquivo = "Base OUTBACK - " & nova_data_fat & ".xlsm"

Else

    If cd_fat2 = "NE" Then
        nome_arquivo = "Base " & cd_fat2 & " - " & nova_data_fat_NE & ".xlsm"
    Else
        nome_arquivo = "Base " & cd_fat2 & " - " & nova_data_fat & ".xlsm"
    End If

End If

' Caminho completo para salvar o arquivo
caminho_consolidado = PortfolioPath(nome_arquivo)

' Verificando se o arquivo já existe e incrementando o nome se necessário
i = 1
novoNome = caminho_consolidado

Do While Dir(novoNome) <> ""

    If cd_fat2 = "NE" Then

        nome_arquivo = _
            "Base " & cd_fat2 & _
            " - " & nova_data_fat_NE & _
            " (" & i & ").xlsm"

    ElseIf InStr(1, UCase(nome_arquivo), "OUTBACK", vbTextCompare) > 0 Then

        nome_arquivo = _
            "Base OUTBACK - " & _
            nova_data_fat & _
            " (" & i & ").xlsm"

    Else

        nome_arquivo = _
            "Base " & cd_fat2 & _
            " - " & nova_data_fat & _
            " (" & i & ").xlsm"

    End If

    novoNome = PortfolioPath(nome_arquivo)

    i = i + 1

Loop

' Passos 6, 7 e 8 permanecem os mesmos
wb2.SaveAs FileName:=novoNome, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
DoEvents
wb2.Close SaveChanges:=False

Application.EnableEvents = False
Set wb3 = Workbooks.Open(novoNome)
DoEvents

'Preenchendo Data de Faturamento (Referência) no arquivo CONSOLIDADO
    wb3.Sheets("Planilha1").Activate
    
    If data_fat_NE = "" Then
    
    Cells(1, 2).Value = data_fat
    
    Else
    
    Cells(1, 2).Value = data_fat_NE
    
    End If
    
    'Preenchendo Sem de Faturamento (Referência) no arquivo CONSOLIDADO
    Call PreencherC1ComValor

'1. preenchendo os dados no arquivo CONSOLIDADO - Copia e Cola

    'Coluna G (DATA FATURAMENTO)
    
    ThisWorkbook.Sheets("Planilha2").Activate
    Range("H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    wb3.Sheets("Planilha1").Activate
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
    'Coluna H (ROTA)

    ThisWorkbook.Sheets("Planilha2").Activate
    Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    wb3.Sheets("Planilha1").Activate
    Range("H2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
    'Coluna I (LOJA)

    ThisWorkbook.Sheets("Planilha2").Activate
    Range("D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    wb3.Sheets("Planilha1").Activate
    Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
    'Coluna J (DATA CARREGAMENTO)

    ThisWorkbook.Sheets("Planilha2").Activate
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    wb3.Sheets("Planilha1").Activate
    Range("J2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
   
    'Coluna K (HORA CARREGAMENTO)

    ThisWorkbook.Sheets("Planilha2").Activate
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    wb3.Sheets("Planilha1").Activate
    Range("K2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
    'Coluna M (DATA ENTREGA LOJA)

    ThisWorkbook.Sheets("Planilha2").Activate
    Range("N2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    wb3.Sheets("Planilha1").Activate
    Range("M2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
    'Coluna N (HORA ENTREGA LOJA)

    ThisWorkbook.Sheets("Planilha2").Activate
    Range("O2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    wb3.Sheets("Planilha1").Activate
    Range("N2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
    
' ===== Normalização da Coluna N (Hora Entrega Loja) para hh:mm:ss =====
Dim wsPlan1 As Worksheet
Dim lastRowN As Long
Dim R As Long
Dim v As Variant
Dim hora As Date

Set wsPlan1 = wb3.Sheets("Planilha1")

' Descobrir a última linha usada na coluna N com base em dados existentes
lastRowN = wb3.Sheets("Planilha1").Cells(wb3.Sheets("Planilha1").Rows.Count, "N").End(xlUp).Row
If lastRowN < 2 Then GoTo FimNormalizaN ' nada a normalizar

' Percorre N2:N(lastRowN), normaliza e garante formato hh:mm:ss
With wsPlan1
    .Range("N2:N" & lastRowN).NumberFormat = "hh:mm:ss" ' formatação de exibição
    For R = 2 To lastRowN
        v = .Cells(R, "N").Value
        If Len(Trim$(CStr(v))) > 0 Then
            On Error Resume Next
            hora = NormalizaHora(v) ' garante Date com segundos
            If Err.Number = 0 Then
                .Cells(R, "N").Value = hora    ' valor consistente para cálculos
                ' A formatação já foi aplicada acima
            Else
                ' Se der erro (texto inválido etc.), mantemos como está
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next R
End With

FimNormalizaN:
' ===== FIM da normalização da coluna N =====

            
    'Coluna Q (CAIXAS)

    ThisWorkbook.Sheets("Planilha2").Activate
    Range("AJ2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    wb3.Sheets("Planilha1").Activate
    Range("Q2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
    
    'Coluna R (ROTAS ÚNICAS)

    'ThisWorkbook.Sheets("Planilha2").Activate
    'Range("AR2").Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.Copy
    
    'wb3.Sheets("Planilha1").Activate
    'Range("R2").Select
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        'Application.CutCopyMode = False
        
    'Coluna W (TRANSPORTADORA)

    ThisWorkbook.Sheets("Planilha2").Activate
    Range("AW2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    wb3.Sheets("Planilha1").Activate
    Range("W2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
    'Coluna R (M³ CONG)

    ThisWorkbook.Sheets("Planilha2").Activate
    Range("AY2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    wb3.Sheets("Planilha1").Activate
    Range("R2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
    'Coluna S (M³ RESF)

    ThisWorkbook.Sheets("Planilha2").Activate
    Range("AZ2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    wb3.Sheets("Planilha1").Activate
    Range("S2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
    'Coluna T (M³ SECO)

    ThisWorkbook.Sheets("Planilha2").Activate
    Range("BA2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    wb3.Sheets("Planilha1").Activate
    Range("T2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
    'Coluna U (M³ TOTAL)

    ThisWorkbook.Sheets("Planilha2").Activate
    Range("BB2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    wb3.Sheets("Planilha1").Activate
    Range("U2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
Application.DisplayAlerts = True

End Sub

Public Sub Etapa2()

Application.DisplayAlerts = False

'reabrindo o arquivo CONSOLIDADO
Set wb3 = Workbooks.Open(novoNome)
DoEvents

'2. preenchendo os dados no modelo_consolidado - Fórmulas

    wb3.Sheets("Planilha1").Activate
    
    ult_linha = Range("J1").End(xlDown).Row
    
    'Coluna L (DTH CARRG PROG)
    
    wb3.Sheets("Planilha1").Activate
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]+RC[-1]"
    Range("L2").Copy
    Range("L3:L" & ult_linha).Select
    ActiveSheet.Paste
    Range("L1").Select
    Application.CutCopyMode = False
    
    'Coluna O (DTH ENTR PROG)
    
    wb3.Sheets("Planilha1").Activate
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]+RC[-1]"
    Range("O2").Copy
    Range("O3:O" & ult_linha).Select
    ActiveSheet.Paste
    Range("O1").Select
    Application.CutCopyMode = False
    
    'Coluna P (EMBARQUE1)
    
    wb3.Sheets("Planilha1").Activate
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=RC[-11]&RC[-8]&TEXT(RC[-9],""dd/mm/aaaa"")"
    Range("P2").Copy
    Range("P3:P" & ult_linha).Select
    ActiveSheet.Paste
    Range("P1").Select
    Application.CutCopyMode = False
    
'2.1 analisando ORIGEM

    'Coluna F (ORIGEM)
    
        Dim lista_cds_origem As Variant
        
        lista_cds_origem = Array("BA", "CE", "DF", "ES", "FT", "GO", "GR", "JC", "JD", "MG", "NE", "PI", "PR", "RJ", "RS")
    
    For linha_origem = 2 To ult_linha
    
        If cd_fat <> "CDFT" And Not IsError(Application.Match(Left(Cells(linha_origem, 8).Value, 2), lista_cds_origem, 0)) Then
       
        cd_origem = "CD" & Left(Cells(linha_origem, 8).Value, 2)
        
        Else
        
        cd_origem = cd_fat
        
            'verificando se é de JD
            If Left(Cells(linha_origem, 8).Value, 2) = "CP" Or Left(Cells(linha_origem, 8).Value, 2) = "PC" Or Left(Cells(linha_origem, 8).Value, 2) = "NT" Or _
            (Left(Cells(linha_origem, 8).Value, 2) = "SO" And Right(Cells(linha_origem, 8).Value, 1) = "F") Then
            
            cd_origem = "CDJD"
            
            GoTo verificaNT
            
            Else
            
            End If
        
        End If
        
verificaNT:
        'verificando NT (FS) - SGT
            
        If Left(Cells(linha_origem, 8).Value, 2) = "NT" And Right(Cells(linha_origem, 8).Value, 1) = "F" Then
            
        cd_origem = "CDGR"
            
        GoTo preencher
            
        Else
        
            'Se for AP, MA ou AP - CDPI
            If Left(Cells(linha_origem, 8).Value, 2) = "AP" Or Left(Cells(linha_origem, 8).Value, 2) = "MA" Or Left(Cells(linha_origem, 8).Value, 2) = "PA" Then
            
                cd_origem = "CDPI"
            
            Else
            
                'Se for SF, SR, RL, JF ou NF - CDRJ
                
                If Left(Cells(linha_origem, 8).Value, 2) = "SF" Or Left(Cells(linha_origem, 8).Value, 2) = "SR" Or Left(Cells(linha_origem, 8).Value, 2) = "RL" Or Left(Cells(linha_origem, 8).Value, 2) = "JF" Or Left(Cells(linha_origem, 8).Value, 2) = "NF" Then
            
                    cd_origem = "CDRJ"
            
                Else
                
                    'Se for MC, IG, BS ou CS - CDMG
                    
                        If Left(Cells(linha_origem, 8).Value, 2) = "MC" Or Left(Cells(linha_origem, 8).Value, 2) = "IG" Or Left(Cells(linha_origem, 8).Value, 2) = "BS" Or Left(Cells(linha_origem, 8).Value, 2) = "CS" Then
                
                            cd_origem = "CDMG"
                
                        Else
                        
                        'Se for SC - CDSC
                    
                            'If Left(Cells(linha_origem, 8).Value, 2) = "SC" Then
                    
                                'cd_origem = "CDSC"
                    
                            'Else
                            
                            'End If
                            
                        'Se for MT, ML - CDMT
                        
                            If Left(Cells(linha_origem, 8).Value, 2) = "MT" Or Left(Cells(linha_origem, 8).Value, 2) = "ML" Then
                            
                                cd_origem = "CDMT"
                            
                            Else
                            
                                'Se for TO - CDDF
                                
                                If Left(Cells(linha_origem, 8).Value, 2) = "TO" Then
                                
                                    cd_origem = "CDDF"
                                
                                End If
                            
                            End If
                            
                            
                        End If
                
                End If
            
            End If
        
        End If
        
preencher:

        wb3.Sheets("Planilha1").Activate
        Range("F" & linha_origem).Select
        Range("F" & linha_origem).Value = cd_origem
       
    Next

Application.DisplayAlerts = True

End Sub

Public Sub Etapa3()

Application.DisplayAlerts = False

'reabrindo o arquivo CONSOLIDADO
Set wb3 = Workbooks.Open(novoNome)
DoEvents

'2.2 analisando CD FATURAMENTO

    'Coluna E (CD FATURAMENTO)
    
        'Dim lista_cds_fat As Variant
        
        'lista_cds_fat = Array("FT", "GR", "JC", "NE", "PR")

    wb3.Sheets("Planilha1").Activate
    Range("E2:E" & ult_linha).Value = cd_fat2
    
    'Coluna D (TIPO OPERACAO)
    
    wb3.Sheets("Planilha1").Activate
    
    Dim valH As String
    
    For linha_origem = 2 To ult_linha
    
        ufDistr = UCase$(Left$(CStr(Cells(linha_origem, "H").Value), 2))

        If (Cells(linha_origem, "F").Value = "CDJC" Or Cells(linha_origem, "F").Value = "CDGR") And _
           (ufDistr = "MS" Or ufDistr = "RO" Or ufDistr = "SM" Or ufDistr = "TM" Or ufDistr = "MD") Then

            ' Se for DIRETA
            Range("D" & linha_origem).Value = "DISTR " & Cells(linha_origem, 6).Value & " - DIR"
            
            ' Ajustar último caractere de H: M -> F (se já for F, mantém)
            'valH = CStr(Cells(linha_origem, "H").Value)
            'If Len(valH) > 0 Then
                'If Right$(valH, 1) = "M" Then
                    'Cells(linha_origem, "H").Value = Left$(valH, Len(valH) - 1) & "F"
                'End If
            'End If
        Else
            ' Se não for DIRETA
            Range("D" & linha_origem).Value = "DISTR " & Cells(linha_origem, 6).Value
        End If
    
    Next
    
            'DISTR OUTBACK
        If InStr(1, novoNome, "OUTBACK") > 0 Then
        
            Range("D2:D" & ult_linha).Value = "DISTR OUTBACK"
        
        End If

    'colar valores na tabela consolidada
    
    'wb3.Sheets("Planilha1").Activate
    'Range("P2:Q" & ult_linha).Select
    'Selection.Copy
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
     '   :=False, Transpose:=False
    'Application.CutCopyMode = False
    'Range("G1").Select
    
   'Arrumando horário carreg. CDJD para 23:00:00
    
    'wb3.Sheets("Planilha1").Activate
    
    'For linha_origem = 2 To ult_linha
    
        'If Cells(linha_origem, 6).Value = "CDJD" Then
        
            'Cells(linha_origem, 11).Value = "23:00:00"
        
        'End If
    
    'Next
    
    'First na coluna Z
    
    wb3.Sheets("Planilha1").Activate
    
    Call EscrevendoFirst
    
    'Removendo entregas FS (exceção GR/NE/PR) - TEMPORARIAMENTE
    
    wb3.Sheets("Planilha1").Activate
    
    'Dim linha_FS As Long
    
    For linha_FS = 2 To ult_linha
        
        'cd_origem_analisado = Cells(linha_FS, 6).Value

        Dim cod As String
        Dim h As String
        
        cod = UCase$(Trim$(cd_fat2))
        h = UCase$(Trim$(Cells(linha_FS, "H").Value2))
        
        If cod <> "GR" And cod <> "JC" Then
            Call RemovendoFS(linha_FS) 'Removendo FS
        End If
        
    'Recalcula após alterações estruturais
    Call EscrevendoFirst
        
         'Fuso horário
         Call FusoHorario
         
        'Verificando DTHCARREG / Direta GO GO*96F/GO*96M
        
        If Not VerificandoDIRETAGO(linha_FS, cd_fat2) Then
            Call VerificandoDTHCARREG(linha_FS, cd_fat2)
        End If

    Next linha_FS
    
    
    'DTH CARREG NE
    'If cd_fat2 = "NE" Then
    
        'Call VerificaCarregNE
    
    'Else
        
        'DTH CARREG JC BRASIL
        'If cd_fat2 = "JC" Then
        
            'Call VerificaCarregJCBrasil
        
        'Else
        
            'DTH CARREG RS
            'If cd_fat2 = "PR" Then
            
                'Call VerificaCarregRS
            
            'End If
        
        'End If
    
   ' End If
    
    'DTH CARREG RJ
        'Dim indice As Long
        'For indice = 2 To ult_linha
            'If Trim(wb3.Sheets("Planilha1").Cells(indice, "F").Value) = "CDRJ" Then
             '   Call VerificaCarregRJ(indice)
            'End If
        'Next indice
    
    'Ajustando possível erro na data de carregamento
    
    ult_linha = Range("H1").End(xlDown).Row
    
    For linha_carregamento = 2 To ult_linha
            Dim dataTexto As String, dataTextoEntrega As String
            Dim parte1 As Integer, parte2 As Integer, ano As Integer
            Dim DataCarregConvertida As Date, dataEntregaConvertida As Date
            Dim novaData As String
        
            ' Data de carregamento (coluna J)
            dataTexto = Trim(Cells(linha_carregamento, "J").Value)
        
            If Len(dataTexto) = 8 Then
                parte1 = Val(Left(dataTexto, 2))
                parte2 = Val(Mid(dataTexto, 4, 2))
                ano = Val(Right(dataTexto, 2))
        
                On Error Resume Next
                DataCarregConvertida = DateSerial(2000 + ano, parte2, parte1)
                On Error GoTo 0
            End If
        
            ' Data de entrega (coluna M)
            dataTextoEntrega = Trim(Cells(linha_carregamento, "M").Value)
        
            If Len(dataTextoEntrega) = 8 Then
                parte1 = Val(Left(dataTextoEntrega, 2))
                parte2 = Val(Mid(dataTextoEntrega, 4, 2))
                ano = Val(Right(dataTextoEntrega, 2))
        
                On Error Resume Next
                dataEntregaConvertida = DateSerial(2000 + ano, parte2, parte1)
                On Error GoTo 0
            End If
        
            ' Se a data de carregamento for maior que a de entrega, inverte dia e mês
            If DataCarregConvertida > dataEntregaConvertida Then
                parte1 = Val(Left(dataTexto, 2))
                parte2 = Val(Mid(dataTexto, 4, 2))
                ano = Val(Right(dataTexto, 2))
                DataCarregConvertida = DateSerial(2000 + ano, parte1, parte2)
            End If
        
            ' Reescreve como texto no formato dd/mm/aa
            novaData = Format(DataCarregConvertida, "dd/mm/yy")
            With Cells(linha_carregamento, "J")
                .NumberFormat = "@"
                .Value = novaData
            End With
    Next linha_carregamento
    
    ' >>> Retorno para o CD na Tbl_Prog <<<

    Dim ehUltimaLoja As Boolean
    Dim ws As Worksheet
    
    Set ws = wb3.Sheets(1) ' ou a planilha correta do seu fluxo
    linha_origem = 3
    
    Do While linha_origem <= ult_linha
        If UCase$(Trim$(ws.Cells(linha_origem, "W").Value)) = "MARTIN BROWER" _
            And UCase$(Trim$(ws.Cells(linha_origem, "D").Value)) <> "DISTR OUTBACK" Then
            
            ' Decide se a linha atual é o FIM do bloco MARTIN BROWER (olhando a próxima)
            Call VerificarUltimaLoja(ws, linha_origem, ehUltimaLoja, ult_linha)
            
            If ehUltimaLoja Then
                ' Insere retorno DEPOIS da linha atual
                Call RetornoCDTblProg(ws, linha_origem, True)
                
                ' Atualiza ult_linha pois foi inserida uma nova linha
                ult_linha = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
                
                ' Pula a linha recém inserida
                linha_origem = linha_origem + 2
            Else
                linha_origem = linha_origem + 1
            End If
            
        Else
            linha_origem = linha_origem + 1
        End If
    Loop
    
    'Fechando Transit Time
    'If Not wbTransitTime Is Nothing Then
     '   wbTransitTime.Close SaveChanges:=False
    'End If
    

    ' formatando como tabela

    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$D$1:$X$" & ult_linha), , xlYes).Name = _
        "Tabela1"
    Range("Tabela1[#All]").Select
    ActiveSheet.ListObjects("Tabela1").TableStyle = "TableStyleLight15"
    
    Range("A1").Select
    
    Call Tratamentos
    
    'Rotas únicas
    Call RotasUnicas
    
Application.EnableEvents = True
Application.DisplayAlerts = True

End Sub
Sub DesativarAtualizacoes()
    Application.Visible = False
    Application.ScreenUpdating = False
End Sub

Sub AtivarAtualizacoes()
    Application.Visible = True
    Application.ScreenUpdating = True
End Sub

Sub ExcluirArquivosExpirados()
    Dim pasta As String
    Dim Arquivo As String
    Dim caminhoCompleto As String
    Dim dataUltimaAlteracao As Date
    Dim agora As Date

    agora = Now
    pasta = PortfolioPath("") & "\"

    ' Analisa somente as bases geradas, incluindo cópias numeradas.
    Arquivo = Dir(pasta & "Base *.xlsm")

    Do While Arquivo <> ""
        caminhoCompleto = pasta & Arquivo

        dataUltimaAlteracao = FileDateTime(caminhoCompleto)

        If DateDiff("h", dataUltimaAlteracao, agora) > 24 Then
            On Error Resume Next
            Kill caminhoCompleto
            On Error GoTo 0
        End If

        Arquivo = Dir()
    Loop
    
    Call ExcluirTemplatesInfolog_14Dias

End Sub

Sub ExcluirTemplatesInfolog_14Dias()

    Dim pasta As String
    Dim Arquivo As String
    Dim caminhoCompleto As String
    Dim dataCriacao As Date
    Dim agora As Date
    Dim nomeArquivo As String

    agora = Now
    pasta = PortfolioPath("Templates\")

    Arquivo = Dir(pasta & "*.xlsm")

    Do While Arquivo <> ""
        nomeArquivo = LCase(Arquivo)
        caminhoCompleto = pasta & Arquivo

        ' Ignorar o modelo
        If nomeArquivo <> "modelo_template.xlsm" Then

            dataCriacao = FileDateTime(caminhoCompleto)

            ' Validade: 14 dias
            If DateDiff("d", dataCriacao, agora) > 14 Then
                On Error Resume Next
                Kill caminhoCompleto
                On Error GoTo 0
            End If

        End If

        Arquivo = Dir
    Loop

End Sub

Sub PreencherC1ComValor()
    Dim wsModelo As Worksheet
    Dim wsCalendario As Worksheet
    Dim valorProcurado As Variant
    Dim resultado As Variant
    Dim col As Integer
    
    'reabrindo o arquivo CONSOLIDADO
    'Set wb3 = Workbooks.Open(novoNome)
    'DoEvents
    
    'reabrindo o arquivo CONSOLIDADO
    Set wb3 = Workbooks.Open(novoNome)
    
    'Verifica se novoNome NÃO contém nenhuma das palavras-chave
    If InStr(1, novoNome, "CROSS DOCKING", vbTextCompare) = 0 And _
       InStr(1, novoNome, "BATATA", vbTextCompare) = 0 And _
       InStr(1, novoNome, "PANI", vbTextCompare) = 0 And _
       InStr(1, novoNome, "COLETAS", vbTextCompare) = 0 Then
        
        DoEvents
        
    End If

    ' Define as planilhas dentro do wb3
    Set wsModelo = wb3.Sheets("Planilha1")
    Set wsCalendario = wb3.Sheets("CALENDARIO")

    ' Pega o valor de B1 no arquivo wb3
    valorProcurado = wsModelo.Range("B1").Value
    resultado = ""

    ' Verifica da coluna 2 (B) até a coluna 8 (H) na planilha CALENDARIO
    For col = 2 To 8
        On Error Resume Next
        resultado = Application.WorksheetFunction.XLookup(valorProcurado, wsCalendario.Columns(col), wsCalendario.Columns(1), "")
        On Error GoTo 0

        If Not IsEmpty(resultado) And resultado <> "" Then
            Exit For
        End If
    Next col

    ' Escreve o resultado diretamente em C1 no arquivo wb3
    wsModelo.Range("C1").Value = resultado
    
    'Formula em C1: =PROCX(B1;CALENDARIO!B:B;CALENDARIO!A:A;
    'PROCX(B1;CALENDARIO!C:C;CALENDARIO!A:A;
    'PROCX(B1;CALENDARIO!D:D;CALENDARIO!A:A;
    'PROCX(B1;CALENDARIO!E:E;CALENDARIO!A:A;
    'PROCX(B1;CALENDARIO!F:F;CALENDARIO!A:A;
    'PROCX(B1;CALENDARIO!G:G;CALENDARIO!A:A;
    'PROCX(B1;CALENDARIO!H:H;CALENDARIO!A:A;"")))))))
End Sub

Sub Tratamentos()
    
        Call CompararDatasProg
        Call ValoresDuplicados
        
        'For linha_origem = 2 To ult_linha
        
         'If (Cells(linha_origem, "F").value = "CDFT" Or Cells(linha_origem, "F").value = "CDJD") And _
    Cells(linha_origem, "W").value = "MARTIN BROWER" Then
         
             'Call FormatarRetornosLoja
         
         'End If
        
        'Next
        
    'Alerta Carregamento > First em 24H
    Call AlertaCarregamento
    
    'Limpando dados das linhas excedentes
    Dim linha_referencia As Long
    
    linha_referencia = ult_linha + 1
    
    Rows(linha_referencia & ":" & 1048576).ClearContents
    
End Sub

Sub ValoresDuplicados()
    Dim ult_linha_tratamento As Long
    Dim cel As Range
    Dim intervalo As Range
    Dim dict As Object
    
    ' Aplica cor nas colunas D:X como formatação condicional
     With Range("D3:X" & ult_linha)
       .FormatConditions.Add Type:=xlExpression, formula1:="=NÚM.CARACT($I3)=4"
       .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(228, 158, 221) ' #E49EDD
     End With

    ' Define a última linha da coluna I
    ult_linha_tratamento = Cells(Rows.Count, "I").End(xlUp).Row
    Set intervalo = Range("I2:I" & ult_linha_tratamento)

    ' Aplica formatação condicional para valores duplicados
    With intervalo
        '.FormatConditions.Delete
        .FormatConditions.AddUniqueValues
        '.FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(.FormatConditions.Count).DupeUnique = xlDuplicate
        With .FormatConditions(.FormatConditions.Count).Font
            .Color = -16383844
            .TintAndShade = 0
        End With
        With .FormatConditions(.FormatConditions.Count).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13551615
            .TintAndShade = 0
        End With
        .FormatConditions(.FormatConditions.Count).StopIfTrue = False
    End With

    ' Cria dicionário para contar valores
    Set dict = CreateObject("Scripting.Dictionary")

    ' Conta as ocorrências de cada valor
    For Each cel In intervalo
        If Not IsEmpty(cel.Value) Then
            If dict.Exists(cel.Value) Then
                dict(cel.Value) = dict(cel.Value) + 1
            Else
                dict.Add cel.Value, 1
            End If
        End If
    Next cel

    ' Marca como "Duplicada" na coluna X
    For Each cel In intervalo
    
        If Not IsEmpty(cel.Value) Then
            If dict(cel.Value) > 1 Then
                If InStr(cel.Offset(0, 15).Value, "Duplicada") = 0 Then
                    cel.Offset(0, 15).Value = cel.Offset(0, 15).Value & IIf(cel.Offset(0, 15).Value <> "", " | ", "") & "Duplicada"
                End If
            End If
        End If
    
    Next cel
    
    ' Remove "Duplicada" se for retorno para o CD
    For linhaRetornoCD = 2 To ult_linha
    
            If Len(Cells(linhaRetornoCD, "I").Value) = 4 Then
                ' Remove apenas a palavra "Duplicada", mantendo o restante
                Cells(linhaRetornoCD, "X").Value = Replace(Cells(linhaRetornoCD, "X").Value, " | Duplicada", "")
                Cells(linhaRetornoCD, "X").Value = Replace(Cells(linhaRetornoCD, "X").Value, "Duplicada", "") ' Caso esteja sozinha
            End If
    Next

End Sub

Sub CompararDatasProg()
    
    Dim ult_linha_tratamento As Long
    Dim i As Long
    Dim celL As Range, celO As Range, celS As Range
    Dim intervalo As Range

    ' Define a última linha com base na coluna I (ou ajuste conforme necessário)
    ult_linha_tratamento = Range("I1").End(xlDown).Row
    Set intervalo = Range("L2:L" & ult_linha_tratamento)

    ' Aplica a formatação condicional: L > O
    With intervalo
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlExpression, formula1:="=L2>O2"
        With .FormatConditions(1).Font
            .Color = -16754788
            .TintAndShade = 0
        End With
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 10284031
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
    End With

    ' Preenche a coluna X com a mensagem, se L > O
    For i = 2 To ult_linha_tratamento
        Set celL = Cells(i, "L")
        Set celO = Cells(i, "O")
        Set celX = Cells(i, "X")

        If IsDate(celL.Value) And IsDate(celO.Value) Then
            
            If celL.Value > celO.Value Then
                If InStr(celX.Value, "DTH CARRG PROG é maior do que DTH ENTR PROG") = 0 Then
                    celX.Value = celX.Value & IIf(celX.Value <> "", " | ", "") & "DTH CARRG PROG é maior do que DTH ENTR PROG"
                End If
            End If
            
        End If
    
    Next i
    
End Sub

Sub AjustandoColunas()

Dim linha_analise_cabecalho As Long
Dim ult_linha_analise_cabecalho As Long

ult_linha_analise_cabecalho = Range("A2").End(xlDown).Row

'Ajustar cabeçalhos p/ coluna E até P - ENTREGA/CX/M3/KG...

For linha_analise_cabecalho = 2 To ult_linha_analise_cabecalho

   If Cells(linha_analise_cabecalho, "A").Value = "ENTREGA" Then
   
    Range("B" & linha_analise_cabecalho & ":M" & linha_analise_cabecalho).Select
    Selection.Cut Destination:=Range("E" & linha_analise_cabecalho & ":P" & linha_analise_cabecalho)
    Range("E" & linha_analise_cabecalho & ":P" & linha_analise_cabecalho).Select
   
   Else
   
   End If

Next

'Verificar novamente o que for Dom

For linha_analise_cabecalho = 2 To ult_linha_analise_cabecalho

   If Left(Cells(linha_analise_cabecalho, "C").Value, 3) = "Dom" Then
    
    'Deslocar mais uma coluna para a direita (a partir da coluna D
    Range("D" & linha_analise_cabecalho & ":O" & linha_analise_cabecalho).Select
    Selection.Cut Destination:=Range("E" & linha_analise_cabecalho & ":P" & linha_analise_cabecalho)
    Range("E" & linha_analise_cabecalho & ":P" & linha_analise_cabecalho).Select
    
    'Separando o horário
    Range("D" & linha_analise_cabecalho).Value = Right(Cells(linha_analise_cabecalho, "C").Value, 5)

    'Formatar E até P como número sem casas decimais
    Columns("E:P").Select
    Selection.NumberFormat = "0"
   
   Else
   
   End If

Next

End Sub

Sub RemovendoFS(linha_FS As Long)

    Set wb3 = Workbooks.Open(novoNome)
    DoEvents
    Dim ws As Worksheet
    Set ws = wb3.Sheets(1)

    Dim valorCelula As String
    valorCelula = ws.Cells(linha_FS, 9).Value ' Coluna I = 9

    If Len(valorCelula) = 7 And Mid(valorCelula, 4, 1) = "-" Then
        
        ws.Rows(linha_FS).Delete
        linha_FS = linha_FS - 1
        
    End If

End Sub

Sub VerificandoTransportadora()

Dim wb1 As Workbook
Dim wsPrincipal As Worksheet

Set wb1 = ThisWorkbook
Set wsPrincipal = wb1.Sheets("Planilha1")

wsPrincipal.Activate

ult_linha_geral = Range("A2").End(xlDown).Row

'Verificando as Transportadoras e preenchendo na BASE DE PROGRAMAÇÃO
For linha_analise = 2 To ult_linha_geral

    wsPrincipal.Activate

    If Cells(linha_analise, "J").Value = "Veículo:" Then
    
        transportadoraSIGLA = Left(Cells(linha_analise, "K").Value, 3)
        
        'Preenchendo na BASE DE PROGRAMAÇÃO
        
        wb1.Sheets("Planilha2").Activate
        
        linha_preenchimento_transportadoraSIGLA = Range("AX1048576").End(xlUp).Row + 1
        
        Cells(linha_preenchimento_transportadoraSIGLA, "AX").Value = transportadoraSIGLA
                
    Else
    
    End If

Next

'Preenchendo Transportadora
Call PreenchendoTransportadora
        

End Sub
Sub PreenchendoTransportadora()

    Dim wb1 As Workbook
    Dim ws As Worksheet
    Dim ult_linha As Long, Linha As Long
    Dim valorProcurado As Variant, sigla As Variant
    Dim transportadora As String

    Set wb1 = ThisWorkbook
    Set ws = wb1.Sheets("Planilha2")

    ult_linha = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    For Linha = 2 To ult_linha
        valorProcurado = ws.Cells(Linha, "C").Value

        If Not IsEmpty(valorProcurado) Then
            On Error Resume Next
            sigla = Application.WorksheetFunction.XLookup(valorProcurado, ws.Range("AR:AR"), ws.Range("AX:AX"))
            On Error GoTo 0

            If Not IsError(sigla) And Not IsEmpty(sigla) Then
                Select Case sigla
                    Case "MBR": transportadora = "MARTIN BROWER"
                    Case "MB-": transportadora = "MARTIN BROWER"
                    Case "MOD": transportadora = "MODERN"
                    Case "TRI": transportadora = "TRIFIT"
                    Case "PRO": transportadora = "PRODELOG"
                    Case "SGT": transportadora = "SGT"
                    Case "YES": transportadora = "YES BRAZIL"
                    Case "LOG": transportadora = "LOGMAM"
                    Case Else: transportadora = "NÃO IDENTIFICADA"
                End Select

                ws.Cells(Linha, "AW").Value = transportadora
            End If
        End If
    Next Linha

End Sub

Sub RotasUnicas()

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim ws As Worksheet
    Set ws = wb3.Sheets(1) ' Altere se necessário
    
    Dim UltimaLinha As Long
    UltimaLinha = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To UltimaLinha
        Dim Valor As Variant
        Valor = ws.Cells(i, "H").Value
        
        If Not IsEmpty(Valor) Then
            If Not dict.Exists(Valor) Then
                dict.Add Valor, Nothing
            End If
        End If
    Next i
    
    ' Escreve os valores únicos na coluna V a partir de V2
    Dim LinhaDestino As Long
    LinhaDestino = 2
    
    Dim Chave As Variant
    For Each Chave In dict.Keys
        ws.Cells(LinhaDestino, "V").Value = Chave
        LinhaDestino = LinhaDestino + 1
    Next Chave

End Sub

Public Sub RetornoCDTblProg(ByVal ws As Worksheet, ByVal linha_origem As Long, _
                            ByVal ehUltimaLoja As Boolean)
    Dim linhaVerifRetorno As Long
    Dim horaRetornoProg As Double
    Dim arrCodigos As Variant
    Dim NomenclaturaRota As String
    
    linhaVerifRetorno = linha_origem
    
    ' Inserir nova linha
    If ehUltimaLoja Then
        ' insere DEPOIS da linha atual
        ws.Rows(linhaVerifRetorno + 1).Insert Shift:=xlDown
        linhaVerifRetorno = linhaVerifRetorno + 1
    Else
        ' (mantido para compatibilidade) insere ANTES da linha atual
        ws.Rows(linhaVerifRetorno).Insert Shift:=xlDown
        ' linhaVerifRetorno permanece apontando para a linha inserida
    End If
    
    ' Copiar D:P e W da linha anterior (fórmulas e formatos)
    ws.Range("D" & linhaVerifRetorno - 1 & ":P" & linhaVerifRetorno - 1).Copy
    ws.Range("D" & linhaVerifRetorno & ":P" & linhaVerifRetorno). _
        PasteSpecial Paste:=xlPasteFormulasAndNumberFormats
    
    ws.Range("W" & linhaVerifRetorno - 1 & ":W" & linhaVerifRetorno - 1).Copy
    ws.Range("W" & linhaVerifRetorno & ":W" & linhaVerifRetorno). _
        PasteSpecial Paste:=xlPasteFormulasAndNumberFormats
    
    Application.CutCopyMode = False
    
    ' Garantir que a linha inserida permaneça na mesma rota (H)
    ws.Cells(linhaVerifRetorno, "H").Value = ws.Cells(linhaVerifRetorno - 1, "H").Value
    
    ' Ajustes específicos da linha de retorno
    ws.Range("I" & linhaVerifRetorno).Value = ws.Range("F" & linhaVerifRetorno).Value ' LOJA = ORIGEM
    ws.Range("Q" & linhaVerifRetorno & ":U" & linhaVerifRetorno).Value = 0           ' Zera caixas e m³
    
    ' Hora Retorno CD
    horaRetornoProg = HoraRetornoCDTblProg(ws, linhaVerifRetorno) ' deve retornar fração de dia
    ws.Range("N" & linhaVerifRetorno).NumberFormat = "hh:mm:ss"
    ws.Range("N" & linhaVerifRetorno).Value = horaRetornoProg
    
    ' Lista de códigos que forçam rolover de data para +1 dia
    arrCodigos = Array("AL", "CR", "JF", "NF", "NO", "NP", "NT", "PA", "PB", "RL", "RN", "RP", "SC", "SE", "SF", "SJ", "SR")
    NomenclaturaRota = UCase$(Trim$(Left$(CStr(ws.Cells(linhaVerifRetorno, "H").Value), 2)))
    
    ' Ajuste da data (coluna M) como texto "dd/mm/yy"
    Dim DataHoraAnterior As Date
    Dim DataHoraRetorno As Date
    
    DataHoraAnterior = _
        CDate(ws.Cells(linhaVerifRetorno - 1, "M").Value) + _
        ws.Cells(linhaVerifRetorno - 1, "N").Value
    
    DataHoraRetorno = _
        DataHoraAnterior + TimeValue("05:00:00")
    
    ws.Cells(linhaVerifRetorno, "M").NumberFormat = "@"
    
    If InArray(NomenclaturaRota, arrCodigos) Then
    
        ws.Cells(linhaVerifRetorno, "M").Value = _
            Format$(CDate(ws.Cells(linhaVerifRetorno - 1, "M").Value) + 1, "dd/mm/yy")
    
    Else
    
        ws.Cells(linhaVerifRetorno, "M").Value = _
            Format$(Int(DataHoraRetorno), "dd/mm/yy")
    
    End If
End Sub

Sub CalculoOffLoad(ws As Worksheet, wsTransitTime As Worksheet, linhaVerifRetorno As Long, ByRef tempoDesc As Date)

    Dim valorBusca As Variant
    Dim caixasPorHora As Double
    Dim qtdCaixas As Double
    Dim celulaBusca As Range
    Dim tempoDecimal As Double

    ' Valor a ser buscado na coluna B da planilha TransitTime
    valorBusca = ws.Cells(linhaVerifRetorno - 1, 9).Value ' Coluna I = 9

    ' Procura o valor na coluna B
    Set celulaBusca = wsTransitTime.Columns(2).Find(What:=valorBusca, LookIn:=xlValues, LookAt:=xlWhole)

    If Not celulaBusca Is Nothing Then
        ' Pega o valor da coluna O (coluna 15) da mesma linha
        caixasPorHora = celulaBusca.Offset(0, 13).Value ' Coluna O está 13 à frente da B

        ' Pega a quantidade de caixas da coluna Q (coluna 17) da linha atual
        qtdCaixas = ws.Cells(linhaVerifRetorno - 1, 17).Value

        ' Calcula o tempo em horas decimais
        tempoDecimal = qtdCaixas / caixasPorHora

        ' Converte para formato de hora (hh:mm:ss)
        tempoDesc = tempoDecimal / 24
    Else
        'MsgBox "Valor não encontrado na planilha TransitTime para cálculo de tempo de descarga.", vbExclamation
        tempoDesc = 0
    End If

End Sub

Sub ArredondarJanela15Min(ByRef horaRef As Double)
    
    Dim totalMinutos As Long
    
    'Converte o valor total da hora para minutos (incluindo dias excedentes)
    totalMinutos = horaRef * 24 * 60
    
    'Arredonda para o próximo múltiplo de 15 minutos
    totalMinutos = ((totalMinutos + 14) \ 15) * 15
    
    'Converte de volta para fração de dia
    horaRef = totalMinutos / (24 * 60)
    
End Sub

Public Sub VerificarUltimaLoja(ByVal ws As Worksheet, ByVal linhaAtual As Long, _
                               ByRef ehUltimaLoja As Boolean, ByVal ult_linha As Long)
    Dim wProx As String
    Dim hAtual As String, hProx As String
    Dim rotaContinua As Boolean
    
    ' Se já está na última linha conhecida, é fim de bloco
    If linhaAtual >= ult_linha Then
        ehUltimaLoja = True
        Exit Sub
    End If
    
    hAtual = Trim$(CStr(ws.Cells(linhaAtual, "H").Value))
    hProx = Trim$(CStr(ws.Cells(linhaAtual + 1, "H").Value))
    wProx = UCase$(Trim$(ws.Cells(linhaAtual + 1, "W").Value))
    
    ' A rota só continua se H_prox == H_atual e não está vazio
    rotaContinua = (hProx <> "") And (hProx = hAtual)
    
    ' É fim do bloco MARTIN BROWER quando:
    ' 1) a rota não continua (mudou H ou ficou vazio), OU
    ' 2) a próxima linha não é MARTIN BROWER
    ehUltimaLoja = (Not rotaContinua) Or (wProx <> "MARTIN BROWER")
End Sub

Sub FormatarRetornosLoja()

    'Dim wb3 As Workbook
    Dim ws As Worksheet
    Dim linhaCadastro As Long
    'Dim ultLinhaCadastro As Long
    Dim intervalo As Range
    Dim formulaCondicional As String

    'Set wb3 = Workbooks("wb3.xlsx") ' ajuste conforme necessário
    wb3.Activate
    Set ws = wb3.Sheets(1)

    'ultLinhaCadastro = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

    ' Define o intervalo da coluna O
    Set intervalo = ws.Range("O3:O" & linha_origem)

    ' Fórmula traduzida para português
    formulaCondicional = "=E(O4=O3;NÃO(ÉCÉL.VAZIA(O3)))"

    ' Adiciona a formatação condicional sem prioridade
    'With intervalo.FormatConditions.Add(Type:=xlExpression, formula1:=formulaCondicional)
     '   .Interior.Color = RGB(233, 113, 50)
      '  .StopIfTrue = False
        ' .SetFirstPriority foi removido
    'End With

    ' Preenche a coluna X com a mensagem
    'For linhaCadastro = 4 To linha_origem
     '   If ws.Cells(linhaCadastro, "O").value = ws.Cells(linhaCadastro - 1, "O").value And _
      '     ws.Cells(linhaCadastro, "O").value <> "" Then

            'ws.Cells(linhaCadastro, "T").Value = "Cadastrar retorno da loja"
       '     ws.Cells(linhaCadastro - 1, "X").value = "Cadastrar retorno da loja"
        'End If
    'Next linhaCadastro

End Sub

Function HoraRetornoCDTblProg(ws As Worksheet, linhaVerifRetorno As Long) As Double
    Dim horaDecimal As Variant
    Dim resultado As Double

    horaDecimal = ws.Range("N" & linhaVerifRetorno - 1).Value

    If IsNumeric(horaDecimal) And horaDecimal > 0 Then
        resultado = CDbl(horaDecimal) + (5 / 24)
        resultado = resultado - Int(resultado)
        HoraRetornoCDTblProg = resultado
    Else
        HoraRetornoCDTblProg = 0 ' ou use CVErr(xlErrValue) se quiser sinalizar erro
    End If
End Function

' Rotina que abre o modelo, atualiza o CALENDARIO com eventos OFF, salva,
' fecha e reabre com eventos ON para rodar o Workbook_Open do wb2.

Public Sub AtualizandoCalendarioModeloConsolidado(ByVal caminho_modelo_consolidado As String)

    Dim wb2 As Workbook
    Dim evtState As Boolean, scrState As Boolean, alertState As Boolean

    ' --- ABERTURA 1: eventos OFF para atualizar o CALENDARIO sem interferência ---
    evtState = Application.EnableEvents: Application.EnableEvents = False
    scrState = Application.ScreenUpdating: Application.ScreenUpdating = False
    alertState = Application.DisplayAlerts: Application.DisplayAlerts = False

    Set wb2 = Workbooks.Open(FileName:=caminho_modelo_consolidado, ReadOnly:=False)

    ' Atualiza o CALENDARIO de wb2 usando a mesma Sub (agora com parâmetro)
    AnalisandoCalendario wb2

    ' Se possível, salvar (não salva se estiver fisicamente read-only)
    If Not wb2.ReadOnly Then wb2.Save

    ' Fecha para reabrir com eventos ON (para executar o Workbook_Open do wb2)
    wb2.Close SaveChanges:=False

    ' Restaura estados
    Application.DisplayAlerts = alertState
    Application.ScreenUpdating = scrState
    Application.EnableEvents = evtState

    ' --- ABERTURA 2: eventos ON para executar o Workbook_Open do wb2 normalmente ---
    Set wb2 = Workbooks.Open(FileName:=caminho_modelo_consolidado)

    ' Se seu fluxo precisar continuar usando wb2 (ex.: copiar dados, consolidar, etc.),
    ' faça aqui. Exemplo:
    ' Call ProcessarPosAbertura(wb2)

    ' Salvar/fechar ao final (ajuste conforme seu fluxo)
    'If Not wb2.ReadOnly Then wb2.Save
    'wb2.Close SaveChanges:=True

End Sub

Private Function InArray(ByVal Value As String, ByVal arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = Value Then
            InArray = True
            Exit Function
        End If
    Next i
End Function

Public Sub AlertaCarregamento()

    'Cells(2, "Z").Value = "FIRST"

    'For linha_origem = 3 To ult_linha
    
        'If Cells(linha_origem, "H").Value <> Cells(linha_origem - 1, "H").Value Then
        
            'Cells(linha_origem, "Z").Value = "FIRST"
        
        'End If
    
    'Next
    
    'Carregamento MA/PA/PI
    'If cd_fat2 = "NE" Then
    
    'Call VerificaCarregNE
    
    'End If
    
    'Formatação condicional de Alerta Carregamento
    Dim ws As Worksheet
    Dim intervaloAlertaCarreg As Range
    Dim formulaCondicionalAlertaCarreg As String

    wb3.Activate
    Set ws = wb3.Sheets(1)

    ' Define o intervalo
    Set intervaloAlertaCarreg = ws.Range("D2:X" & ult_linha)

   
' Fórmula traduzida para português (ancorada em Z2)
formulaCondicionalAlertaCarreg = _
    "=E(" & _
    "CONT.SE($Z$2:$Z2;""FIRST"")>=1;" & _
    "ÍNDICE($O:$O;" & _
        "LIN($Z$2)-1+" & _
        "AGREGAR(15;6;" & _
            "(LIN($Z$2:$Z$" & ult_linha & ")-LIN($Z$2)+1)/($Z$2:$Z$" & ult_linha & "=""FIRST"");" & _
            "CONT.SE($Z$2:$Z2;""FIRST""))" & _
    ")>=" & _
    "ÍNDICE($L:$L;" & _
        "LIN($Z$2)-1+" & _
        "AGREGAR(15;6;" & _
            "(LIN($Z$2:$Z$" & ult_linha & ")-LIN($Z$2)+1)/($Z$2:$Z$" & ult_linha & "=""FIRST"");" & _
            "CONT.SE($Z$2:$Z2;""FIRST""))" & _
    ")+1" & _
    ")"


    ' Adiciona a formatação condicional
    With intervaloAlertaCarreg.FormatConditions.Add(Type:=xlExpression, formula1:=formulaCondicionalAlertaCarreg)
        .Interior.Color = RGB(255, 255, 0)
        .StopIfTrue = False
        .SetFirstPriority
    End With
    
    
' --- Preencher coluna X com o texto do alerta, por blocos entre FIRSTs ---
Dim Linha As Long, inicioBloco As Long, fimBloco As Long
Dim lData As Variant, oData As Variant
Dim textoAlerta As String
Dim pintaBloco As Boolean

textoAlerta = "Alerta: Carregamento programado para 24 horas ou mais antes da primeira entrega"

' (Opcional) limpar X antes do reprocessamento
'ws.Range("X2:X" & ult_linha).ClearContents

Linha = 2
Do While Linha <= ult_linha
    If UCase$(Trim$(CStr(ws.Cells(Linha, "Z").Value))) = "FIRST" Then
        ' início do bloco
        inicioBloco = Linha
        
        ' fim do bloco: vai até a linha anterior ao próximo FIRST (ou ult_linha)
        fimBloco = inicioBloco
        Do While fimBloco + 1 <= ult_linha And _
                 UCase$(Trim$(CStr(ws.Cells(fimBloco + 1, "Z").Value))) <> "FIRST"
            fimBloco = fimBloco + 1
        Loop
        
        ' verifica a condição SOMENTE na linha do FIRST
        lData = ws.Cells(inicioBloco, "L").Value
        oData = ws.Cells(inicioBloco, "O").Value
        
        pintaBloco = False
        If IsDate(lData) And IsDate(oData) Then
            ' 24 horas = +1 dia
            If CDate(oData) >= CDate(lData) + 1 Then
                pintaBloco = True
            End If
        End If
        
        ' se condição verdadeira, escreve texto em X para todo o bloco
        If pintaBloco Then
            ws.Range("X" & inicioBloco & ":X" & fimBloco).Value = textoAlerta
        Else
            ' caso contrário, garante X vazio no bloco (se preferir manter, remova esta linha)
            'ws.Range("X" & inicioBloco & ":X" & fimBloco).ClearContents
        End If
        
        ' avança para após o bloco
        Linha = fimBloco + 1
    Else
        Linha = Linha + 1
    End If
Loop

End Sub

Public Sub ConfigurarValidacoesWBNovo(ByVal wbNovo As Workbook)
    On Error GoTo TrataErro

    If wbNovo Is Nothing Then Err.Raise 91, , "wbNovo está Nothing."

    Dim ws1 As Worksheet, ws2 As Worksheet

    ' Sheet1 precisa existir
    On Error Resume Next
    Set ws1 = wbNovo.Worksheets("Sheet1")
    On Error GoTo 0
    If ws1 Is Nothing Then Err.Raise 9, , "A planilha 'Sheet1' não existe no wbNovo."

    ' Planilha2: cria se não existir
    On Error Resume Next
    Set ws2 = wbNovo.Worksheets("Planilha2")
    On Error GoTo 0
    If ws2 Is Nothing Then
        Set ws2 = wbNovo.Sheets.Add(After:=ws1)
        ws2.Name = "Planilha2"
        DoEvents
    End If

    ' ========= 1) Preencher D com a array =========
    Dim listaD As Variant, i As Long
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
        "Ordem de Carregamento Loja", "Paletização invertida", "POSTERGADA", "Posto Fiscal", "Problema Mecânico", _
        "Problemas de saúde do motorista", "Restrição de Acesso Loja", "Retenção de Balança", _
        "Sinistro", "Sistema / Mobile / BI", "Transbordo de carga", "Transferência/Transbordo", _
        "Trânsito", "Trânsito na Transferência" _
    )
    For i = 0 To UBound(listaD)
        ws2.Cells(i + 1, 4).Value = listaD(i)  ' D1 título; itens a partir de D2
    Next i

    ' ========= 2) Tabela em D =========
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
    DoEvents

    ' ========= 3) Construir H sem "-" (FILTRAR/FILTER) =========
    Dim temFiltro As Boolean
    temFiltro = False

    ws2.Range("H1").Value = "Motivo (sem '-')"

    ' PT-BR
    On Error Resume Next
    ws2.Range("H2").Formula2 = "=FILTRAR(tblMotivos[Motivo]; tblMotivos[Motivo]<>""-"")"
    If Err.Number = 0 Then temFiltro = True
    Err.Clear

    ' EN
    If Not temFiltro Then
        ws2.Range("H2").Formula2 = "=FILTER(tblMotivos[Motivo], tblMotivos[Motivo]<>""-"")"
        If Err.Number = 0 Then temFiltro = True
    End If
    On Error GoTo 0
    DoEvents

    ' ========= 3b) Nomes =========
    On Error Resume Next
    wbNovo.Names("Lista_R_Com_Hifen").Delete
    wbNovo.Names("Lista_R_Sem_Hifen").Delete
    On Error GoTo 0

    If temFiltro Then
        wbNovo.Names.Add Name:="Lista_R_Com_Hifen", RefersTo:="=Planilha2!tblMotivos[Motivo]"
        wbNovo.Names.Add Name:="Lista_R_Sem_Hifen", RefersTo:="=Planilha2!$H$2#"
    Else
        ' Plano B: copia D->H ignorando "-"
        Dim lastD As Long, R As Long, h As Long
        lastD = ws2.Cells(ws2.Rows.Count, "D").End(xlUp).Row
        h = 2
        For R = 2 To lastD
            If ws2.Cells(R, "D").Value <> "-" Then
                ws2.Cells(h, "H").Value = ws2.Cells(R, "D").Value
                h = h + 1
            End If
        Next R

        wbNovo.Names.Add Name:="Lista_R_Com_Hifen", _
            RefersTo:="=Planilha2!$D$2:INDEX(Planilha2!$D:$D, MATCH(""zzz"", Planilha2!$D:$D))"
        wbNovo.Names.Add Name:="Lista_R_Sem_Hifen", _
            RefersTo:="=Planilha2!$H$2:INDEX(Planilha2!$H:$H, MATCH(""zzz"", Planilha2!$H:$H))"
    End If

    ' ========= 4) Validação R (depende de L) =========
    With ws1.Range("R2:R2000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
              formula1:="=INDIRECT(IF($L2="""",""Lista_R_Com_Hifen"",""Lista_R_Sem_Hifen""))"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "Ocorrência inválida"
        .ErrorMessage = "A entrega está com HORA ADICIONAL ATRASO, então precisa ter uma justificativa."
    End With

    ' ========= 5) Validação L (hora só quando R <> "-") =========
    With ws1.Range("L2:L2000").Validation
        .Delete
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
              formula1:="=OR($L2="""", AND($R2<>""-"", ISNUMBER($L2), $L2>=0, $L2<1))"
        .IgnoreBlank = True
        .ErrorTitle = "Hora adicional sem justificativa"
        .ErrorMessage = "Não é possível inserir HORA ADICIONAL ATRASO sem uma justificativa na ocorrência de entrega."
    End With
    ws1.Range("L2:L2000").NumberFormat = "hh:mm:ss"

    ' ========= Q/AF (independente) =========
    With ws1.Range("Q2:Q2000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
              formula1:="=INDIRECT(IF($AF2="""",""Lista_R_Com_Hifen"",""Lista_R_Sem_Hifen""))"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "Ocorrência inválida"
        .ErrorMessage = "O carregamento está com HORA ADICIONAL ATRASO, então precisa ter uma justificativa."
    End With

    With ws1.Range("AF2:AF2000").Validation
        .Delete
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
              formula1:="=OR($AF2="""", AND($Q2<>""-"", ISNUMBER($AF2), $AF2>=0, $AF2<1))"
        .IgnoreBlank = True
        .ErrorTitle = "Hora adicional sem justificativa"
        .ErrorMessage = "Não é possível inserir HORA ADICIONAL ATRASO sem uma justificativa na ocorrência de carregamento."
    End With
    ws1.Range("AF2:AF2000").NumberFormat = "hh:mm:ss"

    Exit Sub

TrataErro:
    MsgBox "Falha ao configurar validações: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ----------------------------------------------------------------------
' Deixa o wbNovo compatível com Shared Workbook (legado), mantendo:
' - Listas em Planilha2 (D com "-" e H sem "-")
' - Validações dinâmicas em Sheet1 (R/L e Q/AF), sem INDIRETO
' - Fallbacks de idioma (EN/PT-BR) e de aplicação (Selection/linha-a-linha)
' ----------------------------------------------------------------------
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

    Dim lastD As Long, lastH As Long, R As Long, h As Long, valD As String
    lastD = ws2.Cells(ws2.Rows.Count, "D").End(xlUp).Row

    h = 2
    For R = 2 To lastD
        valD = Trim$(CStr(ws2.Cells(R, "D").Value))
        If valD <> "" And valD <> "-" Then
            ws2.Cells(h, "H").Value = valD
            h = h + 1
        End If
    Next R
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

' Pressupostos do seu projeto:
' - linha_FS já é Public As Long
' - wb3 já é Public As Workbook

Public Sub FusoHorario()
    Dim wsFuso As Worksheet
    Dim uf As String
    Dim horasAdicionar As Long
    
    Set wsFuso = wb3.Worksheets("Planilha1")
    
    ' 1) Ler os dois caracteres da esquerda da coluna H (rota)
    '    Atenção à ordem dos parênteses: Trim -> CStr -> UCase$ -> Left$
    uf = Left$(UCase$(Trim$(CStr(wsFuso.Cells(linha_FS, "H").Value))), 2)
    
    ' 2) Definir horas a adicionar conforme a sua regra de negócio
    Select Case uf
        Case "MD", "ML", "MT", "MS", "RO", "RR"
            horasAdicionar = 1
        Case "AC"
            horasAdicionar = 2
        Case Else
            Exit Sub   ' não ajusta se não estiver na lista
    End Select
    
    ' 3) Ajustar colunas K e N (somar horas ao que já existe)
    SomarHorasEmCelula wsFuso.Cells(linha_FS, "K"), horasAdicionar
    SomarHorasEmCelula wsFuso.Cells(linha_FS, "N"), horasAdicionar
End Sub

'----------------------------------------------------------
' Soma N horas em uma célula que contém hora.
' Aceita número serial (data/hora) ou texto "hh:mm" / "hh:mm:ss".
' Mantém formato "hh:mm:ss" se a célula estava em Geral.
'----------------------------------------------------------
Private Sub SomarHorasEmCelula(ByVal alvo As Range, ByVal horas As Long)
    Dim v As Variant
    
    If alvo Is Nothing Then Exit Sub
    If IsEmpty(alvo.Value) Then Exit Sub
    
    v = alvo.Value
    
    ' Se vier como texto, tenta converter
    If VarType(v) = vbString Then
        If Len(v) = 0 Then Exit Sub
        If Not IsDate(v) Then
            ' tenta converter via CDate; se falhar, sai
            On Error Resume Next
            v = CDate(v)
            On Error GoTo 0
            If Not IsDate(v) Then Exit Sub
        End If
    End If
    
    ' Se for data/hora ou número (serial), aplica o acréscimo
    If IsDate(v) Or IsNumeric(v) Then
        alvo.Value = CDbl(v) + (horas / 24#)
        
        ' Garante exibição como hora caso estivesse em "Geral"
        If alvo.NumberFormat = "General" Or alvo.NumberFormatLocal = "Geral" Then
            alvo.NumberFormat = "hh:mm:ss"
        End If
    End If
End Sub

' ===== Helpers =====

' Converte texto "dd/mm/aa" (ou "dd/mm/aaaa") para Date com segurança.
' Retorna True se conseguiu converter; False se a string estiver inválida.
Private Function ParseTextDateDMY(ByVal txt As String, ByRef outDate As Date) As Boolean
    On Error GoTo Falha
    Dim p() As String, d As Long, m As Long, Y As Long
    txt = Trim$(CStr(txt))
    If Len(txt) = 0 Then Exit Function

    p = Split(txt, "/")
    If UBound(p) <> 2 Then Exit Function

    d = Val(p(0)): m = Val(p(1)): Y = Val(p(2))
    If Y < 100 Then Y = 2000 + Y ' assume século 2000 (ajuste se necessário)
    outDate = DateSerial(Y, m, d)
    ParseTextDateDMY = True
    Exit Function
Falha:
    ParseTextDateDMY = False
End Function

' Determina a regra de carregamento conforme cd_fat2 (fixo), nomenclatura (H, 2 primeiras letras) e coluna 6 ("CDJD" para FT).
' Retorna True se deve aplicar a regra (dias/hora definidos). False = "Segue preliminar" ou sem regra.

' ============================================================================================
' REGRAS ATUAIS PARA DEFINIÇÃO DA DTHCARREG
'
' FAT. GR / JC
'   BA - Dia anterior à entrega, 23:00
'   DF - Dia anterior à entrega, 23:00
'   TO - Dia anterior à entrega, 23:00
'   ES - Dia anterior à entrega, 23:00
'   GO - Dia anterior à entrega, 23:00
'   LT / SP / VP - Segue preliminar (não alterar DTHCARREG)
'   NO / RP / SJ - Dia anterior à entrega, 13:00
'   GO*96F / GO*96M - 3 dias anteriores à entrega, 23:00
'   CP / PC / SO - Dia anterior à entrega, 23:00
'   MC / CS / TM / SM - Dia anterior à entrega, 23:00
'   BS - 2 Dias anteriores à entrega, 23:00
'   MG / IG - Dia anterior à entrega, 23:00
'   ML - Dia anterior à entrega, 23:00
'   MT - Dia anterior à entrega, 23:00
'   RJ - Dia anterior à entrega, 23:00
'   NF / SF / SR / RL / JF - Dia anterior à entrega, 21:00
'   RO - 3 Dias anteriores à entrega, 23:00
'
' FAT. NE
'   CE - Dia anterior à entrega, 23:00
'   PI - Dia anterior à entrega, 23:00
'   MA / PA / AP / ST - 2 dias anteriores à entrega, 23:00
'   SE - Dia anterior à entrega, 23:00
'
' FAT. PR
'   RS - Dia anterior à entrega, 23:00
'   NP / SC / PN - Dia anterior à entrega, 21:00
'
' FAT. FT
'   Se coluna 6 = "CDJD" ? Dia anterior à entrega, 23:00
'   Caso contrário: segue preliminar (não alterar DTHCARREG)
'
' OBSERVAÇÕES GERAIS:
' - "Dia anterior" = dataEntrega - 1
' - "2 dias anteriores" = dataEntrega - 2
' - Hora deve ser registrada em "hh:mm:ss"
' - Data deve permanecer no formato texto "dd/mm/aa"
' - Para cada rota, utilizar a data de entrega apenas da primeira linha (Z = "FIRST")
' ============================================================================================

Private Function LoadRule(ByVal cd_fat2 As String, ByVal nomen As String, ByVal cd6 As String, _
                          ByRef daysBefore As Long, ByRef hh As Long, ByRef mm As Long) As Boolean
    cd_fat2 = UCase$(Trim$(cd_fat2))
    nomen = UCase$(Trim$(nomen))
    cd6 = UCase$(Trim$(cd6))

    Select Case cd_fat2
        Case "GR", "JC"
            Select Case nomen
                ' Dia anterior 23:00
                Case "BA", "DF", "ES", "GO", "MG", "MT", "RJ", "CP", "PC", "SO", "CS", "MC", "TO", "MC", "BS", "CS", "ML", "TM", "SM", "IG", "MD"
                    daysBefore = 1: hh = 23: mm = 0: LoadRule = True
                ' Dia anterior 13:00
                Case "NO", "RP", "SJ"
                    daysBefore = 1: hh = 13: mm = 0: LoadRule = True
                ' 2 Dias anteriores 23:00
                Case "BS"
                    daysBefore = 2: hh = 23: mm = 0: LoadRule = True
                ' Dia anterior 21:00
                Case "NF", "SF", "SR", "RL", "JF"
                    daysBefore = 1: hh = 21: mm = 0: LoadRule = True
                ' 3 Dias anteriores 23:00
                Case "RO"
                    daysBefore = 3: hh = 23: mm = 0: LoadRule = True
                ' Segue preliminar
                Case "LT", "SP", "VP"
                    LoadRule = False
                Case Else
                    LoadRule = False
            End Select

        Case "NE"
            Select Case nomen
                ' Dia anterior 23:00
                Case "CE", "PI", "SE"
                    daysBefore = 1: hh = 23: mm = 0: LoadRule = True
                ' 2 dias anteriores 23:00
                Case "MA", "PA", "AP", "ST"
                    daysBefore = 2: hh = 23: mm = 0: LoadRule = True
                Case Else
                    LoadRule = False
            End Select

        Case "PR"
            Select Case nomen
                ' Dia anterior 23:00
                Case "RS"
                    daysBefore = 1: hh = 23: mm = 0: LoadRule = True
                ' Dia anterior 21:00
                Case "NP", "SC", "PN"
                    daysBefore = 1: hh = 21: mm = 0: LoadRule = True
                Case Else
                    LoadRule = False
            End Select

        Case "FT"
            ' Se coluna 6 = "CDJD" -> Dia anterior 23:00; senão, não altera (equivalente a "segue preliminar")
            If cd6 = "CDJD" Then
                daysBefore = 1: hh = 23: mm = 0: LoadRule = True
            Else
                LoadRule = False
            End If

        Case Else
            LoadRule = False
    End Select
End Function
' ===== Núcleo =====
' Chamar dentro do seu For: Call VerificandoDTHCARREG(linha_FS, cd_fat2)
' Mantém um cache por rota: pega a primeira entrega (Z="FIRST") e aplica nas linhas seguintes até o próximo "FIRST".
Public Sub VerificandoDTHCARREG(ByVal linha_FS As Long, ByVal cd_fat2 As String)

    Static routeActive As Boolean
    Static cacheDateText As String
    Static cacheTimeValue As Date

    Dim isFirst As Boolean
    Dim nomen As String, cd6 As String
    Dim daysBefore As Long, hh As Long, mm As Long
    Dim entregaTxt As String
    Dim entregaDt As Date, cargaDt As Date

    isFirst = (UCase$(Trim$(Cells(linha_FS, "Z").Value)) = "FIRST")

    If isFirst Then
        ' Sempre reavalia na primeira linha da rota
        nomen = Left$(CStr(Cells(linha_FS, "H").Value), 2)
        cd6 = CStr(Cells(linha_FS, 6).Value)

        If LoadRule(cd_fat2, nomen, cd6, daysBefore, hh, mm) Then
            entregaTxt = CStr(Cells(linha_FS, "M").Value)
            If ParseTextDateDMY(entregaTxt, entregaDt) Then
                cargaDt = DateAdd("d", -daysBefore, entregaDt)

                ' Cache:
                ' J como texto "dd/mm/aa"
                cacheDateText = Format$(cargaDt, "dd/mm/yy")
                ' K como valor de hora e formatação hh:mm:ss
                cacheTimeValue = TimeSerial(hh, mm, 0)

                routeActive = True
            Else
                ' Data inválida -> não altera J/K
                routeActive = False
                cacheDateText = vbNullString
                cacheTimeValue = 0
            End If
        Else
            ' "Segue preliminar" ou sem regra -> não altera J/K nesta rota
            routeActive = False
            cacheDateText = vbNullString
            cacheTimeValue = 0
        End If
    End If

    ' Se a rota tem regra ativa, aplicar nas linhas (inclui a própria FIRST)
    If routeActive Then
        ' Coluna J: texto "dd/mm/aa"
        With Cells(linha_FS, "J")
            .NumberFormat = "@"
            .Value = cacheDateText
        End With
        ' Coluna K: hora como valor com formato hh:mm:ss
        With Cells(linha_FS, "K")
            .NumberFormat = "hh:mm:ss"
            .Value = cacheTimeValue
        End With
    End If
End Sub

Public Sub EscrevendoFirst()

    Range("Z2:Z" & ult_linha).ClearContents

    Cells(2, "Z").Value = "FIRST"

    For linha_origem = 3 To ult_linha

        If UCase$(Trim$(Cells(linha_origem, "H").Value)) <> _
           UCase$(Trim$(Cells(linha_origem - 1, "H").Value)) Then

            Cells(linha_origem, "Z").Value = "FIRST"

        End If

    Next linha_origem

End Sub

' Verifica e aplica a lógica DIRETAGO para rotas do tipo GO*96F (cd_fat2="GR") ou GO*96M (cd_fat2="JC"),
' onde * é um dígito de 2 a 7. Quando aplica:
'   - J = data (texto "dd/mm/aa") = dataEntrega - 2 dias
'   - K = hora (valor "hh:mm:ss") = 09:00:00
'   - D = "DISTR CD" & cd_fat2 & " - DIR"
'   - F = "CD" & cd_fat2
' Usa a data de entrega da PRIMEIRA linha da rota (Z="FIRST") e replica nas seguintes até o próximo "FIRST".
' Retorna True quando aplicou na linha atual (inclui linhas subsequentes da mesma rota).
Public Function VerificandoDIRETAGO(ByVal linha_FS As Long, ByVal cd_fat2 As String) As Boolean
    Static dirRouteActive As Boolean
    Static cacheDateText As String
    Static cacheTimeValue As Date
    Static cacheTipoOper As String
    Static cacheOrigem As String

    Dim isFirst As Boolean
    Dim cFat As String, rotaTxt As String, suf As String, pattern As String
    Dim entregaTxt As String, entregaDt As Date, cargaDt As Date

    isFirst = (UCase$(Trim$(Cells(linha_FS, "Z").Value)) = "FIRST")

    If isFirst Then
        ' Por padrão, desativa para a nova rota
        dirRouteActive = False

        cFat = UCase$(Trim$(cd_fat2))
        ' DIRETAGO só se aplica a GR (F) e JC (M)
        If cFat = "GR" Or cFat = "JC" Then
            rotaTxt = UCase$(CStr(Cells(linha_FS, "H").Value))
            suf = IIf(cFat = "GR", "F", "M")
            ' Padrão: contém GS[2-7]96 + sufixo F/M (ex.: GS296F, GS596M, etc.)
            pattern = "*GS[234567]96" & suf & "*"

            If rotaTxt Like pattern Then
                entregaTxt = CStr(Cells(linha_FS, "M").Value)
                ' Usa o parser já existente para "dd/mm/aa" -> Date
                If ParseTextDateDMY(entregaTxt, entregaDt) Then
                    ' 2 dias anteriores, 22:00
                    cargaDt = DateAdd("d", -2, entregaDt)
                    cacheDateText = Format$(cargaDt, "dd/mm/yy")
                    cacheTimeValue = TimeSerial(22, 0, 0)

                    cacheTipoOper = "DISTR CD" & cFat & " - DIR"
                    cacheOrigem = "CD" & cFat

                    dirRouteActive = True
                Else
                    ' Data inválida na M -> não aplica DIRETAGO nesta rota
                    dirRouteActive = False
                End If
            End If
        End If
    End If

    ' Se a rota DIRETAGO está ativa, aplicar nas linhas da rota (inclui a própria FIRST)
    If dirRouteActive Then
        ' J: texto "dd/mm/aa"
        With Cells(linha_FS, "J")
            .NumberFormat = "@"
            .Value = cacheDateText
        End With
        ' K: hora "hh:mm:ss"
        With Cells(linha_FS, "K")
            .NumberFormat = "hh:mm:ss"
            .Value = cacheTimeValue
        End With
        ' D e F
        Cells(linha_FS, "D").Value = cacheTipoOper
        Cells(linha_FS, "F").Value = cacheOrigem

        VerificandoDIRETAGO = True
    Else
        VerificandoDIRETAGO = False
    End If
End Function

' Rotina que abre o modelo, atualiza o CALENDARIO com eventos OFF, salva,
' fecha e reabre com eventos ON para rodar o Workbook_Open do wb2.

Public Sub AtualizandoCalendarioModeloConsolidadoOUTROS(ByVal caminho_modelo_consolidado_OUTROS As String)

    Dim wb2 As Workbook
    Dim evtState As Boolean, scrState As Boolean, alertState As Boolean

    ' --- ABERTURA 1: eventos OFF para atualizar o CALENDARIO sem interferência ---
    evtState = Application.EnableEvents: Application.EnableEvents = False
    scrState = Application.ScreenUpdating: Application.ScreenUpdating = False
    alertState = Application.DisplayAlerts: Application.DisplayAlerts = False

    Set wb2 = Workbooks.Open(FileName:=caminho_modelo_consolidado_OUTROS, ReadOnly:=False)

    ' Atualiza o CALENDARIO de wb2 usando a mesma Sub (agora com parâmetro)
    AnalisandoCalendario wb2

    ' Se possível, salvar (não salva se estiver fisicamente read-only)
    If Not wb2.ReadOnly Then wb2.Save

    ' Fecha para reabrir com eventos ON (para executar o Workbook_Open do wb2)
    wb2.Close SaveChanges:=False

    ' Restaura estados
    Application.DisplayAlerts = alertState
    Application.ScreenUpdating = scrState
    Application.EnableEvents = evtState

    ' --- ABERTURA 2: eventos ON para executar o Workbook_Open do wb2 normalmente ---
    Set wb2 = Workbooks.Open(FileName:=caminho_modelo_consolidado_OUTROS)

    ' Se seu fluxo precisar continuar usando wb2 (ex.: copiar dados, consolidar, etc.),
    ' faça aqui. Exemplo:
    ' Call ProcessarPosAbertura(wb2)

    ' Salvar/fechar ao final (ajuste conforme seu fluxo)
    'If Not wb2.ReadOnly Then wb2.Save
    'wb2.Close SaveChanges:=True

End Sub
