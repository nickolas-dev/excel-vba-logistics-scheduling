Attribute VB_Name = "Módulo2"
Sub Botão3_Clique()

Call DesativarAtualizacoes

Call ocultandoformato

'2. tratando erros de inicialização da macro

    If Sheets("Planilha1").Cells(2, 1).Value = "" Then
    
        erro1 = MsgBox("Os dados devem ser colados como TEXTO na célula A2", vbCritical)
        
        Call formatooriginal
        
        Call AtivarAtualizacoes
        
        Exit Sub
    
    End If
    
    ActiveSheet.Unprotect 'desprotegendo a planilha
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=LEFT(R[2]C[-1],1)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'protegendo a planilha
    
    If Sheets("Planilha1").Cells(2, 2).Value <> "." Then
    
        erro2 = MsgBox("Os dados devem ser colados como TEXTO na célula A2", vbCritical)
        
        ActiveSheet.Unprotect 'desprotegendo a planilha
        
    Call formatooriginal
    
    
    Range("B2").Clear
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'protegendo a planilha
        
    Call AtivarAtualizacoes
        
        Exit Sub
    
    End If
    
    ActiveSheet.Unprotect 'desprotegendo a planilha
    Range("B2").Clear
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'protegendo a planilha

'UserForm1.Show
    'CD de faturamento
    cd_fat2 = Mid(Range("A3").Value, 5, 2)
    cd_fat = "CD" & cd_fat2
    
    'executar a macro principal
    'Call AtivarAtualizacoes
    criacaoBase

End Sub

Sub criacaoBase()

Call DesativarAtualizacoes
Call ocultandoformato

'1. macro para abrir o .pdf e colar os valores como texto na Planilha1 - PENDENTE
    
'3. reexibindo planilhas

    Sheets("Planilha2").Visible = True
    Sheets("CALENDARIO").Visible = True
    Sheets("Planilha1").Select
    
' Inicializa a barra de progresso (Parte 1 - 0%)
    ShowProgressBar
    DoEvents
    
'4. texto para colunas

    ActiveSheet.Unprotect 'desprotegendo a planilha
    
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:=".", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="kms:", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.TextToColumns Destination:=Range("A2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
        ), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1)), TrailingMinusNumbers:= _
        True
    
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'protegendo a planilha
    Range("A1").Select

'5. arrumando a data de faturamento na célula M2 da Planilha1

    If cd_fat <> "CDNE" Then
    
        ActiveSheet.Unprotect 'desprotegendo a planilha
        
        If Cells(2, 13).Value < Cells(1, 16).Value Then
        
            '1x
            
            Range("M2").Select
            Selection.NumberFormat = "mm/dd/yy;@"
            Range("P3").Value = Cells(2, 13).Value
            
            ' convertendo a data para texto 1x
        
            Range("O3").Select
            ActiveCell.FormulaR1C1 = "=TEXT(R3C16,""dd/mm/aa"")"
            Selection.Copy
            Range("M2").Select
            Selection.PasteSpecial
            Range("O3").Select
            Application.CutCopyMode = False
            Selection.ClearContents
            
            '2x
            
            Range("M2").Select
            Selection.NumberFormat = "mm/dd/yy;@"
            Range("P3").Value = Cells(2, 13).Value
            
            ' convertendo a data para texto 2x
        
            Range("O3").Select
            ActiveCell.FormulaR1C1 = "=TEXT(R3C16,""dd/mm/aa"")"
            Selection.Copy
            Range("M2").Select
            Selection.PasteSpecial
            Range("O3").Select
            Application.CutCopyMode = False
            Selection.ClearContents
    
        Else
        
        End If
        
        'data de faturamento
        data_fat = Cells(2, 13).Value
    
    Else
    
    ActiveSheet.Unprotect 'desprotegendo a planilha
        
        If Cells(2, 12).Value < Cells(1, 12).Value Then
        
            '1x
            
            Range("L2").Select
            Selection.NumberFormat = "mm/dd/yy;@"
            Range("P3").Value = Cells(2, 12).Value
            
            ' convertendo a data para texto 1x
        
            Range("O3").Select
            ActiveCell.FormulaR1C1 = "=TEXT(R3C16,""dd/mm/aa"")"
            Selection.Copy
            Range("L2").Select
            Selection.PasteSpecial
            Range("O3").Select
            Application.CutCopyMode = False
            Selection.ClearContents
            
            '2x
            
            Range("L2").Select
            Selection.NumberFormat = "mm/dd/yy;@"
            Range("P3").Value = Cells(2, 12).Value
            
            ' convertendo a data para texto 2x
        
            Range("O3").Select
            ActiveCell.FormulaR1C1 = "=TEXT(R3C16,""dd/mm/aa"")"
            Selection.Copy
            Range("L2").Select
            Selection.PasteSpecial
            Range("O3").Select
            Application.CutCopyMode = False
            Selection.ClearContents
        
        'data de faturamento NE
        data_fat_NE = Cells(2, 12).Value
        
        Else
        
        'data de faturamento NE
        data_fat_NE = Cells(2, 12).Value
        
        End If
    
    End If
    
'6. tratando erro de dados na coluna incorreta

ult_linha_geral = Sheets("Planilha1").Range("A2").End(xlDown).Row 'última que contém informações na Planilha1

For linha_verificada = 2 To ult_linha_geral

    If Cells(linha_verificada, 2).Value <> "2F" And Cells(linha_verificada, 2).Value <> "3F" And Cells(linha_verificada, 2).Value <> "4F" And Cells(linha_verificada, 2).Value <> "5F" And Cells(linha_verificada, 2).Value <> "6F" And Cells(linha_verificada, 2).Value <> "Sab" Then
    
        GoTo fim_do_teste
        
        Else
        
        Range("B" & linha_verificada).Select
        Range(Selection, Selection.End(xlToRight)).Select
    Selection.Cut Destination:=Range("C" & linha_verificada & ":P" & linha_verificada)
    Range("C" & linha_verificada & ":P" & linha_verificada).Select
        
fim_do_teste:
    
    End If

Next

'6.1 ajustando colunas
    Call AjustandoColunas

'(Parte 2 - 20%)
UserForm5.UpdateProgressBar 20
DoEvents

'7. manipulando as fórmulas na Planilha2

Sheets("Planilha2").Activate

'If cd_fat = "CDFT" Then

'formulaA2 = "=IF(Planilha1!RC=""ENTREGA"",Planilha1!R[-1]C,"""")"

'Else

'formulaA2 = "=IF(AND(LEN(Planilha1!RC)=6,ISTEXT(MID(Planilha1!RC,1,2))," & Chr(10) & "ISNUMBER(MID(Planilha1!RC,3,1)+0),ISNUMBER(MID(Planilha1!RC,4,1)+0)," & Chr(10) & "ISNUMBER(MID(Planilha1!RC,5,1)+0),MID(Planilha1!RC,6,1)=""F""),Planilha1!RC,"""")"

'End If

formulaA2 = "=IF(OR(Planilha1!RC=""ENTREGA"", AND(LEN(Planilha1!RC)=6, ISTEXT(MID(Planilha1!RC,1,2)), ISNUMBER(MID(Planilha1!RC,3,1)+0), ISNUMBER(MID(Planilha1!RC,4,1)+0), ISNUMBER(MID(Planilha1!RC,5,1)+0), MID(Planilha1!RC,6,1)=""F"")), IF(Planilha1!RC=""ENTREGA"", Planilha1!R[-1]C, Planilha1!RC), """")"

    'fórmula na célula A2: copiando e colando na coluna A (Rotas) até a última linha geral
    
    Range("A2").Select
    ActiveCell.FormulaR1C1 = formulaA2
    Range("A2").Copy
    Range("A3:A" & ult_linha_geral).Select
    ActiveSheet.Paste
    Range("A1").Select
    Application.CutCopyMode = False
    
'If cd_fat = "CDFT" Then

'formulaB2 = "=IF(AND(LEN(Planilha1!R[-1]C[-1])=3,Planilha1!R[-1]C[-1]<>""RFG"",Planilha1!R[-1]C[-1]<>""Pao"",Planilha1!R[-1]C[-1]<>""CD:""),Planilha1!R[-1]C[-1],"""")"

'Else

'formulaB2 = "=IF(AND(LEN(Planilha1!RC[-1])=7, MID(Planilha1!RC[-1],4,1)=""-""), Planilha1!RC[-1], """")"

'End If

'formulaB2 = "=IF(OR(AND(LEN(Planilha1!R[-1]C[-1])=3, Planilha1!R[-1]C[-1]<>""RFG"", Planilha1!R[-1]C[-1]<>""Pao"", Planilha1!R[-1]C[-1]<>""CD:""), AND(LEN(Planilha1!RC[-1])=7, MID(Planilha1!RC[-1],4,1)=""-"")), IF(AND(LEN(Planilha1!R[-1]C[-1])=3, Planilha1!R[-1]C[-1]<>""RFG"", Planilha1!R[-1]C[-1]<>""Pao"", Planilha1!R[-1]C[-1]<>""CD:""), Planilha1!R[-1]C[-1], Planilha1!RC[-1]), """")"

'Formula B2 que desconsidera "USO" também
'formulaB2 = "=IF(OR(AND(LEN(Planilha1!R[-1]C[-1])=3, Planilha1!R[-1]C[-1]<>""RFG"", Planilha1!R[-1]C[-1]<>""Pao"", Planilha1!R[-1]C[-1]<>""CD:"", Planilha1!R[-1]C[-1]<>""USO""), AND(LEN(Planilha1!RC[-1])=7, MID(Planilha1!RC[-1],4,1)=""-"")), IF(AND(LEN(Planilha1!R[-1]C[-1])=3, Planilha1!R[-1]C[-1]<>""RFG"", Planilha1!R[-1]C[-1]<>""Pao"", Planilha1!R[-1]C[-1]<>""CD:"", Planilha1!R[-1]C[-1]<>""USO""), Planilha1!R[-1]C[-1], Planilha1!RC[-1]), """")"

'Formula B2 com cenário OUTBACK
formulaB2 = "=IF(OR(AND(LEN(Planilha1!R[-1]C[-1])=3,NOT(ISNUMBER(--Planilha1!R[-1]C[-1])),Planilha1!R[-1]C[-1]<>""RFG"",Planilha1!R[-1]C[-1]<>""Pao"",Planilha1!R[-1]C[-1]<>""CD:"",Planilha1!R[-1]C[-1]<>""USO""),AND(LEN(Planilha1!RC[-1])=7,MID(Planilha1!RC[-1],4,1)=""-"")),IF(AND(LEN(Planilha1!R[-1]C[-1])=3,NOT(ISNUMBER(--Planilha1!R[-1]C[-1])),Planilha1!R[-1]C[-1]<>""RFG"",Planilha1!R[-1]C[-1]<>""Pao"",Planilha1!R[-1]C[-1]<>""CD:"",Planilha1!R[-1]C[-1]<>""USO""),Planilha1!R[-1]C[-1],Planilha1!RC[-1]),"""")"
    
    'fórmula na célula B2: copiando e colando na coluna B (Lojas) até a última linha geral
    
    Range("B2").Select
    ActiveCell.FormulaR1C1 = formulaB2
    Range("B2").Copy
    Range("B3:B" & ult_linha_geral).Select
    ActiveSheet.Paste
    Range("B1").Select
    Application.CutCopyMode = False
    
    
'primeira linha que tem rota na Planilha2

linha_analisada1 = 2

Range("A" & linha_analisada1).Select

Do Until ActiveCell.Value <> ""

Range("A" & linha_analisada1).Select

linha_analisada1 = linha_analisada1 + 1

Loop
    
primeira_linha_rota = ActiveCell.Row

'primeira linha que tem loja na Planilha2

linha_analisada2 = 2

Range("B" & linha_analisada2).Select

Do Until ActiveCell.Value <> ""

Range("B" & linha_analisada2).Select

linha_analisada2 = linha_analisada2 + 1

Loop

'(Parte 3 - 40%)
UserForm5.UpdateProgressBar 40
DoEvents
    
primeira_linha_loja = ActiveCell.Row

    'fórmula na célula E & primeira_linha_loja: copiando e colando na coluna E (ROTA1) até a última linha geral
    
    Range("E" & primeira_linha_loja).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-2]C[-4]<>"""",""FIRST""&R[-2]C[-4],IF(AND(R[1]C[-3]="""",R[1]C[-4]="""",RC[-3]<>""""),""LAST""&RIGHT(R[-1]C,6),RIGHT(R[-1]C,6)))"
    Range("E" & primeira_linha_loja).Copy
    Range("E" & primeira_linha_loja + 1 & ":E" & ult_linha_geral).Select
    ActiveSheet.Paste
    Range("E1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula F & primeira_linha_loja: copiando e colando na coluna F (ROTA2) até a última linha geral
    
    Range("F" & primeira_linha_loja).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(LEFT(R[-1]C[-1],4)=""LAST"","""",IF(LEFT(RC[-1],5)=""FIRST"",RIGHT(RC[-1],6),IF(LEFT(RC[-1],4)=""LAST"",RIGHT(RC[-1],6),IF(R[-1]C="""",RC[-1],RC[-1]))))"
    Range("F" & primeira_linha_loja).Copy
    Range("F" & primeira_linha_loja + 1 & ":F" & ult_linha_geral).Select
    ActiveSheet.Paste
    Range("F1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula G & primeira_linha_loja: copiando e colando na coluna G (ROTA3) até a última linha geral
    
    Range("G" & primeira_linha_loja).Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-5]<>"""",RC[-1],"""")"
    Range("G" & primeira_linha_loja).Copy
    Range("G" & primeira_linha_loja + 1 & ":G" & ult_linha_geral).Select
    ActiveSheet.Paste
    Range("G1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula C2: filtro de rotas
    
    linha_filtro = ult_linha_geral - 2
    
    Range("C2").Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[4]:R[" & linha_filtro & "]C[4],RC[4]:R[" & linha_filtro & "]C[4]<>"""")"
    
    'fórmula na célula D2: filtro de lojas
    
    Range("D2").Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[-2]:R[" & linha_filtro & "]C[-2],RC[-2]:R[" & linha_filtro & "]C[-2]<>"""")"
    

'7.1 erro de rotas x lojas

VerificarNovamente:

    Cells(1, 20).FormulaR1C1 = "=COUNTA(C[-17])"
    Cells(2, 20).FormulaR1C1 = "=COUNTA(C[-16])"
    
    total_rotas = Cells(1, 20).Value
    total_lojas = Cells(2, 20).Value
    
    If total_rotas <> total_lojas Then
    
        Call AjustarRotasELojas
        
        GoTo VerificarNovamente
    
        erro2 = MsgBox("O cálculo rotas x lojas é divergente." & vbNewLine & vbNewLine & "Erro ao criar a base.", vbCritical)
        
        Cells(1, 20).ClearContents
        Cells(2, 20).ClearContents
        
        Call AtivarAtualizacoes
        
        Exit Sub
        
    End If
    
    Cells(1, 20).ClearContents
    Cells(2, 20).ClearContents
    
    'fórmula na célula AR2 (Rotas únicas)
    
    ult_linha_loja = Range("D1").End(xlDown).Row
    
    Range("AR2").Select
    ActiveCell.Formula2R1C1 = "=UNIQUE(R2C3:R" & ult_linha_loja & "C3)"
    
    'fórmula na célula H2: copiando e colando na coluna H (FATURAMENTO) até a última linha loja
    
    ult_linha_loja = Range("D1").End(xlDown).Row
    
    If cd_fat <> "CDNE" Then
    
    formula_faturamento = "=Planilha1!R2C13"
    
    Else
    
    formula_faturamento = "=Planilha1!R2C12"
    
    End If
    
    Range("H2").Select
    ActiveCell.FormulaR1C1 = formula_faturamento
    Range("H2").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("H2:H" & ult_linha_loja).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
        
    'fórmula na célula L2: copiando e colando na coluna L (DATACARREG2) até a última linha loja
    
    Range("L2").Select
    ActiveCell.FormulaR1C1 = _
        "=XLOOKUP(RC[-4],CALENDARIO!C[-10],CALENDARIO!C[-11]," & Chr(10) & "XLOOKUP(RC[-4],CALENDARIO!C[-9],CALENDARIO!C[-11]," & Chr(10) & "XLOOKUP(RC[-4],CALENDARIO!C[-8],CALENDARIO!C[-11]," & Chr(10) & "XLOOKUP(RC[-4],CALENDARIO!C[-7],CALENDARIO!C[-11]," & Chr(10) & "XLOOKUP(RC[-4],CALENDARIO!C[-6],CALENDARIO!C[-11]," & Chr(10) & "XLOOKUP(RC[-4],CALENDARIO!C[-5],CALENDARIO!C[-11]," & Chr(10) & "XLOOKUP(RC[-4],CALENDARIO!C[-4],CALENDARIO!C[-11],"""")))))))"
    Range("L2").Copy
    Range("L3:L" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("L1").Select
    Application.CutCopyMode = False

    'fórmula na célula M2: copiando e colando na coluna M (DATACARREG3) até a última linha loja
    
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=XLOOKUP(RC[-2],CALENDARIO!R2,CALENDARIO!R1,"""")"
    Range("M2").Copy
    Range("M3:M" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("M1").Select
    Application.CutCopyMode = False
    
    
    'fórmula na célula Z2: copiando e colando na coluna Z (DATAFATURAMENTO2) até a última linha loja
    
    Range("Z2").Select
    ActiveCell.FormulaR1C1 = "=WEEKDAY(RC[-18],1)"
    Range("Z2").Copy
    Range("Z3:Z" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("Z1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula AA2: copiando e colando na coluna AA (DATACARREG4) até a última linha loja
    
    Range("AA2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-14]<RC[-1],(RC[-15]+1),RC[-15])"
    Range("AA2").Copy
    Range("AA3:AA" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("AA1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula AB2: copiando e colando na coluna AB (DATAENTREGA4) até a última linha loja
    
    Range("AB2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-10]<RC[-2],(RC[-11]+1),RC[-11])"
    Range("AB2").Copy
    Range("AB3:AB" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("AB1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula AD2: copiando e colando na coluna AD (2FORMATO) até a última linha loja
    
    Range("AD2").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-1],""dd/mm/aa"")"
    Range("AD2").Copy
    Range("AD3:AD" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("AD1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula AE2: copiando e colando na coluna AE (3FORMATO) até a última linha loja
    
    Range("AE2").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-2],""mm/dd/aa"")"
    Range("AE2").Copy
    Range("AE3:AE" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("AE1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula AH2: copiando e colando na coluna AH (ANO FATURAMENTO) até a última linha loja
    
    Range("AH2").Select
    ActiveCell.FormulaR1C1 = "=YEAR(RC[-26])"
    Range("AH2").Copy
    Range("AH3:AH" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("AH1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula AI2: copiando e colando na coluna AI (MÊS FATURAMENTO) até a última linha loja
    
    Range("AI2").Select
    ActiveCell.FormulaR1C1 = "=MONTH(RC[-27])"
    Range("AI2").Copy
    Range("AI3:AI" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("AI1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula AJ2: copiando e colando na coluna AJ (DIA FATURAMENTO) até a última linha loja
    
    Range("AJ2").Select
    ActiveCell.FormulaR1C1 = "=DAY(RC[-28])"
    Range("AJ2").Copy
    Range("AJ3:AJ" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("AJ1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula AF2: copiando e colando na coluna AF (DATAFATURAMENTO3) até a última linha loja
    
    Range("AF2").Select
    ActiveCell.FormulaR1C1 = "=DATE(RC[2],RC[3],RC[4])"
    Range("AF2").Copy
    Range("AF3:AF" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("AF1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula AG2: copiando e colando na coluna AG (1FORMATO<DFAT3) até a última linha loja
    
    Range("AG2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-4]<RC[-1],""Sim"",""Não"")"
    Range("AG2").Copy
    Range("AG3:AG" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("AG1").Select
    Application.CutCopyMode = False
    
'(Parte 4 - 60%)
UserForm5.UpdateProgressBar 60
DoEvents
    
'inserindo informações de carreg. da outra macro

Call infoCarreg

'(Parte 5 - 80%)
UserForm5.UpdateProgressBar 80
    
    'fórmula na célula O2: copiando e colando na coluna O (HORA ENTREGA) até a última linha loja
    
    Range("O2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[3]=1,RIGHT(RC[1],5),XLOOKUP(RC[-11],Planilha1!C[-14],Planilha1!C[-11],""""))"
    Range("O2").Copy
    Range("O3:O" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("O1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula P2: copiando e colando na coluna P (DATAENTREGA1) até a última linha loja
    
    Range("P2").Select
    ActiveCell.FormulaR1C1 = _
        "=XLOOKUP(RC[-12],Planilha1!C[-15],Planilha1!C[-13],"""")"
    Range("P2").Copy
    Range("P3:P" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("P1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula R2: copiando e colando na coluna R (DATAENTREGA3) até a última linha loja
    
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "=XLOOKUP(RC[-2],CALENDARIO!R2,CALENDARIO!R1,1)"
    Range("R2").Copy
    Range("R3:R" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("R1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula Q2: copiando e colando na coluna Q (DATAENTREGA2) até a última linha loja
    
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[1]>=RC[-4],RC[-5],(RC[-5]+1))"
    Range("Q2").Copy
    Range("Q3:Q" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("Q1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula AJ2: copiando e colando na coluna AJ (CAIXAS) até a última linha loja
    
    Range("AJ2").Select
    ActiveCell.FormulaR1C1 = _
        "=XLOOKUP(RC[-32],Planilha1!C[-35],Planilha1!C[-31],"""")"
    Range("AJ2").Copy
    Range("AJ3:AJ" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("AJ1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula AY2: copiando e colando na coluna AY (M³ CONG) até a última linha loja
    
    Range("AY2").Select
    ActiveCell.FormulaR1C1 = _
        "=XLOOKUP(RC[-47],Planilha1!C[-50],Planilha1!C[-42],"""")"
    Range("AY2").Copy
    Range("AY3:AY" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("AY1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula AZ2: copiando e colando na coluna AZ (M³ RESF) até a última linha loja
    
    Range("AZ2").Select
    ActiveCell.FormulaR1C1 = _
        "=XLOOKUP(RC[-48],Planilha1!C[-51],Planilha1!C[-40],"""")"
    Range("AZ2").Copy
    Range("AZ3:AZ" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("AZ1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula BA2: copiando e colando na coluna BA (M³ SECO) até a última linha loja
    
    Range("BA2").Select
    ActiveCell.FormulaR1C1 = _
        "=XLOOKUP(RC[-49],Planilha1!C[-52],Planilha1!C[-38],"""")"
    Range("BA2").Copy
    Range("BA3:BA" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("BA1").Select
    Application.CutCopyMode = False
    
    'fórmula na célula BB2: copiando e colando na coluna BB (M³ TOTAL) até a última linha loja
    
    Range("BB2").Select
    ActiveCell.FormulaR1C1 = _
        "=XLOOKUP(RC[-50],Planilha1!C[-53],Planilha1!C[-48],"""")"
    Range("BB2").Copy
    Range("BB3:BB" & ult_linha_loja).Select
    ActiveSheet.Paste
    Range("BB1").Select
    Application.CutCopyMode = False
    
'inserindo informações de entrega da outra macro

Call infoEntrega

Call VerificandoTransportadora

'fim da macro

Sheets("Planilha1").Activate

Range("A1").Select

Sheets("Planilha2").Visible = False
Sheets("CALENDARIO").Visible = False
Sheets("CROSS DOCKING").Visible = False

'Encerramento da Criação da Base
    'Call AtivarAtualizacoes
    'UserForm2.Show - userForm antigo
    Call InformacoesConsolidadas
    
End Sub

Sub infoCarreg()

linha_analise = 2 'linha na Planilha2 onde estão as entregas

ult_linha_analise = Range("D2").End(xlDown).Row 'linha onde está a última entrega

For linha_analise = 2 To ult_linha_analise
    
    matrizProcurada = Range("Planilha1!A:A") 'coluna onde estão as rotas na Planilha1
    
    matrizProcurada2 = Range("Planilha2!H:H") 'coluna onde estão as rotas únicas na Planilha2
    
    matrizRetorno2 = Range("Planilha2!I:I") 'coluna onde estão a qtd de lojas na Planilha2
    
    Rota = Cells(linha_analise, 3).Value 'rota que está sendo analisada
    
    linha_rota_Plan1 = WorksheetFunction.Match(Rota, matrizProcurada, 0) 'linha onde a rota analisada foi encontrada na Planilha1
    
    'linha onde estão as informações de carregamento na Planilha1
    
    linha_info_carreg_Plan1 = linha_rota_Plan1
        
    Do While Sheets("Planilha1").Cells(linha_info_carreg_Plan1, 5).Value <> "Carregamento:"
    
        linha_info_carreg_Plan1 = linha_info_carreg_Plan1 + 1
    
    Loop
    
    data_carreg1 = Sheets("Planilha1").Cells(linha_info_carreg_Plan1, 6).Value 'data de carregamento da rota analisada
    
    Dim hora_carreg As Date
    
    hora_carreg = NormalizaHora(Sheets("Planilha1").Cells(linha_info_carreg_Plan1, 7).Value) 'hora de carregamento da rota analisada
    
    Sheets("Planilha2").Cells(linha_analise, 10).Value = hora_carreg 'preenchendo hora
    
    Sheets("Planilha2").Cells(linha_analise, 11).Value = data_carreg1 'preenchendo data 1
    
    semana_do_ano = Cells(linha_analise, 27).Value ' semana do ano da data analisada
    
    dia_da_semana = Cells(linha_analise, 13).Value ' dia da semana da data analisada
    
    data_carreg2 = Sheets("CALENDARIO").Cells(semana_do_ano + 3, dia_da_semana + 1).Value 'data de carregamento sem tratamento
    
    Sheets("Planilha2").Cells(linha_analise, 29).Value = data_carreg2 'preenchendo data de carregamento sem tratamento na coluna AC
    
    'colando valores nas colunas AD e AE
    
    Range("AD" & linha_analise & ":AE" & linha_analise).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'analisando datas de carregamento
    
    If Cells(linha_analise, 33).Value = "Sim" Then
    
    Cells(linha_analise, 31).Select
    Selection.Copy
    Cells(linha_analise, 9).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Else

    Cells(linha_analise, 30).Select
    Selection.Copy
    Cells(linha_analise, 9).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    End If
    
Next

End Sub

Sub infoEntrega()

linha_analise = 2 'linha na Planilha2 onde estão as entregas

ult_linha_analise = Range("D2").End(xlDown).Row 'linha onde está a última entrega

For linha_analise = 2 To ult_linha_analise

    'Se DiaSemFat (coluna Z) > DiaSemEnt (coluna R), ajustar próxima semana
    
    If Cells(linha_analise, "Z").Value > Cells(linha_analise, "R").Value Then
    
        semana_do_ano2 = Cells(linha_analise, 17).Value ' semana do ano da data analisada
        
        'Se DiaSemCarreg (coluna M) <= DiaSemEnt (coluna R), a semana é a próxima
        
        If Cells(linha_analise, "M").Value <= Cells(linha_analise, "R").Value Then
        
        semana_do_ano2 = semana_do_ano2 + 1
        
        End If
        
    Else
    
        semana_do_ano2 = Cells(linha_analise, 28).Value ' semana do ano da data analisada
        
    End If
    
    dia_da_semana2 = Cells(linha_analise, 18).Value ' dia da semana da data analisada
    
    Dim celula_data_entrega As Range
    
    Set celula_data_entrega = Sheets("CALENDARIO").Cells(semana_do_ano2 + 3, dia_da_semana2 + 1) 'célula data de entrega
    
    celula_data_entrega.Copy
    
    Sheets("Planilha2").Cells(linha_analise, 14).Select 'preenchendo data de entrega
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

Next

End Sub

Sub ShowProgressBar()
    UserForm5.Show vbModeless
End Sub

Sub RunProcess()
    Dim i As Integer
    Dim totalIterations As Integer
    totalIterations = 5 ' Número total de iterações

    ShowProgressBar

    For i = 0 To totalIterations
        ' Simulação de um processo
        Application.Wait Now + TimeValue("00:00:01") ' Espera 1 segundo

        ' Atualiza a barra de progresso
        UserForm5.UpdateProgressBar (i / totalIterations) * 100
    Next i

    ' Fecha o UserForm ao final do processo
    Unload UserForm5
End Sub

' Coloque em um módulo padrão:
Public Function NormalizaHora(ByVal v As Variant) As Date
    Dim t As String
    Dim parts As Variant
    
    t = Trim$(CStr(v))
    If Len(t) = 0 Then
        NormalizaHora = 0
        Exit Function
    End If
    
    ' Se vier como número Date/Double já de Excel:
    If IsDate(v) And Len(t) > 0 And InStr(t, ":") = 0 Then
        ' Em alguns casos CStr(IsDate) vira só data; reforçamos abaixo via TimeValue
        NormalizaHora = CDate(v)
        Exit Function
    End If
    
    ' Se vier com dois pontos, tratamos os segmentos:
    If InStr(t, ":") > 0 Then
        parts = Split(t, ":")
        Select Case UBound(parts)
            Case 0 ' "8"
                NormalizaHora = TimeSerial(CInt(parts(0)), 0, 0)
            Case 1 ' "08:15"
                NormalizaHora = TimeSerial(CInt(parts(0)), CInt(parts(1)), 0)
            Case 2 ' "08:15:30"
                NormalizaHora = TimeSerial(CInt(parts(0)), CInt(parts(1)), CInt(parts(2)))
            Case Else
                ' Se vier estranho tipo "08:15:30:..." pega os 3 primeiros
                NormalizaHora = TimeSerial(CInt(parts(0)), CInt(parts(1)), CInt(parts(2)))
        End Select
    Else
        ' Sem “:”, tenta converter direto (ex.: 0.5 = 12:00:00)
        If IsNumeric(v) Then
            NormalizaHora = CDate(v)
        ElseIf IsDate(v) Then
            NormalizaHora = CDate(v)
        Else
            ' Último recurso: interpretar como hora inteira
            NormalizaHora = TimeSerial(CInt(t), 0, 0)
        End If
    End If
End Function

Public Sub AjustarRotasELojas()

    Dim wsPlanilha2 As Worksheet

    Dim UltimaLinha As Long
    Dim Linha As Long
    Dim LinhaDestino As Long

    Dim RotaAtual As String

    Set wsPlanilha2 = _
        ThisWorkbook.Sheets("Planilha2")

    UltimaLinha = _
        wsPlanilha2.Cells(wsPlanilha2.Rows.Count, "A").End(xlUp).Row

    'Limpar resultado anterior
    wsPlanilha2.Range("C2:D" & wsPlanilha2.Rows.Count).ClearContents

    LinhaDestino = 2
    RotaAtual = ""

    For Linha = 2 To UltimaLinha

        If Trim(wsPlanilha2.Cells(Linha, "A").Value) <> "" Then

            RotaAtual = _
                Trim(wsPlanilha2.Cells(Linha, "A").Value)

        End If

        If Trim(wsPlanilha2.Cells(Linha, "B").Value) <> "" Then

            wsPlanilha2.Cells(LinhaDestino, "C").Value = RotaAtual

            wsPlanilha2.Cells(LinhaDestino, "D").Value = _
                wsPlanilha2.Cells(Linha, "B").Value

            LinhaDestino = LinhaDestino + 1

        End If

    Next Linha

End Sub


