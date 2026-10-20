Attribute VB_Name = "Módulo3"
Sub aumentandoCalendario()

Application.ScreenUpdating = False

confirma = MsgBox("Confirma que deseja atualizar o calendário até 2099?", vbYesNo + vbQuestion, "Control Tower MB")

If confirma = 7 Then

Exit Sub

End If

ano = 2024

ano = ano + 1

linha_cola = 1

ult_linha = Range("H" & linha_cola).End(xlDown).Row

linha_cola = ult_linha + 2

'copiando modelo
Range("A1:H55").Copy

For ano = 2025 To 2099

    'colando próximo ano
    Range("A" & linha_cola).Select
    Selection.PasteSpecial
    Application.CutCopyMode = False
    
    'alterando informações
    Range("A" & linha_cola).Select
    
    Range("A" & linha_cola).Value = ano
    
    'alterando datas no que foi colado
    
    Dim celula_atual As Range
    
    Dim celula_ult_info As Range
    
    Dim prox_celula As Range
    
    Set celula_atual = Range("B" & linha_cola + 3)
    
    Set celula_ult_info = Range("H" & linha_cola - 2)
    
        linha_ult_semana = Range("A" & linha_cola + 2).End(xlDown).Row
        
        Range("H" & linha_ult_semana).ClearContents
        
    Do Until Cells(linha_ult_semana + 1, 2).Value <> ""
        
        celula_atual.Select
        
        celula_atual.Value = CDate(celula_ult_info.Value) + 1
        
        'convertendo data para texto
        
        linha_atual = celula_atual.Row
        coluna_atual = celula_atual.Column
        coluna_texto = Cells(linha_cola, coluna_atual).Value + 9
        
        Cells(linha_atual, coluna_texto).FormulaR1C1 = "=TEXT(RC[-8],""dd/mm/aa"")"
        
        Cells(linha_atual, coluna_texto).Select
        Cells(linha_atual, coluna_texto).Copy
        celula_atual.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
        'definindo a próxima célula
        
        If celula_atual.Column < 8 Then
        
        Set prox_celula = Cells(linha_atual, celula_atual.Column + 1)
        
        prox_celula.Select
        
        Else
        
        Set prox_celula = Cells(linha_atual + 1, 2)
        
        prox_celula.Select
        
        End If
        
        'preenchendo na próxima célula
        
        Set celula_ult_info = celula_atual
        
        prox_celula.Value = CDate(celula_ult_info.Value) + 1
        
        'próximo caso
        
        Set celula_atual = prox_celula
    
    Loop
    
    'limpar J:P e a célula que encerra o loop
    
    Range("J" & linha_cola + 3 & ":P" & linha_ult_semana).ClearContents
    
    Cells(linha_ult_semana + 1, 2).ClearContents
    
    Range("A1").Select
    
    'continuando
    
    'copiando ano anterior
    
    Range("A" & linha_cola & ":H" & linha_ult_semana).Copy
    
    'próximo ano
    
    linha_cola = linha_ult_semana + 2

Next

Application.ScreenUpdating = True

Range("A1").Select

conclui = MsgBox("O calendário foi atualizado com sucesso!", vbInformation, "Control Tower MB")

End Sub

Sub aumentandoCalendario2()

Application.ScreenUpdating = False

confirma = MsgBox("Confirma que deseja atualizar o calendário até 2099?", vbYesNo + vbQuestion, "Control Tower MB")

If confirma = 7 Then

Exit Sub

End If

'aumentando datas

Dim celula_atual As Range

Set celula_atual = Range("E56")

Set celula_ult_info = Range("D56")

linha_atual = celula_atual.Row
coluna_atual = celula_atual.Column
coluna_texto = Cells(1, coluna_atual).Value + 9

    Do Until Cells(linha_atual, coluna_texto).Value = "31/12/99"
        
        celula_atual.Select
        
        celula_atual.Value = CDate(celula_ult_info.Value) + 1
        
        'convertendo data para texto
        
        linha_atual = celula_atual.Row
        coluna_atual = celula_atual.Column
        coluna_texto = Cells(1, coluna_atual).Value + 9
        
        Cells(linha_atual, coluna_texto).FormulaR1C1 = "=TEXT(RC[-8],""dd/mm/aa"")"
        
        Cells(linha_atual, coluna_texto).Select
        Cells(linha_atual, coluna_texto).Copy
        celula_atual.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
        'definindo a próxima célula
        
        If celula_atual.Column < 8 Then
        
        Set prox_celula = Cells(linha_atual, celula_atual.Column + 1)
        
        prox_celula.Select
        
        Else
        
        Set prox_celula = Cells(linha_atual + 1, 2)
        
        prox_celula.Select
        
        End If
        
        'preenchendo na próxima célula
        
        Set celula_ult_info = celula_atual
        
        prox_celula.Value = CDate(celula_ult_info.Value) + 1
        
        'próximo caso
        
        Set celula_atual = prox_celula
    
    Loop
    
ult_linha = linha_atual
    
'limpar J:P e a célula que encerra o loop
    
    Range("J56:P" & ult_linha).ClearContents
    
    Range("A1").Select
    
'número da semana

Call numerodasemana

Application.ScreenUpdating = True

Range("A1").Select

conclui = MsgBox("O calendário foi atualizado com sucesso!", vbInformation, "Control Tower MB")


End Sub

Sub numerodasemana()

    Dim linha_semana As Long
    Dim ultima_linha As Long
    Dim texto_procurado As String
    Dim coluna As Long
    
    ' Definir o texto a ser procurado
    texto_procurado = "01/01"
    
    ' Definir a última linha a ser verificada
    ultima_linha = 3969
    
    ' Loop através das linhas 4 até 3969
    For linha_semana = 4 To ultima_linha
        ' Loop através das colunas B até H
        For coluna = 2 To 8
            ' Verificar se os 5 primeiros caracteres da célula contêm o texto procurado
            If Left(Cells(linha_semana, coluna).Value, 5) = texto_procurado Then
                ' Definir o valor da coluna A como 1
                Cells(linha_semana, 1).Value = 1
                ' Sair do loop de colunas se o texto for encontrado
                Exit For
            End If
        Next coluna
    Next linha_semana

End Sub

Sub PreencherLinhas()
    Dim linha_semana As Long
    Dim ultima_linha As Long
    Dim valor_anterior As Long
    
    ' Definir a última linha a ser verificada
    ultima_linha = 3969
    
    ' Loop através das linhas 4 até 3969
    For linha_semana = 4 To ultima_linha
        ' Verificar se a célula na coluna A está vazia
        If IsEmpty(Cells(linha_semana, 1).Value) Then
           ' Preencher a célula com o valor da célula acima + 1
            Cells(linha_semana, 1).Value = valor_anterior + 1
            valor_anterior = Cells(linha_semana, 1).Value
        Else
             ' Atualizar o valor anterior
            valor_anterior = Cells(linha_semana, 1).Value
        End If
    Next linha_semana
End Sub

Sub ocultandoformato()

    ActiveSheet.Unprotect
    Cells.Select
    ActiveWindow.DisplayGridlines = False
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("A2").Select
    With Selection.Interior
        .pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub
Sub formatooriginal()

    ActiveSheet.Unprotect
    Cells.Select
    ActiveWindow.DisplayGridlines = True
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
    Range("A2").Select
    With Selection.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("P1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With

End Sub

' Atualiza o CALENDARIO do workbook informado (por padrão, ThisWorkbook)
Public Sub AnalisandoCalendario(Optional ByVal alvo As Workbook = Nothing)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim hojeTxt As String
    Dim primeiraLinha As Long, UltimaLinha As Long
    Dim rngBusca As Range, celEncontrada As Range
    Dim Linha As Long, semana As Variant
    Dim rngSemana As Range, cel0101 As Range
    Dim c As Range
    Dim ano2 As String
    Dim ano4 As Long
    Dim vis As XlSheetVisibility

    ' Se não foi informado, usa ThisWorkbook (funciona como antes)
    If alvo Is Nothing Then
        Set wb = ThisWorkbook
    Else
        Set wb = alvo
    End If

    ' Tenta pegar a planilha CALENDARIO (pode estar oculta; não precisamos exibir)
    On Error Resume Next
    Set ws = wb.Worksheets("CALENDARIO")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub  ' Se não existir, sai silenciosamente

    ' Preserva a visibilidade
    vis = ws.Visible

    ' Texto da data atual no padrão "dd/mm/aa"
    hojeTxt = Format(Date, "dd/mm/yy")

    ' Região de busca: colunas B:H, começando na linha 4
    Const primeiraLinhaConst As Long = 4
    primeiraLinha = primeiraLinhaConst
    UltimaLinha = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If UltimaLinha < primeiraLinha Then GoTo RestaurarVis

    Set rngBusca = ws.Range("B" & primeiraLinha & ":H" & UltimaLinha)

    ' Procura a data de hoje como texto exato
    Set celEncontrada = rngBusca.Find(What:=hojeTxt, LookIn:=xlValues, LookAt:=xlWhole, _
                                      SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If celEncontrada Is Nothing Then GoTo RestaurarVis  ' Não achou a data de hoje, sai

    ' Linha onde a data foi encontrada
    Linha = celEncontrada.Row

    ' Semana do ano informada na coluna A dessa linha
    semana = ws.Cells(Linha, "A").Value

    ' Se a semana não for 1, não faz nada (conforme sua regra)
    If Val(semana) <> 1 Then GoTo RestaurarVis

    ' Na mesma linha, procurar "01/01" nas colunas B:H
    Set rngSemana = ws.Range("B" & Linha & ":H" & Linha)
    Set cel0101 = Nothing
    For Each c In rngSemana.Cells
        Dim txt As String
        txt = Trim$(CStr(c.Value))
        If Len(txt) >= 5 Then
            If Left$(txt, 5) = "01/01" Then
                Set cel0101 = c
                Exit For
            End If
        End If
    Next c

    ' Se não encontrou 01/01, sai
    If cel0101 Is Nothing Then GoTo RestaurarVis

    ' Extrair "aa" (dois dígitos) e montar "20aa"
    ano2 = Right$(Trim$(CStr(cel0101.Value)), 2)
    If Not IsNumeric(ano2) Then GoTo RestaurarVis

    ano4 = CLng("20" & ano2)  ' Ex.: "25" -> 2025

    ' Atualiza A1 com o ano como número
    With ws.Range("A1")
        .Value = ano4
        .NumberFormat = "0"
    End With

    ' Se a linha encontrada não é a primeira (4), excluir da 4 até a anterior
    If Linha > primeiraLinha Then
        ws.Rows(primeiraLinha & ":" & (Linha - 1)).Delete Shift:=xlUp
    End If

RestaurarVis:
    ws.Visible = vis
End Sub


