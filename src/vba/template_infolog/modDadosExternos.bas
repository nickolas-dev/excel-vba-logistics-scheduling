Attribute VB_Name = "modDadosExternos"
Option Explicit

Public dictLojas As Object
Public dictTransportadores As Object
Public chaveDescarga As String

Public Sub Carregar_Dados_Externos()

    Dim wbBase As Workbook
    Dim wsLojas As Worksheet
    Dim wsTransp As Worksheet
    Dim caminho As String
    Dim i As Long
    Dim UltimaLinha As Long

    Set dictLojas = CreateObject("Scripting.Dictionary")
    Set dictTransportadores = CreateObject("Scripting.Dictionary")

    ' Caminho FIXO no Template
    caminho = PortfolioPath("Tbl_Dados_Lojas_Transportador\Tbl_Dados_Lojas_Transportador.xlsx")
    'caminho = PortfolioPath("Tbl_Dados_Lojas_Transportador\Tbl_Dados_Lojas_Transportador.xlsx")

    Application.ScreenUpdating = False

    Set wbBase = Workbooks.Open(caminho, ReadOnly:=True)

    ' -------- LOJAS (Sheets(1)) --------
    Set wsLojas = wbBase.Worksheets(1)
    UltimaLinha = wsLojas.Cells(wsLojas.Rows.Count, "A").End(xlUp).Row

    For i = 2 To UltimaLinha
        'dictLojas(wsLojas.Cells(i, "A").Value) = wsLojas.Cells(i, "B").Value
        
        chaveDescarga = UCase(Trim(wsLojas.Cells(i, "A").Value))
        
        dictLojas(chaveDescarga) = Array(wsLojas.Cells(i, "B").Value, wsLojas.Cells(i, "C").Value, wsLojas.Cells(i, "D").Value)
    Next i

    ' -------- TRANSPORTADORES (Sheets(2)) --------
    Set wsTransp = wbBase.Worksheets(2)
    UltimaLinha = wsTransp.Cells(wsTransp.Rows.Count, "A").End(xlUp).Row

    For i = 2 To UltimaLinha
        dictTransportadores(wsTransp.Cells(i, "A").Value) = wsTransp.Cells(i, "B").Value
    Next i

    wbBase.Close False
    Application.ScreenUpdating = True

End Sub

Public Sub Atualizar_Formulas_Externas(ws As Worksheet, Optional ForcarRecarregar As Boolean = False)

    If ForcarRecarregar Or dictLojas Is Nothing Or dictTransportadores Is Nothing Then
        Carregar_Dados_Externos
    End If

    Dim i As Long
    Dim UltimaLinha As Long

    ' Garante que os dados externos já foram carregados
    If dictLojas Is Nothing Or dictTransportadores Is Nothing Then
        Carregar_Dados_Externos
    End If

    UltimaLinha = Application.WorksheetFunction.Max( _
    ws.Cells(ws.Rows.Count, "P").End(xlUp).Row, _
    ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row, _
    ws.Cells(ws.Rows.Count, "R").End(xlUp).Row)
    
    Dim tipo As String
    Dim caixas As Double
    Dim caixasPadrao As Double
    Dim taxa As Double
    Dim tempo As Double

    For i = 2 To UltimaLinha

        Dim p As String, q As String, r As String
  
        p = Trim(ws.Cells(i, "P").Value)
        q = Trim(ws.Cells(i, "Q").Value)
        r = Trim(ws.Cells(i, "R").Value)
        
        ' ============================================
        ' MARCA SE LINHA JÁ ESTEVE COMPLETA
        ' ============================================
        If p <> "" And q <> "" And r <> "" Then
            ws.Cells(i, "Z").Value = "OK"
        End If
        
        ' ============================================
        ' REGRA DE LIMPEZA ENTRE P, Q, R
        ' ============================================
        
        If ws.Cells(i, "Z").Value = "OK" Then
        
            If r = "" And (q <> "" Or p <> "") Then
                ws.Cells(i, "P").ClearContents
                ws.Cells(i, "Q").ClearContents
            
            ElseIf q = "" And (p <> "" Or r <> "") Then
                ws.Cells(i, "P").ClearContents
                ws.Cells(i, "R").ClearContents
            
            ElseIf p = "" And (q <> "" Or r <> "") Then
                ws.Cells(i, "Q").ClearContents
                ws.Cells(i, "R").ClearContents
            
            End If
            
            ' Atualiza variáveis depois da limpeza
            p = Trim(ws.Cells(i, "P").Value)
            q = Trim(ws.Cells(i, "Q").Value)
            r = Trim(ws.Cells(i, "R").Value)
            
        End If

        ' ============================================
        ' REGRA FORTE:
        ' Se APAGAR qualquer P, Q ou R
        ' ? limpa A, K, M e B, H, I, J
        ' ============================================
        If p = "" Or q = "" Or r = "" Then
            ws.Cells(i, "A").ClearContents
            ws.Cells(i, "K").ClearContents
            ws.Cells(i, "M").ClearContents
    
            ws.Cells(i, "B").ClearContents
            ws.Cells(i, "H").ClearContents
            ws.Cells(i, "I").ClearContents
            ws.Cells(i, "J").ClearContents
            ws.Cells(i, "O").ClearContents
            ws.Cells(i, "Z").ClearContents
    
            GoTo ProximaLinha
        End If
    
        ' ============================================
        ' A, K, M (dados externos)
        ' ============================================
        If dictTransportadores.Exists(p) Then
            ws.Cells(i, "A").Value = dictTransportadores(p)
        Else
            ws.Cells(i, "A").Value = "-"
        End If
    
        If dictLojas.Exists(q) Then
            ws.Cells(i, "K").Value = dictLojas(q)(0)
        Else
            ws.Cells(i, "K").Value = "-"
        End If
    
        If dictLojas.Exists(r) Then
            ws.Cells(i, "M").Value = dictLojas(r)(0)
        Else
            ws.Cells(i, "M").Value = "-"
        End If
    
        ' ============================================
        ' B, H, I, J (valores fixos)
        ' ============================================
        ws.Cells(i, "B").Value = "RASTREADO"
        ws.Cells(i, "H").Value = 1237
        ws.Cells(i, "I").Value = 1257
        ws.Cells(i, "J").Value = 1260
        
        ' ============================================
        ' CÁLCULO TEMPO DESCARGA (coluna O)
        ' ============================================
        
        chaveDescarga = r
        
        If dictLojas.Exists(chaveDescarga) Then
        
            tipo = Trim(dictLojas(chaveDescarga)(1))
            caixasPadrao = dictLojas(chaveDescarga)(2)
        
            ' Regra da coluna S
            If ws.Cells(i, "V").Value = "" Or ws.Cells(i, "V").Value = 0 Then
                caixas = caixasPadrao
            Else
                caixas = ws.Cells(i, "V").Value
            End If
        
            Select Case tipo
        
                Case "CD"
                    ws.Cells(i, "O").Value = TimeSerial(5, 0, 0)
        
                Case "Drive-Thru"
                    taxa = 230
        
                Case "In Store"
                    taxa = 180
        
                Case "Shopping", "Outros"
                    taxa = 140
        
                Case Else
                    taxa = 140 ' segurança
        
            End Select
        
            If tipo <> "CD" Then
                tempo = caixas / taxa / 24
                ws.Cells(i, "O").Value = tempo
            End If
        
        Else
            ws.Cells(i, "O").ClearContents
        End If
        
        ws.Cells(i, "O").NumberFormat = "hh:mm:ss"
        
    
ProximaLinha:
    Next i

End Sub
