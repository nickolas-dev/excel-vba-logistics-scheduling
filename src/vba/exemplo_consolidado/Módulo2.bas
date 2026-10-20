Attribute VB_Name = "Módulo2"
Public wbTemplateCriado As Workbook
Public FiltroLojaNaoCadastrada As Boolean

Sub NovoTemplateInfolog()

    Dim wbOrigem As Workbook
    Dim wbModelo As Workbook
    Dim wbDestino As Workbook
    Dim wsOrigem As Worksheet
    Dim wsFormulas As Worksheet

    Dim caminhoModelo As String
    Dim pastaTemplates As String
    Dim caminhoDestino As String
    Dim nomeBase As String
    Dim nomeTemplateBase As String
    Dim nomeTemplateFinal As String

    Dim contador As Long
    Dim UltimaLinha As Long
    Dim ultimaLinhaFormulas As Long
    Dim linhaDestino As Long
    Dim i As Long

    Dim celulaValor As String
    Dim embarqueAtual As String
    Static ultimoEmbarque As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbOrigem = ThisWorkbook
    Set wsOrigem = wbOrigem.Sheets(1)

    ' === CAMINHOS ===
    caminhoModelo = PortfolioPath("modelo_template.xlsm")

    'pastaTemplates = PortfolioPath("Templates\")
    pastaTemplates = PortfolioPath("Templates\")

    ' Nome base do Template (sem contador)
    nomeBase = wbOrigem.Name
    nomeTemplateBase = Replace(nomeBase, "Base", "Template", , 1, vbTextCompare)

    ' === VERIFICAR SE JÁ EXISTE ===
    contador = 0
    nomeTemplateFinal = nomeTemplateBase
    caminhoDestino = pastaTemplates & nomeTemplateFinal

    Do While Dir(caminhoDestino) <> ""
        contador = contador + 1
        nomeTemplateFinal = Replace(nomeTemplateBase, ".xlsm", " (" & contador & ").xlsm")
        caminhoDestino = pastaTemplates & nomeTemplateFinal
    Loop

    ' === CRIAR CÓPIA DO MODELO ===
    Set wbModelo = Workbooks.Open(caminhoModelo, ReadOnly:=True)
    wbModelo.SaveCopyAs caminhoDestino
    wbModelo.Close False

    ' === ABRIR TEMPLATE ===
    Set wbDestino = Workbooks.Open(caminhoDestino)
    Set wbTemplateCriado = wbDestino
    Application.Run "'" & wbDestino.Name & "'!SetTemplateEmCarga", True
    Set wsFormulas = wbDestino.Sheets("Formulas")

    ' Limpar dados antigos
    wsFormulas.Range("2:" & wsFormulas.Rows.Count).ClearContents

    ' Formatação básica
    With wsFormulas.Range("A2:R" & wsFormulas.Rows.Count)
        .Font.Name = "Calibri"
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Última linha da origem
    UltimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "I").End(xlUp).Row
    linhaDestino = 2

    For i = 2 To UltimaLinha

        wsFormulas.Cells(linhaDestino, "R").Value = wsOrigem.Cells(i, "I").Value
        wsFormulas.Cells(linhaDestino, "Q").Value = wsOrigem.Cells(i, "F").Value
        wsFormulas.Cells(linhaDestino, "P").Value = wsOrigem.Cells(i, "W").Value
        wsFormulas.Cells(linhaDestino, "N").Value = wsOrigem.Cells(i, "O").Value
        wsFormulas.Cells(linhaDestino, "L").Value = wsOrigem.Cells(i, "L").Value
        wsFormulas.Cells(linhaDestino, "D").Value = wsOrigem.Cells(i, "P").Value
        
        'Caixas (para cálculo)
        If wsFormulas.Cells(linhaDestino, "M").Value <> "-" Then
            wsFormulas.Cells(linhaDestino, "V").Value = wsOrigem.Cells(i, "Q").Value
        Else
            wsFormulas.Cells(linhaDestino, "V").Value = TimeSerial(1, 0, 0)
        End If
        

        wsFormulas.Cells(linhaDestino, "J").Value = 1260
        wsFormulas.Cells(linhaDestino, "I").Value = 1257
        wsFormulas.Cells(linhaDestino, "H").Value = 1237
        wsFormulas.Cells(linhaDestino, "B").Value = "RASTREADO"

        ' Regra de DISTRIBUICAO por embarque
        embarqueAtual = Trim$(CStr(wsOrigem.Cells(i, "P").Text))
        celulaValor = Trim$(CStr(wsOrigem.Cells(i, "I").Text))

        If linhaDestino > 2 And embarqueAtual = ultimoEmbarque Then
            wsFormulas.Cells(linhaDestino, "C").Value = wsFormulas.Cells(linhaDestino - 1, "C").Value
        Else
            Select Case Len(celulaValor)
                Case 7
                    wsFormulas.Cells(linhaDestino, "C").Value = "DISTRIBUICAO FS"
                Case 3
                    wsFormulas.Cells(linhaDestino, "C").Value = "DISTRIBUICAO MCD"
                Case 4
                    wsFormulas.Cells(linhaDestino, "C").Value = wsFormulas.Cells(linhaDestino - 1, "C").Value
                Case Else
                    wsFormulas.Cells(linhaDestino, "C").Value = ""
            End Select
        End If

        ultimoEmbarque = embarqueAtual
        linhaDestino = linhaDestino + 1

    Next i

    ' === VALIDAÇÃO DA COLUNA M ===
    Dim Celula As Range
    Dim erroEncontrado As Boolean

    ultimaLinhaFormulas = wsFormulas.Cells(wsFormulas.Rows.Count, "R").End(xlUp).Row

    For Each Celula In wsFormulas.Range("M2:M" & ultimaLinhaFormulas)
        If Trim(Celula.Value) = "-" Then
            erroEncontrado = True
            Exit For
        End If
    Next Celula

    If erroEncontrado Then
        MsgBox "Por favor verificar registros com loja não cadastrada.", _
               vbCritical, "Erro de validação"
    
        ' Aplicar filtro na coluna M
        wsFormulas.Range("A1:R" & ultimaLinhaFormulas).AutoFilter
        wsFormulas.Range("A1:R" & ultimaLinhaFormulas).AutoFilter Field:=13, Criteria1:="-"
        
        FiltroLojaNaoCadastrada = True
    
        ' === Abrir arquivo de dados de lojas ===
        Dim wbLojas As Workbook
        Dim wsLojas As Worksheet
        Dim caminhoLojas As String
        Dim ultimaLinhaLojas As Long
    
        caminhoLojas = PortfolioPath("Tbl_Dados_Lojas_Transportador\Tbl_Dados_Lojas_Transportador.xlsx")
    
        ' Verifica se já está aberto
        On Error Resume Next
        Set wbLojas = Workbooks("Tbl_Dados_Lojas_Transportador.xlsx")
        On Error GoTo 0
    
        If wbLojas Is Nothing Then
            Set wbLojas = Workbooks.Open(caminhoLojas)
        End If
    
        Set wsLojas = wbLojas.Sheets("Dados_Lojas")
    
        wsLojas.Activate
    
        ' Encontrar próxima linha vazia na coluna A
        ultimaLinhaLojas = wsLojas.Cells(wsLojas.Rows.Count, "A").End(xlUp).Row + 1
    
        wsLojas.Cells(ultimaLinhaLojas, "A").Select
    
    End If
    
    Application.Run "'" & wbDestino.Name & "'!Atualizar_Formulas_Externas", wsFormulas, True
    Application.Run "'" & wbDestino.Name & "'!Atualizar_Ativacao", wsFormulas, wbDestino.Sheets("Ativação")
    
    Application.Run "'" & wbDestino.Name & "'!SetTemplateEmCarga", False
    
    'Limpando caixas
    'Range("V2:V" & ultimaLinhaFormulas).ClearContents

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Public Sub PreencherTurnosPontoFocal(wbTemplate As Workbook)

    Dim wsTemplate As Worksheet
    Dim wbPF As Workbook
    Dim wsPF As Worksheet
    Dim dictPF As Object

    Dim caminhoBase As String
    Dim caminhoArquivo As String
    Dim anoFat As String
    Dim semanaFat As String
    Dim dataFat As Date
    Dim dataTexto As String
    Dim cdFat As String

    Dim ultimaLinhaPF As Long
    Dim ultimaLinhaTemplate As Long
    Dim i As Long
    Dim Chave As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wsTemplate = wbTemplate.Worksheets("Formulas")
    Set dictPF = CreateObject("Scripting.Dictionary")

    ' =====================================================
    ' Extrair informações do nome do Template
    ' Ex: Template FT - 23.03.26.xlsm
    ' =====================================================
    cdFat = Mid(wbTemplate.Name, 10, 2)
    dataTexto = Mid(wbTemplate.Name, InStr(wbTemplate.Name, "-") + 2, 8)

    dataFat = DateSerial(2000 + Right(dataTexto, 2), Mid(dataTexto, 4, 2), Left(dataTexto, 2))
    anoFat = Year(dataFat)
    semanaFat = "Sem" & Format(DatePart("ww", dataFat, vbMonday, vbFirstFourDays), "00")

    ' =====================================================
    ' Montar caminho do arquivo Tbl_PontoFocal
    ' =====================================================
    caminhoBase = PortfolioPath("Tbl_PontoFocal\")
    caminhoArquivo = caminhoBase & anoFat & "\" & semanaFat & _
                     "\Tbl_PontoFocal_Fat" & Format(dataFat, "dd_mm_yyyy") & ".xlsx"

    ' =====================================================
    ' Abrir arquivo de Ponto Focal
    ' =====================================================
    Set wbPF = Workbooks.Open(caminhoArquivo, ReadOnly:=True)
    Set wsPF = wbPF.Sheets(1)

    ultimaLinhaPF = wsPF.Cells(wsPF.Rows.Count, "A").End(xlUp).Row

    ' =====================================================
    ' Carregar dados em Dictionary
    ' =====================================================
    For i = 2 To ultimaLinhaPF
        Chave = wsPF.Cells(i, "A").Value & _
                wsPF.Cells(i, "C").Value & _
                Format(wsPF.Cells(i, "B").Value, "dd/mm/yyyy")

        dictPF(Chave) = Array( _
            wsPF.Cells(i, "D").Value, _
            wsPF.Cells(i, "E").Value, _
            wsPF.Cells(i, "F").Value)
    Next i

    wbPF.Close False

    ' =====================================================
    ' Preencher Turnos no Template
    ' =====================================================
    ultimaLinhaTemplate = wsTemplate.Cells(wsTemplate.Rows.Count, "D").End(xlUp).Row

    For i = 2 To ultimaLinhaTemplate
        Chave = wsTemplate.Cells(i, "D").Value

        If dictPF.Exists(Chave) Then
            wsTemplate.Cells(i, "S").Value = dictPF(Chave)(0)
            wsTemplate.Cells(i, "T").Value = dictPF(Chave)(1)
            wsTemplate.Cells(i, "U").Value = dictPF(Chave)(2)
        End If
        
        'LEANDRO -> VENANCIO
        If wsTemplate.Cells(i, "U").Value = "LEANDRO" Then
        
            wsTemplate.Cells(i, "U").Value = "VENANCIO"
        
        End If
           
    Next i

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    If FiltroLojaNaoCadastrada Then
        With wsTemplate
            If Not .AutoFilterMode Then
                .Range("A1:R" & .Cells(.Rows.Count, "R").End(xlUp).Row).AutoFilter
            End If
    
            .Range("A1:R" & .Cells(.Rows.Count, "R").End(xlUp).Row) _
                .AutoFilter Field:=13, Criteria1:="-"
        End With
    End If

End Sub
