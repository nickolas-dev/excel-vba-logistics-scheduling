Attribute VB_Name = "modIndicePontoFocal"
Public Sub InicializarIndicePontoFocal()

    'Produção
    'Const CAMINHO_PF As String = _
    PortfolioPath("Tbl_PontoFocal\")
    
    'Teste
    Dim CAMINHO_PF As String
    CAMINHO_PF = PortfolioPath("Tbl_PontoFocal\")

    Dim CAMINHO_INDICE As String

    CAMINHO_INDICE = PortfolioPath("Tbl_PontoFocal_Index.xlsx")

    Dim wbIndice As Workbook
    Dim wsIndice As Worksheet
    Dim ProximaLinha As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Abre o índice
    Set wbIndice = Workbooks.Open(CAMINHO_INDICE)
    Set wsIndice = wbIndice.Sheets("Index")

    ProximaLinha = wsIndice.Cells(wsIndice.Rows.Count, "A").End(xlUp).Row + 1

    ' Processa todas as subpastas recursivamente
    Call ProcessarPastaPF(CAMINHO_PF, wsIndice, ProximaLinha)

    wbIndice.Save
    wbIndice.Close

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Índice inicial de Ponto Focal criado com sucesso!", vbInformation

End Sub

Private Sub ProcessarPastaPF( _
    caminhoAtual As String, _
    wsIndice As Worksheet, _
    ByRef ProximaLinha As Long)

    Dim fso As Object
    Dim pasta As Object
    Dim subpasta As Object
    Dim Arquivo As Object

    Dim wbPF As Workbook
    Dim wsPF As Worksheet
    Dim UltLinha As Long
    Dim i As Long

    Dim Chave As String
    Dim DataFat As Date

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set pasta = fso.GetFolder(caminhoAtual)

    ' Primeiro: processa arquivos da pasta atual
    For Each Arquivo In pasta.Files

        If LCase(fso.GetExtensionName(Arquivo.Name)) = "xlsx" Then

            Set wbPF = Workbooks.Open(Arquivo.Path)
            Set wsPF = wbPF.Sheets(1)

            UltLinha = wsPF.Cells(wsPF.Rows.Count, "A").End(xlUp).Row

            For i = 2 To UltLinha

                If wsPF.Cells(i, "A").Value <> "" Then

                    DataFat = wsPF.Cells(i, "B").Value

                    Chave = _
                        wsPF.Cells(i, "A").Value & _
                        wsPF.Cells(i, "C").Value & _
                        Format(DataFat, "dd/mm/yyyy")

                    wsIndice.Cells(ProximaLinha, "A").Value = Chave
                    wsIndice.Cells(ProximaLinha, "B").Value = DataFat
                    ProximaLinha = ProximaLinha + 1

                End If

            Next i

            wbPF.Close SaveChanges:=False

        End If

    Next Arquivo

    ' Depois: entra recursivamente nas subpastas (Ano ? Semana ? etc.)
    For Each subpasta In pasta.SubFolders
        Call ProcessarPastaPF(subpasta.Path, wsIndice, ProximaLinha)
    Next subpasta

End Sub

Public Function CarregarIndicePontoFocal() As Object

    Dim CAMINHO_INDICE As String

    CAMINHO_INDICE = PortfolioPath("Tbl_PontoFocal_Index.xlsx")

    Dim wbIndice As Workbook
    Dim wsIndice As Worksheet
    Dim dict As Object
    Dim UltLinha As Long
    Dim i As Long

    Set dict = CreateObject("Scripting.Dictionary")

    Set wbIndice = Workbooks.Open(CAMINHO_INDICE)
    Set wsIndice = wbIndice.Sheets("Index")

    UltLinha = wsIndice.Cells(wsIndice.Rows.Count, "A").End(xlUp).Row

    For i = 2 To UltLinha
        dict(wsIndice.Cells(i, "A").Value) = True
    Next i

    wbIndice.Close SaveChanges:=False

    Set CarregarIndicePontoFocal = dict

End Function

Public Sub AdicionarChavesNoIndice(arrChaves As Variant, DataFat As Date)

    Dim CAMINHO_INDICE As String

    CAMINHO_INDICE = PortfolioPath("Tbl_PontoFocal_Index.xlsx")

    Dim wbIndice As Workbook
    Dim wsIndice As Worksheet
    Dim ProximaLinha As Long
    Dim i As Long

    Application.DisplayAlerts = False

    Set wbIndice = Workbooks.Open(CAMINHO_INDICE)
    Set wsIndice = wbIndice.Sheets("Index")

    ProximaLinha = wsIndice.Cells(wsIndice.Rows.Count, "A").End(xlUp).Row + 1

    For i = LBound(arrChaves) To UBound(arrChaves)
        wsIndice.Cells(ProximaLinha, "A").Value = arrChaves(i)
        wsIndice.Cells(ProximaLinha, "B").Value = DataFat
        ProximaLinha = ProximaLinha + 1
    Next i

    wbIndice.Save
    wbIndice.Close

    Application.DisplayAlerts = True

End Sub
