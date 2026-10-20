Attribute VB_Name = "UserForm6"
Attribute VB_Base = "0{340636C8-BCEF-4E5C-8ECF-F2D02D6D6143}{E2095ED4-8F42-4521-A5CA-FB3F2C8E12F7}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CommandButton1_Click()

Unload UserForm6

CriarTemplateInfologModelo

End Sub

Private Sub CommandButton2_Click()

    Unload UserForm6

    Dim fd As FileDialog
    Dim caminhoTemplate As String
    Dim wbTemplate As Workbook
    Dim wsTemplate As Worksheet

    Dim ultimaLinhaTemplate As Long
    Dim i As Long

    Dim chaveTemplate As String
    Dim dataTexto As String
    Dim DataFat As Date
    Dim AnoFat As String
    Dim semanaFat As String

    Dim caminhoPF As String
    Dim wbPF As Workbook
    Dim wsPF As Worksheet

    Dim dictPF As Object
    Dim ultimaLinhaPF As Long
    Dim j As Long
    Dim chavePF As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' ================= FILE DIALOG =================
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Selecionar Template Infolog"
        .InitialFileName = PortfolioPath("Templates\")
        .Filters.Clear
        .Filters.Add "Templates", "*.xlsm"
        .AllowMultiSelect = False

        If .Show <> -1 Then
            Unload Me
            GoTo Fim
        End If

        caminhoTemplate = .SelectedItems(1)
    End With

    ' ================= ABRIR TEMPLATE =================
    Set wbTemplate = Workbooks.Open(caminhoTemplate)
    Set wsTemplate = wbTemplate.Worksheets("Formulas")

    ultimaLinhaTemplate = wsTemplate.Cells(wsTemplate.Rows.Count, "D").End(xlUp).Row

    ' ================= PROCESSAR LINHAS =================
    For i = 2 To ultimaLinhaTemplate

        chaveTemplate = Trim(wsTemplate.Cells(i, "D").Value)

        If Len(chaveTemplate) < 10 Then GoTo ProximaLinha

        ' Extrair data (últimos 10 caracteres)
        dataTexto = Right(chaveTemplate, 10)
        If Not IsDate(dataTexto) Then GoTo ProximaLinha

        DataFat = CDate(dataTexto)
        AnoFat = Year(DataFat)
        semanaFat = "Sem" & Format(DatePart("ww", DataFat, vbMonday, vbFirstFourDays), "00")

        ' Caminho do Ponto Focal
        caminhoPF = PortfolioPath("Tbl_PontoFocal\") & _
                    AnoFat & "\" & semanaFat & "\Tbl_PontoFocal_Fat" & _
                    Format(DataFat, "dd_mm_yyyy") & ".xlsx"

        If Dir(caminhoPF) = "" Then GoTo ProximaLinha

        ' Abrir PF
        Set wbPF = Workbooks.Open(caminhoPF, ReadOnly:=True)
        Set wsPF = wbPF.Sheets(1)
        Set dictPF = CreateObject("Scripting.Dictionary")

        ultimaLinhaPF = wsPF.Cells(wsPF.Rows.Count, "A").End(xlUp).Row

        ' Carregar PF em Dictionary
        For j = 2 To ultimaLinhaPF

            chavePF = _
                wsPF.Cells(j, "A").Value & _
                wsPF.Cells(j, "C").Value & _
                Format(wsPF.Cells(j, "B").Value, "dd/mm/yyyy")

            If Not dictPF.Exists(chavePF) Then
                dictPF(chavePF) = Array( _
                    wsPF.Cells(j, "D").Value, _
                    wsPF.Cells(j, "E").Value, _
                    wsPF.Cells(j, "F").Value)
            End If

        Next j

        wbPF.Close False

        ' Preencher S:T:U
        If dictPF.Exists(chaveTemplate) Then
            wsTemplate.Cells(i, "S").Value = dictPF(chaveTemplate)(0)
            wsTemplate.Cells(i, "T").Value = dictPF(chaveTemplate)(1)
            wsTemplate.Cells(i, "U").Value = dictPF(chaveTemplate)(2)
        End If
        
        'LEANDRO -> VENANCIO
        If wsTemplate.Cells(i, "U").Value = "LEANDRO" Then
        
            wsTemplate.Cells(i, "U").Value = "VENANCIO"
        
        End If

ProximaLinha:
    Next i

    MsgBox "Pontos focais carregados com sucesso.", _
           vbInformation, "Control Tower MB"

Fim:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Unload Me

End Sub

Private Sub UserForm_Initialize()
    CommandButton2.Caption = "Carregar Ponto Focal" & vbCrLf & " para o Template"
    
    CommandButton1.SetFocus
    
End Sub

