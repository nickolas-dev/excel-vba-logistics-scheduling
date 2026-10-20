Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{C05564B4-33C7-41C1-95E3-7F9658B67874}{085021D9-F846-444A-8673-A3A33C9F82A7}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub ComboBox1_Change()

End Sub

Private Sub UserForm_Initialize()

        With ComboBox1
            .AddItem "Tbl_Distribuição + Template Infolog"
            .AddItem "Tbl_Distribuição"
            .AddItem "Template Infolog"
            ' Adicione mais conforme necessário
        End With
    
CommandButton1.SetFocus
    
End Sub

Private Sub CommandButton1_Click()

    Dim resposta As VbMsgBoxResult
    Dim destino As String

    destino = ComboBox1.Value

    If destino = "" Then
        MsgBox "Por favor, selecione uma opção antes de continuar.", vbCritical
        Exit Sub
    End If

    resposta = MsgBox("Deseja exportar as informações para " & destino & "?", vbYesNo + vbQuestion, "Confirmação")

    If resposta = vbYes Then
    
    Application.ScreenUpdating = False
    Application.Calculation = False
    
        Select Case destino
        Case "Tbl_Distribuição + Template Infolog"
            Unload Me
            Call Tbl_Distribuicao
            Call NovoTemplateInfolog
            Call PreencherTurnosPontoFocal(wbTemplateCriado)
        Case "Tbl_Distribuição"
            Unload Me
            Call Tbl_Distribuicao
        Case "Template Infolog"
            Unload Me
            Call NovoTemplateInfolog
        Case Else
        Unload Me
        MsgBox "Opção inválida.", vbCritical
End Select
        MsgBox "Exportação concluída para " & destino & ".", vbInformation
    Else
        MsgBox "Exportação cancelada.", vbCritical
        Unload Me
    End If
    
    Application.ScreenUpdating = True
    Application.Calculation = True
    
End Sub


