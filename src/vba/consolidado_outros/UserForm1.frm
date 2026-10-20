Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{635E2E83-32B5-4472-88A9-3C385EB29084}{2931E639-B0C2-4059-94B5-B64453B16070}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub UserForm_Initialize()

        With ComboBox1
            .AddItem "Tbl_Prog + Template Infolog"
            .AddItem "Tbl_Prog"
            .AddItem "Template Infolog"
            ' Adicione mais conforme necessário
        End With
    
CommandButton1.SetFocus
    
End Sub

Private Sub CommandButton1_Click() 'Botao OK

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
                Case "Tbl_Prog + Template Infolog"
                    Unload Me
                    Call Tbl_Programacao
                    Call TemplateInfologOUTROS
                    Call PreencherTurnosPontoFocal(wbTemplateCriado)
                Case "Tbl_Prog"
                    Unload Me
                    Call Tbl_Programacao
                Case "Template Infolog"
                    Unload Me
                    Call TemplateInfologOUTROS
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


