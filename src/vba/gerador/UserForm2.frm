Attribute VB_Name = "UserForm2"
Attribute VB_Base = "0{B7C4931A-88CD-48D7-8202-1DB02E82C1D6}{C4EDB43F-660D-4CF4-B33F-7BFC342A7F8E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub UserForm_Initialize()

Me.Picture = LoadPicture(PortfolioPath("assets\fundo_userform.jpg"), 240, 180)
Me.PictureSizeMode = fmPictureSizeModeStretch

lblEtapas.BackStyle = fmBackStyleTransparent
lblPercent.BackStyle = fmBackStyleTransparent

Call UpdateProgressBar(0)

End Sub

Public Sub UpdateProgressBar(percent As Integer)
    lblProgress.Width = (percent / 100) * fraProgress.Width
    lblPercent.Caption = percent & "%"
    lblPercent.Left = (fraProgress.Width - lblPercent.Width) / 2

    ' Atualiza o texto de etapas
    Select Case percent
        Case 0 To 25
            lblEtapas.Caption = "Criando um novo arquivo..."
        Case 26 To 50
            lblEtapas.Caption = "Formatando linhas e colunas..."
        Case 51 To 74
            lblEtapas.Caption = "Inserindo fórmulas..."
        Case 75 To 90
            lblEtapas.Caption = "Ajustando validações de dados..."
        Case 91 To 100
            lblEtapas.Caption = "Concluindo..."
    End Select

    DoEvents
End Sub
