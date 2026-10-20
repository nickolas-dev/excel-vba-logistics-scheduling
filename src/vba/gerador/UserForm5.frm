Attribute VB_Name = "UserForm5"
Attribute VB_Base = "0{3A177AD5-26A5-485D-BE06-FD0E38111660}{77BD4B26-3F46-47D0-9CD1-3B2843B0CDA7}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub UserForm_Click()

End Sub

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
        Case 0 To 20
            lblEtapas.Caption = "Convertendo os dados da Preliminar Definitiva..."
        Case 21 To 40
            lblEtapas.Caption = "Calculando fórmulas..."
        Case 41 To 60
            lblEtapas.Caption = "Analisando dados de carregamento e de entregas..."
        Case 61 To 80
            lblEtapas.Caption = "Consolidando as informações..."
        Case 81 To 100
            lblEtapas.Caption = "Consolidando as informações..."
    End Select

    DoEvents
End Sub

