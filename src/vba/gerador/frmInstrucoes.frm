Attribute VB_Name = "frmInstrucoes"
Attribute VB_Base = "0{216716F1-07D8-4506-9D1A-CE120C39E55A}{2BE4A6BD-D2FB-49CB-897C-472196DE2734}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub UserForm_Initialize()
    
    Me.Picture = LoadPicture(PortfolioPath("assets\fundo_userform.jpg"), 240, 180)
    Me.PictureSizeMode = fmPictureSizeModeStretch
    
    Label1.BackStyle = fmBackStyleTransparent

    optEtapa1.Value = True
    Call AtualizarTexto
    

End Sub

Private Sub optEtapa1_Click()
    Call AtualizarTexto
End Sub

Private Sub optEtapa2_Click()
    Call AtualizarTexto
End Sub

Private Sub AtualizarTexto()
    If optEtapa1.Value = True Then
        txtInstrucoes.Text = _
            "1. Abra o PDF da Preliminar Definitiva no aplicativo Acrobat Reader." & vbCrLf & vbCrLf & _
            "2. Selecione e copie os dados (CTRL+A e CTRL+C)." & vbCrLf & vbCrLf & _
            "3. Cole como TEXTO na célula A2." & vbCrLf & vbCrLf & _
            "4. Clique em 'Criar'."
            
    ElseIf optEtapa2.Value = True Then
        txtInstrucoes.Text = _
            "1. Confirme se deseja criar um novo arquivo Tbl_Prog." & vbCrLf & vbCrLf & _
            "2. Selecione se gostaria de importar dados ao criar ou seguir sem importar dados." & vbCrLf & vbCrLf & _
            "   Importando dados:" & vbCrLf & _
            "   - Selecione um arquivo (outra Tbl_Prog) para carregar as informações." & vbCrLf & vbCrLf & _
            "   Sem importar dados:" & vbCrLf & _
            "   - Será gerado um arquivo Tbl_Prog_Modelo somente com a formatação, as fórmulas e as validações de dados." & vbCrLf & vbCrLf & _
            "3. A nova Tbl_Prog será salva na mesma pasta deste arquivo."
    End If
End Sub


Private Sub btnTutorial_Click()
    MsgBox "O video de tutorial nao esta incluido nesta demonstracao.", vbInformation, "Tutorial"
End Sub


