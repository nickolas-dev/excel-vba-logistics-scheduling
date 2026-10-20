Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_Open()
    ' Exemplo oficial permanente: sem contador, procura da pasta ou expiracao.
    ' Para executar o fluxo, use o gerador na pasta app.
End Sub

Private Sub Workbook_SheetChange( _
    ByVal Sh As Object, _
    ByVal Target As Range)

    Dim cel As Range

    If Sh.Name <> "Planilha1" Then Exit Sub

    On Error GoTo Fim

    Application.EnableEvents = False

    Application.Calculate

    For Each cel In Target.Cells

        Select Case cel.Column

            ' LOJA
            Case 9

                AtualizarObsDuplicada

            ' J, K, M, N
            Case 10, 11, 13, 14

                AtualizarObsDTHCarreg cel.Row

                AtualizarObsAlertaCarregamento cel.Row

        End Select

    Next cel

Fim:

    Application.EnableEvents = True

End Sub
