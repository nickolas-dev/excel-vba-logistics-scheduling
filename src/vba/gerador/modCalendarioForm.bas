Attribute VB_Name = "modCalendarioForm"
Public DataAnteriorDP As Date

Option Explicit

Public Sub AbrirFrmCross()

    If frmPaniColetas.Visible Then

        MsgBox _
            "Feche o formulário Panifresh e Coletas antes de abrir Cross.", _
            vbExclamation

        Exit Sub

    End If

    Unload frmCross

    frmCross.Show vbModeless

End Sub
