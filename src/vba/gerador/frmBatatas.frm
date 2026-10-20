Attribute VB_Name = "frmBatatas"
Attribute VB_Base = "0{B77F8082-1004-4F67-A865-CCF4B31B1763}{9CBEB97C-DF44-417D-90B7-4C4A9536DFD2}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private UltimaDataSelecionada As Date

Private Sub cmdCriarProgramacao_Click()

    Application.ScreenUpdating = False
    
    Call Etapa1Batatas
    Call Etapa2Batatas
    Call Etapa3Batatas
    Call Etapa4Batatas
    Call Etapa5Batatas
    Call Etapa6Batatas
    Call Etapa7Batatas

    Application.ScreenUpdating = True

End Sub

Private Sub imgCalendario_Click()

    DataAnteriorDP = gDP_SelectedDate

    On Error Resume Next
    DP_Close
    On Error GoTo 0

    DP_Show

End Sub

Private Sub UserForm_Initialize()

    txtDataFat.Value = Format(Date, "dd/mm/yyyy")
    txtDataFat.Locked = True
    txtDataFat.TabStop = False
    
    'Batatas
    optBatatasSim.Value = True

End Sub

Public Function PrefixoDiaSemana() As String

    Select Case Weekday(CDate(txtDataFat.Value), vbSunday)

        Case 1
            PrefixoDiaSemana = "0"

        Case 2
            PrefixoDiaSemana = "2"

        Case 3
            PrefixoDiaSemana = "3"

        Case 4
            PrefixoDiaSemana = "4"

        Case 5
            PrefixoDiaSemana = "5"

        Case 6
            PrefixoDiaSemana = "6"

        Case 7
            PrefixoDiaSemana = "7"

    End Select

End Function

Private Sub AtualizarStatusCampo( _
    ByVal Valor As Long, _
    ByVal LabelControle As MSForms.Label, _
    ByVal TextControle As MSForms.TextBox)

    If Valor = 0 Then

        LabelControle.ForeColor = vbRed
        TextControle.ForeColor = vbRed

    Else

        LabelControle.ForeColor = vbBlack
        TextControle.ForeColor = vbBlack

    End If

End Sub

Private Sub spnBatatasxFT_SpinUp()

    txtBatatasxFT.Value = _
        Format(CLng(txtBatatasxFT.Value) + 1, "00")
        
    AtualizarStatusBatatas
    AtualizarResumoBatatas

End Sub

Private Sub spnBatatasxFT_SpinDown()

    If CLng(txtBatatasxFT.Value) > 0 Then

        txtBatatasxFT.Value = _
            Format(CLng(txtBatatasxFT.Value) - 1, "00")

    End If
    
    ValidarBatatasAtivas
    AtualizarStatusBatatas
    AtualizarResumoBatatas

End Sub

Private Sub spnBatatasxJC_SpinUp()

    txtBatatasxJC.Value = _
        Format(CLng(txtBatatasxJC.Value) + 1, "00")
        
    AtualizarStatusBatatas
    AtualizarResumoBatatas

End Sub

Private Sub spnBatatasxJC_SpinDown()

    If CLng(txtBatatasxJC.Value) > 0 Then

        txtBatatasxJC.Value = _
            Format(CLng(txtBatatasxJC.Value) - 1, "00")

    End If
    
    ValidarBatatasAtivas
    AtualizarStatusBatatas
    AtualizarResumoBatatas

End Sub

Private Sub spnBatatasxRJ_SpinUp()

    txtBatatasxRJ.Value = _
        Format(CLng(txtBatatasxRJ.Value) + 1, "00")
        
    AtualizarStatusBatatas
    AtualizarResumoBatatas

End Sub

Private Sub spnBatatasxRJ_SpinDown()

    If CLng(txtBatatasxRJ.Value) > 0 Then

        txtBatatasxRJ.Value = _
            Format(CLng(txtBatatasxRJ.Value) - 1, "00")

    End If
    
    ValidarBatatasAtivas
    AtualizarStatusBatatas
    AtualizarResumoBatatas


End Sub

Public Sub AtualizarStatusBatatas()

    If Not fraBatatas.Enabled Then Exit Sub
    
    'FT

    If CLng(txtBatatasxFT.Value) = 0 Then

        lblBatatasxFT.ForeColor = vbRed
        txtBatatasxFT.ForeColor = vbRed

    Else

        lblBatatasxFT.ForeColor = vbBlack
        txtBatatasxFT.ForeColor = vbBlack

    End If

    'JC

    If CLng(txtBatatasxJC.Value) = 0 Then

        lblBatatasxJC.ForeColor = vbRed
        txtBatatasxJC.ForeColor = vbRed

    Else

        lblBatatasxJC.ForeColor = vbBlack
        txtBatatasxJC.ForeColor = vbBlack

    End If

    'RJ

    If CLng(txtBatatasxRJ.Value) = 0 Then

        lblBatatasxRJ.ForeColor = vbRed
        txtBatatasxRJ.ForeColor = vbRed

    Else

        lblBatatasxRJ.ForeColor = vbBlack
        txtBatatasxRJ.ForeColor = vbBlack

    End If
    
    If CLng(txtBatatasxFT.Value) = 0 _
        And CLng(txtBatatasxJC.Value) = 0 _
        And CLng(txtBatatasxRJ.Value) = 0 Then
        
            cmdCriarProgramacao.Enabled = False
        
        Else
        
            cmdCriarProgramacao.Enabled = True
        
    End If


End Sub

Private Sub AtualizarFrameBatatas(ByVal Ativar As Boolean)

    fraBatatas.Enabled = Ativar

    If Ativar Then

        lblBatatasxFT.ForeColor = vbBlack
        lblBatatasxJC.ForeColor = vbBlack
        lblBatatasxRJ.ForeColor = vbBlack

        txtBatatasxFT.BackColor = vbWhite
        txtBatatasxJC.BackColor = vbWhite
        txtBatatasxRJ.BackColor = vbWhite

    Else

        lblBatatasxFT.ForeColor = RGB(150, 150, 150)
        lblBatatasxJC.ForeColor = RGB(150, 150, 150)
        lblBatatasxRJ.ForeColor = RGB(150, 150, 150)

        txtBatatasxFT.BackColor = RGB(240, 240, 240)
        txtBatatasxJC.BackColor = RGB(240, 240, 240)
        txtBatatasxRJ.BackColor = RGB(240, 240, 240)

        txtBatatasxFT.ForeColor = RGB(120, 120, 120)
        txtBatatasxJC.ForeColor = RGB(120, 120, 120)
        txtBatatasxRJ.ForeColor = RGB(120, 120, 120)

    End If

End Sub

Private Sub optBatatasNao_Click()

    txtBatatasxFT.Value = "00"
    txtBatatasxJC.Value = "00"
    txtBatatasxRJ.Value = "00"

    AtualizarFrameBatatas False

    cmdCriarProgramacao.Enabled = False
    
    AtualizarResumoBatatas

End Sub

Private Sub optBatatasSim_Click()

    AtualizarFrameBatatas True

    lblBatatasxFT.ForeColor = vbBlack
    lblBatatasxJC.ForeColor = vbBlack
    lblBatatasxRJ.ForeColor = vbBlack

    txtBatatasxFT.ForeColor = vbBlack
    txtBatatasxJC.ForeColor = vbBlack
    txtBatatasxRJ.ForeColor = vbBlack

    cmdCriarProgramacao.Enabled = True
    
    AtualizarStatusBatatas
    AtualizarResumoBatatas

End Sub

Private Sub ValidarBatatasAtivas()

    If CLng(txtBatatasxFT.Value) = 0 _
    And CLng(txtBatatasxJC.Value) = 0 _
    And CLng(txtBatatasxRJ.Value) = 0 Then

        optBatatasNao.Value = True

    End If

End Sub

Public Sub AtualizarResumoBatatas()

    Dim Resumo As String
    Dim i As Long

    If optBatatasNao.Value Then

        txtResumo.Value = ""

        Exit Sub

    End If

    If Not IsDate(txtDataFat.Value) Then

        txtResumo.Value = ""

        Exit Sub

    End If

    If CLng(txtBatatasxFT.Value) > 0 _
    Or CLng(txtBatatasxJC.Value) > 0 _
    Or CLng(txtBatatasxRJ.Value) > 0 Then

        Resumo = _
            "• Batatas:" & _
            vbCrLf & vbCrLf

    End If

    '=========================
    ' FT
    '=========================

    For i = 1 To CLng(txtBatatasxFT.Value)

        Resumo = Resumo & _
            "   - BATATAFT" & _
            PrefixoDiaSemana & i & _
            " (" & txtDataFat.Value & ")" & _
            vbCrLf

    Next i

    If CLng(txtBatatasxFT.Value) > 0 Then

        Resumo = Resumo & vbCrLf

    End If

    '=========================
    ' JC
    '=========================

    For i = 1 To CLng(txtBatatasxJC.Value)

        Resumo = Resumo & _
            "   - BATATAJC" & _
            PrefixoDiaSemana & i & _
            " (" & txtDataFat.Value & ")" & _
            vbCrLf

    Next i

    If CLng(txtBatatasxJC.Value) > 0 Then

        Resumo = Resumo & vbCrLf

    End If

    '=========================
    ' RJ
    '=========================

    For i = 1 To CLng(txtBatatasxRJ.Value)

        Resumo = Resumo & _
            "   - BATATARJ" & _
            PrefixoDiaSemana & i & _
            " (" & txtDataFat.Value & ")" & _
            vbCrLf

    Next i

    txtResumo.Value = Trim(Resumo)

End Sub
