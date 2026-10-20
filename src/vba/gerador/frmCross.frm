Attribute VB_Name = "frmCross"
Attribute VB_Base = "0{4C58152E-DB82-4761-A47B-F80105F907BF}{E6A229DC-2BED-43A9-994B-6CDB0CDE38F1}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private UltimaDataSelecionada As Date

Private Sub imgCalendario_Click()

    DataAnteriorDP = gDP_SelectedDate

    On Error Resume Next
    DP_Close
    On Error GoTo 0

    DP_Show

End Sub

Private Sub optAmbos_Click()

    'Atualizar Resumo
    AtualizarResumo

End Sub

Private Sub optGR_Click()

    'Atualizar Resumo
    AtualizarResumo

End Sub

Private Sub optJC_Click()

    'Atualizar Resumo
    AtualizarResumo

End Sub

Private Sub optNao_Click()

    'Atualizar Resumo
    AtualizarResumo

End Sub

Private Sub spnFTxJD_Change()

    txtFTxJD.Value = Format(spnFTxJD.Value, "00")

    AtualizarIndicadoresCross

    If CLng(txtFTxJD.Value) = 0 And _
       CLng(txtGRxJD.Value) = 0 Then
    
        optCrossJDNao.Value = True
    
    Else
    
        optCrossJDSim.Value = True
    
    End If

    AtualizarResumo

End Sub

Private Sub spnGRxJD_Change()

    txtGRxJD.Value = Format(spnGRxJD.Value, "00")

    AtualizarIndicadoresCross

    If CLng(txtFTxJD.Value) = 0 And _
       CLng(txtGRxJD.Value) = 0 Then
    
        optCrossJDNao.Value = True
    
    Else
    
        optCrossJDSim.Value = True
    
    End If

    AtualizarResumo

End Sub

Private Sub spnNExCE_Change()

    txtNExCE.Value = Format(spnNExCE.Value, "00")

    AtualizarIndicadoresCross

    If CLng(txtNExCE.Value) = 0 And _
       CLng(txtNExPI.Value) = 0 Then
    
        optCrossNENao.Value = True
    
    Else
    
        optCrossNESim.Value = True
    
    End If

    AtualizarResumo

End Sub

Private Sub spnNExPI_Change()

    txtNExPI.Value = Format(spnNExPI.Value, "00")

    AtualizarIndicadoresCross

    If CLng(txtNExCE.Value) = 0 And _
       CLng(txtNExPI.Value) = 0 Then
    
        optCrossNENao.Value = True
    
    Else
    
        optCrossNESim.Value = True
    
    End If

    AtualizarResumo

End Sub

Private Sub spnPRxRS_Change()

    txtPRxRS.Value = Format(spnPRxRS.Value, "00")
    
    'Verificar Frames
    VerificarFramesZerados
    
    'Atualizar Resumo
    AtualizarResumo

End Sub

Private Sub UserForm_Initialize()

    txtDataFat.Value = Format(Date, "dd/mm/yyyy")
    txtDataFat.Locked = True
    txtDataFat.TabStop = False
    
    optAmbos.Value = True
    
    txtFTxJD.Value = "04"
    spnFTxJD.Value = 4
    
    txtGRxJD.Value = "02"
    spnGRxJD.Value = 2
    
    txtNExCE.Value = "01"
    spnNExCE.Value = 1
    
    txtNExPI.Value = "02"
    spnNExPI.Value = 2
    
    txtPRxRS.Value = "04"
    spnPRxRS.Value = 4
    
    'Cross JD
    optCrossJDSim.Value = True
    AtualizarFrameJD True
    
    'Cross NE
    optCrossNESim.Value = True
    AtualizarFrameNE True
    
    'Cross PR
    optCrossPRSim.Value = True
    AtualizarFramePR True
    
    'Atualizar Resumo
    AtualizarResumo
    
    'Atualizar Indicadores Cross
    AtualizarIndicadoresCross
    
    'Criar Programação
    AtualizarBotaoCriar

End Sub

Private Sub txtFTxJD_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii

        Case 48 To 57
            'permite 0 a 9

        Case 8
            'Backspace

        Case Else
            KeyAscii = 0

    End Select

End Sub

Private Sub txtFTxJD_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not fraCrossJD.Enabled Then Exit Sub

    If Trim(txtFTxJD.Value) = "" Then

        txtFTxJD.Value = "04"
        spnFTxJD.Value = 4

    Else

        txtFTxJD.Value = Format(CLng(txtFTxJD.Value), "00")
        spnFTxJD.Value = CLng(txtFTxJD.Value)

    End If
    
    If CLng(txtFTxJD.Value) = 0 And _
       CLng(txtGRxJD.Value) = 0 Then
    
        optCrossJDNao.Value = True
    
    End If
    
    'Verificar Frames
    VerificarFramesZerados
    
    'Atualizar Resumo
    AtualizarResumo
    
    'Atualizar Indicadores Cross
    AtualizarIndicadoresCross

End Sub

Private Sub txtGRxJD_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Trim(txtGRxJD.Value) = "" Then

        txtFTxJD.Value = "02"
        spnFTxJD.Value = 2

    Else

        txtFTxJD.Value = Format(CLng(txtFTxJD.Value), "00")
        spnFTxJD.Value = CLng(txtFTxJD.Value)

    End If
    
    If CLng(txtFTxJD.Value) = 0 And _
       CLng(txtGRxJD.Value) = 0 Then
    
        optCrossJDNao.Value = True
    
    End If
    
    'Verificar Frames
    VerificarFramesZerados

    'Atualizar Resumo
    AtualizarResumo
    
    'Atualizar Indicadores Cross
    AtualizarIndicadoresCross

End Sub

Private Sub txtGRxJD_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii

        Case 48 To 57
            'permite 0 a 9

        Case 8
            'Backspace

        Case Else
            KeyAscii = 0

    End Select

End Sub

Private Sub txtNExCE_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii

        Case 48 To 57
            'permite 0 a 9

        Case 8
            'Backspace

        Case Else
            KeyAscii = 0

    End Select
    
End Sub

Private Sub txtNExCE_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not fraCrossNE.Enabled Then Exit Sub

    If Trim(txtNExCE.Value) = "" Then

        txtNExCE.Value = "00"
        spnNExCE.Value = 0

    Else

        txtNExCE.Value = Format(CLng(txtNExCE.Value), "00")
        spnNExCE.Value = CLng(txtNExCE.Value)

    End If

    AtualizarIndicadoresCross

    If CLng(txtNExCE.Value) = 0 And _
       CLng(txtNExPI.Value) = 0 Then

        optCrossNENao.Value = True

    Else

        optCrossNESim.Value = True

    End If

    AtualizarResumo

End Sub

Private Sub txtNExPI_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii

        Case 48 To 57
            'permite 0 a 9

        Case 8
            'Backspace

        Case Else
            KeyAscii = 0

    End Select

End Sub

Private Sub txtNExPI_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not fraCrossNE.Enabled Then Exit Sub

    If Trim(txtNExPI.Value) = "" Then

        txtNExPI.Value = "00"
        spnNExPI.Value = 0

    Else

        txtNExPI.Value = Format(CLng(txtNExPI.Value), "00")
        spnNExPI.Value = CLng(txtNExPI.Value)

    End If

    AtualizarIndicadoresCross

    If CLng(txtNExCE.Value) = 0 And _
       CLng(txtNExPI.Value) = 0 Then

        optCrossNENao.Value = True

    Else

        optCrossNESim.Value = True

    End If

    AtualizarResumo

End Sub

Private Sub txtPRxRS_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii

        Case 48 To 57
            'permite 0 a 9

        Case 8
            'Backspace

        Case Else
            KeyAscii = 0

    End Select

End Sub

Private Sub txtPRxRS_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Trim(txtPRxRS.Value) = "" Then

        txtPRxRS.Value = "04"
        spnPRxRS.Value = 4

    Else

        txtPRxRS.Value = Format(CLng(txtPRxRS.Value), "00")
        spnPRxRS.Value = CLng(txtPRxRS.Value)

    End If
    
    'Verificar Frames
    VerificarFramesZerados
    
    'Atualizar Resumo
    AtualizarResumo

End Sub

Private Sub optCrossJDSim_Click()

    AtualizarFrameJD True

    spnFTxJD.Value = CLng(txtFTxJD.Value)
    spnGRxJD.Value = CLng(txtGRxJD.Value)

    AtualizarIndicadoresCross
    AtualizarResumo

End Sub

Private Sub optCrossJDNao_Click()

    txtFTxJD.Value = "00"
    txtGRxJD.Value = "00"

    AtualizarFrameJD False
    
    'Atualizar Resumo
    AtualizarResumo

End Sub

Private Sub optCrossNESim_Click()

    AtualizarFrameNE True

    spnNExCE.Value = CLng(txtNExCE.Value)
    spnNExPI.Value = CLng(txtNExPI.Value)

    AtualizarIndicadoresCross
    AtualizarResumo

End Sub

Private Sub optCrossNENao_Click()

    txtNExCE.Value = "00"
    txtNExPI.Value = "00"

    AtualizarFrameNE False
    
    'Atualizar Resumo
    AtualizarResumo
    
End Sub

Private Sub optCrossPRSim_Click()

    AtualizarFramePR True

    If txtPRxRS.Value = "00" Then txtPRxRS.Value = "04"
    
    'Atualizar Resumo
    AtualizarResumo
    
End Sub

Private Sub optCrossPRNao_Click()

    txtPRxRS.Value = "00"
    
    'Atualizar Resumo
    AtualizarResumo

    AtualizarFramePR False

End Sub

Private Sub AtualizarFrameJD(ByVal Ativo As Boolean)

    fraCrossJD.Enabled = Ativo

    If Ativo Then

        lblFTxJD.ForeColor = vbBlack
        lblGRxJD.ForeColor = vbBlack

        txtFTxJD.BackColor = vbWhite
        txtGRxJD.BackColor = vbWhite

        txtFTxJD.ForeColor = vbBlack
        txtGRxJD.ForeColor = vbBlack

    Else

        lblFTxJD.ForeColor = RGB(150, 150, 150)
        lblGRxJD.ForeColor = RGB(150, 150, 150)

        txtFTxJD.BackColor = RGB(240, 240, 240)
        txtGRxJD.BackColor = RGB(240, 240, 240)

        txtFTxJD.ForeColor = RGB(120, 120, 120)
        txtGRxJD.ForeColor = RGB(120, 120, 120)

    End If

End Sub

Private Sub AtualizarFrameNE(ByVal Ativo As Boolean)

    fraCrossNE.Enabled = Ativo

    If Ativo Then

        lblNExCE.ForeColor = vbBlack
        lblNExPI.ForeColor = vbBlack

        txtNExCE.BackColor = vbWhite
        txtNExPI.BackColor = vbWhite

        txtNExCE.ForeColor = vbBlack
        txtNExPI.ForeColor = vbBlack

    Else

        lblNExCE.ForeColor = RGB(150, 150, 150)
        lblNExPI.ForeColor = RGB(150, 150, 150)

        txtNExCE.BackColor = RGB(240, 240, 240)
        txtNExPI.BackColor = RGB(240, 240, 240)

        txtNExCE.ForeColor = RGB(120, 120, 120)
        txtNExPI.ForeColor = RGB(120, 120, 120)

    End If

End Sub

Private Sub AtualizarFramePR(ByVal Ativo As Boolean)

    fraCrossPR.Enabled = Ativo

    If Ativo Then

        lblPRxRS.ForeColor = vbBlack

        txtPRxRS.BackColor = vbWhite

        txtPRxRS.ForeColor = vbBlack

    Else

        lblPRxRS.ForeColor = RGB(150, 150, 150)

        txtPRxRS.BackColor = RGB(240, 240, 240)

        txtPRxRS.ForeColor = RGB(120, 120, 120)

    End If

End Sub

Public Sub AtualizarResumo()

    Dim DataFat As String

    DataFat = txtDataFat.Value

    Dim Resumo As String

    Resumo = "• Monta Pallet:" & vbCrLf

    If optAmbos.Value Then

        Resumo = "• Monta Pallet:" & vbCrLf & _
                 "   - JC (" & DataFat & ")" & vbCrLf & _
                 "   - GR (" & DataFat & ")"
    
    ElseIf optJC.Value Then
    
        Resumo = "• Monta Pallet:" & vbCrLf & _
                 "   - JC (" & DataFat & ")"
    
    ElseIf optGR.Value Then
    
        Resumo = "• Monta Pallet:" & vbCrLf & _
                 "   - GR (" & DataFat & ")"
    
    ElseIf optNao.Value Then
    
        Resumo = ""
    
    End If
    
    Dim Prefixo As String
    Dim i As Long
    
    Prefixo = PrefixoDiaSemana()
    
    'Cross JD
    If optCrossJDSim.Value And _
   (CLng(txtFTxJD.Value) > 0 Or CLng(txtGRxJD.Value) > 0) Then

        Resumo = Resumo & vbCrLf & vbCrLf & _
                 "• Cross JD:" & vbCrLf
    
        For i = 1 To CLng(txtFTxJD.Value)
        
            Resumo = Resumo & _
                     "   - TRJD" & _
                     Prefixo & Format(i, "0") & _
                     " (FTxJD)" & vbCrLf
        
        Next i
        
        Resumo = Resumo & vbCrLf
        
        For i = 1 To CLng(txtGRxJD.Value)
        
            Resumo = Resumo & _
                     "   - TRJD" & _
                     Prefixo & Format(i, "0") & _
                     " (GRxJD)" & vbCrLf
        
        Next i
    
    End If
    
    'Cross NE
    If optCrossNESim.Value And _
       (CLng(txtNExCE.Value) > 0 Or CLng(txtNExPI.Value) > 0) Then
    
        Resumo = Resumo & vbCrLf & vbCrLf & _
                 "• Cross NE:" & vbCrLf
    
        For i = 1 To CLng(txtNExCE.Value)
    
            Resumo = Resumo & _
                     "   - TRCE" & _
                     Prefixo & Format(i, "0") & _
                     " (NExCE)" & vbCrLf
    
        Next i
    
        Resumo = Resumo & vbCrLf
    
        For i = 1 To CLng(txtNExPI.Value)
    
            Resumo = Resumo & _
                     "   - TRPI" & _
                     Prefixo & Format(i, "0") & _
                     " (NExPI)" & vbCrLf
    
        Next i
    
    End If
    
    'Cross PR
    If optCrossPRSim.Value And _
       CLng(txtPRxRS.Value) > 0 Then
    
        Resumo = Resumo & vbCrLf & vbCrLf & _
                 "• Cross PR:" & vbCrLf
    
        For i = 1 To CLng(txtPRxRS.Value)
    
            Resumo = Resumo & _
                     "   - TRRS" & _
                     Prefixo & Format(i, "0") & _
                     " (PRxRS)" & vbCrLf
    
        Next i
    
    End If
    
    
    'Gerando o Resumo
    txtResumo.Value = Resumo
    
    'Criar Programação
    AtualizarBotaoCriar

End Sub

Private Function PrefixoDiaSemana() As String

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

Private Sub AtualizarIndicadoresCross()

    If fraCrossJD.Enabled Then

        AtualizarStatusCampo CLng(txtFTxJD.Value), lblFTxJD, txtFTxJD
        AtualizarStatusCampo CLng(txtGRxJD.Value), lblGRxJD, txtGRxJD

    End If

    If fraCrossNE.Enabled Then

        AtualizarStatusCampo CLng(txtNExCE.Value), lblNExCE, txtNExCE
        AtualizarStatusCampo CLng(txtNExPI.Value), lblNExPI, txtNExPI

    End If

End Sub

Private Sub VerificarFramesZerados()

    'If CLng(txtFTxJD.Value) = 0 And _
       CLng(txtGRxJD.Value) = 0 Then

        'optCrossJDNao.Value = True

    'End If

    'If CLng(txtNExCE.Value) = 0 And _
       CLng(txtNExPI.Value) = 0 Then

        'optCrossNENao.Value = True

    'End If

    'If CLng(txtPRxRS.Value) = 0 Then

        'optCrossPRNao.Value = True

    'End If

End Sub

Private Sub AtualizarBotaoCriar()

    cmdCriarProgramacao.Enabled = _
        (Trim(txtResumo.Value) <> "")

End Sub

Private Sub cmdCriarProgramacao_Click()

    Application.ScreenUpdating = False

    Call Etapa1CrossDocking
    Call Etapa2CrossDocking
    Call Etapa3CrossDocking
    Call Etapa4CrossDocking
    Call Etapa5CrossDocking
    Call Etapa6ACrossDocking
    Call Etapa6BCrossDocking
    Call Etapa7CrossDocking
    Call Etapa8CrossDocking
    
    Application.ScreenUpdating = True

End Sub
