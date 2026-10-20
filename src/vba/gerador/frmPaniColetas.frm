Attribute VB_Name = "frmPaniColetas"
Attribute VB_Base = "0{7FB006FB-9D4C-4D4F-B125-C1212EDA852D}{7DF91A73-E69E-4618-9399-F0A66E8EC4CA}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private UltimaDataSelecionada As Date

Private Sub cmdCriarProgramacao_Click()

    Application.ScreenUpdating = False

    Call Etapa1PaniColetas
    Call Etapa2PaniColetas
    Call Etapa3APaniColetas
    Call Etapa3BPaniColetas
    Call Etapa3CPaniColetas
    Call Etapa4APaniColetas
    Call Etapa4BPaniColetas
    Call Etapa5APaniColetas
    Call Etapa5BPaniColetas
    Call Etapa5CPaniColetas
    Call Etapa6PaniColetas
    Call Etapa7PaniColetas

    Application.ScreenUpdating = True

End Sub

Private Sub imgCalendario_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub imgCalendario_Click()

    DataAnteriorDP = gDP_SelectedDate

    On Error Resume Next
    DP_Close
    On Error GoTo 0

    DP_Show

End Sub

Private Sub optColetasSim_Click()

    fraColetas.Enabled = True

    'Atualizar Botão
    AtualizarBotaoCriar

    AtualizarResumoPaniColetas

End Sub

Private Sub UserForm_Initialize()

    txtDataFat.Value = Format(Date, "dd/mm/yyyy")
    txtDataFat.Locked = True
    txtDataFat.TabStop = False
    
    'Pani
    optPaniSim.Value = True
    AtualizarFrame fraPani, True
    
    'Coletas
    'optColetasSim.Value = True
    'AtualizarFrame fraColetas, True
    
    'Bom Sabor
    'optBomSaborSim.Value = True
    'AtualizarFrame fraBomSabor, True
    
    'Cellier
    'optCellierSim.Value = True
    'AtualizarFrame fraCellier, True
    
    'Dohler
    'optDohlerSim.Value = True
    'AtualizarFrame fraDohler, True
    
    'Kerry
    'optKerrySim.Value = True
    'AtualizarFrame fraKerry, True
    
    'Omamori
    'optOmamoriSim.Value = True
    'AtualizarFrame fraOmamori, True
    
    'Poty
    'optPotySim.Value = True
    'AtualizarFrame fraPoty, True
    
    'Transf JC
    'optTransfJCSim.Value = True
    'AtualizarFrame fraTransfJC, True
    
    'Desabilitando Coletas
    
    optColetasNao.Value = True
    
    optBomSaborNao.Value = True
    optCellierNao.Value = True
    optDohlerNao.Value = True
    optKerryNao.Value = True
    optOmamoriNao.Value = True
    optPotyNao.Value = True
    optTransfJCNao.Value = True
    
    AtualizarFrame fraBomSabor, False
    AtualizarFrame fraCellier, False
    AtualizarFrame fraDohler, False
    AtualizarFrame fraKerry, False
    AtualizarFrame fraOmamori, False
    AtualizarFrame fraPoty, False
    AtualizarFrame fraTransfJC, False
    
    'Atualizar Resumo
    'AtualizarResumo
    
    'Atualizar Indicadores Cross
    'AtualizarIndicadoresCross
    
    'Criar Programação
    'AtualizarBotaoCriar

End Sub

Private Sub AtualizarQuantidade( _
    ByRef txt As MSForms.TextBox, _
    ByRef spn As MSForms.SpinButton)

    txt.Value = Format(spn.Value, "00")

    DoEvents

    'Atualizar Indicadores
    AtualizarIndicadoresPaniColetas

    'Atualizar Status
    AtualizarStatusOperacoes
    
    'Atualizar Botão
    AtualizarBotaoCriar

    'Atualizar Resumo
    AtualizarResumoPaniColetas

End Sub

Public Sub AtualizarResumoPaniColetas()

    Dim Resumo As String
    Dim SufixoDia As String
    Dim DataFat As String

    Dim i As Long

    SufixoDia = PrefixoDiaSemana()
    DataFat = txtDataFat.Value

    '=========================
    ' PANIFRESH
    '=========================

    If CLng(txtPanixMG.Value) > 0 _
    Or CLng(txtPanixRJ.Value) > 0 _
    Or CLng(txtPanixSP.Value) > 0 Then

        Resumo = Resumo & _
            "• Panifresh:" & vbCrLf

        ' MG
        For i = 1 To CLng(txtPanixMG.Value)
        
            Resumo = Resumo & _
                "   - PANIMG" & SufixoDia & i & _
                " (" & DataFat & ")" & vbCrLf
        
        Next i
        
        If CLng(txtPanixMG.Value) > 0 Then
            Resumo = Resumo & vbCrLf
        End If
        
        ' RJ
        For i = 1 To CLng(txtPanixRJ.Value)
        
            Resumo = Resumo & _
                "   - PANIRJ" & SufixoDia & i & _
                " (" & DataFat & ")" & vbCrLf
        
        Next i
        
        If CLng(txtPanixRJ.Value) > 0 Then
            Resumo = Resumo & vbCrLf
        End If

        ' SP
        Select Case CLng(txtPanixSP.Value)
        
            Case 1
        
                Resumo = Resumo & _
                    "   - PANIJC" & _
                    SufixoDia & "1" & _
                    " (" & DataFat & ")" & vbCrLf
        
            Case Is >= 2
        
                ' Primeira carga
                Resumo = Resumo & _
                    "   - PANIJC" & _
                    SufixoDia & "1" & _
                    " (" & DataFat & ")" & vbCrLf
        
                ' Segunda carga
                Resumo = Resumo & _
                    "   - PANIJC" & _
                    SufixoDia & "2" & _
                    " (" & DataFat & ")" & vbCrLf
        
                ' Da terceira em diante
                For i = 1 To CLng(txtPanixSP.Value) - 2
        
                    Resumo = Resumo & _
                        "   - PANIPRO" & _
                        SufixoDia & i & _
                        " (" & DataFat & ")" & vbCrLf
        
                Next i
        
        End Select
        
        If CLng(txtPanixSP.Value) > 0 Then
            Resumo = Resumo & vbCrLf
        End If

    End If
    
    '=========================
    ' COLETAS
    '=========================
    
    If _
        (fraBomSabor.Enabled And _
            (CLng(txtColBomSabor.Value) > 0 Or CLng(txtEntBomSabor.Value) > 0)) _
    Or _
        (fraCellier.Enabled And _
            (CLng(txtColCellier.Value) > 0 Or CLng(txtEntCellier.Value) > 0)) _
    Or _
        (fraDohler.Enabled And _
            (CLng(txtColDohler.Value) > 0 Or CLng(txtEntDohler.Value) > 0)) _
    Or _
        (fraKerry.Enabled And _
            (CLng(txtColKerry.Value) > 0 Or CLng(txtEntKerry.Value) > 0)) _
    Or _
        (fraOmamori.Enabled And _
            (CLng(txtColOmamori.Value) > 0 Or CLng(txtEntOmamori.Value) > 0)) _
    Or _
        (fraPoty.Enabled And _
            (CLng(txtColPoty.Value) > 0 Or CLng(txtEntPoty.Value) > 0)) _
    Or _
        (fraTransfJC.Enabled And _
            CLng(txtTransfJC.Value) > 0) Then

    Resumo = Resumo & _
        "• Coletas:" & vbCrLf & vbCrLf

    '=========================
    ' BOM SABOR
    '=========================

    If fraBomSabor.Enabled Then

        If CLng(txtColBomSabor.Value) > 0 _
        Or CLng(txtEntBomSabor.Value) > 0 Then

            Resumo = Resumo & _
                "   • Bom Sabor:" & vbCrLf

            For i = 1 To CLng(txtColBomSabor.Value)

                Resumo = Resumo & _
                    "      - COLBOMSABOR" & _
                    SufixoDia & i & _
                    " (" & DataFat & ")" & vbCrLf

            Next i

            For i = 1 To CLng(txtEntBomSabor.Value)

                Resumo = Resumo & _
                    "      - ENTBOMSABOR" & _
                    SufixoDia & i & _
                    " (" & DataFat & ")" & vbCrLf

            Next i

            Resumo = Resumo & vbCrLf

        End If

    End If
    
    '=========================
    ' CELLIER
    '=========================

    If fraCellier.Enabled Then

        If CLng(txtColCellier.Value) > 0 _
        Or CLng(txtEntCellier.Value) > 0 Then

            Resumo = Resumo & _
                "   • Cellier:" & vbCrLf

            For i = 1 To CLng(txtColCellier.Value)

                Resumo = Resumo & _
                    "      - COLCELLIER" & _
                    SufixoDia & i & _
                    " (" & DataFat & ")" & vbCrLf

            Next i

            For i = 1 To CLng(txtEntCellier.Value)

                Resumo = Resumo & _
                    "      - ENTCELLIER" & _
                    SufixoDia & i & _
                    " (" & DataFat & ")" & vbCrLf

            Next i

            Resumo = Resumo & vbCrLf

        End If

    End If

    '=========================
    ' DOHLER
    '=========================

    If fraDohler.Enabled Then

        If CLng(txtColDohler.Value) > 0 _
        Or CLng(txtEntDohler.Value) > 0 Then

            Resumo = Resumo & _
                "   • Dohler:" & vbCrLf

            For i = 1 To CLng(txtColDohler.Value)

                Resumo = Resumo & _
                    "      - COLDOHLER" & _
                    SufixoDia & i & _
                    " (" & DataFat & ")" & vbCrLf

            Next i

            For i = 1 To CLng(txtEntDohler.Value)

                Resumo = Resumo & _
                    "      - ENTDOHLER" & _
                    SufixoDia & i & _
                    " (" & DataFat & ")" & vbCrLf

            Next i

            Resumo = Resumo & vbCrLf

        End If

    End If

    '=========================
    ' KERRY
    '=========================

    If fraKerry.Enabled Then

        If CLng(txtColKerry.Value) > 0 _
        Or CLng(txtEntKerry.Value) > 0 Then

            Resumo = Resumo & _
                "   • Kerry:" & vbCrLf

            For i = 1 To CLng(txtColKerry.Value)

                Resumo = Resumo & _
                    "      - COLKERRY" & _
                    SufixoDia & i & _
                    " (" & DataFat & ")" & vbCrLf

            Next i

            For i = 1 To CLng(txtEntKerry.Value)

                Resumo = Resumo & _
                    "      - ENTKERRY" & _
                    SufixoDia & i & _
                    " (" & DataFat & ")" & vbCrLf

            Next i

            Resumo = Resumo & vbCrLf

        End If

    End If

    '=========================
    ' OMAMORI
    '=========================

    If fraOmamori.Enabled Then

        If CLng(txtColOmamori.Value) > 0 _
        Or CLng(txtEntOmamori.Value) > 0 Then

            Resumo = Resumo & _
                "   • Omamori:" & vbCrLf

            For i = 1 To CLng(txtColOmamori.Value)

                Resumo = Resumo & _
                    "      - COLOMAMORI" & _
                    SufixoDia & i & _
                    " (" & DataFat & ")" & vbCrLf

            Next i

            For i = 1 To CLng(txtEntOmamori.Value)

                Resumo = Resumo & _
                    "      - ENTOMAMORI" & _
                    SufixoDia & i & _
                    " (" & DataFat & ")" & vbCrLf

            Next i

            Resumo = Resumo & vbCrLf

        End If

    End If

    '=========================
    ' POTY
    '=========================

    If fraPoty.Enabled Then

        If CLng(txtColPoty.Value) > 0 _
        Or CLng(txtEntPoty.Value) > 0 Then

            Resumo = Resumo & _
                "   • Poty:" & vbCrLf

            For i = 1 To CLng(txtColPoty.Value)

                Resumo = Resumo & _
                    "      - COLPOTY" & _
                    SufixoDia & i & _
                    " (" & DataFat & ")" & vbCrLf

            Next i

            For i = 1 To CLng(txtEntPoty.Value)

                Resumo = Resumo & _
                    "      - ENTPOTY" & _
                    SufixoDia & i & _
                    " (" & DataFat & ")" & vbCrLf

            Next i

            Resumo = Resumo & vbCrLf

        End If

    End If

    '=========================
    ' TRANSFERÊNCIA JC
    '=========================

    If fraTransfJC.Enabled Then

        If CLng(txtTransfJC.Value) > 0 Then

            Resumo = Resumo & _
                "   • Transferência JC:" & vbCrLf

            For i = 1 To CLng(txtTransfJC.Value)
            
                Resumo = Resumo & _
                    "      - JC" & _
                    SufixoDia & _
                    Format(i, "000") & _
                    " (" & DataFat & ")" & vbCrLf
            
            Next i

            Resumo = Resumo & vbCrLf

        End If

    End If

End If

    txtResumo.Value = Resumo

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

Private Sub spnPanixMG_Change()
    AtualizarQuantidade txtPanixMG, spnPanixMG
End Sub

Private Sub spnPanixRJ_Change()
    AtualizarQuantidade txtPanixRJ, spnPanixRJ
End Sub

Private Sub spnPanixSP_Change()
    AtualizarQuantidade txtPanixSP, spnPanixSP
End Sub

Private Sub spnColBomSabor_Change()
    AtualizarQuantidade txtColBomSabor, spnColBomSabor
End Sub

Private Sub spnEntBomSabor_Change()
    AtualizarQuantidade txtEntBomSabor, spnEntBomSabor
End Sub

Private Sub spnColCellier_Change()
    AtualizarQuantidade txtColCellier, spnColCellier
End Sub

Private Sub spnEntCellier_Change()
    AtualizarQuantidade txtEntCellier, spnEntCellier
End Sub

Private Sub spnColDohler_Change()
    AtualizarQuantidade txtColDohler, spnColDohler
End Sub

Private Sub spnEntDohler_Change()
    AtualizarQuantidade txtEntDohler, spnEntDohler
End Sub

Private Sub spnColKerry_Change()
    AtualizarQuantidade txtColKerry, spnColKerry
End Sub

Private Sub spnEntKerry_Change()
    AtualizarQuantidade txtEntKerry, spnEntKerry
End Sub

Private Sub spnColOmamori_Change()
    AtualizarQuantidade txtColOmamori, spnColOmamori
End Sub

Private Sub spnEntOmamori_Change()
    AtualizarQuantidade txtEntOmamori, spnEntOmamori
End Sub

Private Sub spnColPoty_Change()
    AtualizarQuantidade txtColPoty, spnColPoty
End Sub

Private Sub spnEntPoty_Change()
    AtualizarQuantidade txtEntPoty, spnEntPoty
End Sub

Private Sub spnTransfJC_Change()
    AtualizarQuantidade txtTransfJC, spnTransfJC
End Sub

Private Sub AtualizarFrame( _
    ByRef FrameDestino As MSForms.Frame, _
    ByVal Ativar As Boolean)

    FrameDestino.Enabled = Ativar

    AtualizarResumoPaniColetas

End Sub

Private Sub AtualizarFramePani(ByVal Ativo As Boolean)

    fraPani.Enabled = Ativo

    If Ativo Then
    
        txtPanixMG.BackColor = vbWhite
        txtPanixRJ.BackColor = vbWhite
        txtPanixSP.BackColor = vbWhite
    
    Else
    
        lblPanixMG.ForeColor = RGB(150, 150, 150)
        txtPanixMG.BackColor = RGB(240, 240, 240)
        txtPanixMG.ForeColor = RGB(120, 120, 120)
    
        lblPanixRJ.ForeColor = RGB(150, 150, 150)
        txtPanixRJ.BackColor = RGB(240, 240, 240)
        txtPanixRJ.ForeColor = RGB(120, 120, 120)
    
        lblPanixSP.ForeColor = RGB(150, 150, 150)
        txtPanixSP.BackColor = RGB(240, 240, 240)
        txtPanixSP.ForeColor = RGB(120, 120, 120)
    
    End If

End Sub

Private Sub optPaniSim_Click()

    AtualizarFramePani True

    AtualizarIndicadoresPaniColetas

    'Atualizar Botão
    AtualizarBotaoCriar

    AtualizarResumoPaniColetas

End Sub

Private Sub optPaniNao_Click()

    txtPanixMG.Value = "00"
    txtPanixRJ.Value = "00"
    txtPanixSP.Value = "00"

    spnPanixMG.Value = 0
    spnPanixRJ.Value = 0
    spnPanixSP.Value = 0

    AtualizarFramePani False
    
    'Atualizar Botão
    AtualizarBotaoCriar

    AtualizarResumoPaniColetas

End Sub

Private Sub optColetasNao_Click()

    optBomSaborNao.Value = True
    optCellierNao.Value = True
    optDohlerNao.Value = True
    optKerryNao.Value = True
    optOmamoriNao.Value = True
    optPotyNao.Value = True
    optTransfJCNao.Value = True
    
    fraColetas.Enabled = False
    
    'Atualizar Botão
    AtualizarBotaoCriar

    AtualizarResumoPaniColetas

End Sub

Private Sub optBomSaborSim_Click()

    AtualizarFrameBomSabor True

    AtualizarIndicadoresPaniColetas

    AtualizarResumoPaniColetas

End Sub

Private Sub optBomSaborNao_Click()

    txtColBomSabor.Value = "00"
    txtEntBomSabor.Value = "00"

    spnColBomSabor.Value = 0
    spnEntBomSabor.Value = 0

    AtualizarFrameBomSabor False

    AtualizarResumoPaniColetas

End Sub

Private Sub optCellierSim_Click()

    AtualizarFrameCellier True

    AtualizarIndicadoresPaniColetas

    AtualizarResumoPaniColetas

End Sub

Private Sub optCellierNao_Click()

    txtColCellier.Value = "00"
    txtEntCellier.Value = "00"

    spnColCellier.Value = 0
    spnEntCellier.Value = 0

    AtualizarFrameCellier False

    AtualizarResumoPaniColetas

End Sub

Private Sub optDohlerNao_Click()

    txtColDohler.Value = "00"
    txtEntDohler.Value = "00"

    spnColDohler.Value = 0
    spnEntDohler.Value = 0

    AtualizarFrameDohler False

    AtualizarResumoPaniColetas

End Sub

Private Sub optDohlerSim_Click()

    AtualizarFrameDohler True

    AtualizarIndicadoresPaniColetas

    AtualizarResumoPaniColetas

End Sub

Private Sub optKerryNao_Click()

    txtColKerry.Value = "00"
    txtEntKerry.Value = "00"

    spnColKerry.Value = 0
    spnEntKerry.Value = 0

    AtualizarFrameKerry False

    AtualizarResumoPaniColetas

End Sub

Private Sub optKerrySim_Click()

    AtualizarFrameKerry True

    AtualizarIndicadoresPaniColetas

    AtualizarResumoPaniColetas

End Sub

Private Sub optOmamoriNao_Click()

    txtColOmamori.Value = "00"
    txtEntOmamori.Value = "00"

    spnColOmamori.Value = 0
    spnEntOmamori.Value = 0

    AtualizarFrameOmamori False

    AtualizarResumoPaniColetas

End Sub

Private Sub optOmamoriSim_Click()

    AtualizarFrameOmamori True

    AtualizarIndicadoresPaniColetas

    AtualizarResumoPaniColetas

End Sub

Private Sub optPotyNao_Click()

    txtColPoty.Value = "00"
    txtEntPoty.Value = "00"

    spnColPoty.Value = 0
    spnEntPoty.Value = 0

    AtualizarFramePoty False

    AtualizarResumoPaniColetas

End Sub

Private Sub optPotySim_Click()

    AtualizarFramePoty True

    AtualizarIndicadoresPaniColetas

    AtualizarResumoPaniColetas

End Sub

Private Sub optTransfJCNao_Click()

    txtTransfJC.Value = "00"

    spnTransfJC.Value = 0

    AtualizarFrameTransfJC False

    AtualizarResumoPaniColetas

End Sub

Private Sub optTransfJCSim_Click()

    AtualizarFrameTransfJC True

    AtualizarIndicadoresPaniColetas

    AtualizarResumoPaniColetas

End Sub

Private Sub AtualizarStatusOperacoes()

    '=========================
    ' PANIFRESH
    '=========================

    If CLng(txtPanixMG.Value) = 0 _
    And CLng(txtPanixRJ.Value) = 0 _
    And CLng(txtPanixSP.Value) = 0 Then

        optPaniNao.Value = True
        AtualizarFramePani False

    Else

        optPaniSim.Value = True
        AtualizarFramePani True

    End If

    '=========================
    ' BOM SABOR
    '=========================

    If CLng(txtColBomSabor.Value) = 0 _
    And CLng(txtEntBomSabor.Value) = 0 Then

        optBomSaborNao.Value = True
        AtualizarFrameBomSabor False

    Else

        optBomSaborSim.Value = True
        AtualizarFrameBomSabor True

    End If

    '=========================
    ' CELLIER
    '=========================

    If CLng(txtColCellier.Value) = 0 _
    And CLng(txtEntCellier.Value) = 0 Then

        optCellierNao.Value = True
        AtualizarFrameCellier False

    Else

        optCellierSim.Value = True
        AtualizarFrameCellier True

    End If

    '=========================
    ' DOHLER
    '=========================

    If CLng(txtColDohler.Value) = 0 _
    And CLng(txtEntDohler.Value) = 0 Then

        optDohlerNao.Value = True
        AtualizarFrameDohler False

    Else

        optDohlerSim.Value = True
        AtualizarFrameDohler True

    End If

    '=========================
    ' KERRY
    '=========================

    If CLng(txtColKerry.Value) = 0 _
    And CLng(txtEntKerry.Value) = 0 Then

        optKerryNao.Value = True
        AtualizarFrameKerry False

    Else

        optKerrySim.Value = True
        AtualizarFrameKerry True

    End If

    '=========================
    ' OMAMORI
    '=========================

    If CLng(txtColOmamori.Value) = 0 _
    And CLng(txtEntOmamori.Value) = 0 Then

        optOmamoriNao.Value = True
        AtualizarFrameOmamori False

    Else

        optOmamoriSim.Value = True
        AtualizarFrameOmamori True

    End If

    '=========================
    ' POTY
    '=========================

    If CLng(txtColPoty.Value) = 0 _
    And CLng(txtEntPoty.Value) = 0 Then

        optPotyNao.Value = True
        AtualizarFramePoty False

    Else

        optPotySim.Value = True
        AtualizarFramePoty True

    End If

    '=========================
    ' TRANSFERÊNCIA JC
    '=========================

    If CLng(txtTransfJC.Value) = 0 Then

        optTransfJCNao.Value = True
        AtualizarFrameTransfJC False

    Else

        optTransfJCSim.Value = True
        AtualizarFrameTransfJC True

    End If

    '=========================
    ' PULO DO GATO
    ' COLETAS
    '=========================

    If Not fraBomSabor.Enabled _
    And Not fraCellier.Enabled _
    And Not fraDohler.Enabled _
    And Not fraKerry.Enabled _
    And Not fraOmamori.Enabled _
    And Not fraPoty.Enabled _
    And Not fraTransfJC.Enabled Then

        optColetasNao.Value = True

    Else

        optColetasSim.Value = True

    End If

End Sub

Private Sub AtualizarIndicadoresPaniColetas()

    '=========================
    ' PANIFRESH
    '=========================
    
    If fraPani.Enabled Then
    
        AtualizarStatusCampo _
            CLng(txtPanixMG.Value), _
            lblPanixMG, _
            txtPanixMG
    
        AtualizarStatusCampo _
            CLng(txtPanixRJ.Value), _
            lblPanixRJ, _
            txtPanixRJ
    
        AtualizarStatusCampo _
            CLng(txtPanixSP.Value), _
            lblPanixSP, _
            txtPanixSP
    
    End If

    '=========================
    ' BOM SABOR
    '=========================
    
    If fraBomSabor.Enabled Then
    
        AtualizarStatusCampo _
            CLng(txtColBomSabor.Value), _
            lblColBomSabor, _
            txtColBomSabor
    
        AtualizarStatusCampo _
            CLng(txtEntBomSabor.Value), _
            lblEntBomSabor, _
            txtEntBomSabor
    
    End If
    
    '=========================
    ' CELLIER
    '=========================
    
    If fraCellier.Enabled Then
    
        AtualizarStatusCampo _
            CLng(txtColCellier.Value), _
            lblColCellier, _
            txtColCellier
    
        AtualizarStatusCampo _
            CLng(txtEntCellier.Value), _
            lblEntCellier, _
            txtEntCellier
    
    End If
    
    '=========================
    ' DOHLER
    '=========================
    
    If fraDohler.Enabled Then
    
        AtualizarStatusCampo _
            CLng(txtColDohler.Value), _
            lblColDohler, _
            txtColDohler
    
        AtualizarStatusCampo _
            CLng(txtEntDohler.Value), _
            lblEntDohler, _
            txtEntDohler
    
    End If
    
    '=========================
    ' KERRY
    '=========================
    
    If fraKerry.Enabled Then
    
        AtualizarStatusCampo _
            CLng(txtColKerry.Value), _
            lblColKerry, _
            txtColKerry
    
        AtualizarStatusCampo _
            CLng(txtEntKerry.Value), _
            lblEntKerry, _
            txtEntKerry
    
    End If
    
    '=========================
    ' OMAMORI
    '=========================
    
    If fraOmamori.Enabled Then
    
        AtualizarStatusCampo _
            CLng(txtColOmamori.Value), _
            lblColOmamori, _
            txtColOmamori
    
        AtualizarStatusCampo _
            CLng(txtEntOmamori.Value), _
            lblEntOmamori, _
            txtEntOmamori
    
    End If
    
    '=========================
    ' POTY
    '=========================
    
    If fraPoty.Enabled Then
    
        AtualizarStatusCampo _
            CLng(txtColPoty.Value), _
            lblColPoty, _
            txtColPoty
    
        AtualizarStatusCampo _
            CLng(txtEntPoty.Value), _
            lblEntPoty, _
            txtEntPoty
    
    End If
    
    '=========================
    ' TRANSF JC
    '=========================
    
    If fraTransfJC.Enabled Then
    
        AtualizarStatusCampo _
            CLng(txtTransfJC.Value), _
            lblTransfJC, _
            txtTransfJC
    
    End If

End Sub

Private Sub AtualizarFrameBomSabor(ByVal Ativo As Boolean)

    fraBomSabor.Enabled = Ativo

    If Ativo Then

        txtColBomSabor.BackColor = vbWhite
        txtEntBomSabor.BackColor = vbWhite

    Else

        lblColBomSabor.ForeColor = RGB(150, 150, 150)
        txtColBomSabor.BackColor = RGB(240, 240, 240)
        txtColBomSabor.ForeColor = RGB(120, 120, 120)

        lblEntBomSabor.ForeColor = RGB(150, 150, 150)
        txtEntBomSabor.BackColor = RGB(240, 240, 240)
        txtEntBomSabor.ForeColor = RGB(120, 120, 120)

    End If

End Sub

Private Sub AtualizarFrameCellier(ByVal Ativo As Boolean)

    fraCellier.Enabled = Ativo

    If Ativo Then

        txtColCellier.BackColor = vbWhite
        txtEntCellier.BackColor = vbWhite

    Else

        lblColCellier.ForeColor = RGB(150, 150, 150)
        txtColCellier.BackColor = RGB(240, 240, 240)
        txtColCellier.ForeColor = RGB(120, 120, 120)

        lblEntCellier.ForeColor = RGB(150, 150, 150)
        txtEntCellier.BackColor = RGB(240, 240, 240)
        txtEntCellier.ForeColor = RGB(120, 120, 120)

    End If

End Sub

Private Sub AtualizarFrameDohler(ByVal Ativo As Boolean)

    fraDohler.Enabled = Ativo

    If Ativo Then

        txtColDohler.BackColor = vbWhite
        txtEntDohler.BackColor = vbWhite

    Else

        lblColDohler.ForeColor = RGB(150, 150, 150)
        txtColDohler.BackColor = RGB(240, 240, 240)
        txtColDohler.ForeColor = RGB(120, 120, 120)

        lblEntDohler.ForeColor = RGB(150, 150, 150)
        txtEntDohler.BackColor = RGB(240, 240, 240)
        txtEntDohler.ForeColor = RGB(120, 120, 120)

    End If

End Sub

Private Sub AtualizarFrameKerry(ByVal Ativo As Boolean)

    fraKerry.Enabled = Ativo

    If Ativo Then

        txtColKerry.BackColor = vbWhite
        txtEntKerry.BackColor = vbWhite

    Else

        lblColKerry.ForeColor = RGB(150, 150, 150)
        txtColKerry.BackColor = RGB(240, 240, 240)
        txtColKerry.ForeColor = RGB(120, 120, 120)

        lblEntKerry.ForeColor = RGB(150, 150, 150)
        txtEntKerry.BackColor = RGB(240, 240, 240)
        txtEntKerry.ForeColor = RGB(120, 120, 120)

    End If

End Sub

Private Sub AtualizarFrameOmamori(ByVal Ativo As Boolean)

    fraOmamori.Enabled = Ativo

    If Ativo Then

        txtColOmamori.BackColor = vbWhite
        txtEntOmamori.BackColor = vbWhite

    Else

        lblColOmamori.ForeColor = RGB(150, 150, 150)
        txtColOmamori.BackColor = RGB(240, 240, 240)
        txtColOmamori.ForeColor = RGB(120, 120, 120)

        lblEntOmamori.ForeColor = RGB(150, 150, 150)
        txtEntOmamori.BackColor = RGB(240, 240, 240)
        txtEntOmamori.ForeColor = RGB(120, 120, 120)

    End If

End Sub

Private Sub AtualizarFramePoty(ByVal Ativo As Boolean)

    fraPoty.Enabled = Ativo

    If Ativo Then

        txtColPoty.BackColor = vbWhite
        txtEntPoty.BackColor = vbWhite

    Else

        lblColPoty.ForeColor = RGB(150, 150, 150)
        txtColPoty.BackColor = RGB(240, 240, 240)
        txtColPoty.ForeColor = RGB(120, 120, 120)

        lblEntPoty.ForeColor = RGB(150, 150, 150)
        txtEntPoty.BackColor = RGB(240, 240, 240)
        txtEntPoty.ForeColor = RGB(120, 120, 120)

    End If

End Sub

Private Sub AtualizarFrameTransfJC(ByVal Ativo As Boolean)

    fraTransfJC.Enabled = Ativo

    If Ativo Then

        txtTransfJC.BackColor = vbWhite

    Else

        lblTransfJC.ForeColor = RGB(150, 150, 150)

        txtTransfJC.BackColor = RGB(240, 240, 240)
        txtTransfJC.ForeColor = RGB(120, 120, 120)

    End If

End Sub

Private Sub AtualizarBotaoCriar()

    Dim TemPani As Boolean
    Dim TemColetas As Boolean

    '=========================
    ' PANIFRESH
    '=========================

    TemPani = _
        CLng(txtPanixMG.Value) > 0 _
        Or CLng(txtPanixRJ.Value) > 0 _
        Or CLng(txtPanixSP.Value) > 0

    '=========================
    ' COLETAS
    '=========================

    TemColetas = _
        CLng(txtColBomSabor.Value) > 0 _
        Or CLng(txtEntBomSabor.Value) > 0 _
        Or CLng(txtColCellier.Value) > 0 _
        Or CLng(txtEntCellier.Value) > 0 _
        Or CLng(txtColDohler.Value) > 0 _
        Or CLng(txtEntDohler.Value) > 0 _
        Or CLng(txtColKerry.Value) > 0 _
        Or CLng(txtEntKerry.Value) > 0 _
        Or CLng(txtColOmamori.Value) > 0 _
        Or CLng(txtEntOmamori.Value) > 0 _
        Or CLng(txtColPoty.Value) > 0 _
        Or CLng(txtEntPoty.Value) > 0 _
        Or CLng(txtTransfJC.Value) > 0

    '=========================
    ' BOTÃO
    '=========================

    If TemPani And TemColetas Then

        cmdCriarProgramacao.Caption = _
            "CRIAR PROGRAMAÇÃO PANIFRESH E COLETAS"

        cmdCriarProgramacao.Enabled = True

    ElseIf TemPani Then

        cmdCriarProgramacao.Caption = _
            "CRIAR PROGRAMAÇÃO PANIFRESH"

        cmdCriarProgramacao.Enabled = True

    ElseIf TemColetas Then

        cmdCriarProgramacao.Caption = _
            "CRIAR PROGRAMAÇÃO COLETAS"

        cmdCriarProgramacao.Enabled = True

    Else

        cmdCriarProgramacao.Caption = _
            "CRIAR PROGRAMAÇÃO PANIFRESH E COLETAS"

        cmdCriarProgramacao.Enabled = False

    End If

End Sub

