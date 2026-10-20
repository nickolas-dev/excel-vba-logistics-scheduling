Attribute VB_Name = "modExportar"
Public Function ObterArquivoTbl( _
    embarque As String, _
    mesExtenso As String) As String

    embarque = UCase(Trim(embarque))

    If InStr(embarque, "PANI") > 0 Then

        ObterArquivoTbl = _
            "Tbl_Coletas_PANIFRESH_" & _
            mesExtenso & ".xlsx"

    ElseIf InStr(embarque, "BATATA") > 0 Then

        ObterArquivoTbl = _
            "Tbl_Transfer_Batatas_" & _
            mesExtenso & ".xlsx"

    ElseIf InStr(embarque, "COL") > 0 _
       Or InStr(embarque, "ENT") > 0 _
       Or InStr(embarque, "FTJC") > 0 Then

        ObterArquivoTbl = _
            "Tbl_Coletas_MB_" & _
            mesExtenso & ".xlsx"

    Else

        ObterArquivoTbl = ""

    End If

End Function

Public Sub Tbl_Programacao()

    Dim nomeArquivo As String

    nomeArquivo = UCase(ThisWorkbook.Name)

    If InStr(nomeArquivo, "CROSS") > 0 Then

        Call Tbl_CrossDocking

    ElseIf InStr(nomeArquivo, "PANI") > 0 Then

        Call Tbl_PaniColetas

    ElseIf InStr(nomeArquivo, "BATATAS") > 0 Then

        Call Tbl_Batatas

    Else

        MsgBox _
            "Tipo de programação não identificado.", _
            vbCritical

    End If

End Sub


