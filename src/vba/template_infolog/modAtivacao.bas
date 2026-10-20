Attribute VB_Name = "modAtivacao"
Option Explicit

Public Sub Atualizar_Ativacao(wsF As Worksheet, wsA As Worksheet)

    Dim dict As Object
    Dim tbl As ListObject
    Dim i As Long
    Dim UltimaLinha As Long
    Dim chave As String

    Set dict = CreateObject("Scripting.Dictionary")
    Set tbl = wsA.ListObjects("tbl_Ativacao")

    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.DataBodyRange.Rows.Count
            chave = Trim(tbl.DataBodyRange.Cells(i, 1).Value)

            If Len(chave) > 0 Then
                dict(chave) = Array( _
                    tbl.DataBodyRange.Cells(i, 3).Value, _
                    tbl.DataBodyRange.Cells(i, 4).Value, _
                    tbl.DataBodyRange.Cells(i, 5).Value)
            End If
        Next i
    End If

    UltimaLinha = wsF.Cells(wsF.Rows.Count, "D").End(xlUp).Row

    For i = 2 To UltimaLinha
        chave = wsF.Cells(i, "D").Value

        If dict.Exists(chave) Then
            wsF.Cells(i, "E").Value = dict(chave)(0)
            wsF.Cells(i, "F").Value = dict(chave)(1)
            wsF.Cells(i, "G").Value = dict(chave)(2)
        Else
            wsF.Cells(i, "E").ClearContents
            wsF.Cells(i, "F").ClearContents
            wsF.Cells(i, "G").ClearContents
        End If
    Next i

End Sub

