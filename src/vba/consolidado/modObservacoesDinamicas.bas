Attribute VB_Name = "modObservacoesDinamicas"
Public Sub AdicionarObservacao( _
    ByVal Celula As Range, _
    ByVal Texto As String)

    If InStr(1, Celula.Value, Texto, vbTextCompare) = 0 Then

        Celula.Value = _
            Celula.Value & _
            IIf(Celula.Value <> "", " | ", "") & _
            Texto

    End If

End Sub

Public Sub RemoverObservacao( _
    ByVal Celula As Range, _
    ByVal Texto As String)

    Celula.Value = Replace( _
        Celula.Value, _
        " | " & Texto, _
        "")

    Celula.Value = Replace( _
        Celula.Value, _
        Texto & " | ", _
        "")

    Celula.Value = Replace( _
        Celula.Value, _
        Texto, _
        "")

    Do While InStr(Celula.Value, " |  | ") > 0

        Celula.Value = _
            Replace(Celula.Value, " |  | ", " | ")

    Loop

    If Trim(Celula.Value) = "|" Then
        Celula.Value = ""
    End If

End Sub

Public Sub AtualizarObsDTHCarreg(ByVal Linha As Long)

    Dim Texto As String

    Texto = "DTH CARRG PROG é maior do que DTH ENTR PROG"

    If Linha < 2 Then Exit Sub

    If IsDate(Cells(Linha, "L").Value) _
       And IsDate(Cells(Linha, "O").Value) Then

        If Cells(Linha, "L").Value > _
           Cells(Linha, "O").Value Then

            AdicionarObservacao _
                Cells(Linha, "X"), _
                Texto

        Else

            RemoverObservacao _
                Cells(Linha, "X"), _
                Texto

        End If

    Else

        RemoverObservacao _
            Cells(Linha, "X"), _
            Texto

    End If

End Sub

Public Sub AtualizarObsDuplicada()

    Dim UltimaLinha As Long

    Dim cel As Range
    Dim Intervalo As Range

    Dim Dict As Object

    Dim Texto As String
    Dim Chave As String

    Texto = "Duplicada"

    Set Dict = CreateObject("Scripting.Dictionary")

    UltimaLinha = _
        Cells(Rows.Count, "I").End(xlUp).Row

    Set Intervalo = _
        Range("I2:I" & UltimaLinha)

    ' Contagem
    For Each cel In Intervalo

        If Trim(cel.Value) <> "" Then

            Chave = _
                CStr(Cells(cel.Row, "P").Value) & "|" & _
                CStr(Cells(cel.Row, "I").Value)

            If Dict.Exists(Chave) Then

                Dict(Chave) = Dict(Chave) + 1

            Else

                Dict.Add Chave, 1

            End If

        End If

    Next cel

    ' Atualização das observações
    For Each cel In Intervalo

        If Trim(cel.Value) <> "" Then

            Chave = _
                CStr(Cells(cel.Row, "P").Value) & "|" & _
                CStr(Cells(cel.Row, "I").Value)

            If Dict(Chave) > 1 Then

                If Len(cel.Value) <> 4 Then

                    AdicionarObservacao _
                        Cells(cel.Row, "X"), _
                        Texto

                End If

            Else

                RemoverObservacao _
                    Cells(cel.Row, "X"), _
                    Texto

            End If

        Else

            RemoverObservacao _
                Cells(cel.Row, "X"), _
                Texto

        End If

    Next cel

End Sub

Public Sub AtualizarObsAlertaCarregamento(ByVal Linha As Long)

    Dim LinhaFirst As Long
    Dim LinhaFim As Long

    Dim lData As Variant
    Dim oData As Variant

    Dim Texto As String

    Texto = _
        "Alerta: Carregamento programado para 24 horas ou mais antes da primeira entrega"

    '=================================
    ' LOCALIZA O FIRST DO BLOCO
    '=================================

    LinhaFirst = Linha

    Do While LinhaFirst > 2

        If UCase$(Trim$(Cells(LinhaFirst, "Z").Value)) = "FIRST" Then
            Exit Do
        End If

        LinhaFirst = LinhaFirst - 1

    Loop

    '=================================
    ' LOCALIZA O FIM DO BLOCO
    '=================================

    LinhaFim = LinhaFirst

    Do While LinhaFim + 1 <= Cells(Rows.Count, "H").End(xlUp).Row

        If UCase$(Trim$(Cells(LinhaFim + 1, "Z").Value)) = "FIRST" Then
            Exit Do
        End If

        LinhaFim = LinhaFim + 1

    Loop

    lData = Cells(LinhaFirst, "L").Value
    oData = Cells(LinhaFirst, "O").Value

    If IsDate(lData) And IsDate(oData) Then

        If CDate(oData) >= CDate(lData) + 1 Then

            For Linha = LinhaFirst To LinhaFim

                AdicionarObservacao _
                    Cells(Linha, "X"), _
                    Texto

            Next Linha

        Else

            For Linha = LinhaFirst To LinhaFim

                RemoverObservacao _
                    Cells(Linha, "X"), _
                    Texto

            Next Linha

        End If

    Else

        For Linha = LinhaFirst To LinhaFim

            RemoverObservacao _
                Cells(Linha, "X"), _
                Texto

        Next Linha

    End If

End Sub

