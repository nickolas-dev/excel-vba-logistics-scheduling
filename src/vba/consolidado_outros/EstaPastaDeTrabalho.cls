Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_Open()
    Dim nomeArquivo As String
    Dim dataCriacao As Date
    Dim tempoRestante As Date
    Dim ws As Worksheet
    Dim contador As Long
    Dim caminhoArquivo As String
    Dim usuario As String

    Set ws = ThisWorkbook.Sheets(1)
    nomeArquivo = ThisWorkbook.Name
    caminhoArquivo = ThisWorkbook.FullName
    usuario = Environ("USERNAME")

    ' Se for o modelo original
    If StrComp(nomeArquivo, "modelo_consolidado.xlsm", vbTextCompare) = 0 Or _
   StrComp(nomeArquivo, "modelo_consolidado_OUTROS.xlsm", vbTextCompare) = 0 Then
        ' Verifica se o usuário é autorizado
        If LCase(usuario) <> "ncandido" Then
            'MsgBox "Você está abrindo este arquivo como somente leitura.", vbInformation
            ThisWorkbook.ChangeFileAccess Mode:=xlReadOnly
        End If

        ' Zera o contador e sai
        ws.Range("AA1").Value = 0
        Exit Sub
    End If

    ' Atualiza contador de aberturas na célula AA1
    If IsNumeric(ws.Range("AA1").Value) Then
        contador = ws.Range("AA1").Value + 1
    Else
        contador = 1
    End If
    ws.Range("AA1").Value = contador
    ThisWorkbook.Save

    ' Verifica validade do arquivo
    dataCriacao = FileDateTime(caminhoArquivo)
    tempoRestante = (dataCriacao + 1) - Now ' 24 horas

    If tempoRestante > 0 Then
        If contador > 1 Then
            MsgBox "Este arquivo irá expirar em aproximadamente " & Format(tempoRestante, "hh:mm:ss") & ".", vbInformation, "Validade do Arquivo"
        End If
    Else
        MsgBox "Este arquivo já passou do prazo de 24 horas e será excluído.", vbCritical, "Arquivo Expirado"
        ThisWorkbook.Saved = True
        Application.DisplayAlerts = False
        ThisWorkbook.Close SaveChanges:=False
        Kill caminhoArquivo
    End If
End Sub



