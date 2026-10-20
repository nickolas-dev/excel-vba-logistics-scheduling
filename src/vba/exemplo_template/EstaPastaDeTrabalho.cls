Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    If TemplateEmCarga Then Exit Sub

    On Error GoTo Sair

    ' =========================
    ' ABA: FORMULAS
    ' =========================
    If Sh.Name = "Formulas" Then

        Application.EnableEvents = False

        ' Atualiza A, K, M (dados externos)
        If Not Intersect(Target, Sh.Range("P:R")) Is Nothing Then
            Atualizar_Formulas_Externas Sh
        End If

        ' Atualiza E, F, G (Ativação)
        If Not Intersect(Target, Sh.Range("D:D")) Is Nothing Then
            Atualizar_Ativacao _
                ThisWorkbook.Worksheets("Formulas"), _
                ThisWorkbook.Worksheets("Ativação")
        End If

    ' =========================
    ' ABA: ATIVAÇÃO
    ' =========================
    ElseIf Sh.Name = "Ativação" Then

        Dim tbl As ListObject
        Dim rngValido As Range

        Set tbl = Sh.ListObjects("tbl_Ativacao")

        ' ?? Ignora qualquer coisa fora do corpo da tabela
        Set rngValido = Intersect(Target, tbl.DataBodyRange)
        If rngValido Is Nothing Then Exit Sub

        Application.EnableEvents = False

        ' Normaliza formatação SOMENTE no corpo da tabela
        NormalizarFormatoAtivacao rngValido

        ' Dinâmica A ? B + merge de duplicados
        If Not Intersect(rngValido, Sh.Range("A:B")) Is Nothing Then
            SincronizarDocumentoAtivacao Sh, rngValido
        End If

        ' Qualquer alteração na Ativação reflete em E:F:G
        Atualizar_Ativacao _
            ThisWorkbook.Worksheets("Formulas"), _
            ThisWorkbook.Worksheets("Ativação")
            
        ' ? UX FINAL
        AjustarTabelaAtivacao Sh
        
    End If

Sair:
    Application.EnableEvents = True
End Sub

