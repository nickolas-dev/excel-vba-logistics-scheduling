Attribute VB_Name = "modtemplate"
Public Sub CriarTemplateInfologModelo()

    Dim wbModelo As Workbook
    Dim caminhoBase As String
    Dim caminhoDestino As String
    Dim nomeBase As String
    Dim nomeFinal As String
    Dim contador As Long
    
    confirmacao = MsgBox("Deseja criar um novo arquivo Template Infolog?", vbYesNo + vbQuestion, "Confirmação")
    
    If confirmacao = 6 Then

        Dim MODO_TESTE As Boolean
        MODO_TESTE = True  ' <<< troque para True quando quiser testar
    
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
    
        ' -------------------------------------------------------------------------
        ' Caminho do MODELO BASE (único, sempre o mesmo)
        ' -------------------------------------------------------------------------
        caminhoBase = PortfolioPath("modelo_template.xlsm")
    
        ' -------------------------------------------------------------------------
        ' Caminho de DESTINO
        ' -------------------------------------------------------------------------
        If MODO_TESTE Then
            caminhoDestino = PortfolioPath("Templates\")
        Else
            caminhoDestino = PortfolioPath("Templates\")
        End If
    
        ' -------------------------------------------------------------------------
        ' Nome base do arquivo (com InputBox)
        ' -------------------------------------------------------------------------
        Dim nomeUsuario As String
        
        nomeUsuario = InputBox("Digite o nome do arquivo:", "Nome do Template", "Template_Infolog_Modelo")
        
        ' Se cancelar ? sai da macro
        If nomeUsuario = "" Then Exit Sub
        
        ' Remove espaços extras
        nomeUsuario = Trim(nomeUsuario)
        
        ' Remove extensão se o usuário digitar
        nomeUsuario = Replace(nomeUsuario, ".xlsm", "", , , vbTextCompare)
        
        ' Se ficou vazio depois de tratar ? usa padrão
        If nomeUsuario = "" Then
            nomeUsuario = "Infolog_Modelo"
        End If
        
        ' Remove "Template" do começo (se já tiver)
        If LCase(Left(nomeUsuario, 8)) = "template" Then
            nomeUsuario = Trim(Mid(nomeUsuario, 9))
        End If
        
        ' Monta nome final garantido
        nomeBase = "Template " & nomeUsuario & ".xlsm"
        
        nomeFinal = nomeBase
        contador = 0
    
        ' Nome incremental se já existir
        Do While Dir(caminhoDestino & nomeFinal) <> ""
            contador = contador + 1
            nomeFinal = Replace(nomeBase, ".xlsm", " (" & contador & ").xlsm")
        Loop
    
        ' -------------------------------------------------------------------------
        ' Criar cópia
        ' -------------------------------------------------------------------------
        Dim wbNovo As Workbook
        
        Set wbModelo = Workbooks.Open(caminhoBase, ReadOnly:=True)
        wbModelo.SaveCopyAs caminhoDestino & nomeFinal
        wbModelo.Close False
        
        ' --- ABRIR A CÓPIA CRIADA ---
        Set wbNovo = Workbooks.Open(caminhoDestino & nomeFinal)
        
        MsgBox "O novo arquivo Template Infolog foi criado com sucesso!", _
               vbInformation, "Control Tower MB"
        
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True

    Else
    
    End If

End Sub

Public Sub CarregarUserForm6()

UserForm6.Show

End Sub
