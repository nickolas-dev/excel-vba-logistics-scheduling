Attribute VB_Name = "Módulo1"
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Public wb3 As Workbook

Public Sub ExecutarAntesDeFechar()

'ocultando a planilha CALENDARIO

    ThisWorkbook.Activate
    
    Sheets("CALENDARIO").Visible = True
    Sheets("CALENDARIO").Visible = False
    

'ocultando a planilha CROSS DOCKING

    ThisWorkbook.Activate
    
    Sheets("CROSS DOCKING").Visible = True
    Sheets("CROSS DOCKING").Visible = False
    
'ocultando a planilha PANI e COLETAS

    ThisWorkbook.Activate
    
    Sheets("PANI e COLETAS").Visible = True
    Sheets("PANI e COLETAS").Visible = False
    
'ocultando a planilha PARAMETROS_PF_CACHE

    ThisWorkbook.Activate
    
    Sheets("PARAMETROS_PF_CACHE").Visible = True
    Sheets("PARAMETROS_PF_CACHE").Visible = False


'limpando e ocultando a Planilha2

    On Error GoTo Proximo
    
    Sheets("Planilha2").Visible = True
    Sheets("Planilha2").Select
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A1").Select
    Sheets("Planilha2").Visible = False

'limpando a Planilha1
    
Proximo:
    
    ActiveSheet.Unprotect 'desprotegendo a planilha
    Sheets("Planilha1").Select
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=3
    Selection.ClearContents
    
'Call formatooriginal

ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'protegendo a planilha

Range("A1").Select

End Sub

Public Sub InformacoesConsolidadas()

Application.ScreenUpdating = False

    ' Etapa 1 - com subetapas
    Call Etapa1

    ' Etapa 2 - com subetapas
    Call Etapa2

    ' Etapa 3 - com subetapas
    Call Etapa3
    
    'Total de rotas e de lojas para msgbox
    'total_lojas = Range("I1").End(xlDown).Row - 1
    total_lojas = Application.WorksheetFunction.CountIf(Range("I2:I" & Rows.Count), "???") + _
              Application.WorksheetFunction.CountIf(Range("I2:I" & Rows.Count), "???????")
              
    total_rotas = Range("V1").End(xlDown).Row - 1

Application.ScreenUpdating = True

    '(Fim - 100%)
    UserForm5.UpdateProgressBar 100
    DoEvents

    ' Fecha o UserForm ao final do processo
    Unload UserForm5
    DoEvents
    Application.Wait Now + TimeValue("00:00:01")
    
    'reabrindo o arquivo CONSOLIDADO
    Set wb3 = Workbooks.Open(novoNome)
    wb3.Save
    
    Fim = MsgBox("A Base de Distribuição foi criada com sucesso!" & vbNewLine & vbNewLine & "Total de Rotas: " & total_rotas & vbNewLine & "Total de Lojas: " & total_lojas, vbInformation, "Control Tower MB")
    
    ' Chama a sub pública que executa o código antes de fechar
    Call ExecutarAntesDeFechar
    
    'Fechar o arquivo com a Macro sem pedir confirmação
    ThisWorkbook.Close SaveChanges = True
        
End Sub

Public Sub instrucoes()

frmInstrucoes.Show

'instrucao = MsgBox("Olá!" & vbNewLine & vbNewLine & "Seja bem vindo ao arquivo de Criação da Base de Distribuição da Control Tower MB!" & vbNewLine & vbNewLine & "Siga o passo a passo abaixo para criar a base:" & vbNewLine & "1. Abra o arquivo .PDF da Preliminar Definitiva através do aplicativo Acrobat Reader e copie os dados (CTRL+A e CTRL+C)." & vbNewLine & "2. Insira os dados como TEXTO na célula A2: Clique na seta em Colar, selecione 'Colar Especial...' e selecione a opção 'Texto'." & vbNewLine & "3. Clique no botão 'Criar'." & vbNewLine & vbNewLine & "ALERTA: É recomendado fechar todos os demais arquivos Excel antes de prosseguir.", vbInformation, "Control Tower MB")

End Sub

Public Sub CriarArquivoDoZeroNaBase()

confirmacao = MsgBox("Deseja criar um novo arquivo da Tbl_Prog?", vbYesNo + vbQuestion, "Confirmação")

    If confirmacao = 7 Then
    
       Exit Sub
    
    Else
    
        UserForm1.Show
    
    End If

End Sub

