Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_Open()

Application.ScreenUpdating = False

Call ExcluirArquivosExpirados

Call AnaliseMoverParaBackup

Call formatooriginal

Call AnalisandoCalendario

Sheets("Planilha1").Activate

ActiveSheet.Unprotect 'desprotegendo a planilha

'preenchendo a fórmula
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "=TEXT(R[1]C,""mm/dd/aa"")"
    
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "=TEXT(R[1]C,""mm/dd/aa"")"
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'protegendo a planilha

    Sheets("Planilha2").Visible = True
    Sheets("Planilha2").Select
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A1").Select
    Sheets("Planilha2").Visible = False

Sheets("Planilha1").Cells(2, 1).Select

'boas_vindas = MsgBox("Olá!" & vbNewLine & vbNewLine & "Seja bem vindo ao arquivo de Criação da Base de Distribuição da Control Tower MB!" & vbNewLine & vbNewLine & "Siga o passo a passo abaixo para criar a base:" & vbNewLine & "1. Abra o arquivo .PDF da Preliminar Definitiva através do aplicativo Acrobat Reader e copie os dados (CTRL+A e CTRL+C)." & vbNewLine & "2. Insira os dados como TEXTO na célula A2: Clique na seta em Colar, selecione 'Colar Especial...' e selecione a opção 'Texto'." & vbNewLine & "3. Clique no botão 'Criar'." & vbNewLine & vbNewLine & "ALERTA: É recomendado fechar todos os demais arquivos Excel antes de prosseguir.", vbInformation, "Control Tower MB")

Application.ScreenUpdating = True

End Sub


