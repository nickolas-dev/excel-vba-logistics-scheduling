Attribute VB_Name = "Módulo5"
Option Explicit

' =============================================================================
' CONFIGURAÇÃO DE AMBIENTE
' =============================================================================
' True = Teste | False = Produção
Private Const MODO_TESTE As Boolean = True
' Em teste, ignorar o gate diário/lock para permitir várias execuções no mesmo dia
Private Const BYPASS_GATE_EM_TESTE As Boolean = True

' --- PROD (definitivo) ---
' Path resolved at runtime by PROD_caminhoTblProg.
' Path resolved at runtime by PROD_caminhoTblProgBackupBase.

' --- TESTE ---
' Path resolved at runtime by TESTE_Raiz.
' Path resolved at runtime by TESTE_caminhoTblProg.
' Path resolved at runtime by TESTE_caminhoTblProgBackupBase.

' =============================================================================
' ARQUIVOS DE CONTROLE (marker/lock)
' =============================================================================
Private Const nomeMarker As String = "AnaliseMoverParaBackup.last"
Private Const nomeLock As String = "AnaliseMoverParaBackup.lock"

' =============================================================================
' BARRA DE PROGRESSO (UserForm3) - APENAS BARRA (sem texto)
' =============================================================================
Private Type tProgressCtx
    total As Long
    current As Long
    frmWidth As Single
    active As Boolean
End Type

Private gProg As tProgressCtx  ' <<<<< AGORA DECLARADO EM NÍVEL DE MÓDULO

' Inicializa a barra: conta .xlsx (sem temporários) e abre o UserForm3 (modeless)
Private Sub Progress_Init(ByVal origemPasta As String)
    On Error Resume Next
    Dim fso As Object, pasta As Object, arq As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set pasta = fso.GetFolder(EnsureTrailingBackslash(origemPasta))
    
    gProg.total = 0
    gProg.current = 0
    gProg.active = False
    
    ' Conta apenas arquivos .xlsx que não sejam temporários (~$)
    For Each arq In pasta.Files
        If LCase$(fso.GetExtensionName(arq.Name)) = "xlsx" Then
            If Left$(arq.Name, 2) <> "~$" Then
                gProg.total = gProg.total + 1
            End If
        End If
    Next arq
    
    If gProg.total = 0 Then Exit Sub  ' nada a mostrar
    
    ' Abre o form modeless e zera a barra
    UserForm3.Show vbModeless
    DoEvents
    
    gProg.frmWidth = UserForm3.fraProgress.InsideWidth
    UserForm3.lblAbertura.Width = 0
    gProg.active = True
End Sub

' Avança a barra um passo (sem texto)
Private Sub Progress_Step()
    On Error Resume Next
    If Not gProg.active Then Exit Sub
    gProg.current = gProg.current + 1
    
    Dim pct As Double
    pct = 0
    If gProg.total > 0 Then pct = gProg.current / gProg.total
    If pct < 0 Then pct = 0
    If pct > 1 Then pct = 1
    
    UserForm3.lblAbertura.Width = gProg.frmWidth * pct
    DoEvents
End Sub

' Fecha o UserForm
Private Sub Progress_Finalize()
    On Error Resume Next
    If gProg.active Then
        Unload UserForm3
        gProg.active = False
    End If
End Sub

' =============================================================================
' FUNÇÃO PRINCIPAL
' =============================================================================
Public Sub AnaliseMoverParaBackup(Optional ByVal ForcarExecucao As Boolean = False)
    Dim prevScr As Boolean, prevDisp As Boolean, prevEv As Boolean
    prevScr = Application.ScreenUpdating
    prevDisp = Application.DisplayAlerts
    prevEv = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    On Error GoTo Fim
    
    ' Resolve caminhos
    Dim caminhoOrigem As String, caminhoBackupBase As String
    ResolveCaminhos caminhoOrigem, caminhoBackupBase
    
    ' Base de backup
    EnsureFolderRecursive caminhoBackupBase
    
    ' Gate diário + lock
    Dim precisaGate As Boolean
    precisaGate = Not (MODO_TESTE And BYPASS_GATE_EM_TESTE) And Not ForcarExecucao
    
    If precisaGate Then
        If Not PodeExecutarHoje(caminhoBackupBase) Then GoTo Fim
        If Not AdquirirLock(caminhoBackupBase) Then GoTo Fim
    Else
        LimparLockSeExistir caminhoBackupBase
    End If
    
    ' INICIA BARRA (somente na 1ª execução do dia e se houver arquivos)
    If precisaGate Then
        Progress_Init caminhoOrigem
        If gProg.active Then Application.Visible = False
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim nomeArq As String
    nomeArq = Dir(EnsureTrailingBackslash(caminhoOrigem) & "*.xlsx")
    
    Do While Len(nomeArq) > 0
        ' Ignora temporários do Excel
        If Left$(nomeArq, 2) = "~$" Then GoTo Proximo
        
        Dim CaminhoArquivo As String
        CaminhoArquivo = EnsureTrailingBackslash(caminhoOrigem) & nomeArq
        
        ' Pula se estiver aberto (usa FSO; NÃO usa Dir aqui!)
        If IsWorkbookBloqueado(CaminhoArquivo) Then GoTo Proximo
        
        ' Abre silenciosamente (ReadOnly)
        Dim wb As Workbook
        On Error Resume Next
        Set wb = Workbooks.Open(FileName:=CaminhoArquivo, UpdateLinks:=False, ReadOnly:=True, _
                                IgnoreReadOnlyRecommended:=True, Notify:=False)
        If wb Is Nothing Then
            On Error GoTo Fim
            GoTo Proximo
        End If
        On Error GoTo Fim
        
        ' === SOMENTE A ABA "Sheet1" ===
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = wb.Worksheets("Sheet1")
        On Error GoTo Fim
        
        If ws Is Nothing Then
            wb.Close SaveChanges:=False
            GoTo Proximo
        End If
        
        
        ' >>> Limpar filtros (cabeçalho na linha 1)
        ClearFilters ws

        
        ' Maior data (somente data) na coluna N (a partir de N2)
        Dim maxDataN As Date, temDataN As Boolean
        temDataN = MaxDateFromColumn_Sheet1(ws, "N", 2, maxDataN)
        If Not temDataN Then
            wb.Close SaveChanges:=False
            GoTo Proximo
        End If
        
        ' Condição: maior data N <= ontem
        If maxDataN <= (Date - 1) Then
            ' Define ano pela maior data na coluna D (fallback: N)
            Dim maxDataD As Date, temDataD As Boolean
            temDataD = MaxDateFromColumn_Sheet1(ws, "D", 2, maxDataD)
            
            Dim anoDestino As Long
            If temDataD Then
                anoDestino = Year(maxDataD)
            Else
                anoDestino = Year(maxDataN)
            End If
            
            ' Garante subpasta de ano
            Dim pastaAno As String
            pastaAno = EnsureTrailingBackslash(caminhoBackupBase) & CStr(anoDestino) & "\"
            EnsureFolderRecursive pastaAno
            
            ' Caminho destino
            Dim caminhoDestino As String
            caminhoDestino = pastaAno & nomeArq
            
            ' Fecha antes de copiar/apagar
            wb.Close SaveChanges:=False
            Set wb = Nothing
            
            ' Copia e apaga, sem sobrescrever (se já existe, pula)
            'If Not fso.FileExists(caminhoDestino) Then
             '   FileCopy caminhoArquivo, caminhoDestino
              '  Kill caminhoArquivo
            'End If
            
            ' Sempre move (sobrescrevendo se necessário)
            If fso.FileExists(caminhoDestino) Then
                Kill caminhoDestino ' apaga o arquivo existente no backup
            End If
            FileCopy CaminhoArquivo, caminhoDestino ' copia o novo
            Kill CaminhoArquivo ' apaga o original
        Else
            wb.Close SaveChanges:=False
            Set wb = Nothing
        End If
        
Proximo:
        ' AVANÇA BARRA UMA VEZ POR ARQUIVO
        If precisaGate Then Progress_Step
        
        nomeArq = Dir() ' próximo arquivo
    Loop
    
    ' >>> SNAPSHOT DIÁRIO (somente PRODUÇÃO e 1ª execução do dia)
    If precisaGate And Not MODO_TESTE Then
        SnapshotCopiaDia caminhoOrigem, caminhoBackupBase
    End If
    
    ' Marca execução do dia (SOMENTE depois de tudo)
    If precisaGate Then AtualizarMarkerELiberarLock caminhoBackupBase
    
Fim:
    ' FECHA BARRA E REEXIBE EXCEL (sempre: fim ou erro)
    Progress_Finalize
    Application.Visible = True
    
    ' Restaura estados
    Application.ScreenUpdating = prevScr
    Application.DisplayAlerts = prevDisp
    Application.EnableEvents = prevEv
End Sub

' =============================================================================
' APOIO: CAMINHOS / PASTAS
' =============================================================================
Private Sub ResolveCaminhos(ByRef caminhoOrigem As String, ByRef caminhoBackupBase As String)
    If MODO_TESTE Then
        caminhoOrigem = EnsureTrailingBackslash(TESTE_caminhoTblProg)
        caminhoBackupBase = EnsureTrailingBackslash(TESTE_caminhoTblProgBackupBase)
    Else
        caminhoOrigem = EnsureTrailingBackslash(PROD_caminhoTblProg)
        caminhoBackupBase = EnsureTrailingBackslash(PROD_caminhoTblProgBackupBase)
    End If
End Sub

Private Function EnsureTrailingBackslash(ByVal p As String) As String
    If Len(p) = 0 Then
        EnsureTrailingBackslash = ""
    ElseIf Right$(p, 1) = "\" Then
        EnsureTrailingBackslash = p
    Else
        EnsureTrailingBackslash = p & "\"
    End If
End Function

' Cria pastas recursivamente (todos os níveis do caminho)
Private Sub EnsureFolderRecursive(ByVal fullPath As String)
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(fullPath) Then Exit Sub
    
    Dim pathNorm As String: pathNorm = Replace(fullPath, "/", "\")
    pathNorm = EnsureTrailingBackslash(pathNorm)
    
    Dim parts() As String: parts = Split(pathNorm, "\")
    Dim build As String, i As Long
    build = ""
    If InStr(parts(0), ":") > 0 Then
        build = parts(0) & "\"
        i = 1
    Else
        i = 0
    End If
    
    For i = i To UBound(parts)
        If Len(parts(i)) > 0 Then
            build = build & parts(i) & "\"
            If Not fso.FolderExists(build) Then fso.CreateFolder build
        End If
    Next i
End Sub

' =============================================================================
' APOIO: GATE DIÁRIO E LOCK
' =============================================================================
Private Function PodeExecutarHoje(ByVal backupBase As String) As Boolean
    On Error Resume Next
    Dim markerPath As String
    markerPath = EnsureTrailingBackslash(backupBase) & nomeMarker
    
    If Len(Dir(markerPath)) > 0 Then
        Dim dt As Date
        dt = FileDateTime(markerPath)
        If DateValue(dt) = Date Then
            PodeExecutarHoje = False
            Exit Function
        End If
    End If
    PodeExecutarHoje = True
End Function

Private Function AdquirirLock(ByVal backupBase As String) As Boolean
    On Error Resume Next
    Dim lockPath As String
    lockPath = EnsureTrailingBackslash(backupBase) & nomeLock
    
    If Len(Dir(lockPath)) > 0 Then
        AdquirirLock = False
        Exit Function
    End If
    
    Dim ff As Integer: ff = FreeFile
    Open lockPath For Output As #ff
    Print #ff, "lock " & Now
    Close #ff
    
    AdquirirLock = True
End Function

Private Sub AtualizarMarkerELiberarLock(ByVal backupBase As String)
    On Error Resume Next
    Dim markerPath As String
    markerPath = EnsureTrailingBackslash(backupBase) & nomeMarker
    
    Dim ff As Integer: ff = FreeFile
    Open markerPath For Output As #ff
    Print #ff, "ultima_execucao=" & Now
    Close #ff
    
    Dim lockPath As String
    lockPath = EnsureTrailingBackslash(backupBase) & nomeLock
    If Len(Dir(lockPath)) > 0 Then Kill lockPath
End Sub

Private Sub LimparLockSeExistir(ByVal backupBase As String)
    On Error Resume Next
    Dim lockPath As String
    lockPath = EnsureTrailingBackslash(backupBase) & nomeLock
    If Len(Dir(lockPath)) > 0 Then Kill lockPath
End Sub

' =============================================================================
' APOIO: ARQUIVO BLOQUEADO (sem usar Dir para não resetar enumeração)
' =============================================================================
Private Function IsWorkbookBloqueado(ByVal CaminhoArquivo As String) As Boolean
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(CaminhoArquivo) Then
        IsWorkbookBloqueado = False
        Exit Function
    End If
    
    Dim pasta As String, nome As String, lockPath As String
    pasta = fso.GetParentFolderName(CaminhoArquivo)
    nome = fso.GetFileName(CaminhoArquivo)
    lockPath = EnsureTrailingBackslash(pasta) & "~$" & nome
    
    IsWorkbookBloqueado = fso.FileExists(lockPath)
End Function

' =============================================================================
' APOIO: DATAS (Sheet1)
' =============================================================================
' Normaliza um texto de data comum: troca "-" por "/", remove "h" no fim, tira duplos espaços/nbsp
Private Function NormalizeDateText(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    t = Replace(t, Chr(160), " ")           ' NBSP
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    t = Replace(t, "-", "/")
    If Right$(t, 1) = "h" Then t = Trim$(Left$(t, Len(t) - 1))
    NormalizeDateText = t
End Function

' Tenta obter a data (sem hora) de uma célula pela ordem: Text(normalizado) -> Value(Date) -> Value2(numérico)
Private Function TryGetDateOnlyFromCell(ByVal c As Range, ByRef dOnly As Date) As Boolean
    On Error Resume Next
    Dim s As String, dt As Date
    
    ' 1) Texto exibido
    s = NormalizeDateText(c.Text)
    If Len(s) > 0 And IsDate(s) Then
        dt = CDate(s)
        dOnly = DateSerial(Year(dt), Month(dt), Day(dt))
        TryGetDateOnlyFromCell = True
        Exit Function
    End If
    
    ' 2) Value (tipo Date)
    If VarType(c.Value) = vbDate Then
        dt = c.Value
        dOnly = DateSerial(Year(dt), Month(dt), Day(dt))
        TryGetDateOnlyFromCell = True
        Exit Function
    End If
    
    ' 3) Value2 numérico (serial Excel)
    If IsNumeric(c.Value2) Then
        dt = CDate(c.Value2)
        dOnly = DateSerial(Year(dt), Month(dt), Day(dt))
        TryGetDateOnlyFromCell = True
        Exit Function
    End If
    
    TryGetDateOnlyFromCell = False
    If Err.Number <> 0 Then Err.Clear
End Function

' Retorna a maior data (sem hora) em uma coluna da "Sheet1" a partir de uma linha inicial
Private Function MaxDateFromColumn_Sheet1(ByVal ws As Worksheet, ByVal columnLetter As String, ByVal startRow As Long, ByRef outDate As Date) As Boolean
    On Error Resume Next
    Dim lastRow As Long, R As Long
    Dim dOnly As Date
    Dim hasData As Boolean
    
    outDate = 0
    hasData = False
    
    lastRow = ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row
    If lastRow < startRow Then
        MaxDateFromColumn_Sheet1 = False
        Exit Function
    End If
    
    For R = startRow To lastRow
        If TryGetDateOnlyFromCell(ws.Cells(R, columnLetter), dOnly) Then
            If Not hasData Or dOnly > outDate Then
                outDate = dOnly
                hasData = True
            End If
        End If
    Next R
    
    MaxDateFromColumn_Sheet1 = hasData
End Function

' Limpa filtros (mostra tudo) e garante setinhas no cabeçalho (linha 1)
Private Sub ClearFilters(ByVal ws As Worksheet)
    On Error Resume Next
    
    ' 1) Se houver filtro aplicado (dados ocultos), mostra tudo
    If ws.FilterMode Then
        ws.ShowAllData
    End If
    
    ' 2) Garante setinhas na linha 1, se ainda não houver AutoFilter
    If Not ws.AutoFilterMode Then
        Dim lastCol As Long
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        If lastCol >= 1 Then
            ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).AutoFilter
        End If
    End If
    
    On Error GoTo 0
End Sub

' Retorna True se o nome for exatamente "dd.mm" (dia.mês) válido
Private Function IsDateFolderName(ByVal nm As String) As Boolean
    On Error Resume Next
    IsDateFolderName = False
    If Len(nm) <> 5 Then Exit Function
    If Mid$(nm, 3, 1) <> "." Then Exit Function
    
    Dim dd As Integer, mm As Integer
    If Not IsNumeric(Left$(nm, 2)) Then Exit Function
    If Not IsNumeric(Right$(nm, 2)) Then Exit Function
    
    dd = CInt(Left$(nm, 2))
    mm = CInt(Right$(nm, 2))
    
    If dd >= 1 And dd <= 31 And mm >= 1 And mm <= 12 Then
        IsDateFolderName = True
    End If
End Function

' Cria a pasta dd.mm do dia em Backup e copia os arquivos .xlsx restantes de origem
Private Sub SnapshotCopiaDia(ByVal origemPasta As String, ByVal backupBase As String)
    On Error Resume Next
    
    ' Só roda em PRODUÇÃO (em teste, sai)
    If MODO_TESTE Then Exit Sub
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim base As String: base = EnsureTrailingBackslash(backupBase)
    
    ' 1) Apaga pastas antigas no formato dd.mm (ex.: "12.03")
    Dim PastaBase As Object, subf As Object
    Set PastaBase = fso.GetFolder(base)
    For Each subf In PastaBase.SubFolders
        If IsDateFolderName(subf.Name) Then
            subf.Delete True       ' apaga recursivamente
        End If
    Next subf
    
    ' 2) Cria pasta de hoje
    Dim nomeHoje As String
    nomeHoje = Format$(Date, "dd.mm")
    Dim pastaHoje As String
    pastaHoje = base & nomeHoje & "\"
    EnsureFolderRecursive pastaHoje
    
    ' 3) Copia os arquivos restantes de origem
    Dim pastaOrigem As Object, arq As Object
    Set pastaOrigem = fso.GetFolder(EnsureTrailingBackslash(origemPasta))
    
    For Each arq In pastaOrigem.Files
        If LCase$(fso.GetExtensionName(arq.Name)) = "xlsx" Then
            If Left$(arq.Name, 2) <> "~$" Then
                Dim fullSrc As String: fullSrc = arq.Path
                
                ' Pula se estiver bloqueado/aberto
                If Not IsWorkbookBloqueado(fullSrc) Then
                    Dim dest As String: dest = pastaHoje & arq.Name
                    ' Copia (pasta do dia acabou de ser criada, então não deve existir conflito)
                    FileCopy fullSrc, dest
                End If
                
                ' Mantém a UI fluida (mesmo que a barra já esteja cheia)
                Progress_Step
            End If
        End If
    Next arq
End Sub


Private Function PROD_caminhoTblProg() As String
    PROD_caminhoTblProg = PortfolioPath("Tbl_Prog\")
End Function

Private Function PROD_caminhoTblProgBackupBase() As String
    PROD_caminhoTblProgBackupBase = PortfolioPath("Backup\Tbl_Prog - Backup\")
End Function

Private Function TESTE_Raiz() As String
    TESTE_Raiz = PortfolioPath("Tbl_Prog_Testes\")
End Function

Private Function TESTE_caminhoTblProg() As String
    TESTE_caminhoTblProg = TESTE_Raiz
End Function

Private Function TESTE_caminhoTblProgBackupBase() As String
    TESTE_caminhoTblProgBackupBase = TESTE_Raiz & "Backup\Tbl_Prog - Backup\"
End Function

