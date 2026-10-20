Attribute VB_Name = "modPortfolioPaths"
Option Explicit

' Resolve paths relative to the project, including generated workbooks and backups.
Public Function PortfolioPath(ByVal relativePath As String) As String
    Dim fso As Object
    Dim folder As String
    Dim parent As String
    Dim oneDriveRoot As String
    Dim accountEnd As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = ThisWorkbook.Path
    ' Excel may return the personal OneDrive URL even for a local file.
    If LCase$(Left$(folder, 24)) = "https://d.docs.live.net/" Then
        oneDriveRoot = Environ$("OneDriveConsumer")
        If Len(oneDriveRoot) = 0 Then oneDriveRoot = Environ$("OneDrive")
        accountEnd = InStr(25, folder, "/")
        If Len(oneDriveRoot) > 0 And accountEnd > 0 Then
            folder = fso.BuildPath(oneDriveRoot, Replace(Mid$(folder, accountEnd + 1), "/", "\"))
            folder = Replace(folder, "%20", " ")
        End If
    End If
    Do While Len(folder) > 0
        If fso.FileExists(fso.BuildPath(folder, "criacao_prog_v5_6_testes.xlsm")) And _
           fso.FileExists(fso.BuildPath(folder, "modelo_template.xlsm")) Then
            PortfolioPath = fso.GetAbsolutePathName(fso.BuildPath(folder, relativePath))
            If Right$(relativePath, 1) = "\" And Right$(PortfolioPath, 1) <> "\" Then
                PortfolioPath = PortfolioPath & "\"
            End If
            Exit Function
        End If
        parent = fso.GetParentFolderName(folder)
        If parent = folder Then Exit Do
        folder = parent
    Loop
    Err.Raise vbObjectError + 2056, "PortfolioPath", _
        "Project folder not found. Keep the project files together."
End Function


