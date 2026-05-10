'===============================================================================
' MODULE: modLogging.bas
' DESCRIPTION: Gestion des logs et des erreurs pour le gestionnaire SEPA XML
'===============================================================================
Option Explicit

' Collection globale pour stocker les erreurs courantes
Public g_ErrorCollection As Collection

' =============================================================================
' SUB: InitializeErrorLogging
' DESCRIPTION: Initialise la collection d'erreurs
' =============================================================================
Public Sub InitializeErrorLogging()
    On Error Resume Next
    Set g_ErrorCollection = New Collection
    On Error GoTo 0
End Sub

' =============================================================================
' SUB: ClearErrorLog
' DESCRIPTION: Efface tous les logs d'erreurs
' =============================================================================
Public Sub ClearErrorLog()
    On Error Resume Next
    
    If Not g_ErrorCollection Is Nothing Then
        Set g_ErrorCollection = New Collection
    End If
    
    ' Effacer la feuille Erreurs
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_ERREURS)
    If Not ws Is Nothing Then
        ws.Cells.Clear
    End If
    On Error GoTo 0
End Sub

' =============================================================================
' SUB: LogError
' DESCRIPTION: Enregistre une erreur dans le log et optionnellement dans Excel
' =============================================================================
Public Sub LogError(ByVal ErrorCode As Long, _
                    ByVal Description As String, _
                    ByVal Source As String, _
                    Optional ByVal TransactionRef As String = "", _
                    Optional ByVal CellReference As String = "", _
                    Optional ByVal DisplayError As Boolean = False, _
                    Optional ByVal ErrorLevel As String = "")
    
    Dim ErrInfo As TErrorInfo
    Dim Level As String
    
    ' Déterminer le niveau d'erreur
    If ErrorLevel = "" Then
        If ErrorCode >= 3000 Then
            Level = ERR_LEVEL_CRITICAL
        ElseIf ErrorCode >= 2000 Then
            Level = ERR_LEVEL_WARNING
        Else
            Level = ERR_LEVEL_INFO
        End If
    Else
        Level = ErrorLevel
    End If
    
    ' Remplir la structure d'erreur
    ErrInfo.ErrorCode = ErrorCode
    ErrInfo.ErrorLevel = Level
    ErrInfo.Description = Description
    ErrInfo.Source = Source
    ErrInfo.TransactionRef = TransactionRef
    ErrInfo.CellReference = CellReference
    ErrInfo.Timestamp = Now
    
    ' Ajouter à la collection
    If g_ErrorCollection Is Nothing Then
        Set g_ErrorCollection = New Collection
    End If
    
    On Error Resume Next
    g_ErrorCollection.Add ErrInfo, "ERR_" & Format(Now, "yyyymmddhhmmss") & "_" & CStr(ErrorCode)
    On Error GoTo 0
    
    ' Écrire dans la feuille Erreurs
    WriteErrorToSheet ErrInfo
    
    ' Afficher un message si demandé
    If DisplayError Then
        MsgBox "Erreur " & Level & vbCrLf & _
               "Code: " & ErrorCode & vbCrLf & _
               "Description: " & Description & vbCrLf & _
               "Source: " & Source, _
               vbExclamation + vbOKOnly, "SEPA XML Manager - Erreur"
    End If
End Sub

' =============================================================================
' SUB: WriteErrorToSheet
' DESCRIPTION: Écrit une erreur dans la feuille Excel "Erreurs"
' =============================================================================
Private Sub WriteErrorToSheet(ByRef ErrInfo As TErrorInfo)
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim NextRow As Long
    
    ' Récupérer ou créer la feuille Erreurs
    Set ws = GetWorksheetByName(SHEET_ERREURS)
    
    ' Initialiser les en-têtes si nécessaire
    If ws.Range("A1").Value = "" Then
        Call SetupErrorSheet(ws)
    End If
    
    ' Trouver la prochaine ligne libre
    NextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If NextRow < 2 Then NextRow = 2
    
    ' Écrire les informations d'erreur
    ws.Cells(NextRow, 1).Value = ErrInfo.Timestamp
    ws.Cells(NextRow, 2).Value = ErrInfo.ErrorLevel
    ws.Cells(NextRow, 3).Value = ErrInfo.ErrorCode
    ws.Cells(NextRow, 4).Value = ErrInfo.Description
    ws.Cells(NextRow, 5).Value = ErrInfo.Source
    ws.Cells(NextRow, 6).Value = ErrInfo.TransactionRef
    ws.Cells(NextRow, 7).Value = ErrInfo.CellReference
    
    ' Formater la ligne
    With ws.Range(ws.Cells(NextRow, 1), ws.Cells(NextRow, 7))
        .Font.Name = "Calibri"
        .Font.Size = 9
        
        ' Couleur selon le niveau
        Select Case ErrInfo.ErrorLevel
            Case ERR_LEVEL_CRITICAL
                .Interior.Color = RGB(255, 200, 200)
                .Font.Bold = True
            Case ERR_LEVEL_WARNING
                .Interior.Color = RGB(255, 255, 200)
            Case ERR_LEVEL_INFO
                .Interior.Color = RGB(200, 255, 200)
        End Select
    End With
    
    On Error GoTo 0
End Sub

' =============================================================================
' SUB: SetupErrorSheet
' DESCRIPTION: Configure la feuille Erreurs avec les en-têtes
' =============================================================================
Public Sub SetupErrorSheet(ws As Worksheet)
    On Error Resume Next
    
    ' Effacer la feuille
    ws.Cells.Clear
    
    ' Définir les en-têtes
    ws.Range("A1:G1").Font.Bold = True
    ws.Range("A1:G1").Interior.Color = RGB(200, 200, 200)
    ws.Range("A1:G1").HorizontalAlignment = xlCenter
    ws.Range("A1:G1").WrapText = True
    
    ws.Cells(1, 1).Value = "Horodatage"
    ws.Cells(1, 2).Value = "Niveau"
    ws.Cells(1, 3).Value = "Code"
    ws.Cells(1, 4).Value = "Description"
    ws.Cells(1, 5).Value = "Source"
    ws.Cells(1, 6).Value = "Transaction"
    ws.Cells(1, 7).Value = "Cellule"
    
    ' Ajuster la largeur des colonnes
    ws.Columns("A:A").ColumnWidth = 18
    ws.Columns("B:B").ColumnWidth = 12
    ws.Columns("C:C").ColumnWidth = 10
    ws.Columns("D:D").ColumnWidth = 45
    ws.Columns("E:E").ColumnWidth = 25
    ws.Columns("F:F").ColumnWidth = 20
    ws.Columns("G:G").ColumnWidth = 12
    
    ' Figer la première ligne
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    On Error GoTo 0
End Sub

' =============================================================================
' SUB: LogInfo
' DESCRIPTION: Enregistre une information dans le log
' =============================================================================
Public Sub LogInfo(ByVal Message As String, Optional ByVal Source As String = "")
    LogError ERR_NONE, Message, Source, , , False, ERR_LEVEL_INFO
End Sub

' =============================================================================
' SUB: LogWarning
' DESCRIPTION: Enregistre un avertissement dans le log
' =============================================================================
Public Sub LogWarning(ByVal Message As String, Optional ByVal Source As String = "")
    LogError 0, Message, Source, , , False, ERR_LEVEL_WARNING
End Sub

' =============================================================================
' FUNCTION: GetErrorCount
' DESCRIPTION: Retourne le nombre d'erreurs enregistrées
' =============================================================================
Public Function GetErrorCount() As Long
    On Error Resume Next
    
    If g_ErrorCollection Is Nothing Then
        GetErrorCount = 0
        Exit Function
    End If
    
    GetErrorCount = g_ErrorCollection.Count
    
    On Error GoTo 0
End Function

' =============================================================================
' FUNCTION: HasCriticalErrors
' DESCRIPTION: Vérifie s'il y a des erreurs critiques
' =============================================================================
Public Function HasCriticalErrors() As Boolean
    Dim ErrItem As Variant
    Dim ErrInfo As TErrorInfo
    
    HasCriticalErrors = False
    
    If g_ErrorCollection Is Nothing Then Exit Function
    
    On Error Resume Next
    For Each ErrItem In g_ErrorCollection
        If TypeName(ErrItem) = "TErrorInfo" Then
            ErrInfo = ErrItem
            If ErrInfo.ErrorLevel = ERR_LEVEL_CRITICAL Then
                HasCriticalErrors = True
                Exit Function
            End If
        End If
    Next ErrItem
    On Error GoTo 0
End Function

' =============================================================================
' SUB: ExportErrorLog
' DESCRIPTION: Exporte le log d'erreurs vers un fichier texte
' =============================================================================
Public Sub ExportErrorLog(Optional ByVal FilePath As String = "")
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim fNum As Integer
    Dim i As Long, j As Long
    Dim LastRow As Long
    Dim LineText As String
    
    Set ws = ThisWorkbook.Worksheets(SHEET_ERREURS)
    If ws Is Nothing Then Exit Sub
    
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If LastRow < 2 Then
        MsgBox "Aucune erreur à exporter.", vbInformation
        Exit Sub
    End If
    
    ' Si aucun chemin fourni, demander à l'utilisateur
    If FilePath = "" Then
        FilePath = Application.GetSaveAsFilename( _
            InitialFileName:="SEPA_ErrorLog_" & Format(Now, "yyyymmdd_hhmmss") & ".txt", _
            FileFilter:="Text Files (*.txt), *.txt")
        
        If FilePath = "False" Then Exit Sub
    End If
    
    ' Ouvrir le fichier
    fNum = FreeFile
    Open FilePath For Output As #fNum
    
    ' Écrire l'en-tête
    Print #fNum, "LOG D'ERREURS - SEPA XML Manager"
    Print #fNum, "Généré le: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Print #fNum, String(100, "-")
    Print #fNum, ""
    
    ' Écrire les en-têtes de colonnes
    Print #fNum, "Date" & vbTab & _
                 "Niveau" & vbTab & _
                 "Code" & vbTab & _
                 "Description" & vbTab & _
                 "Source" & vbTab & _
                 "Transaction" & vbTab & _
                 "Cellule"
    Print #fNum, String(100, "-")
    
    ' Écrire chaque ligne
    For i = 2 To LastRow
        LineText = ""
        For j = 1 To 7
            If j > 1 Then LineText = LineText & vbTab
            LineText = LineText & CStr(ws.Cells(i, j).Value)
        Next j
        Print #fNum, LineText
    Next i
    
    Close #fNum
    
    MsgBox "Log d'erreurs exporté vers:" & vbCrLf & FilePath, vbInformation
    
    On Error GoTo 0
End Sub

' =============================================================================
' SUB: ShowErrorSummary
' DESCRIPTION: Affiche un résumé des erreurs
' =============================================================================
Public Sub ShowErrorSummary()
    Dim CriticalCount As Long
    Dim WarningCount As Long
    Dim InfoCount As Long
    Dim Msg As String
    
    CountErrorsByLevel CriticalCount, WarningCount, InfoCount
    
    Msg = "RÉSUMÉ DES ERREURS" & vbCrLf & vbCrLf
    Msg = Msg & "Critiques: " & CriticalCount & vbCrLf
    Msg = Msg & "Avertissements: " & WarningCount & vbCrLf
    Msg = Msg & "Informations: " & InfoCount & vbCrLf
    Msg = Msg & "Total: " & GetErrorCount()
    
    If CriticalCount > 0 Then
        MsgBox Msg, vbCritical + vbOKOnly, "SEPA XML Manager - Résumé"
    ElseIf WarningCount > 0 Then
        MsgBox Msg, vbExclamation + vbOKOnly, "SEPA XML Manager - Résumé"
    Else
        MsgBox Msg, vbInformation + vbOKOnly, "SEPA XML Manager - Résumé"
    End If
End Sub

' =============================================================================
' SUB: CountErrorsByLevel
' DESCRIPTION: Compte les erreurs par niveau
' =============================================================================
Private Sub CountErrorsByLevel(ByRef CriticalCount As Long, _
                               ByRef WarningCount As Long, _
                               ByRef InfoCount As Long)
    Dim ErrItem As Variant
    Dim ErrInfo As TErrorInfo
    
    CriticalCount = 0
    WarningCount = 0
    InfoCount = 0
    
    If g_ErrorCollection Is Nothing Then Exit Sub
    
    On Error Resume Next
    For Each ErrItem In g_ErrorCollection
        If TypeName(ErrItem) = "TErrorInfo" Then
            ErrInfo = ErrItem
            Select Case ErrInfo.ErrorLevel
                Case ERR_LEVEL_CRITICAL
                    CriticalCount = CriticalCount + 1
                Case ERR_LEVEL_WARNING
                    WarningCount = WarningCount + 1
                Case ERR_LEVEL_INFO
                    InfoCount = InfoCount + 1
            End Select
        End If
    Next ErrItem
    On Error GoTo 0
End Sub
