'===============================================================================
' MODULE: modUI.bas
' DESCRIPTION: Interface utilisateur pour le gestionnaire SEPA XML
'===============================================================================
Option Explicit

' =============================================================================
' SUB: CreateMainMenu
' DESCRIPTION: Crée un menu principal dans la feuille Principal
' =============================================================================
Public Sub CreateMainMenu()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim BtnRange As Range
    Dim i As Long
    
    Set ws = GetWorksheetByName(SHEET_PRINCIPAL)
    
    ' Supprimer les boutons existants
    Dim btn As Button
    For Each btn In ws.Buttons
        btn.Delete
    Next btn
    
    ' Créer les boutons dans la colonne P (à droite)
    i = 2
    
    ' Bouton Importer XML
    Set BtnRange = ws.Range("P" & i & ":R" & i)
    CreateButton ws, BtnRange, "Importer XML", "Main_ImportXML"
    i = i + 2
    
    ' Bouton Valider
    Set BtnRange = ws.Range("P" & i & ":R" & i)
    CreateButton ws, BtnRange, "Valider Données", "Main_ValidateData"
    i = i + 2
    
    ' Bouton Exporter XML
    Set BtnRange = ws.Range("P" & i & ":R" & i)
    CreateButton ws, BtnRange, "Exporter XML", "Main_ExportXML"
    i = i + 2
    
    ' Bouton Recalculer Totaux
    Set BtnRange = ws.Range("P" & i & ":R" & i)
    CreateButton ws, BtnRange, "Recalculer Totaux", "Main_RecalculateTotals"
    i = i + 2
    
    ' Bouton Effacer Erreurs
    Set BtnRange = ws.Range("P" & i & ":R" & i)
    CreateButton ws, BtnRange, "Effacer Erreurs", "Main_ClearErrors"
    i = i + 2
    
    ' Bouton Afficher Résumé Erreurs
    Set BtnRange = ws.Range("P" & i & ":R" & i)
    CreateButton ws, BtnRange, "Résumé Erreurs", "Main_ShowErrorSummary"
    
    LogInfo "Menu principal créé", "modUI.CreateMainMenu"
    
    On Error GoTo 0
End Sub

' =============================================================================
' SUB: CreateButton
' DESCRIPTION: Crée un bouton dans une plage donnée
' =============================================================================
Private Sub CreateButton(ByRef ws As Worksheet, ByRef TargetRange As Range, _
                         ByVal Caption As String, ByVal MacroName As String)
    On Error Resume Next
    
    Dim btn As Button
    
    Set btn = ws.Buttons.Add(Left:=TargetRange.Left, Top:=TargetRange.Top, _
                             Width:=TargetRange.Width, Height:=TargetRange.Height)
    
    With btn
        .Caption = Caption
        .OnAction = MacroName
        .Font.Bold = True
        .Font.Size = 9
        .Interior.Color = RGB(79, 129, 189)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    On Error GoTo 0
End Sub

' =============================================================================
' SUB: ShowImportDialog
' DESCRIPTION: Affiche la boîte de dialogue d'import
' =============================================================================
Public Sub ShowImportDialog()
    On Error Resume Next
    
    Dim FilePath As String
    
    FilePath = Application.GetOpenFilename( _
        FileFilter:="XML Files (*.xml), *.xml", _
        Title:="Sélectionner un fichier XML SEPA à importer")
    
    If FilePath = "False" Then Exit Sub
    
    ' Importer le fichier
    If ImportXMLFile(FilePath) Then
        MsgBox "Import réussi!" & vbCrLf & _
               "Le fichier a été chargé dans la feuille '" & SHEET_PRINCIPAL & "'.", _
               vbInformation + vbOKOnly, "SEPA XML Manager - Succès"
        
        ' Créer le menu
        Call CreateMainMenu
        
        ' Recalculer les totaux
        Call RecalculateTotals
    Else
        MsgBox "Échec de l'import." & vbCrLf & _
               "Consultez la feuille '" & SHEET_ERREURS & "' pour plus de détails.", _
               vbCritical + vbOKOnly, "SEPA XML Manager - Erreur"
    End If
    
    On Error GoTo 0
End Sub

' =============================================================================
' SUB: ShowExportDialog
' DESCRIPTION: Affiche la boîte de dialogue d'export
' =============================================================================
Public Sub ShowExportDialog()
    On Error Resume Next
    
    Dim FilePath As String
    Dim Success As Boolean
    
    ' Demander le chemin du fichier
    FilePath = Application.GetSaveAsFilename( _
        InitialFileName:="SEPA_" & Format(Now, "yyyymmdd_hhmmss") & ".xml", _
        FileFilter:="XML Files (*.xml), *.xml", _
        Title:="Enregistrer le fichier XML SEPA")
    
    If FilePath = "False" Then Exit Sub
    
    ' Exporter le fichier
    Success = ExportXMLFile(FilePath, True)
    
    If Success Then
        MsgBox "Export réussi!" & vbCrLf & _
               "Fichier XML: " & FilePath & vbCrLf & _
               "Fichier TXT: " & Left(FilePath, InStrRev(FilePath, ".")) & "txt", _
               vbInformation + vbOKOnly, "SEPA XML Manager - Succès"
    Else
        MsgBox "Échec de l'export." & vbCrLf & _
               "Consultez la feuille '" & SHEET_ERREURS & "' pour plus de détails.", _
               vbCritical + vbOKOnly, "SEPA XML Manager - Erreur"
    End If
    
    On Error GoTo 0
End Sub

' =============================================================================
' SUB: ShowValidationResults
' DESCRIPTION: Affiche les résultats de la validation
' =============================================================================
Public Sub ShowValidationResults()
    On Error Resume Next
    
    Dim IsValid As Boolean
    Dim Msg As String
    
    IsValid = ValidateAllData()
    
    If IsValid Then
        Msg = "✓ Toutes les validations ont réussi!" & vbCrLf & vbCrLf
        Msg = Msg & "Le fichier est prêt à être exporté."
        MsgBox Msg, vbInformation + vbOKOnly, "SEPA XML Manager - Validation"
    Else
        Msg = "✗ Des erreurs de validation ont été détectées." & vbCrLf & vbCrLf
        Msg = Msg & "Nombre total d'erreurs: " & GetErrorCount() & vbCrLf & vbCrLf
        Msg = Msg & "Veuillez consulter la feuille '" & SHEET_ERREURS & "'" & vbCrLf
        Msg = Msg & "et corriger les problèmes avant l'export."
        MsgBox Msg, vbExclamation + vbOKOnly, "SEPA XML Manager - Validation"
        
        ' Afficher la feuille Erreurs
        ThisWorkbook.Worksheets(SHEET_ERREURS).Activate
    End If
    
    On Error GoTo 0
End Sub

' =============================================================================
' SUB: AddTransactionRow
' DESCRIPTION: Ajoute une nouvelle ligne de transaction
' =============================================================================
Public Sub AddTransactionRow()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim NewRow As Long
    
    Set ws = ThisWorkbook.Worksheets(SHEET_PRINCIPAL)
    If ws Is Nothing Then
        MsgBox "Feuille Principal introuvable!", vbCritical
        Exit Sub
    End If
    
    ' Trouver la dernière ligne
    LastRow = ws.Cells(ws.Rows.Count, COL_TX_ENDTOENDID).End(xlUp).Row
    
    If LastRow < ROW_TRANSACTION_START Then
        NewRow = ROW_TRANSACTION_START
    Else
        NewRow = LastRow + 1
    End If
    
    ' Insérer une nouvelle ligne avec des valeurs par défaut
    With ws
        .Cells(NewRow, COL_TX_ENDTOENDID).Value = "ENDTOEND_" & Format(Now, "hhmmss")
        .Cells(NewRow, COL_TX_INSTRID).Value = ""
        .Cells(NewRow, COL_TX_CDTR_NM).Value = "Nouveau Bénéficiaire"
        .Cells(NewRow, COL_TX_CDTR_STRTNM).Value = ""
        .Cells(NewRow, COL_TX_CDTR_PSTCD).Value = ""
        .Cells(NewRow, COL_TX_CDTR_TWNNM).Value = ""
        .Cells(NewRow, COL_TX_CDTR_CTRY).Value = "TN"
        .Cells(NewRow, COL_TX_CDTR_IBAN).Value = ""
        .Cells(NewRow, COL_TX_CDTR_BIC).Value = ""
        .Cells(NewRow, COL_TX_CDTR_BANKNM).Value = ""
        .Cells(NewRow, COL_TX_AMT).Value = 0
        .Cells(NewRow, COL_TX_CCY).Value = SafeCStr(.Range(CELL_DEVISE).Value)
        .Cells(NewRow, COL_TX_USTRD).Value = SafeCStr(.Range(CELL_PMTINFID).Value)
    End With
    
    ' Sélectionner la nouvelle ligne
    ws.Rows(NewRow).Select
    
    ' Recalculer les totaux
    Call RecalculateTotals
    
    LogInfo "Nouvelle ligne de transaction ajoutée: " & NewRow, "modUI.AddTransactionRow"
    
    On Error GoTo 0
End Sub

' =============================================================================
' SUB: DeleteSelectedTransaction
' DESCRIPTION: Supprime la ligne de transaction sélectionnée
' =============================================================================
Public Sub DeleteSelectedTransaction()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim SelectedRow As Long
    
    Set ws = ThisWorkbook.ActiveSheet
    
    If ws.Name <> SHEET_PRINCIPAL Then
        MsgBox "Veuillez sélectionner une ligne dans la feuille '" & SHEET_PRINCIPAL & "'.", _
               vbExclamation
        Exit Sub
    End If
    
    SelectedRow = ActiveCell.Row
    
    If SelectedRow < ROW_TRANSACTION_START Then
        MsgBox "Veuillez sélectionner une ligne de transaction (ligne " & _
               ROW_TRANSACTION_START & " ou supérieure).", vbExclamation
        Exit Sub
    End If
    
    ' Confirmer la suppression
    If MsgBox("Voulez-vous vraiment supprimer la ligne " & SelectedRow & "?", _
              vbYesNo + vbQuestion, "Confirmation") = vbYes Then
        
        ws.Rows(SelectedRow).Delete Shift:=xlUp
        
        ' Recalculer les totaux
        Call RecalculateTotals
        
        LogInfo "Ligne de transaction supprimée: " & SelectedRow, "modUI.DeleteSelectedTransaction"
    End If
    
    On Error GoTo 0
End Sub

' =============================================================================
' SUB: SetupQuickAccessToolbar
' DESCRIPTION: Configure les raccourcis clavier et conseils
' =============================================================================
Public Sub SetupQuickAccessToolbar()
    On Error Resume Next
    
    ' Afficher les instructions dans une cellule en haut
    Dim ws As Worksheet
    Set ws = GetWorksheetByName(SHEET_PRINCIPAL)
    
    With ws.Range("T1:T10")
        .ClearContents
        .Font.Size = 9
        .Font.Italic = True
    End With
    
    ws.Range("T1").Value = "=== RACCOURCIS ==="
    ws.Range("T2").Value = "Ctrl+I: Importer XML"
    ws.Range("T3").Value = "Ctrl+E: Exporter XML"
    ws.Range("T4").Value = "Ctrl+V: Valider"
    ws.Range("T5").Value = "Ctrl+T: Recalculer"
    ws.Range("T6").Value = "Ctrl+N: Nouvelle ligne"
    ws.Range("T7").Value = "Ctrl+D: Supprimer ligne"
    ws.Range("T9").Value = "=== INFO ==="
    ws.Range("T10").Value = "Format: pain.001.001.03"
    
    On Error GoTo 0
End Sub

' =============================================================================
' SUB: HighlightErrors
' DESCRIPTION: Met en surbrillance les cellules avec des erreurs
' =============================================================================
Public Sub HighlightErrors()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim ErrRange As Range
    Dim Cell As Range
    
    Set ws = ThisWorkbook.Worksheets(SHEET_PRINCIPAL)
    If ws Is Nothing Then Exit Sub
    
    ' Effacer le formatage précédent
    ws.Cells.Interior.ColorIndex = xlNone
    
    ' Parser la feuille Erreurs pour trouver les cellules concernées
    Dim ErrWs As Worksheet
    Set ErrWs = ThisWorkbook.Worksheets(SHEET_ERREURS)
    If ErrWs Is Nothing Then Exit Sub
    
    Dim LastRow As Long
    Dim i As Long
    
    LastRow = ErrWs.Cells(ErrWs.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
        Dim CellRef As String
        CellRef = SafeCStr(ErrWs.Cells(i, 7).Value)
        
        If Len(CellRef) > 0 And InStr(CellRef, "$") > 0 Then
            On Error Resume Next
            Set ErrRange = ws.Range(CellRef)
            If Not ErrRange Is Nothing Then
                ErrRange.Interior.Color = RGB(255, 200, 200)
            End If
            Set ErrRange = Nothing
            On Error GoTo 0
        End If
    Next i
    
    LogInfo "Surbrillance des erreurs appliquée", "modUI.HighlightErrors"
    
    On Error GoTo 0
End Sub
