' ====================================================================================================
' MODULE: modUI
' DESCRIPTION: Interface utilisateur et gestion des événements
' ====================================================================================================
Option Explicit

' ====================================================================================================
' SUB: CreateRibbonMenu
' DESCRIPTION: Crée un menu personnalisé dans le ruban (à implémenter via CustomUI)
' ====================================================================================================
Public Sub CreateRibbonMenu()
    ' Note: Le ruban personnalisé doit être créé via un fichier customUI.xml
    ' Ce code fournit les callbacks pour les boutons
    
    MsgBox "Pour ajouter un ruban personnalisé:" & vbCrLf & _
           "1. Créez un dossier 'customUI' dans le fichier Excel" & vbCrLf & _
           "2. Ajoutez un fichier customUI.xml" & vbCrLf & _
           "3. Importez-le avec l'éditeur Office Open XML", vbInformation
End Sub

' ====================================================================================================
' SUB: ShowMainMenu
' DESCRIPTION: Affiche le menu principal sous forme de UserForm ou MsgBox
' ====================================================================================================
Public Sub ShowMainMenu()
    On Error GoTo ErrorHandler
    
    Dim choice As VbMsgBoxResult
    Dim msg As String
    
    msg = "=== SEPA XML Manager ===" & vbCrLf & vbCrLf & _
          "Choisissez une action:" & vbCrLf & vbCrLf & _
          "1. Importer un fichier XML" & vbCrLf & _
          "2. Valider les données" & vbCrLf & _
          "3. Exporter vers XML" & vbCrLf & _
          "4. Exporter vers TXT" & vbCrLf & _
          "5. Recalculer les totaux" & vbCrLf & _
          "6. Ajouter une transaction" & vbCrLf & _
          "7. Aide"
    
    choice = MsgBox(msg, vbOKCancel + vbQuestion, "SEPA XML Manager")
    
    If choice = vbOK Then
        Call ShowActionMenu
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError "modUI.ShowMainMenu", Err.Number, Err.Description
End Sub

' ====================================================================================================
' SUB: ShowActionMenu
' DESCRIPTION: Affiche le menu d'actions détaillé
' ====================================================================================================
Public Sub ShowActionMenu()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim filePath As String
    Dim result As Boolean
    
    Set ws = ThisWorkbook.Worksheets(SHEET_PRINCIPAL)
    
    ' Créer une interface simple avec des boutons
    Dim frm As Object
    Dim btnImport As Object, btnValidate As Object, btnExport As Object
    Dim btnExportTxt As Object, btnRecalc As Object, btnAddTx As Object
    
    ' Pour simplifier, utiliser des InputBox et MsgBox
    Dim choice As String
    
    choice = InputBox( _
        "Tapez le numéro de l'action:" & vbCrLf & vbCrLf & _
        "1 - Importer XML" & vbCrLf & _
        "2 - Valider" & vbCrLf & _
        "3 - Exporter XML" & vbCrLf & _
        "4 - Exporter TXT" & vbCrLf & _
        "5 - Recalculer" & vbCrLf & _
        "6 - Ajouter Transaction" & vbCrLf & _
        "7 - Quitter", _
        "Action", _
        "1")
    
    Select Case choice
        Case "1"
            Call ImportXML_Action
        Case "2"
            Call Validate_Action
        Case "3"
            Call ExportXML_Action
        Case "4"
            Call ExportTXT_Action
        Case "5"
            Call Recalculate_Action
        Case "6"
            Call AddTransaction_Action
        Case "7"
            Exit Sub
        Case Else
            MsgBox "Action non reconnue.", vbExclamation
    End Select
    
    Exit Sub
    
ErrorHandler:
    LogError "modUI.ShowActionMenu", Err.Number, Err.Description
End Sub

' ====================================================================================================
' SUB: ImportXML_Action
' DESCRIPTION: Action d'import XML
' ====================================================================================================
Public Sub ImportXML_Action()
    On Error GoTo ErrorHandler
    
    Dim filePath As String
    Dim result As Boolean
    
    ' Ouvrir la boîte de dialogue de sélection de fichier
    filePath = SelectFileDialog("Fichiers XML|*.xml|Tous les fichiers|*.*", "Sélectionner un fichier XML")
    
    If filePath = "" Then
        Exit Sub
    End If
    
    ' Importer le fichier
    result = ImportXMLFile(filePath)
    
    If result Then
        ' Mettre à jour l'affichage
        Call RefreshDisplay
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError "modUI.ImportXML_Action", Err.Number, Err.Description
    MsgBox "Erreur lors de l'import: " & Err.Description, vbCritical
End Sub

' ====================================================================================================
' SUB: Validate_Action
' DESCRIPTION: Action de validation
' ====================================================================================================
Public Sub Validate_Action()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim result As Boolean
    
    Set ws = ThisWorkbook.Worksheets(SHEET_PRINCIPAL)
    
    If ws Is Nothing Then
        MsgBox "La feuille Principal n'existe pas.", vbCritical
        Exit Sub
    End If
    
    result = ValidateAll(ws)
    
    Exit Sub
    
ErrorHandler:
    LogError "modUI.Validate_Action", Err.Number, Err.Description
    MsgBox "Erreur lors de la validation: " & Err.Description, vbCritical
End Sub

' ====================================================================================================
' SUB: ExportXML_Action
' DESCRIPTION: Action d'export XML
' ====================================================================================================
Public Sub ExportXML_Action()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim filePath As String
    Dim result As Boolean
    
    Set ws = ThisWorkbook.Worksheets(SHEET_PRINCIPAL)
    
    If ws Is Nothing Then
        MsgBox "La feuille Principal n'existe pas.", vbCritical
        Exit Sub
    End If
    
    ' Valider d'abord
    If Not ValidateAll(ws) Then
        Dim continueChoice As VbMsgBoxResult
        continueChoice = MsgBox("Des erreurs ont été détectées. Voulez-vous continuer quand même?", _
                                vbYesNo + vbExclamation, "Continuer?")
        If continueChoice = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Demander le chemin de sortie
    filePath = SaveFileDialog("Fichiers XML|*.xml|Tous les fichiers|*.*", "sepa-export.xml")
    
    If filePath = "" Then
        Exit Sub
    End If
    
    ' Exporter
    result = ExportToXML(filePath)
    
    Exit Sub
    
ErrorHandler:
    LogError "modUI.ExportXML_Action", Err.Number, Err.Description
    MsgBox "Erreur lors de l'export: " & Err.Description, vbCritical
End Sub

' ====================================================================================================
' SUB: ExportTXT_Action
' DESCRIPTION: Action d'export TXT
' ====================================================================================================
Public Sub ExportTXT_Action()
    On Error GoTo ErrorHandler
    
    Dim filePath As String
    Dim result As Boolean
    
    filePath = SaveFileDialog("Fichiers TXT|*.txt|Tous les fichiers|*.*", "sepa-export.txt")
    
    If filePath = "" Then
        Exit Sub
    End If
    
    result = ExportToTXT(filePath)
    
    If result Then
        MsgBox "Export TXT réussi!", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError "modUI.ExportTXT_Action", Err.Number, Err.Description
End Sub

' ====================================================================================================
' SUB: Recalculate_Action
' DESCRIPTION: Action de recalcul des totaux
' ====================================================================================================
Public Sub Recalculate_Action()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim txCount As Long
    Dim ctrlSum As Double
    
    Set ws = ThisWorkbook.Worksheets(SHEET_PRINCIPAL)
    
    If ws Is Nothing Then
        Exit Sub
    End If
    
    ' Calculer le nombre de transactions
    txCount = CountNonEmptyTransactions(ws)
    
    ' Calculer la somme de contrôle
    ctrlSum = CalculateCtrlSumFromExcel(ws)
    
    ' Mettre à jour les cellules
    Application.ScreenUpdating = False
    
    ws.Range(RANGE_NBFTXS).Value = txCount
    ws.Range(RANGE_CTRL_SUM_PI).Value = FormatAmount(ctrlSum)
    ws.Range(RANGE_CTRLSUM).Value = FormatAmount(ctrlSum)
    
    Application.ScreenUpdating = True
    
    MsgBox "Totaux recalculés:" & vbCrLf & _
           "Nombre de transactions: " & txCount & vbCrLf & _
           "Somme totale: " & FormatAmount(ctrlSum) & " EUR", vbInformation
    
    Exit Sub
    
ErrorHandler:
    LogError "modUI.Recalculate_Action", Err.Number, Err.Description
End Sub

' ====================================================================================================
' SUB: AddTransaction_Action
' DESCRIPTION: Action d'ajout de transaction
' ====================================================================================================
Public Sub AddTransaction_Action()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim newId As String
    
    Set ws = ThisWorkbook.Worksheets(SHEET_PRINCIPAL)
    
    If ws Is Nothing Then
        Exit Sub
    End If
    
    ' Trouver la dernière ligne
    lastRow = ws.Cells(ws.Rows.Count, TX_COL_ENDTOENDID).End(xlUp).Row
    
    ' Si aucune transaction, commencer à TX_START_ROW
    If lastRow < TX_START_ROW Then
        lastRow = TX_START_ROW - 1
    End If
    
    ' Générer un EndToEndId unique
    newId = GetUniqueEndToEndId()
    
    ' Insérer une nouvelle ligne
    ws.Rows(lastRow + 1).Insert Shift:=xlDown
    
    ' Remplir avec des valeurs par défaut
    ws.Cells(lastRow + 1, TX_COL_ENDTOENDID).Value = newId
    ws.Cells(lastRow + 1, TX_COL_AMOUNT).Value = "0.00"
    ws.Cells(lastRow + 1, TX_COL_EXECUTION_DATE).Value = ws.Range(RANGE_REQUESTED_EXECUTION_DATE).Value
    
    ' Sélectionner la nouvelle ligne
    ws.Rows(lastRow + 1).Select
    
    MsgBox "Nouvelle transaction ajoutée." & vbCrLf & _
           "Veuillez remplir les informations du bénéficiaire.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    LogError "modUI.AddTransaction_Action", Err.Number, Err.Description
End Sub

' ====================================================================================================
' FUNCTION: SelectFileDialog
' DESCRIPTION: Ouvre une boîte de dialogue de sélection de fichier
' PARAMETERS: filter, initialFileName
' RETURNS: String - Chemin du fichier sélectionné
' ====================================================================================================
Public Function SelectFileDialog(ByVal filter As String, Optional ByVal initialFileName As String = "") As String
    On Error Resume Next
    
    Dim fd As Object
    Dim retVal As Integer
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Sélectionner un fichier"
        .InitialFileName = initialFileName
        .Filters.Clear
        
        ' Parser le filter
        Dim parts() As String
        parts = Split(filter, "|")
        
        Dim i As Long
        For i = 0 To UBound(parts) Step 2
            If i + 1 <= UBound(parts) Then
                .Filters.Add parts(i), parts(i + 1)
            End If
        Next i
        
        .AllowMultiSelect = False
        
        retVal = .Show
        
        If retVal = -1 Then
            SelectFileDialog = .SelectedItems(1)
        Else
            SelectFileDialog = ""
        End If
    End With
    
    Set fd = Nothing
End Function

' ====================================================================================================
' FUNCTION: SaveFileDialog
' DESCRIPTION: Ouvre une boîte de dialogue d'enregistrement de fichier
' PARAMETERS: filter, initialFileName
' RETURNS: String - Chemin du fichier à enregistrer
' ====================================================================================================
Public Function SaveFileDialog(ByVal filter As String, Optional ByVal initialFileName As String = "") As String
    On Error Resume Next
    
    Dim fd As Object
    Dim retVal As Integer
    
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    
    With fd
        .Title = "Enregistrer sous"
        .InitialFileName = initialFileName
        .Filters.Clear
        
        Dim parts() As String
        parts = Split(filter, "|")
        
        Dim i As Long
        For i = 0 To UBound(parts) Step 2
            If i + 1 <= UBound(parts) Then
                .Filters.Add parts(i), parts(i + 1)
            End If
        Next i
        
        retVal = .Show
        
        If retVal = -1 Then
            SaveFileDialog = .SelectedItems(1)
        Else
            SaveFileDialog = ""
        End If
    End With
    
    Set fd = Nothing
End Function

' ====================================================================================================
' SUB: RefreshDisplay
' DESCRIPTION: Rafraîchit l'affichage Excel
' ====================================================================================================
Public Sub RefreshDisplay()
    On Error Resume Next
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_PRINCIPAL)
    
    If Not ws Is Nothing Then
        ws.Calculate
        ws.Columns.AutoFit
    End If
    
    Application.ScreenUpdating = True
End Sub

' ====================================================================================================
' SUB: SetupWorksheet
' DESCRIPTION: Configure la feuille Principal avec formatage et protection
' ====================================================================================================
Public Sub SetupWorksheet()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim exists As Boolean
    
    ' Vérifier si la feuille existe
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_PRINCIPAL)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_PRINCIPAL
    End If
    
    ' Effacer et configurer
    ws.Cells.Clear
    
    ' Créer les en-têtes
    Call CreateExcelHeaders(ws)
    
    ' Formater les colonnes
    ws.Columns("A:A").ColumnWidth = 25
    ws.Columns("B:B").ColumnWidth = 35
    ws.Columns("C:C").ColumnWidth = 20
    ws.Columns("D:D").ColumnWidth = 30
    ws.Columns("E:E").ColumnWidth = 30
    ws.Columns("F:F").ColumnWidth = 15
    ws.Columns("G:G").ColumnWidth = 15
    ws.Columns("H:H").ColumnWidth = 40
    ws.Columns("I:I").ColumnWidth = 15
    
    ' Format des montants
    ws.Columns("G").NumberFormat = "0.00"
    
    ' Format des dates
    ws.Columns("I").NumberFormat = "yyyy-mm-dd"
    
    MsgBox "Feuille Principal configurée avec succès!", vbInformation
    
    Exit Sub
    
ErrorHandler:
    LogError "modUI.SetupWorksheet", Err.Number, Err.Description
End Sub

' ====================================================================================================
' SUB: ShowHelp
' DESCRIPTION: Affiche l'aide
' ====================================================================================================
Public Sub ShowHelp()
    Dim helpText As String
    
    helpText = "=== SEPA XML Manager - Aide ===" & vbCrLf & vbCrLf & _
               "1. IMPORTER UN FICHIER XML:" & vbCrLf & _
               "   - Cliquez sur 'Importer XML'" & vbCrLf & _
               "   - Sélectionnez votre fichier XML SEPA" & vbCrLf & _
               "   - Les données sont chargées dans Excel" & vbCrLf & vbCrLf & _
               "2. MODIFIER LES DONNÉES:" & vbCrLf & _
               "   - Modifiez les cellules directement" & vbCrLf & _
               "   - Ajoutez/supprimez des transactions" & vbCrLf & vbCrLf & _
               "3. VALIDER:" & vbCrLf & _
               "   - Vérifie IBAN, BIC, montants" & vbCrLf & _
               "   - Détecte les doublons" & vbCrLf & vbCrLf & _
               "4. EXPORTER:" & vbCrLf & _
               "   - Génère un XML conforme au modèle" & vbCrLf & _
               "   - Conserve namespaces et structure" & vbCrLf & vbCrLf & _
               "FORMATS SUPPORTÉS:" & vbCrLf & _
               "- pain.001.001.03 (virement SEPA)" & vbCrLf & _
               "- ISO 20022 XML" & vbCrLf & vbCrLf & _
               "CONTACT: Support technique"
    
    MsgBox helpText, vbInformation, "Aide SEPA XML Manager"
End Sub
