'===============================================================================
' MODULE: modMain.bas
' DESCRIPTION: Point d'entrée principal et procédures macro pour le gestionnaire SEPA XML
'===============================================================================
Option Explicit

' =============================================================================
' SUB: Main_ImportXML
' DESCRIPTION: Macro principale pour l'import XML (appelée depuis les boutons)
' =============================================================================
Public Sub Main_ImportXML()
    On Error Resume Next
    Application.ScreenUpdating = False
    
    Call ShowImportDialog
    
    Application.ScreenUpdating = True
End Sub

' =============================================================================
' SUB: Main_ExportXML
' DESCRIPTION: Macro principale pour l'export XML (appelée depuis les boutons)
' =============================================================================
Public Sub Main_ExportXML()
    On Error Resume Next
    Application.ScreenUpdating = False
    
    Call ShowExportDialog
    
    Application.ScreenUpdating = True
End Sub

' =============================================================================
' SUB: Main_ValidateData
' DESCRIPTION: Macro principale pour la validation (appelée depuis les boutons)
' =============================================================================
Public Sub Main_ValidateData()
    On Error Resume Next
    Application.ScreenUpdating = False
    
    Call ShowValidationResults
    
    Application.ScreenUpdating = True
End Sub

' =============================================================================
' SUB: Main_RecalculateTotals
' DESCRIPTION: Macro principale pour recalculer les totaux
' =============================================================================
Public Sub Main_RecalculateTotals()
    On Error Resume Next
    Application.ScreenUpdating = False
    
    Call RecalculateTotals
    
    MsgBox "Totaux recalculés avec succès!", vbInformation, "SEPA XML Manager"
    
    Application.ScreenUpdating = True
End Sub

' =============================================================================
' SUB: Main_ClearErrors
' DESCRIPTION: Efface tous les logs d'erreurs
' =============================================================================
Public Sub Main_ClearErrors()
    On Error Resume Next
    
    If MsgBox("Voulez-vous vraiment effacer tous les logs d'erreurs?", _
              vbYesNo + vbQuestion, "Confirmation") = vbYes Then
        
        ClearErrorLog
        MsgBox "Logs d'erreurs effacés.", vbInformation, "SEPA XML Manager"
    End If
End Sub

' =============================================================================
' SUB: Main_ShowErrorSummary
' DESCRIPTION: Affiche le résumé des erreurs
' =============================================================================
Public Sub Main_ShowErrorSummary()
    On Error Resume Next
    Call ShowErrorSummary
End Sub

' =============================================================================
' SUB: Main_AddTransaction
' DESCRIPTION: Ajoute une nouvelle transaction
' =============================================================================
Public Sub Main_AddTransaction()
    On Error Resume Next
    Application.ScreenUpdating = False
    
    Call AddTransactionRow
    
    Application.ScreenUpdating = True
End Sub

' =============================================================================
' SUB: Main_DeleteTransaction
' DESCRIPTION: Supprime une transaction
' =============================================================================
Public Sub Main_DeleteTransaction()
    On Error Resume Next
    Application.ScreenUpdating = False
    
    Call DeleteSelectedTransaction
    
    Application.ScreenUpdating = True
End Sub

' =============================================================================
' SUB: Main_HighlightErrors
' DESCRIPTION: Met en surbrillance les erreurs
' =============================================================================
Public Sub Main_HighlightErrors()
    On Error Resume Next
    Application.ScreenUpdating = False
    
    Call HighlightErrors
    
    Application.ScreenUpdating = True
End Sub

' =============================================================================
' SUB: Workbook_Open
' DESCRIPTION: S'exécute à l'ouverture du classeur (à placer dans ThisWorkbook)
' =============================================================================
' NOTE: Cette procédure doit être copiée dans le module ThisWorkbook
'
' Private Sub Workbook_Open()
'     On Error Resume Next
'     
'     ' Initialiser le système de logging
'     InitializeErrorLogging
'     
'     ' Créer la feuille d'erreurs si nécessaire
'     Dim ws As Worksheet
'     Set ws = GetWorksheetByName(SHEET_ERREURS)
'     If Not ws Is Nothing Then
'         SetupErrorSheet ws
'     End If
'     
'     ' Afficher un message de bienvenue
'     MsgBox "Bienvenue dans SEPA XML Manager!" & vbCrLf & _
'            "Format supporté: pain.001.001.03" & vbCrLf & vbCrLf & _
'            "Cliquez sur 'Importer XML' pour commencer.", _
'            vbInformation, "SEPA XML Manager"
' End Sub

' =============================================================================
' SUB: InitializeApplication
' DESCRIPTION: Initialise l'application (à exécuter au démarrage)
' =============================================================================
Public Sub InitializeApplication()
    On Error Resume Next
    
    Application.ScreenUpdating = False
    
    ' Initialiser le logging
    InitializeErrorLogging
    
    ' Créer les feuilles nécessaires
    Dim ws As Worksheet
    
    Set ws = GetWorksheetByName(SHEET_ERREURS)
    If Not ws Is Nothing Then
        SetupErrorSheet ws
    End If
    
    Set ws = GetWorksheetByName(SHEET_PRINCIPAL)
    If Not ws Is Nothing Then
        SetupPrincipalSheet
    End If
    
    ' Configurer l'interface
    Call SetupQuickAccessToolbar
    
    Application.ScreenUpdating = True
    
    LogInfo "Application initialisée", "modMain.InitializeApplication"
End Sub

' =============================================================================
' SUB: TestImportFromSample
' DESCRIPTION: Importe le fichier modèle xml-Ok.xml pour test
' =============================================================================
Public Sub TestImportFromSample()
    On Error Resume Next
    
    Dim FilePath As String
    Dim TestPath As String
    
    ' Chemins possibles pour le fichier de test
    TestPath = ThisWorkbook.Path & "\xml-Ok.xml"
    
    If Dir(TestPath) = "" Then
        TestPath = "/workspace/xml-Ok.xml"
    End If
    
    If Dir(TestPath) = "" Then
        ' Demander à l'utilisateur
        FilePath = Application.GetOpenFilename( _
            FileFilter:="XML Files (*.xml), *.xml", _
            Title:="Sélectionner le fichier xml-Ok.xml")
        
        If FilePath = "False" Then Exit Sub
    Else
        FilePath = TestPath
    End If
    
    ' Importer
    If ImportXMLFile(FilePath) Then
        MsgBox "Test réussi! Le fichier modèle a été importé." & vbCrLf & _
               "Vous pouvez maintenant modifier les données et exporter.", _
               vbInformation, "SEPA XML Manager - Test"
        
        Call CreateMainMenu
        Call RecalculateTotals
    Else
        MsgBox "Échec du test d'import.", vbCritical, "SEPA XML Manager - Erreur"
    End If
End Sub

' =============================================================================
' SUB: GenerateSampleData
' DESCRIPTION: Génère des données de test pour démonstration
' =============================================================================
Public Sub GenerateSampleData()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = SetupPrincipalSheet()
    
    ' Remplir l'en-tête avec des données de démo
    ws.Range(CELL_MSGID).Value = "DEMO_" & Format(Now, "yyyymmddhhmmss")
    ws.Range(CELL_CREDTTM).Value = Now
    ws.Range(CELL_NBOFTXS).Value = 2
    ws.Range(CELL_CTRLSUM).Value = 452.5
    
    ' Initiating Party
    ws.Range(CELL_INITGPTY_NM).Value = "SOCIETE DEMO"
    ws.Range(CELL_INITGPTY_STRTNM).Value = "Avenue Principale"
    ws.Range(CELL_INITGPTY_PSTCD).Value = "1000"
    ws.Range(CELL_INITGPTY_TWNNM).Value = "Tunis"
    ws.Range(CELL_INITGPTY_CTRY).Value = "TN"
    
    ' Debtor
    ws.Range(CELL_DBTR_NM).Value = "SOCIETE DEMO"
    ws.Range(CELL_DBTR_STRTNM).Value = "Avenue Principale"
    ws.Range(CELL_DBTR_PSTCD).Value = "1000"
    ws.Range(CELL_DBTR_TWNNM).Value = "Tunis"
    ws.Range(CELL_DBTR_CTRY).Value = "TN"
    ws.Range(CELL_DBTR_IBAN).Value = "TN5912003000330070607942"
    ws.Range(CELL_DBTR_BIC).Value = "UBCITNTTXXX"
    
    ' Payment Information
    ws.Range(CELL_PMTINFID).Value = "PAIEMENT_DEMO"
    ws.Range(CELL_PMTMTD).Value = "TRF"
    ws.Range(CELL_BTCHBOOKG).Value = XML_FALSE
    ws.Range(CELL_REQDEXCTNDT).Value = Date + 7
    ws.Range(CELL_CHRGBR).Value = CHRG_BR_DEBT
    ws.Range(CELL_DEVISE).Value = "TND"
    
    ' Transactions de démo
    Dim TxData As Variant
    TxData = Array( _
        Array("TX001", "EMPLOYE UN", "RUE DE LA PAIX", "1001", "Tunis", "TN", _
              "TN5904202061001400054173", "BSTUTNTTXXX", "TIJARI BANK", 301.5, "TND", "SALAIRE"), _
        Array("TX002", "EMPLOYE DEUX", "AVENUE HABIB", "1002", "Sousse", "TN", _
              "TN5901028070111006804463", "ATBKTNTTXXX", "BIAT", 151.0, "TND", "SALAIRE") _
    )
    
    For i = 0 To UBound(TxData)
        With ws
            .Cells(ROW_TRANSACTION_START + i, COL_TX_ENDTOENDID).Value = TxData(i)(0)
            .Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_NM).Value = TxData(i)(1)
            .Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_STRTNM).Value = TxData(i)(2)
            .Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_PSTCD).Value = TxData(i)(3)
            .Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_TWNNM).Value = TxData(i)(4)
            .Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_CTRY).Value = TxData(i)(5)
            .Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_IBAN).Value = TxData(i)(6)
            .Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_BIC).Value = TxData(i)(7)
            .Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_BANKNM).Value = TxData(i)(8)
            .Cells(ROW_TRANSACTION_START + i, COL_TX_AMT).Value = TxData(i)(9)
            .Cells(ROW_TRANSACTION_START + i, COL_TX_CCY).Value = TxData(i)(10)
            .Cells(ROW_TRANSACTION_START + i, COL_TX_USTRD).Value = TxData(i)(11)
        End With
    Next i
    
    ' Créer le menu
    Call CreateMainMenu
    Call SetupQuickAccessToolbar
    
    MsgBox "Données de démonstration générées!" & vbCrLf & _
           "Vous pouvez modifier les valeurs et exporter vers XML.", _
           vbInformation, "SEPA XML Manager"
    
    LogInfo "Données de démonstration générées", "modMain.GenerateSampleData"
End Sub

' =============================================================================
' SUB: Help_About
' DESCRIPTION: Affiche les informations sur l'application
' =============================================================================
Public Sub Help_About()
    Dim Msg As String
    
    Msg = "SEPA XML Manager" & vbCrLf
    Msg = Msg & String(50, "=") & vbCrLf & vbCrLf
    Msg = Msg & "Version: 1.0" & vbCrLf
    Msg = Msg & "Format: ISO 20022 pain.001.001.03" & vbCrLf
    Msg = Msg & "Date: " & Format(Date, "dd/mm/yyyy") & vbCrLf & vbCrLf
    Msg = Msg & "Fonctionnalités:" & vbCrLf
    Msg = Msg & "  • Import de fichiers XML SEPA" & vbCrLf
    Msg = Msg & "  • Visualisation et édition dans Excel" & vbCrLf
    Msg = Msg & "  • Validation IBAN/BIC (MOD-97)" & vbCrLf
    Msg = Msg & "  • Détection de doublons" & vbCrLf
    Msg = Msg & "  • Export XML conforme au modèle" & vbCrLf
    Msg = Msg & "  • Export TXT UTF-8" & vbCrLf & vbCrLf
    Msg = Msg & "Développé pour la compatibilité bancaire SEPA."
    
    MsgBox Msg, vbInformation + vbOKOnly, "À propos - SEPA XML Manager"
End Sub
