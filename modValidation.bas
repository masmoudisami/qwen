' ====================================================================================================
' MODULE: modValidation
' DESCRIPTION: Validation des données avant export XML
' ====================================================================================================
Option Explicit

' ====================================================================================================
' VARIABLES GLOBALES POUR LES ERREURS
' ====================================================================================================
Public gErrors As Collection
Public gErrorCount As Long
Public gWarningCount As Long

' ====================================================================================================
' SUB: InitializeValidation
' DESCRIPTION: Initialise la collection d'erreurs
' ====================================================================================================
Public Sub InitializeValidation()
    On Error Resume Next
    
    Set gErrors = New Collection
    gErrorCount = 0
    gWarningCount = 0
End Sub

' ====================================================================================================
' FUNCTION: ValidateAll
' DESCRIPTION: Exécute toutes les validations
' PARAMETERS: ws - Feuille Excel à valider
' RETURNS: Boolean - True si aucune erreur critique
' ====================================================================================================
Public Function ValidateAll(ByVal ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    Dim isValid As Boolean
    
    ' Initialiser
    Call InitializeValidation
    Call ClearErrorSheet
    
    isValid = True
    
    ' Valider l'en-tête
    If Not ValidateHeader(ws) Then
        isValid = False
    End If
    
    ' Valider les transactions
    If Not ValidateTransactions(ws) Then
        isValid = False
    End If
    
    ' Valider la structure XML
    If Not ValidateXMLStructure(ws) Then
        isValid = False
    End If
    
    ' Afficher les erreurs dans la feuille Erreurs
    Call DisplayErrors
    
    ' Résultat
    If gErrorCount > 0 Then
        MsgBox "Validation échouée!" & vbCrLf & _
               "Erreurs critiques: " & gErrorCount & vbCrLf & _
               "Avertissements: " & gWarningCount, vbCritical
        ValidateAll = False
    Else
        If gWarningCount > 0 Then
            MsgBox "Validation réussie avec avertissements." & vbCrLf & _
                   "Avertissements: " & gWarningCount, vbExclamation
        Else
            MsgBox "Validation réussie! Aucune erreur.", vbInformation
        End If
        ValidateAll = True
    End If
    
    Exit Function
    
ErrorHandler:
    ValidateAll = False
    LogError "modValidation.ValidateAll", Err.Number, Err.Description
End Function

' ====================================================================================================
' FUNCTION: ValidateHeader
' DESCRIPTION: Valide les données de l'en-tête
' PARAMETERS: ws - Feuille Excel
' RETURNS: Boolean
' ====================================================================================================
Public Function ValidateHeader(ByVal ws As Worksheet) As Boolean
    On Error Resume Next
    
    Dim isValid As Boolean
    Dim value As String
    
    isValid = True
    
    ' MsgId (obligatoire ou généré)
    value = Trim(CStr(ws.Range(RANGE_MSGID).Value))
    If value = "" Then
        ' Sera généré automatiquement - warning seulement
        Call AddError("INFO", "MsgId sera généré automatiquement", "", RANGE_MSGID, EL_INFO)
    End If
    
    ' CreDtTm (obligatoire ou généré)
    value = Trim(CStr(ws.Range(RANGE_CREDTTM).Value))
    If value = "" Then
        Call AddError("INFO", "CreDtTm sera généré automatiquement", "", RANGE_CREDTTM, EL_INFO)
    End If
    
    ' InitiatingParty Name (obligatoire)
    value = Trim(CStr(ws.Range(RANGE_INITIATING_PARTY_NAME).Value))
    If value = "" Then
        Call AddError(ERR_MANDATORY_MISSING, "InitiatingParty Name est obligatoire", "", RANGE_INITIATING_PARTY_NAME, EL_CRITICAL)
        isValid = False
        gErrorCount = gErrorCount + 1
    End If
    
    ' Debtor Name (obligatoire)
    value = Trim(CStr(ws.Range(RANGE_DEBTOR_NAME).Value))
    If value = "" Then
        Call AddError(ERR_MANDATORY_MISSING, "Debtor Name est obligatoire", "", RANGE_DEBTOR_NAME, EL_CRITICAL)
        isValid = False
        gErrorCount = gErrorCount + 1
    End If
    
    ' Debtor IBAN (obligatoire et doit être valide)
    value = Trim(CStr(ws.Range(RANGE_DEBTOR_IBAN).Value))
    If value = "" Then
        Call AddError(ERR_MANDATORY_MISSING, "Debtor IBAN est obligatoire", "", RANGE_DEBTOR_IBAN, EL_CRITICAL)
        isValid = False
        gErrorCount = gErrorCount + 1
    ElseIf Not ValidateIBAN(value) Then
        Call AddError(ERR_IBAN_INVALID, "Debtor IBAN invalide: " & value, "", RANGE_DEBTOR_IBAN, EL_CRITICAL)
        isValid = False
        gErrorCount = gErrorCount + 1
    End If
    
    ' Debtor BIC (obligatoire et doit être valide)
    value = Trim(CStr(ws.Range(RANGE_DEBTOR_BIC).Value))
    If value = "" Then
        Call AddError(ERR_MANDATORY_MISSING, "Debtor BIC est obligatoire", "", RANGE_DEBTOR_BIC, EL_CRITICAL)
        isValid = False
        gErrorCount = gErrorCount + 1
    ElseIf Not ValidateBIC(value) Then
        Call AddError(ERR_BIC_INVALID, "Debtor BIC invalide: " & value, "", RANGE_DEBTOR_BIC, EL_CRITICAL)
        isValid = False
        gErrorCount = gErrorCount + 1
    End If
    
    ' RequestedExecutionDate
    value = Trim(CStr(ws.Range(RANGE_REQUESTED_EXECUTION_DATE).Value))
    If value = "" Then
        Call AddError("INFO", "RequestedExecutionDate sera défini à demain", "", RANGE_REQUESTED_EXECUTION_DATE, EL_INFO)
    End If
    
    ValidateHeader = isValid
End Function

' ====================================================================================================
' FUNCTION: ValidateTransactions
' DESCRIPTION: Valide toutes les transactions
' PARAMETERS: ws - Feuille Excel
' RETURNS: Boolean
' ====================================================================================================
Public Function ValidateTransactions(ByVal ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    Dim i As Long
    Dim isValid As Boolean
    Dim hasValidTransaction As Boolean
    
    isValid = True
    hasValidTransaction = False
    lastRow = ws.Cells(ws.Rows.Count, TX_COL_ENDTOENDID).End(xlUp).Row
    
    ' Vérifier qu'il y a au moins une transaction
    If lastRow < TX_START_ROW Then
        Call AddError(ERR_MANDATORY_MISSING, "Aucune transaction trouvée", "", "B" & TX_START_ROW, EL_CRITICAL)
        ValidateTransactions = False
        gErrorCount = gErrorCount + 1
        Exit Function
    End If
    
    ' Collection pour détecter les doublons
    Dim endToEndIds As Collection
    Set endToEndIds = New Collection
    
    For i = TX_START_ROW To lastRow
        ' Vérifier si la ligne contient des données
        If Trim(CStr(ws.Cells(i, TX_COL_ENDTOENDID).Value)) <> "" Then
            
            ' EndToEndId (obligatoire et unique)
            Dim e2eId As String
            e2eId = Trim(CStr(ws.Cells(i, TX_COL_ENDTOENDID).Value))
            
            If e2eId = "" Then
                Call AddError(ERR_MANDATORY_MISSING, "EndToEndId est obligatoire", "Ligne " & i, "B" & i, EL_CRITICAL)
                isValid = False
                gErrorCount = gErrorCount + 1
            Else
                ' Vérifier les doublons
                On Error Resume Next
                endToEndIds.Add e2eId, e2eId
                If Err.Number <> 0 Then
                    Call AddError(ERR_DUPLICATE_FOUND, "EndToEndId en double: " & e2eId, "Ligne " & i, "B" & i, EL_CRITICAL)
                    isValid = False
                    gErrorCount = gErrorCount + 1
                    Err.Clear
                End If
                On Error GoTo 0
            End If
            
            ' Beneficiary Name (obligatoire)
            Dim benName As String
            benName = Trim(CStr(ws.Cells(i, TX_COL_BENEFICIARY_NAME).Value))
            If benName = "" Then
                Call AddError(ERR_MANDATORY_MISSING, "Beneficiary Name est obligatoire", "Ligne " & i, "D" & i, EL_CRITICAL)
                isValid = False
                gErrorCount = gErrorCount + 1
            End If
            
            ' Beneficiary IBAN (obligatoire et valide)
            Dim benIban As String
            benIban = Trim(CStr(ws.Cells(i, TX_COL_BENEFICIARY_IBAN).Value))
            If benIban = "" Then
                Call AddError(ERR_MANDATORY_MISSING, "Beneficiary IBAN est obligatoire", "Ligne " & i, "E" & i, EL_CRITICAL)
                isValid = False
                gErrorCount = gErrorCount + 1
            ElseIf Not ValidateIBAN(benIban) Then
                Call AddError(ERR_IBAN_INVALID, "Beneficiary IBAN invalide: " & benIban, "Ligne " & i, "E" & i, EL_CRITICAL)
                isValid = False
                gErrorCount = gErrorCount + 1
            End If
            
            ' Beneficiary BIC (obligatoire et valide)
            Dim benBic As String
            benBic = Trim(CStr(ws.Cells(i, TX_COL_BENEFICIARY_BIC).Value))
            If benBic = "" Then
                Call AddError(ERR_MANDATORY_MISSING, "Beneficiary BIC est obligatoire", "Ligne " & i, "F" & i, EL_CRITICAL)
                isValid = False
                gErrorCount = gErrorCount + 1
            ElseIf Not ValidateBIC(benBic) Then
                Call AddError(ERR_BIC_INVALID, "Beneficiary BIC invalide: " & benBic, "Ligne " & i, "F" & i, EL_CRITICAL)
                isValid = False
                gErrorCount = gErrorCount + 1
            End If
            
            ' Amount (obligatoire et positif)
            Dim amount As Variant
            amount = ws.Cells(i, TX_COL_AMOUNT).Value
            If IsEmpty(amount) Or amount = "" Then
                Call AddError(ERR_MANDATORY_MISSING, "Amount est obligatoire", "Ligne " & i, "G" & i, EL_CRITICAL)
                isValid = False
                gErrorCount = gErrorCount + 1
            ElseIf Not IsNumeric(amount) Then
                Call AddError(ERR_AMOUNT_INVALID, "Amount doit être numérique: " & CStr(amount), "Ligne " & i, "G" & i, EL_CRITICAL)
                isValid = False
                gErrorCount = gErrorCount + 1
            ElseIf CDbl(amount) <= 0 Then
                Call AddError(ERR_AMOUNT_INVALID, "Amount doit être positif: " & CStr(amount), "Ligne " & i, "G" & i, EL_CRITICAL)
                isValid = False
                gErrorCount = gErrorCount + 1
            ElseIf CDbl(amount) > 999999999.99 Then
                Call AddError(ERR_AMOUNT_INVALID, "Amount trop élevé (max 999999999.99): " & CStr(amount), "Ligne " & i, "G" & i, EL_WARNING)
                gWarningCount = gWarningCount + 1
            End If
            
            hasValidTransaction = True
        End If
    Next i
    
    If Not hasValidTransaction Then
        Call AddError(ERR_MANDATORY_MISSING, "Aucune transaction valide trouvée", "", "B" & TX_START_ROW, EL_CRITICAL)
        isValid = False
        gErrorCount = gErrorCount + 1
    End If
    
    ValidateTransactions = isValid
    
    Exit Function
    
ErrorHandler:
    ValidateTransactions = False
    LogError "modValidation.ValidateTransactions", Err.Number, Err.Description
End Function

' ====================================================================================================
' FUNCTION: ValidateXMLStructure
' DESCRIPTION: Valide la structure XML globale
' PARAMETERS: ws - Feuille Excel
' RETURNS: Boolean
' ====================================================================================================
Public Function ValidateXMLStructure(ByVal ws As Worksheet) As Boolean
    On Error Resume Next
    
    Dim isValid As Boolean
    
    isValid = True
    
    ' Vérifier le namespace
    If gNamespaceUri = "" Then
        Call AddError(ERR_XML_NAMESPACE, "Namespace URI non défini. Importez d'abord un fichier XML modèle.", "", "A1", EL_CRITICAL)
        isValid = False
        gErrorCount = gErrorCount + 1
    End If
    
    ' Vérifier la cohérence NbOfTxs
    Dim expectedTxCount As Long
    expectedTxCount = CountNonEmptyTransactions(ws)
    
    ' Vérifier la cohérence CtrlSum
    Dim expectedCtrlSum As Double
    expectedCtrlSum = CalculateCtrlSumFromExcel(ws)
    
    ' Ces valeurs seront recalculées automatiquement lors de l'export
    Call AddError("INFO", "NbOfTxs calculé: " & CStr(expectedTxCount), "", RANGE_NBFTXS, EL_INFO)
    Call AddError("INFO", "CtrlSum calculé: " & FormatAmount(expectedCtrlSum), "", RANGE_CTRLSUM, EL_INFO)
    
    ValidateXMLStructure = isValid
End Function

' ====================================================================================================
' SUB: AddError
' DESCRIPTION: Ajoute une erreur à la collection
' PARAMETERS: errorType, description, transactionRef, cellRef, level
' ====================================================================================================
Public Sub AddError(ByVal errorType As String, _
                    ByVal description As String, _
                    ByVal transactionRef As String, _
                    ByVal cellRef As String, _
                    ByVal level As ErrorLevel)
    On Error Resume Next
    
    Dim err As ValidationError
    
    err.ErrorType = errorType
    err.Description = description
    err.TransactionRef = transactionRef
    err.CellReference = cellRef
    err.Level = level
    
    gErrors.Add err
End Sub

' ====================================================================================================
' SUB: ClearErrorSheet
' DESCRIPTION: Efface la feuille d'erreurs
' ====================================================================================================
Public Sub ClearErrorSheet()
    On Error Resume Next
    
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(SHEET_ERREURS)
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_ERREURS
        
        ' Créer les en-têtes
        ws.Range("A1").Value = "Type"
        ws.Range("B1").Value = "Description"
        ws.Range("C1").Value = "Transaction"
        ws.Range("D1").Value = "Cellule"
        ws.Range("E1").Value = "Niveau"
        
        ws.Range("A1:E1").Font.Bold = True
        ws.Range("A1:E1").Interior.Color = RGB(255, 200, 200)
    Else
        ws.Range("A2:E" & ws.Rows.Count).ClearContents
    End If
End Sub

' ====================================================================================================
' SUB: DisplayErrors
' DESCRIPTION: Affiche les erreurs dans la feuille Erreurs
' ====================================================================================================
Public Sub DisplayErrors()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim err As ValidationError
    Dim row As Long
    
    Set ws = ThisWorkbook.Worksheets(SHEET_ERREURS)
    
    If ws Is Nothing Then
        Exit Sub
    End If
    
    row = 2
    
    For Each err In gErrors
        ws.Cells(row, ERR_COL_TYPE).Value = err.ErrorType
        ws.Cells(row, ERR_COL_DESCRIPTION).Value = err.Description
        ws.Cells(row, ERR_COL_TRANSACTION).Value = err.TransactionRef
        ws.Cells(row, ERR_COL_CELLULE).Value = err.CellReference
        
        Select Case err.Level
            Case EL_CRITICAL
                ws.Cells(row, ERR_COL_NIVEAU).Value = "CRITIQUE"
                ws.Cells(row, ERR_COL_NIVEAU).Interior.Color = RGB(255, 150, 150)
            Case EL_WARNING
                ws.Cells(row, ERR_COL_NIVEAU).Value = "AVERTISSEMENT"
                ws.Cells(row, ERR_COL_NIVEAU).Interior.Color = RGB(255, 255, 150)
            Case EL_INFO
                ws.Cells(row, ERR_COL_NIVEAU).Value = "INFO"
                ws.Cells(row, ERR_COL_NIVEAU).Interior.Color = RGB(200, 255, 200)
        End Select
        
        row = row + 1
    Next err
    
    ws.Columns.AutoFit
End Sub

' ====================================================================================================
' FUNCTION: ValidateSingleIBAN
' DESCRIPTION: Valide un IBAN individuel
' PARAMETERS: iban
' RETURNS: Boolean
' ====================================================================================================
Public Function ValidateSingleIBAN(ByVal iban As String) As Boolean
    On Error Resume Next
    ValidateSingleIBAN = ValidateIBAN(iban)
End Function

' ====================================================================================================
' FUNCTION: ValidateSingleBIC
' DESCRIPTION: Valide un BIC individuel
' PARAMETERS: bic
' RETURNS: Boolean
' ====================================================================================================
Public Function ValidateSingleBIC(ByVal bic As String) As Boolean
    On Error Resume Next
    ValidateSingleBIC = ValidateBIC(bic)
End Function

' ====================================================================================================
' FUNCTION: CheckDuplicateEndToEndId
' DESCRIPTION: Vérifie les doublons d'EndToEndId
' PARAMETERS: ws, newId, excludeRow
' RETURNS: Boolean - True si doublon trouvé
' ====================================================================================================
Public Function CheckDuplicateEndToEndId(ByVal ws As Worksheet, _
                                         ByVal newId As String, _
                                         Optional ByVal excludeRow As Long = 0) As Boolean
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    Dim i As Long
    Dim existingId As String
    
    CheckDuplicateEndToEndId = False
    lastRow = ws.Cells(ws.Rows.Count, TX_COL_ENDTOENDID).End(xlUp).Row
    
    For i = TX_START_ROW To lastRow
        If i <> excludeRow Then
            existingId = Trim(CStr(ws.Cells(i, TX_COL_ENDTOENDID).Value))
            If existingId = newId And existingId <> "" Then
                CheckDuplicateEndToEndId = True
                Exit Function
            End If
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    CheckDuplicateEndToEndId = False
    LogError "modValidation.CheckDuplicateEndToEndId", Err.Number, Err.Description
End Function
