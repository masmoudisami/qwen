'===============================================================================
' MODULE: modValidation.bas
' DESCRIPTION: Validation des données pour le gestionnaire SEPA XML
'===============================================================================
Option Explicit

' =============================================================================
' FONCTION: ValidateAllData
' DESCRIPTION: Valide toutes les données avant export
' =============================================================================
Public Function ValidateAllData() As Boolean
    On Error Resume Next
    
    Dim IsValid As Boolean
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(SHEET_PRINCIPAL)
    If ws Is Nothing Then
        LogError ERR_VALIDATION_MANDATORY, "Feuille Principal introuvable", _
                 "modValidation.ValidateAllData", , , True
        ValidateAllData = False
        Exit Function
    End If
    
    IsValid = True
    
    ' Valider l'en-tête
    If Not ValidateHeader(ws) Then
        IsValid = False
    End If
    
    ' Valider les transactions
    If Not ValidateTransactions(ws) Then
        IsValid = False
    End If
    
    ' Vérifier les doublons
    If Not CheckDuplicates(ws) Then
        IsValid = False
    End If
    
    ValidateAllData = IsValid
    
    If IsValid Then
        LogInfo "Toutes les validations ont réussi", "modValidation.ValidateAllData"
    Else
        LogWarning "Certaines validations ont échoué - voir feuille Erreurs", "modValidation.ValidateAllData"
    End If
    
    On Error GoTo 0
End Function

' =============================================================================
' FONCTION: ValidateHeader
' DESCRIPTION: Valide les données de l'en-tête
' =============================================================================
Private Function ValidateHeader(ByRef ws As Worksheet) As Boolean
    On Error Resume Next
    
    Dim IsValid As Boolean
    Dim HasError As Boolean
    
    IsValid = True
    HasError = False
    
    ' MsgId (obligatoire)
    If Len(Trim(SafeCStr(ws.Range(CELL_MSGID).Value))) = 0 Then
        LogError ERR_VALIDATION_MANDATORY, "MsgId est obligatoire", _
                 "modValidation.ValidateHeader", , CELL_MSGID, False
        HasError = True
    End If
    
    ' CreDtTm (obligatoire)
    If IsEmpty(ws.Range(CELL_CREDTTM).Value) Then
        LogError ERR_VALIDATION_MANDATORY, "CreDtTm est obligatoire", _
                 "modValidation.ValidateHeader", , CELL_CREDTTM, False
        HasError = True
    End If
    
    ' Initiating Party - Nom (obligatoire)
    If Len(Trim(SafeCStr(ws.Range(CELL_INITGPTY_NM).Value))) = 0 Then
        LogError ERR_VALIDATION_MANDATORY, "Nom de l'Initiating Party est obligatoire", _
                 "modValidation.ValidateHeader", , CELL_INITGPTY_NM, False
        HasError = True
    End If
    
    ' Debtor - IBAN (obligatoire et à valider)
    Dim Iban As String
    Iban = SafeCStr(ws.Range(CELL_DBTR_IBAN).Value)
    
    If Len(Trim(Iban)) = 0 Then
        LogError ERR_VALIDATION_MANDATORY, "IBAN du débiteur est obligatoire", _
                 "modValidation.ValidateHeader", , CELL_DBTR_IBAN, False
        HasError = True
    ElseIf Not ValidateIBAN(Iban) Then
        LogError ERR_VALIDATION_IBAN, "IBAN du débiteur invalide: " & Iban, _
                 "modValidation.ValidateHeader", , CELL_DBTR_IBAN, False
        HasError = True
    End If
    
    ' Debtor - BIC (obligatoire et à valider)
    Dim Bic As String
    Bic = SafeCStr(ws.Range(CELL_DBTR_BIC).Value)
    
    If Len(Trim(Bic)) = 0 Then
        LogError ERR_VALIDATION_MANDATORY, "BIC du débiteur est obligatoire", _
                 "modValidation.ValidateHeader", , CELL_DBTR_BIC, False
        HasError = True
    ElseIf Not ValidateBIC(Bic) Then
        LogError ERR_VALIDATION_BIC, "BIC du débiteur invalide: " & Bic, _
                 "modValidation.ValidateHeader", , CELL_DBTR_BIC, False
        HasError = True
    End If
    
    ' Payment Information - PmtInfId (obligatoire)
    If Len(Trim(SafeCStr(ws.Range(CELL_PMTINFID).Value))) = 0 Then
        LogError ERR_VALIDATION_MANDATORY, "PmtInfId est obligatoire", _
                 "modValidation.ValidateHeader", , CELL_PMTINFID, False
        HasError = True
    End If
    
    ' ReqdExctnDt (obligatoire)
    If IsEmpty(ws.Range(CELL_REQDEXCTNDT).Value) Then
        LogError ERR_VALIDATION_DATE, "ReqdExctnDt est obligatoire", _
                 "modValidation.ValidateHeader", , CELL_REQDEXCTNDT, False
        HasError = True
    End If
    
    ' ChrgBr (obligatoire)
    Dim ChrgBr As String
    ChrgBr = SafeCStr(ws.Range(CELL_CHRGBR).Value)
    
    If Len(Trim(ChrgBr)) = 0 Then
        LogError ERR_VALIDATION_MANDATORY, "ChrgBr est obligatoire (DEBT, CRED ou SHAR)", _
                 "modValidation.ValidateHeader", , CELL_CHRGBR, False
        HasError = True
    ElseIf ChrgBr <> CHRG_BR_DEBT And ChrgBr <> CHRG_BR_CRED And ChrgBr <> CHRG_BR_SHAR Then
        LogError ERR_VALIDATION_MANDATORY, "ChrgBr doit être DEBT, CRED ou SHAR", _
                 "modValidation.ValidateHeader", , CELL_CHRGBR, False
        HasError = True
    End If
    
    ' Devise (obligatoire)
    If Len(Trim(SafeCStr(ws.Range(CELL_DEVISE).Value))) = 0 Then
        LogError ERR_VALIDATION_MANDATORY, "La devise est obligatoire", _
                 "modValidation.ValidateHeader", , CELL_DEVISE, False
        HasError = True
    End If
    
    If HasError Then
        IsValid = False
    End If
    
    ValidateHeader = IsValid
    
    On Error GoTo 0
End Function

' =============================================================================
' FONCTION: ValidateTransactions
' DESCRIPTION: Valide toutes les transactions
' =============================================================================
Private Function ValidateTransactions(ByRef ws As Worksheet) As Boolean
    On Error Resume Next
    
    Dim IsValid As Boolean
    Dim LastRow As Long
    Dim i As Long
    Dim HasError As Boolean
    
    IsValid = True
    HasError = False
    
    ' Trouver la dernière ligne de transaction
    LastRow = ws.Cells(ws.Rows.Count, COL_TX_ENDTOENDID).End(xlUp).Row
    
    If LastRow < ROW_TRANSACTION_START Then
        LogWarning "Aucune transaction trouvée", "modValidation.ValidateTransactions"
        ValidateTransactions = True  ' Pas d'erreur si aucune transaction
        Exit Function
    End If
    
    ' Valider chaque transaction
    For i = ROW_TRANSACTION_START To LastRow
        If Not ValidateSingleTransaction(ws, i) Then
            HasError = True
        End If
    Next i
    
    If HasError Then
        IsValid = False
    End If
    
    ValidateTransactions = IsValid
    
    On Error GoTo 0
End Function

' =============================================================================
' FONCTION: ValidateSingleTransaction
' DESCRIPTION: Valide une transaction individuelle
' =============================================================================
Private Function ValidateSingleTransaction(ByRef ws As Worksheet, ByVal RowNum As Long) As Boolean
    On Error Resume Next
    
    Dim IsValid As Boolean
    Dim HasError As Boolean
    Dim CellRef As String
    Dim Value As String
    
    IsValid = True
    HasError = False
    
    ' EndToEndId (obligatoire)
    CellRef = ws.Cells(RowNum, COL_TX_ENDTOENDID).Address
    Value = SafeCStr(ws.Cells(RowNum, COL_TX_ENDTOENDID).Value)
    
    If Len(Trim(Value)) = 0 Then
        LogError ERR_VALIDATION_MANDATORY, "EndToEndId est obligatoire pour la ligne " & RowNum, _
                 "modValidation.ValidateSingleTransaction", , CellRef, False
        HasError = True
    End If
    
    ' Cdtr - Nom (obligatoire)
    CellRef = ws.Cells(RowNum, COL_TX_CDTR_NM).Address
    Value = SafeCStr(ws.Cells(RowNum, COL_TX_CDTR_NM).Value)
    
    If Len(Trim(Value)) = 0 Then
        LogError ERR_VALIDATION_MANDATORY, "Nom du bénéficiaire est obligatoire pour la ligne " & RowNum, _
                 "modValidation.ValidateSingleTransaction", , CellRef, False
        HasError = True
    End If
    
    ' Cdtr - IBAN (obligatoire et à valider)
    CellRef = ws.Cells(RowNum, COL_TX_CDTR_IBAN).Address
    Value = SafeCStr(ws.Cells(RowNum, COL_TX_CDTR_IBAN).Value)
    
    If Len(Trim(Value)) = 0 Then
        LogError ERR_VALIDATION_MANDATORY, "IBAN du bénéficiaire est obligatoire pour la ligne " & RowNum, _
                 "modValidation.ValidateSingleTransaction", , CellRef, False
        HasError = True
    ElseIf Not ValidateIBAN(Value) Then
        LogError ERR_VALIDATION_IBAN, "IBAN du bénéficiaire invalide: " & Value, _
                 "modValidation.ValidateSingleTransaction", , CellRef, False
        HasError = True
    End If
    
    ' Cdtr - BIC (obligatoire et à valider)
    CellRef = ws.Cells(RowNum, COL_TX_CDTR_BIC).Address
    Value = SafeCStr(ws.Cells(RowNum, COL_TX_CDTR_BIC).Value)
    
    If Len(Trim(Value)) = 0 Then
        LogError ERR_VALIDATION_MANDATORY, "BIC de la banque bénéficiaire est obligatoire pour la ligne " & RowNum, _
                 "modValidation.ValidateSingleTransaction", , CellRef, False
        HasError = True
    ElseIf Not ValidateBIC(Value) Then
        LogError ERR_VALIDATION_BIC, "BIC de la banque bénéficiaire invalide: " & Value, _
                 "modValidation.ValidateSingleTransaction", , CellRef, False
        HasError = True
    End If
    
    ' Montant (obligatoire et > 0)
    CellRef = ws.Cells(RowNum, COL_TX_AMT).Address
    Dim Amount As Double
    Amount = SafeCDbl(ws.Cells(RowNum, COL_TX_AMT).Value)
    
    If Amount <= 0 Then
        LogError ERR_VALIDATION_AMOUNT, "Le montant doit être > 0 pour la ligne " & RowNum, _
                 "modValidation.ValidateSingleTransaction", , CellRef, False
        HasError = True
    ElseIf Amount > MAX_AMOUNT Then
        LogError ERR_VALIDATION_AMOUNT, "Le montant dépasse le maximum autorisé (" & MAX_AMOUNT & ") pour la ligne " & RowNum, _
                 "modValidation.ValidateSingleTransaction", , CellRef, False
        HasError = True
    End If
    
    ' Devise (obligatoire)
    CellRef = ws.Cells(RowNum, COL_TX_CCY).Address
    Value = SafeCStr(ws.Cells(RowNum, COL_TX_CCY).Value)
    
    If Len(Trim(Value)) = 0 Then
        LogError ERR_VALIDATION_MANDATORY, "La devise est obligatoire pour la ligne " & RowNum, _
                 "modValidation.ValidateSingleTransaction", , CellRef, False
        HasError = True
    ElseIf Len(Value) <> 3 Then
        LogError ERR_VALIDATION_MANDATORY, "La devise doit être un code ISO à 3 lettres pour la ligne " & RowNum, _
                 "modValidation.ValidateSingleTransaction", , CellRef, False
        HasError = True
    End If
    
    ' Pays du bénéficiaire (optionnel mais recommandé)
    CellRef = ws.Cells(RowNum, COL_TX_CDTR_CTRY).Address
    Value = SafeCStr(ws.Cells(RowNum, COL_TX_CDTR_CTRY).Value)
    
    If Len(Trim(Value)) > 0 And Len(Value) <> 2 Then
        LogError ERR_VALIDATION_MANDATORY, "Le code pays doit être à 2 lettres pour la ligne " & RowNum, _
                 "modValidation.ValidateSingleTransaction", , CellRef, False
        HasError = True
    End If
    
    ' Vérifier les caractères spéciaux interdits
    Call ValidateCharacters(ws, RowNum)
    
    If HasError Then
        IsValid = False
    End If
    
    ValidateSingleTransaction = IsValid
    
    On Error GoTo 0
End Function

' =============================================================================
' SUB: ValidateCharacters
' DESCRIPTION: Vérifie les caractères autorisés dans les champs texte
' =============================================================================
Private Sub ValidateCharacters(ByRef ws As Worksheet, ByVal RowNum As Long)
    On Error Resume Next
    
    Dim Col As Long
    Dim Value As String
    Dim i As Integer
    Dim Char As String
    
    ' Colonnes texte à vérifier
    Dim TextCols As Variant
    TextCols = Array(COL_TX_ENDTOENDID, COL_TX_INSTRID, COL_TX_CDTR_NM, _
                     COL_TX_CDTR_STRTNM, COL_TX_CDTR_TWNNM, COL_TX_USTRD)
    
    For Each Col In TextCols
        Value = SafeCStr(ws.Cells(RowNum, Col).Value)
        
        If Len(Value) > 0 Then
            For i = 1 To Len(Value)
                Char = Mid(Value, i, 1)
                ' Ignorer les caractères étendus (accentués, etc.)
                If Asc(Char) < 128 Then
                    If InStr(ALLOWED_CHARS, Char) = 0 Then
                        LogError ERR_VALIDATION_CHARS, "Caractère non autorisé '" & Char & _
                                 "' dans la colonne " & Col & " ligne " & RowNum, _
                                 "modValidation.ValidateCharacters", , _
                                 ws.Cells(RowNum, Col).Address, False
                    End If
                End If
            Next i
        End If
    Next Col
    
    On Error GoTo 0
End Sub

' =============================================================================
' FONCTION: CheckDuplicates
' DESCRIPTION: Vérifie les doublons d'EndToEndId
' =============================================================================
Private Function CheckDuplicates(ByRef ws As Worksheet) As Boolean
    On Error Resume Next
    
    Dim Dict As Object
    Dim LastRow As Long
    Dim i As Long
    Dim EndToEndId As String
    Dim HasDuplicates As Boolean
    
    Set Dict = CreateObject("Scripting.Dictionary")
    HasDuplicates = False
    
    LastRow = ws.Cells(ws.Rows.Count, COL_TX_ENDTOENDID).End(xlUp).Row
    
    If LastRow < ROW_TRANSACTION_START Then
        CheckDuplicates = True
        Exit Function
    End If
    
    For i = ROW_TRANSACTION_START To LastRow
        EndToEndId = Trim(SafeCStr(ws.Cells(i, COL_TX_ENDTOENDID).Value))
        
        If Len(EndToEndId) > 0 Then
            If Dict.Exists(EndToEndId) Then
                LogError ERR_VALIDATION_DUPLICATE, "Doublon d'EndToEndId: " & EndToEndId, _
                         "modValidation.CheckDuplicates", EndToEndId, _
                         ws.Cells(i, COL_TX_ENDTOENDID).Address, False
                HasDuplicates = True
            Else
                Dict.Add EndToEndId, i
            End If
        End If
    Next i
    
    CheckDuplicates = Not HasDuplicates
    
    On Error GoTo 0
End Function

' =============================================================================
' FONCTION: RecalculateTotals
' DESCRIPTION: Recalcule NbOfTxs et CtrlSum depuis Excel
' =============================================================================
Public Sub RecalculateTotals()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim TxCount As Long
    Dim TotalAmount As Double
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets(SHEET_PRINCIPAL)
    If ws Is Nothing Then Exit Sub
    
    ' Compter les transactions
    LastRow = ws.Cells(ws.Rows.Count, COL_TX_ENDTOENDID).End(xlUp).Row
    
    If LastRow < ROW_TRANSACTION_START Then
        TxCount = 0
        TotalAmount = 0#
    Else
        TxCount = LastRow - ROW_TRANSACTION_START + 1
        
        ' Calculer le total des montants
        For i = ROW_TRANSACTION_START To LastRow
            TotalAmount = TotalAmount + SafeCDbl(ws.Cells(i, COL_TX_AMT).Value)
        Next i
    End If
    
    ' Mettre à jour les cellules
    ws.Range(CELL_NBOFTXS).Value = TxCount
    ws.Range(CELL_CTRLSUM).Value = Format(TotalAmount, FORMAT_AMOUNT)
    
    LogInfo "Totaux recalculés: NbOfTxs=" & TxCount & ", CtrlSum=" & Format(TotalAmount, FORMAT_AMOUNT), _
            "modValidation.RecalculateTotals"
    
    On Error GoTo 0
End Sub

' =============================================================================
' FONCTION: ValidateIBANFormat
' DESCRIPTION: Valide le format d'un IBAN (wrapper pour compatibilité)
' =============================================================================
Public Function ValidateIBANFormat(ByVal IBAN As String) As Boolean
    ValidateIBANFormat = ValidateIBAN(IBAN)
End Function

' =============================================================================
' FONCTION: ValidateBICFormat
' DESCRIPTION: Valide le format d'un BIC (wrapper pour compatibilité)
' =============================================================================
Public Function ValidateBICFormat(ByVal BIC As String) As Boolean
    ValidateBICFormat = ValidateBIC(BIC)
End Function
