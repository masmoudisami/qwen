' ====================================================================================================
' MODULE: modUtils
' DESCRIPTION: Fonctions utilitaires pour l'application SEPA XML
' ====================================================================================================
Option Explicit

' ====================================================================================================
' FUNCTION: FormatAmount
' DESCRIPTION: Formate un montant selon le standard SEPA (pas de séparateur de milliers, point décimal)
' PARAMETERS: amount - Le montant à formater
' RETURNS: String - Le montant formaté
' ====================================================================================================
Public Function FormatAmount(ByVal amount As Double) As String
    On Error GoTo ErrorHandler
    
    ' Format SEPA: pas de séparateur de milliers, 2 décimales, point comme séparateur décimal
    FormatAmount = Replace(Format(amount, "0.00"), ",", ".")
    
    Exit Function
    
ErrorHandler:
    FormatAmount = "0.00"
    LogError "modUtils.FormatAmount", Err.Number, Err.Description
End Function

' ====================================================================================================
' FUNCTION: ParseAmount
' DESCRIPTION: Parse un montant depuis une chaîne de caractères
' PARAMETERS: amountStr - La chaîne contenant le montant
' RETURNS: Double - Le montant parsé
' ====================================================================================================
Public Function ParseAmount(ByVal amountStr As String) As Double
    On Error GoTo ErrorHandler
    
    Dim cleaned As String
    
    ' Nettoyer la chaîne: supprimer espaces, remplacer virgule par point
    cleaned = Trim(amountStr)
    cleaned = Replace(cleaned, " ", "")
    cleaned = Replace(cleaned, ",", ".")
    
    ' Convertir en Double
    ParseAmount = CDbl(cleaned)
    
    Exit Function
    
ErrorHandler:
    ParseAmount = 0#
    LogError "modUtils.ParseAmount", Err.Number, Err.Description
End Function

' ====================================================================================================
' FUNCTION: FormatDateISO
' DESCRIPTION: Formate une date au format ISO 8601 (YYYY-MM-DD)
' PARAMETERS: dt - La date à formater
' RETURNS: String - La date formatée
' ====================================================================================================
Public Function FormatDateISO(ByVal dt As Date) As String
    On Error GoTo ErrorHandler
    
    FormatDateISO = Year(dt) & "-" & Right("0" & Month(dt), 2) & "-" & Right("0" & Day(dt), 2)
    
    Exit Function
    
ErrorHandler:
    FormatDateISO = ""
    LogError "modUtils.FormatDateISO", Err.Number, Err.Description
End Function

' ====================================================================================================
' FUNCTION: FormatDateTimeISO
' DESCRIPTION: Formate une date/heure au format ISO 8601 complet (YYYY-MM-DDThh:mm:ss)
' PARAMETERS: dt - La date/heure à formater
' RETURNS: String - La date/heure formatée
' ====================================================================================================
Public Function FormatDateTimeISO(ByVal dt As Date) As String
    On Error GoTo ErrorHandler
    
    FormatDateTimeISO = Year(dt) & "-" & _
                        Right("0" & Month(dt), 2) & "-" & _
                        Right("0" & Day(dt), 2) & "T" & _
                        Right("0" & Hour(dt), 2) & ":" & _
                        Right("0" & Minute(dt), 2) & ":" & _
                        Right("0" & Second(dt), 2)
    
    Exit Function
    
ErrorHandler:
    FormatDateTimeISO = ""
    LogError "modUtils.FormatDateTimeISO", Err.Number, Err.Description
End Function

' ====================================================================================================
' FUNCTION: ParseDateISO
' DESCRIPTION: Parse une date depuis une chaîne au format ISO
' PARAMETERS: dateStr - La chaîne contenant la date
' RETURNS: Date - La date parsée
' ====================================================================================================
Public Function ParseDateISO(ByVal dateStr As String) As Date
    On Error GoTo ErrorHandler
    
    Dim parts() As String
    Dim dateParts() As String
    
    ' Gérer le format avec heure (T separator)
    If InStr(dateStr, "T") > 0 Then
        parts = Split(dateStr, "T")
        dateStr = parts(0)
    End If
    
    ' Parser YYYY-MM-DD
    dateParts = Split(dateStr, "-")
    
    If UBound(dateParts) >= 2 Then
        ParseDateISO = DateSerial(CInt(dateParts(0)), CInt(dateParts(1)), CInt(dateParts(2)))
    Else
        ParseDateISO = CDate(dateStr)
    End If
    
    Exit Function
    
ErrorHandler:
    ParseDateISO = #1/1/1900#
    LogError "modUtils.ParseDateISO", Err.Number, Err.Description
End Function

' ====================================================================================================
' FUNCTION: ValidateIBAN
' DESCRIPTION: Valide un IBAN selon l'algorithme MOD-97
' PARAMETERS: iban - L'IBAN à valider
' RETURNS: Boolean - True si valide, False sinon
' ====================================================================================================
Public Function ValidateIBAN(ByVal iban As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim cleanIBAN As String
    Dim rearranged As String
    Dim numericStr As String
    Dim i As Long
    Dim chunk As String
    Dim remainder As Double
    Dim charCode As Long
    Dim c As String
    
    ' Nettoyer l'IBAN: supprimer espaces, mettre en majuscules
    cleanIBAN = UCase(Replace(iban, " ", ""))
    
    ' Vérifier longueur minimale
    If Len(cleanIBAN) < 15 Then
        ValidateIBAN = False
        Exit Function
    End If
    
    ' Vérifier format: 2 lettres + chiffres
    If Not Left(cleanIBAN, 2) Like "[A-Z][A-Z]" Then
        ValidateIBAN = False
        Exit Function
    End If
    
    ' Déplacer les 4 premiers caractères à la fin
    rearranged = Mid(cleanIBAN, 5) & Left(cleanIBAN, 4)
    
    ' Convertir lettres en chiffres (A=10, B=11, ..., Z=35)
    numericStr = ""
    For i = 1 To Len(rearranged)
        c = Mid(rearranged, i, 1)
        If c Like "[A-Z]" Then
            charCode = Asc(c) - Asc("A") + 10
            numericStr = numericStr & CStr(charCode)
        Else
            numericStr = numericStr & c
        End If
    Next i
    
    ' Calculer MOD-97 en traitant par chunks de 9 chiffres
    remainder = 0
    For i = 1 To Len(numericStr) Step 9
        chunk = Mid(numericStr, i, 9)
        remainder = (remainder * (10 ^ Len(chunk)) + CDbl(chunk)) Mod 97
    Next i
    
    ' Si remainder = 1, l'IBAN est valide
    ValidateIBAN = (remainder = 1)
    
    Exit Function
    
ErrorHandler:
    ValidateIBAN = False
    LogError "modUtils.ValidateIBAN", Err.Number, Err.Description
End Function

' ====================================================================================================
' FUNCTION: ValidateBIC
' DESCRIPTION: Valide un code BIC/SWIFT
' PARAMETERS: bic - Le code BIC à valider
' RETURNS: Boolean - True si valide, False sinon
' ====================================================================================================
Public Function ValidateBIC(ByVal bic As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim cleanBIC As String
    
    ' Nettoyer le BIC: supprimer espaces, mettre en majuscules
    cleanBIC = UCase(Trim(bic))
    
    ' Vérifier longueur (8 ou 11 caractères)
    If Len(cleanBIC) <> 8 And Len(cleanBIC) <> 11 Then
        ValidateBIC = False
        Exit Function
    End If
    
    ' Vérifier format: 4 lettres + 2 lettres/chiffres + 2 lettres/chiffres + optionnel 3 lettres/chiffres
    If Not Left(cleanBIC, 4) Like "[A-Z][A-Z][A-Z][A-Z]" Then
        ValidateBIC = False
        Exit Function
    End If
    
    If Not Mid(cleanBIC, 5, 2) Like "[A-Z0-9][A-Z0-9]" Then
        ValidateBIC = False
        Exit Function
    End If
    
    If Len(cleanBIC) = 11 Then
        If Not Mid(cleanBIC, 7, 4) Like "[A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9]" Then
            ValidateBIC = False
            Exit Function
        End If
    End If
    
    ValidateBIC = True
    
    Exit Function
    
ErrorHandler:
    ValidateBIC = False
    LogError "modUtils.ValidateBIC", Err.Number, Err.Description
End Function

' ====================================================================================================
' FUNCTION: SanitizeXMLText
' DESCRIPTION: Échappe les caractères spéciaux XML
' PARAMETERS: text - Le texte à échapper
' RETURNS: String - Le texte échappé
' ====================================================================================================
Public Function SanitizeXMLText(ByVal text As String) As String
    On Error GoTo ErrorHandler
    
    ' Échapper dans l'ordre correct
    text = Replace(text, "&", "&amp;")
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    text = Replace(text, """", "&quot;")
    text = Replace(text, "'", "&apos;")
    
    SanitizeXMLText = text
    
    Exit Function
    
ErrorHandler:
    SanitizeXMLText = ""
    LogError "modUtils.SanitizeXMLText", Err.Number, Err.Description
End Function

' ====================================================================================================
' FUNCTION: UnescapeXMLText
' DESCRIPTION: Déséchappe les caractères spéciaux XML
' PARAMETERS: text - Le texte échappé
' RETURNS: String - Le texte original
' ====================================================================================================
Public Function UnescapeXMLText(ByVal text As String) As String
    On Error GoTo ErrorHandler
    
    ' Déséchapper dans l'ordre inverse
    text = Replace(text, "&apos;", "'")
    text = Replace(text, "&quot;", """")
    text = Replace(text, "&gt;", ">")
    text = Replace(text, "&lt;", "<")
    text = Replace(text, "&amp;", "&")
    
    UnescapeXMLText = text
    
    Exit Function
    
ErrorHandler:
    UnescapeXMLText = ""
    LogError "modUtils.UnescapeXMLText", Err.Number, Err.Description
End Function

' ====================================================================================================
' FUNCTION: GenerateMsgId
' DESCRIPTION: Génère un identifiant de message unique
' PARAMETERS: prefix - Préfixe optionnel
' RETURNS: String - L'identifiant généré
' ====================================================================================================
Public Function GenerateMsgId(Optional ByVal prefix As String = "MSG") As String
    On Error GoTo ErrorHandler
    
    Dim timestamp As String
    Dim randomPart As String
    
    timestamp = Format(Now, "yyyymmddhhmmss")
    Randomize
    randomPart = Right("000" & CStr(Int(Rnd * 1000)), 3)
    
    GenerateMsgId = prefix & "-" & timestamp & "-" & randomPart
    
    Exit Function
    
ErrorHandler:
    GenerateMsgId = "MSG-" & Format(Now, "yyyymmddhhmmss")
    LogError "modUtils.GenerateMsgId", Err.Number, Err.Description
End Function

' ====================================================================================================
' FUNCTION: GetUniqueEndToEndId
' DESCRIPTION: Génère un EndToEndId unique
' PARAMETERS: baseId - Identifiant de base
' RETURNS: String - L'EndToEndId généré
' ====================================================================================================
Public Function GetUniqueEndToEndId(Optional ByVal baseId As String = "E2E") As String
    On Error GoTo ErrorHandler
    
    Dim timestamp As String
    Dim randomPart As String
    
    timestamp = Format(Now, "yyyymmdd")
    Randomize
    randomPart = Right("00000" & CStr(Int(Rnd * 100000)), 5)
    
    GetUniqueEndToEndId = baseId & "-" & timestamp & "-" & randomPart
    
    Exit Function
    
ErrorHandler:
    GetUniqueEndToEndId = baseId & "-" & Format(Now, "yyyymmdd")
    LogError "modUtils.GetUniqueEndToEndId", Err.Number, Err.Description
End Function

' ====================================================================================================
' SUB: LogError
' DESCRIPTION: Enregistre une erreur dans le log
' PARAMETERS: moduleName, errNumber, errDescription
' ====================================================================================================
Public Sub LogError(ByVal moduleName As String, ByVal errNumber As Long, ByVal errDescription As String)
    On Error Resume Next
    
    Dim wsLog As Worksheet
    Dim nextRow As Long
    Dim timestamp As String
    
    timestamp = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    ' Tenter d'écrire dans la feuille Log si elle existe
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets(SHEET_LOG)
    On Error GoTo ErrorHandler
    
    If wsLog Is Nothing Then
        ' Créer la feuille Log si elle n'existe pas
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsLog.Name = SHEET_LOG
        
        ' En-têtes
        wsLog.Range("A1").Value = "Timestamp"
        wsLog.Range("B1").Value = "Module"
        wsLog.Range("C1").Value = "Error Number"
        wsLog.Range("D1").Value = "Description"
    End If
    
    ' Trouver la prochaine ligne libre
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Écrire l'erreur
    wsLog.Cells(nextRow, 1).Value = timestamp
    wsLog.Cells(nextRow, 2).Value = moduleName
    wsLog.Cells(nextRow, 3).Value = errNumber
    wsLog.Cells(nextRow, 4).Value = errDescription
    
    Exit Sub
    
ErrorHandler:
    ' Silencieux - éviter les boucles d'erreurs
End Sub

' ====================================================================================================
' FUNCTION: CleanString
' DESCRIPTION: Nettoie une chaîne de caractères (supprime caractères interdits SEPA)
' PARAMETERS: text - Le texte à nettoyer
' RETURNS: String - Le texte nettoyé
' ====================================================================================================
Public Function CleanString(ByVal text As String) As String
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim c As String
    Dim result As String
    
    ' Caractères autorisés SEPA: A-Z, a-z, 0-9, et certains caractères spéciaux
    result = ""
    
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        
        ' Vérifier si le caractère est autorisé
        If c Like "[A-Za-z0-9]" Or c = " " Or c = "." Or c = "," Or c = "-" Or _
           c = "(" Or c = ")" Or c = "/" Or c = "'" Or c = "+" Or c = "?" Or _
           c = ":" Or c = ";" Or c = "!" Or c = "_" Or c = """" Or c = "@" Or _
           c = "&" Or c = "*" Or c = "%" Or c = "#" Or c = "€" Or c = "$" Or _
           c = "£" Or c = "¥" Or c = "ç" Or c = "à" Or c = "é" Or c = "è" Or _
           c = "ù" Or c = "â" Or c = "ê" Or c = "î" Or c = "ô" Or c = "û" Or _
           c = "ä" Or c = "ë" Or c = "ï" Or c = "ö" Or c = "ü" Or c = "œ" Or _
           c = "æ" Then
            result = result & c
        End If
    Next i
    
    CleanString = result
    
    Exit Function
    
ErrorHandler:
    CleanString = ""
    LogError "modUtils.CleanString", Err.Number, Err.Description
End Function

' ====================================================================================================
' FUNCTION: CalculateCtrlSum
' DESCRIPTION: Calcule la somme de contrôle des transactions
' PARAMETERS: transactions - Tableau de transactions
' RETURNS: Double - La somme totale
' ====================================================================================================
Public Function CalculateCtrlSum(transactions() As TransactionDetails) As Double
    On Error GoTo ErrorHandler
    
    Dim total As Double
    Dim i As Long
    
    total = 0#
    
    For i = LBound(transactions) To UBound(transactions)
        total = total + transactions(i).Amount
    Next i
    
    CalculateCtrlSum = total
    
    Exit Function
    
ErrorHandler:
    CalculateCtrlSum = 0#
    LogError "modUtils.CalculateCtrlSum", Err.Number, Err.Description
End Function

' ====================================================================================================
' FUNCTION: CountNonEmptyTransactions
' DESCRIPTION: Compte le nombre de transactions non vides
' PARAMETERS: ws - Feuille Excel contenant les transactions
' RETURNS: Long - Le nombre de transactions
' ====================================================================================================
Public Function CountNonEmptyTransactions(ws As Worksheet) As Long
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    Dim i As Long
    Dim count As Long
    
    lastRow = ws.Cells(ws.Rows.Count, TX_COL_ENDTOENDID).End(xlUp).Row
    
    count = 0
    
    For i = TX_START_ROW To lastRow
        If Trim(ws.Cells(i, TX_COL_ENDTOENDID).Value) <> "" Then
            count = count + 1
        End If
    Next i
    
    CountNonEmptyTransactions = count
    
    Exit Function
    
ErrorHandler:
    CountNonEmptyTransactions = 0
    LogError "modUtils.CountNonEmptyTransactions", Err.Number, Err.Description
End Function
