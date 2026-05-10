'===============================================================================
' MODULE: modUtils.bas
' DESCRIPTION: Fonctions utilitaires pour le gestionnaire SEPA XML
'===============================================================================
Option Explicit

' =============================================================================
' FONCTION: FormatAmountForXML
' DESCRIPTION: Formate un montant selon la norme ISO 20022 (3 décimales)
' =============================================================================
Public Function FormatAmountForXML(ByVal Amount As Double) As String
    On Error GoTo ErrorHandler
    
    ' Formater avec 3 décimales comme dans le modèle
    FormatAmountForXML = Format(Amount, "0.000")
    
    Exit Function
    
ErrorHandler:
    FormatAmountForXML = "0.000"
    LogError ERR_VALIDATION_AMOUNT, "Erreur de formatage du montant: " & Amount, _
             "modUtils.FormatAmountForXML", , , True
End Function

' =============================================================================
' FONCTION: ParseAmountFromXML
' DESCRIPTION: Parse un montant depuis une chaîne XML
' =============================================================================
Public Function ParseAmountFromXML(ByVal AmountStr As String) As Double
    On Error GoTo ErrorHandler
    
    Dim CleanStr As String
    
    ' Remplacer les virgules par des points si nécessaire
    CleanStr = Replace(Trim(AmountStr), ",", ".")
    
    ' Convertir en Double
    ParseAmountFromXML = CDbl(CleanStr)
    
    Exit Function
    
ErrorHandler:
    ParseAmountFromXML = 0#
    LogError ERR_VALIDATION_AMOUNT, "Erreur de parsing du montant: " & AmountStr, _
             "modUtils.ParseAmountFromXML", , , True
End Function

' =============================================================================
' FONCTION: FormatDateForXML
' DESCRIPTION: Formate une date selon le format ISO 8601
' =============================================================================
Public Function FormatDateForXML(ByVal dt As Date, Optional ByVal IncludeTime As Boolean = True) As String
    On Error GoTo ErrorHandler
    
    If IncludeTime Then
        ' Format complet: YYYY-MM-DDThh:mm:ss
        FormatDateForXML = Format(dt, "yyyy-mm-dd") & "T" & Format(dt, "hh:mm:ss")
    Else
        ' Format court: YYYY-MM-DD
        FormatDateForXML = Format(dt, "yyyy-mm-dd")
    End If
    
    Exit Function
    
ErrorHandler:
    FormatDateForXML = Format(Date, "yyyy-mm-dd")
    LogError ERR_VALIDATION_DATE, "Erreur de formatage de la date", _
             "modUtils.FormatDateForXML", , , True
End Function

' =============================================================================
' FONCTION: ParseDateFromXML
' DESCRIPTION: Parse une date depuis une chaîne XML
' =============================================================================
Public Function ParseDateFromXML(ByVal DateStr As String) As Date
    On Error GoTo ErrorHandler
    
    Dim CleanStr As String
    Dim DatePart As String
    Dim TimePart As String
    
    CleanStr = Trim(DateStr)
    
    ' Vérifier si le format inclut le temps (T separator)
    If InStr(CleanStr, "T") > 0 Then
        DatePart = Split(CleanStr, "T")(0)
        TimePart = Split(CleanStr, "T")(1)
        
        ' Parser la date et l'heure
        ParseDateFromXML = CDate(Replace(DatePart, "-", "/"))
        If Len(TimePart) >= 8 Then
            ParseDateFromXML = ParseDateFromXML + TimeValue(Left(TimePart, 8))
        End If
    Else
        ' Format date seule
        ParseDateFromXML = CDate(Replace(CleanStr, "-", "/"))
    End If
    
    Exit Function
    
ErrorHandler:
    ParseDateFromXML = Date
    LogError ERR_VALIDATION_DATE, "Erreur de parsing de la date: " & DateStr, _
             "modUtils.ParseDateFromXML", , , True
End Function

' =============================================================================
' FONCTION: ValidateIBAN
' DESCRIPTION: Valide un IBAN en utilisant l'algorithme MOD-97
' =============================================================================
Public Function ValidateIBAN(ByVal IBAN As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim CleanIBAN As String
    Dim RearrangedIBAN As String
    Dim NumericIBAN As String
    Dim Remainder As Long
    Dim i As Integer
    Dim Chunk As String
    Dim CheckDigits As String
    
    If Len(Trim(IBAN)) = 0 Then
        ValidateIBAN = False
        Exit Function
    End If
    
    ' Nettoyer l'IBAN (supprimer espaces, mettre en majuscules)
    CleanIBAN = UCase(Replace(Replace(IBAN, " ", ""), vbTab, ""))
    
    ' Vérifier la longueur minimale
    If Len(CleanIBAN) < IBAN_MIN_LENGTH Or Len(CleanIBAN) > IBAN_MAX_LENGTH Then
        ValidateIBAN = False
        Exit Function
    End If
    
    ' Vérifier que les 2 premiers caractères sont des lettres
    If Not IsAlpha(Mid(CleanIBAN, 1, 2)) Then
        ValidateIBAN = False
        Exit Function
    End If
    
    ' Déplacer les 4 premiers caractères à la fin
    RearrangedIBAN = Mid(CleanIBAN, 5) & Left(CleanIBAN, 4)
    
    ' Convertir les lettres en chiffres (A=10, B=11, ..., Z=35)
    NumericIBAN = ""
    For i = 1 To Len(RearrangedIBAN)
        Dim Char As String
        Char = Mid(RearrangedIBAN, i, 1)
        
        If IsNumeric(Char) Then
            NumericIBAN = NumericIBAN & Char
        ElseIf IsAlpha(Char) Then
            NumericIBAN = NumericIBAN & CStr(Asc(Char) - Asc("A") + 10)
        Else
            NumericIBAN = NumericIBAN & Char
        End If
    Next i
    
    ' Calculer MOD-97 en traitant par chunks pour éviter les dépassements
    Remainder = 0
    i = 1
    Do While i <= Len(NumericIBAN)
        ' Prendre jusqu'à 9 chiffres à la fois
        If Len(CStr(Remainder)) < 9 Then
            Chunk = Remainder & Mid(NumericIBAN, i, 9 - Len(CStr(Remainder)))
            i = i + (9 - Len(CStr(Remainder)))
        Else
            Chunk = CStr(Remainder)
        End If
        
        Remainder = CLng(Chunk) Mod 97
    Loop
    
    ' Si le reste est 1, l'IBAN est valide
    ValidateIBAN = (Remainder = 1)
    
    Exit Function
    
ErrorHandler:
    ValidateIBAN = False
    LogError ERR_VALIDATION_IBAN, "Erreur de validation IBAN: " & IBAN, _
             "modUtils.ValidateIBAN", , , True
End Function

' =============================================================================
' FONCTION: ValidateBIC
' DESCRIPTION: Valide un code BIC/SWIFT
' =============================================================================
Public Function ValidateBIC(ByVal BIC As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim CleanBIC As String
    
    If Len(Trim(BIC)) = 0 Then
        ValidateBIC = False
        Exit Function
    End If
    
    ' Nettoyer le BIC
    CleanBIC = UCase(Replace(Replace(BIC, " ", ""), vbTab, ""))
    
    ' Vérifier la longueur (8 ou 11 caractères)
    If Len(CleanBIC) <> BIC_LENGTH And Len(CleanBIC) <> BIC_LENGTH_WITH_BRANCH Then
        ValidateBIC = False
        Exit Function
    End If
    
    ' Les 4 premiers caractères doivent être des lettres (code banque)
    If Not IsAlpha(Left(CleanBIC, 4)) Then
        ValidateBIC = False
        Exit Function
    End If
    
    ' Les 2 caractères suivants doivent être des lettres (code pays)
    If Not IsAlpha(Mid(CleanBIC, 5, 2)) Then
        ValidateBIC = False
        Exit Function
    End If
    
    ' Les 2 caractères suivants doivent être alphanumériques (code lieu)
    Dim LocationCode As String
    LocationCode = Mid(CleanBIC, 7, 2)
    If Not IsAlphaNumeric(LocationCode) Then
        ValidateBIC = False
        Exit Function
    End If
    
    ' Si 11 caractères, les 3 derniers doivent être alphanumériques (code agence)
    If Len(CleanBIC) = BIC_LENGTH_WITH_BRANCH Then
        Dim BranchCode As String
        BranchCode = Right(CleanBIC, 3)
        If Not IsAlphaNumeric(BranchCode) Then
            ValidateBIC = False
            Exit Function
        End If
    End If
    
    ValidateBIC = True
    
    Exit Function
    
ErrorHandler:
    ValidateBIC = False
    LogError ERR_VALIDATION_BIC, "Erreur de validation BIC: " & BIC, _
             "modUtils.ValidateBIC", , , True
End Function

' =============================================================================
' FONCTION: IsAlpha
' DESCRIPTION: Vérifie si une chaîne ne contient que des lettres
' =============================================================================
Public Function IsAlpha(ByVal str As String) As Boolean
    Dim i As Integer
    Dim Char As String
    
    IsAlpha = True
    For i = 1 To Len(str)
        Char = Mid(str, i, 1)
        If Char < "A" Or Char > "Z" Then
            If Char < "a" Or Char > "z" Then
                IsAlpha = False
                Exit Function
            End If
        End If
    Next i
End Function

' =============================================================================
' FONCTION: IsAlphaNumeric
' DESCRIPTION: Vérifie si une chaîne est alphanumérique
' =============================================================================
Public Function IsAlphaNumeric(ByVal str As String) As Boolean
    Dim i As Integer
    Dim Char As String
    
    IsAlphaNumeric = True
    For i = 1 To Len(str)
        Char = Mid(str, i, 1)
        If Not ((Char >= "A" And Char <= "Z") Or _
                (Char >= "a" And Char <= "z") Or _
                (Char >= "0" And Char <= "9")) Then
            IsAlphaNumeric = False
            Exit Function
        End If
    Next i
End Function

' =============================================================================
' FONCTION: SanitizeStringForXML
' DESCRIPTION: Nettoie une chaîne pour l'export XML (caractères spéciaux)
' =============================================================================
Public Function SanitizeStringForXML(ByVal str As String) As String
    On Error GoTo ErrorHandler
    
    Dim Result As String
    
    If Len(Trim(str)) = 0 Then
        SanitizeStringForXML = ""
        Exit Function
    End If
    
    Result = str
    
    ' Encoder les caractères spéciaux XML
    Result = Replace(Result, "&", "&amp;")
    Result = Replace(Result, "<", "&lt;")
    Result = Replace(Result, ">", "&gt;")
    Result = Replace(Result, """", "&quot;")
    Result = Replace(Result, "'", "&apos;")
    
    ' Supprimer les caractères non autorisés SEPA
    Dim i As Integer
    Dim CleanResult As String
    Dim Char As String
    
    CleanResult = ""
    For i = 1 To Len(Result)
        Char = Mid(Result, i, 1)
        If InStr(ALLOWED_CHARS, Char) > 0 Or Asc(Char) > 127 Then
            CleanResult = CleanResult & Char
        End If
    Next i
    
    SanitizeStringForXML = CleanResult
    
    Exit Function
    
ErrorHandler:
    SanitizeStringForXML = ""
    LogError ERR_VALIDATION_CHARS, "Erreur de nettoyage de chaîne", _
             "modUtils.SanitizeStringForXML", , , True
End Function

' =============================================================================
' FONCTION: GenerateUniqueMsgId
' DESCRIPTION: Génère un identifiant de message unique
' =============================================================================
Public Function GenerateUniqueMsgId(Optional ByVal Prefix As String = "MSG") As String
    Dim Timestamp As String
    Dim RandomNum As String
    
    Timestamp = Format(Now, "yyyymmddhhmmss")
    RandomNum = Format(Int(Rnd * 10000), "0000")
    
    GenerateUniqueMsgId = Prefix & "_" & Timestamp & "_" & RandomNum
End Function

' =============================================================================
' FONCTION: CalculateCtrlSum
' DESCRIPTION: Calcule la somme de contrôle des transactions
' =============================================================================
Public Function CalculateCtrlSum(Transactions() As TTransaction) As Double
    On Error GoTo ErrorHandler
    
    Dim Total As Double
    Dim i As Long
    
    Total = 0#
    
    If Not IsArray(Transactions) Then
        CalculateCtrlSum = 0#
        Exit Function
    End If
    
    For i = LBound(Transactions) To UBound(Transactions)
        Total = Total + Transactions(i).Amt.Amount
    Next i
    
    CalculateCtrlSum = Total
    
    Exit Function
    
ErrorHandler:
    CalculateCtrlSum = 0#
    LogError ERR_VALIDATION_AMOUNT, "Erreur de calcul de la somme de contrôle", _
             "modUtils.CalculateCtrlSum", , , True
End Function

' =============================================================================
' FONCTION: GetWorksheetByName
' DESCRIPTION: Récupère une feuille Excel par son nom, la crée si nécessaire
' =============================================================================
Public Function GetWorksheetByName(ByVal SheetName As String) As Worksheet
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName)
    
    If ws Is Nothing Then
        ' Créer la feuille si elle n'existe pas
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SheetName
    End If
    
    Set GetWorksheetByName = ws
    
    On Error GoTo 0
End Function

' =============================================================================
' FONCTION: ClearWorksheetRange
' DESCRIPTION: Efface le contenu d'une plage de cellules
' =============================================================================
Public Sub ClearWorksheetRange(ws As Worksheet, StartRow As Long, EndRow As Long, _
                               StartCol As Long, EndCol As Long)
    On Error GoTo ErrorHandler
    
    If EndRow < StartRow Or EndCol < StartCol Then Exit Sub
    
    ws.Range(ws.Cells(StartRow, StartCol), ws.Cells(EndRow, EndCol)).ClearContents
    
    Exit Sub
    
ErrorHandler:
    LogError ERR_NONE, "Erreur lors du nettoyage de la plage", _
             "modUtils.ClearWorksheetRange", , , True
End Sub

' =============================================================================
' FONCTION: ArrayCount
' DESCRIPTION: Compte le nombre d'éléments dans un tableau
' =============================================================================
Public Function ArrayCount(arr As Variant) As Long
    On Error Resume Next
    
    If Not IsArray(arr) Then
        ArrayCount = 0
        Exit Function
    End If
    
    ArrayCount = UBound(arr) - LBound(arr) + 1
    
    If Err.Number <> 0 Then
        ArrayCount = 0
    End If
    
    On Error GoTo 0
End Function

' =============================================================================
' FONCTION: IsEmptyVariant
' DESCRIPTION: Vérifie si une variante est vide ou null
' =============================================================================
Public Function IsEmptyVariant(v As Variant) As Boolean
    IsEmptyVariant = IsEmpty(v) Or IsNull(v) Or (VarType(v) = vbString And Len(Trim(CStr(v))) = 0)
End Function

' =============================================================================
' FONCTION: SafeCStr
' DESCRIPTION: Conversion sécurisée en chaîne
' =============================================================================
Public Function SafeCStr(v As Variant) As String
    On Error Resume Next
    
    If IsEmptyVariant(v) Then
        SafeCStr = ""
        Exit Function
    End If
    
    SafeCStr = CStr(v)
    
    If Err.Number <> 0 Then
        SafeCStr = ""
    End If
    
    On Error GoTo 0
End Function

' =============================================================================
' FONCTION: SafeCDbl
' DESCRIPTION: Conversion sécurisée en Double
' =============================================================================
Public Function SafeCDbl(v As Variant) As Double
    On Error Resume Next
    
    If IsEmptyVariant(v) Then
        SafeCDbl = 0#
        Exit Function
    End If
    
    SafeCDbl = CDbl(v)
    
    If Err.Number <> 0 Then
        SafeCDbl = 0#
    End If
    
    On Error GoTo 0
End Function

' =============================================================================
' FONCTION: SafeCDate
' DESCRIPTION: Conversion sécurisée en Date
' =============================================================================
Public Function SafeCDate(v As Variant) As Date
    On Error Resume Next
    
    If IsEmptyVariant(v) Then
        SafeCDate = Date
        Exit Function
    End If
    
    SafeCDate = CDate(v)
    
    If Err.Number <> 0 Then
        SafeCDate = Date
    End If
    
    On Error GoTo 0
End Function

' =============================================================================
' FONCTION: SafeCLng
' DESCRIPTION: Conversion sécurisée en Long
' =============================================================================
Public Function SafeCLng(v As Variant) As Long
    On Error Resume Next
    
    If IsEmptyVariant(v) Then
        SafeCLng = 0
        Exit Function
    End If
    
    SafeCLng = CLng(v)
    
    If Err.Number <> 0 Then
        SafeCLng = 0
    End If
    
    On Error GoTo 0
End Function
