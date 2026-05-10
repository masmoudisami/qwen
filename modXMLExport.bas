' ====================================================================================================
' MODULE: modXMLExport
' DESCRIPTION: Export des données Excel vers fichier XML SEPA conforme
' ====================================================================================================
Option Explicit

' ====================================================================================================
' SUB: ExportToXML
' DESCRIPTION: Exporte les données Excel vers un fichier XML SEPA
' PARAMETERS: filePath - Chemin complet du fichier de sortie
' RETURNS: Boolean - True si succès
' ====================================================================================================
Public Function ExportToXML(ByVal filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim xmlDoc As MSXML2.DOMDocument60
    Dim rootElem As MSXML2.IXMLDOMElement
    Dim ws As Worksheet
    Dim success As Boolean
    
    ' Initialiser
    ExportToXML = False
    Set ws = ThisWorkbook.Worksheets(SHEET_PRINCIPAL)
    
    If ws Is Nothing Then
        MsgBox "La feuille '" & SHEET_PRINCIPAL & "' n'existe pas.", vbCritical
        Exit Function
    End If
    
    ' Créer le document XML
    Set xmlDoc = New MSXML2.DOMDocument60
    
    With xmlDoc
        .async = False
        .validateOnParse = False
        .preserveWhiteSpace = True
        .resolveExternals = False
    End With
    
    ' Créer l'élément racine avec namespace
    Set rootElem = xmlDoc.createElement("Document")
    rootElem.setAttribute "xmlns", gNamespaceUri
    rootElem.setAttribute "xmlns:xsi", NS_XSI
    
    xmlDoc.appendChild rootElem
    
    ' Construire la structure XML complète
    Call BuildGroupHeader(xmlDoc, rootElem, ws)
    Call BuildPaymentInformation(xmlDoc, rootElem, ws)
    
    ' Sauvegarder le fichier avec indentation
    success = SaveXMLFile(xmlDoc, filePath)
    
    If Not success Then
        LogError "modXMLExport.ExportToXML", 0, "Erreur lors de la sauvegarde du fichier XML"
        Exit Function
    End If
    
    ExportToXML = True
    
    MsgBox "Export XML réussi!" & vbCrLf & "Fichier: " & filePath, vbInformation
    
    Exit Function
    
ErrorHandler:
    ExportToXML = False
    LogError "modXMLExport.ExportToXML", Err.Number, Err.Description
    MsgBox "Erreur lors de l'export XML:" & vbCrLf & Err.Description, vbCritical
End Function

' ====================================================================================================
' SUB: BuildGroupHeader
' DESCRIPTION: Construit le bloc GroupHeader
' PARAMETERS: xmlDoc, parentElem, ws
' ====================================================================================================
Private Sub BuildGroupHeader(ByVal xmlDoc As MSXML2.DOMDocument60, _
                             ByVal parentElem As MSXML2.IXMLDOMElement, _
                             ByVal ws As Worksheet)
    On Error Resume Next
    
    Dim grpHdrElem As MSXML2.IXMLDOMElement
    Dim childElem As MSXML2.IXMLDOMElement
    Dim value As String
    
    ' Créer GrpHdr
    Set grpHdrElem = xmlDoc.createElement("GrpHdr")
    parentElem.appendChild grpHdrElem
    
    ' MsgId
    value = CStr(ws.Range(RANGE_MSGID).Value)
    If value = "" Then value = GenerateMsgId()
    Set childElem = CreateElementWithValue(xmlDoc, "MsgId", value)
    grpHdrElem.appendChild childElem
    
    ' CreDtTm
    value = CStr(ws.Range(RANGE_CREDTTM).Value)
    If value = "" Then value = FormatDateTimeISO(Now)
    Set childElem = CreateElementWithValue(xmlDoc, "CreDtTm", value)
    grpHdrElem.appendChild childElem
    
    ' NbOfTxs (calculé automatiquement)
    value = CStr(CountNonEmptyTransactions(ws))
    Set childElem = CreateElementWithValue(xmlDoc, "NbOfTxs", value)
    grpHdrElem.appendChild childElem
    
    ' CtrlSum (calculé automatiquement)
    value = FormatAmount(CalculateCtrlSumFromExcel(ws))
    Set childElem = CreateElementWithValue(xmlDoc, "CtrlSum", value)
    grpHdrElem.appendChild childElem
    
    ' InitgPty (InitiatingParty)
    Set childElem = xmlDoc.createElement("InitgPty")
    grpHdrElem.appendChild childElem
    
    value = CStr(ws.Range(RANGE_INITIATING_PARTY_NAME).Value)
    Dim nmElem As MSXML2.IXMLDOMElement
    Set nmElem = CreateElementWithValue(xmlDoc, "Nm", SanitizeXMLText(value))
    childElem.appendChild nmElem
End Sub

' ====================================================================================================
' SUB: BuildPaymentInformation
' DESCRIPTION: Construit le bloc PaymentInformation avec toutes les transactions
' PARAMETERS: xmlDoc, parentElem, ws
' ====================================================================================================
Private Sub BuildPaymentInformation(ByVal xmlDoc As MSXML2.DOMDocument60, _
                                    ByVal parentElem As MSXML2.IXMLDOMElement, _
                                    ByVal ws As Worksheet)
    On Error Resume Next
    
    Dim pmtInfElem As MSXML2.IXMLDOMElement
    Dim childElem As MSXML2.IXMLDOMElement
    Dim value As String
    Dim lastRow As Long
    Dim i As Long
    Dim txCount As Long
    Dim totalAmount As Double
    
    ' Créer PmtInf
    Set pmtInfElem = xmlDoc.createElement("PmtInf")
    parentElem.appendChild pmtInfElem
    
    ' PmtInfId
    value = CStr(ws.Range(RANGE_PAYMENT_INFO_ID).Value)
    If value = "" Then value = GenerateMsgId("PMT")
    Set childElem = CreateElementWithValue(xmlDoc, "PmtInfId", value)
    pmtInfElem.appendChild childElem
    
    ' PmtMtd
    value = CStr(ws.Range(RANGE_PAYMENT_METHOD).Value)
    If value = "" Then value = "TRF"
    Set childElem = CreateElementWithValue(xmlDoc, "PmtMtd", value)
    pmtInfElem.appendChild childElem
    
    ' BtchBookg
    value = CStr(ws.Range(RANGE_BATCH_BOOKING).Value)
    If value = "" Then value = "true"
    Set childElem = CreateElementWithValue(xmlDoc, "BtchBookg", value)
    pmtInfElem.appendChild childElem
    
    ' NbOfTxs (dans PmtInf)
    txCount = CountNonEmptyTransactions(ws)
    Set childElem = CreateElementWithValue(xmlDoc, "NbOfTxs", CStr(txCount))
    pmtInfElem.appendChild childElem
    
    ' CtrlSum (dans PmtInf)
    totalAmount = CalculateCtrlSumFromExcel(ws)
    Set childElem = CreateElementWithValue(xmlDoc, "CtrlSum", FormatAmount(totalAmount))
    pmtInfElem.appendChild childElem
    
    ' RqdExctnDt
    value = CStr(ws.Range(RANGE_REQUESTED_EXECUTION_DATE).Value)
    If value = "" Then value = FormatDateISO(Date + 1)
    Set childElem = CreateElementWithValue(xmlDoc, "ReqdExctnDt", value)
    pmtInfElem.appendChild childElem
    
    ' Dbtr (Debtor)
    Set childElem = xmlDoc.createElement("Dbtr")
    pmtInfElem.appendChild childElem
    
    value = CStr(ws.Range(RANGE_DEBTOR_NAME).Value)
    Dim nmElem As MSXML2.IXMLDOMElement
    Set nmElem = CreateElementWithValue(xmlDoc, "Nm", SanitizeXMLText(value))
    childElem.appendChild nmElem
    
    ' DbtrAcct (Debtor Account)
    Set childElem = xmlDoc.createElement("DbtrAcct")
    pmtInfElem.appendChild childElem
    
    Dim idElem As MSXML2.IXMLDOMElement
    Set idElem = xmlDoc.createElement("Id")
    childElem.appendChild idElem
    
    value = CStr(ws.Range(RANGE_DEBTOR_IBAN).Value)
    Set nmElem = CreateElementWithValue(xmlDoc, "IBAN", value)
    idElem.appendChild nmElem
    
    ' DbtrAgt (Debtor Agent)
    Set childElem = xmlDoc.createElement("DbtrAgt")
    pmtInfElem.appendChild childElem
    
    Dim finInstnElem As MSXML2.IXMLDOMElement
    Set finInstnElem = xmlDoc.createElement("FinInstnId")
    childElem.appendChild finInstnElem
    
    value = CStr(ws.Range(RANGE_DEBTOR_BIC).Value)
    Set nmElem = CreateElementWithValue(xmlDoc, "BIC", value)
    finInstnElem.appendChild nmElem
    
    ' ChrgBr
    value = CStr(ws.Range(RANGE_CHARGE_BEARER).Value)
    If value = "" Then value = "SLEV"
    Set childElem = CreateElementWithValue(xmlDoc, "ChrgBr", value)
    pmtInfElem.appendChild childElem
    
    ' Currency
    Dim currency As String
    currency = CStr(ws.Range(RANGE_DEVISE).Value)
    If currency = "" Then currency = "EUR"
    
    ' Construire les transactions (CdtTrfTxInf)
    lastRow = ws.Cells(ws.Rows.Count, TX_COL_ENDTOENDID).End(xlUp).Row
    
    For i = TX_START_ROW To lastRow
        If Trim(ws.Cells(i, TX_COL_ENDTOENDID).Value) <> "" Then
            Call BuildCreditTransferTransaction(xmlDoc, pmtInfElem, ws, i, currency)
        End If
    Next i
End Sub

' ====================================================================================================
' SUB: BuildCreditTransferTransaction
' DESCRIPTION: Construit une transaction individuelle CdtTrfTxInf
' PARAMETERS: xmlDoc, parentElem, ws, row, currency
' ====================================================================================================
Private Sub BuildCreditTransferTransaction(ByVal xmlDoc As MSXML2.DOMDocument60, _
                                           ByVal parentElem As MSXML2.IXMLDOMElement, _
                                           ByVal ws As Worksheet, _
                                           ByVal row As Long, _
                                           ByVal currency As String)
    On Error Resume Next
    
    Dim cdtTrfTxInfElem As MSXML2.IXMLDOMElement
    Dim childElem As MSXML2.IXMLDOMElement
    Dim value As String
    
    ' Créer CdtTrfTxInf
    Set cdtTrfTxInfElem = xmlDoc.createElement("CdtTrfTxInf")
    parentElem.appendChild cdtTrfTxInfElem
    
    ' PmtId
    Set childElem = xmlDoc.createElement("PmtId")
    cdtTrfTxInfElem.appendChild childElem
    
    ' InstrId (optionnel)
    value = CStr(ws.Cells(row, TX_COL_INSTRID).Value)
    If value <> "" Then
        Dim instrElem As MSXML2.IXMLDOMElement
        Set instrElem = CreateElementWithValue(xmlDoc, "InstrId", SanitizeXMLText(value))
        childElem.appendChild instrElem
    End If
    
    ' EndToEndId (obligatoire)
    value = CStr(ws.Cells(row, TX_COL_ENDTOENDID).Value)
    If value = "" Then value = GetUniqueEndToEndId()
    Dim e2eElem As MSXML2.IXMLDOMElement
    Set e2eElem = CreateElementWithValue(xmlDoc, "EndToEndId", SanitizeXMLText(value))
    childElem.appendChild e2eElem
    
    ' Amt (Amount)
    Set childElem = xmlDoc.createElement("Amt")
    cdtTrfTxInfElem.appendChild childElem
    
    Dim instdAmtElem As MSXML2.IXMLDOMElement
    Set instdAmtElem = xmlDoc.createElement("InstdAmt")
    instdAmtElem.setAttribute "Ccy", currency
    value = FormatAmount(CDbl(ws.Cells(row, TX_COL_AMOUNT).Value))
    instdAmtElem.Text = value
    childElem.appendChild instdAmtElem
    
    ' CdtrAgt (Creditor Agent)
    Set childElem = xmlDoc.createElement("CdtrAgt")
    cdtTrfTxInfElem.appendChild childElem
    
    Dim finInstnElem As MSXML2.IXMLDOMElement
    Set finInstnElem = xmlDoc.createElement("FinInstnId")
    childElem.appendChild finInstnElem
    
    value = CStr(ws.Cells(row, TX_COL_BENEFICIARY_BIC).Value)
    Dim bicElem As MSXML2.IXMLDOMElement
    Set bicElem = CreateElementWithValue(xmlDoc, "BIC", value)
    finInstnElem.appendChild bicElem
    
    ' Cdtr (Creditor)
    Set childElem = xmlDoc.createElement("Cdtr")
    cdtTrfTxInfElem.appendChild childElem
    
    value = CStr(ws.Cells(row, TX_COL_BENEFICIARY_NAME).Value)
    Dim nmElem As MSXML2.IXMLDOMElement
    Set nmElem = CreateElementWithValue(xmlDoc, "Nm", SanitizeXMLText(value))
    childElem.appendChild nmElem
    
    ' CdtrAcct (Creditor Account)
    Set childElem = xmlDoc.createElement("CdtrAcct")
    cdtTrfTxInfElem.appendChild childElem
    
    Dim idElem As MSXML2.IXMLDOMElement
    Set idElem = xmlDoc.createElement("Id")
    childElem.appendChild idElem
    
    value = CStr(ws.Cells(row, TX_COL_BENEFICIARY_IBAN).Value)
    Set nmElem = CreateElementWithValue(xmlDoc, "IBAN", value)
    idElem.appendChild nmElem
    
    ' RmtInf (Remittance Information)
    value = CStr(ws.Cells(row, TX_COL_REMITTANCE_INFO).Value)
    If value <> "" Then
        Set childElem = xmlDoc.createElement("RmtInf")
        cdtTrfTxInfElem.appendChild childElem
        
        Dim ustrdElem As MSXML2.IXMLDOMElement
        Set ustrdElem = CreateElementWithValue(xmlDoc, "Ustrd", SanitizeXMLText(value))
        childElem.appendChild ustrdElem
    End If
End Sub

' ====================================================================================================
' FUNCTION: CreateElementWithValue
' DESCRIPTION: Crée un élément XML avec une valeur texte
' PARAMETERS: xmlDoc, elementName, elementValue
' RETURNS: IXMLDOMElement
' ====================================================================================================
Public Function CreateElementWithValue(ByVal xmlDoc As MSXML2.DOMDocument60, _
                                       ByVal elementName As String, _
                                       ByVal elementValue As String) As MSXML2.IXMLDOMElement
    On Error Resume Next
    
    Dim elem As MSXML2.IXMLDOMElement
    
    Set elem = xmlDoc.createElement(elementName)
    elem.Text = elementValue
    
    Set CreateElementWithValue = elem
End Function

' ====================================================================================================
' FUNCTION: CalculateCtrlSumFromExcel
' DESCRIPTION: Calcule la somme de contrôle depuis Excel
' PARAMETERS: ws - Feuille Excel
' RETURNS: Double
' ====================================================================================================
Public Function CalculateCtrlSumFromExcel(ByVal ws As Worksheet) As Double
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    Dim i As Long
    Dim total As Double
    Dim amount As Variant
    
    total = 0#
    lastRow = ws.Cells(ws.Rows.Count, TX_COL_AMOUNT).End(xlUp).Row
    
    For i = TX_START_ROW To lastRow
        amount = ws.Cells(i, TX_COL_AMOUNT).Value
        If IsNumeric(amount) And amount <> "" Then
            total = total + CDbl(amount)
        End If
    Next i
    
    CalculateCtrlSumFromExcel = total
    
    Exit Function
    
ErrorHandler:
    CalculateCtrlSumFromExcel = 0#
    LogError "modXMLExport.CalculateCtrlSumFromExcel", Err.Number, Err.Description
End Function

' ====================================================================================================
' FUNCTION: SaveXMLFile
' DESCRIPTION: Sauvegarde le fichier XML avec formatage propre
' PARAMETERS: xmlDoc, filePath
' RETURNS: Boolean
' ====================================================================================================
Public Function SaveXMLFile(ByVal xmlDoc As MSXML2.DOMDocument60, ByVal filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim xmlString As String
    Dim formattedXml As String
    Dim fileNum As Integer
    
    ' Obtenir le XML brut
    xmlString = xmlDoc.xml
    
    ' Formater le XML avec indentation
    formattedXml = FormatXMLString(xmlString)
    
    ' Sauvegarder en UTF-8
    fileNum = FreeFile
    
    Open filePath For Output As #fileNum
    Print #fileNum, formattedXml
    Close #fileNum
    
    ' Sauvegarder également en .txt UTF-8
    Dim txtPath As String
    txtPath = Left(filePath, InStrRev(filePath, ".")) & "txt"
    
    Open txtPath For Output As #fileNum
    Print #fileNum, formattedXml
    Close #fileNum
    
    SaveXMLFile = True
    
    Exit Function
    
ErrorHandler:
    SaveXMLFile = False
    LogError "modXMLExport.SaveXMLFile", Err.Number, Err.Description
End Function

' ====================================================================================================
' FUNCTION: FormatXMLString
' DESCRIPTION: Formate une chaîne XML avec indentation
' PARAMETERS: xmlString
' RETURNS: String
' ====================================================================================================
Public Function FormatXMLString(ByVal xmlString As String) As String
    On Error Resume Next
    
    Dim result As String
    Dim indentLevel As Long
    Dim i As Long
    Dim currentChar As String
    Dim nextChar As String
    Dim inTag As Boolean
    Dim tagContent As String
    
    indentLevel = 0
    result = ""
    inTag = False
    tagContent = ""
    
    ' Ajouter l'en-tête XML
    result = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    
    i = 1
    
    ' Ignorer l'éventuel en-tête existant
    If Left(xmlString, 5) = "<?xml" Then
        Do While i <= Len(xmlString) And Mid(xmlString, i, 1) <> ">"
            i = i + 1
        Loop
        i = i + 1
        ' Sauter la nouvelle ligne après l'en-tête
        Do While i <= Len(xmlString) And (Mid(xmlString, i, 1) = vbCrLf Or Mid(xmlString, i, 1) = vbCr Or Mid(xmlString, i, 1) = vbLf)
            i = i + 1
        Loop
    End If
    
    ' Parser caractère par caractère
    Do While i <= Len(xmlString)
        currentChar = Mid(xmlString, i, 1)
        
        If i < Len(xmlString) Then
            nextChar = Mid(xmlString, i + 1, 1)
        Else
            nextChar = ""
        End If
        
        If currentChar = "<" Then
            ' Début d'une balise
            inTag = True
            tagContent = "<"
            
            ' Vérifier si c'est une balise fermante
            If nextChar = "/" Then
                indentLevel = indentLevel - 1
                If indentLevel < 0 Then indentLevel = 0
                result = result & String(indentLevel * 2, " ")
            ElseIf nextChar = "?" Then
                ' Balise de traitement
            Else
                ' Balise ouvrante
                result = result & String(indentLevel * 2, " ")
            End If
            
        ElseIf currentChar = ">" And inTag Then
            ' Fin d'une balise
            tagContent = tagContent & ">"
            result = result & tagContent & vbCrLf
            inTag = False
            tagContent = ""
            
            ' Vérifier si c'est une balise auto-fermante ou ouvrante
            If Mid(result, Len(result) - 2, 1) = "/" Then
                ' Balise auto-fermante, ne pas incrémenter
            ElseIf Mid(xmlString, i - 1, 1) <> "/" Then
                ' Balise ouvrante normale
                If nextChar <> "/" Then
                    indentLevel = indentLevel + 1
                End If
            End If
            
        ElseIf inTag Then
            tagContent = tagContent & currentChar
            
        ElseIf currentChar = " " Or currentChar = vbTab Then
            ' Ignorer les espaces entre les balises
            
        Else
            ' Contenu texte
            If currentChar <> vbCrLf And currentChar <> vbCr And currentChar <> vbLf Then
                result = result & String(indentLevel * 2, " ") & currentChar & vbCrLf
            End If
        End If
        
        i = i + 1
    Loop
    
    FormatXMLString = result
End Function

' ====================================================================================================
' SUB: ExportToTXT
' DESCRIPTION: Exporte les données en fichier TXT UTF-8
' PARAMETERS: filePath - Chemin du fichier
' RETURNS: Boolean
' ====================================================================================================
Public Function ExportToTXT(ByVal filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim fileNum As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim line As String
    
    Set ws = ThisWorkbook.Worksheets(SHEET_PRINCIPAL)
    
    If ws Is Nothing Then
        ExportToTXT = False
        Exit Function
    End If
    
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    ' En-têtes
    Print #fileNum, "EndToEndId;InstrId;BeneficiaryName;BeneficiaryIBAN;BeneficiaryBIC;Amount;RemittanceInfo;ExecutionDate"
    
    ' Données
    lastRow = ws.Cells(ws.Rows.Count, TX_COL_ENDTOENDID).End(xlUp).Row
    
    For i = TX_START_ROW To lastRow
        If Trim(ws.Cells(i, TX_COL_ENDTOENDID).Value) <> "" Then
            line = ws.Cells(i, TX_COL_ENDTOENDID).Value & ";" & _
                   ws.Cells(i, TX_COL_INSTRID).Value & ";" & _
                   ws.Cells(i, TX_COL_BENEFICIARY_NAME).Value & ";" & _
                   ws.Cells(i, TX_COL_BENEFICIARY_IBAN).Value & ";" & _
                   ws.Cells(i, TX_COL_BENEFICIARY_BIC).Value & ";" & _
                   FormatAmount(CDbl(ws.Cells(i, TX_COL_AMOUNT).Value)) & ";" & _
                   ws.Cells(i, TX_COL_REMITTANCE_INFO).Value & ";" & _
                   ws.Range(RANGE_REQUESTED_EXECUTION_DATE).Value
            
            Print #fileNum, line
        End If
    Next i
    
    Close #fileNum
    
    ExportToTXT = True
    
    Exit Function
    
ErrorHandler:
    ExportToTXT = False
    LogError "modXMLExport.ExportToTXT", Err.Number, Err.Description
End Function
