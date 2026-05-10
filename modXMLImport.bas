' ====================================================================================================
' MODULE: modXMLImport
' DESCRIPTION: Import et parsing du fichier XML SEPA
' ====================================================================================================
Option Explicit

' ====================================================================================================
' VARIABLES GLOBALES POUR LE STOCKAGE DE LA STRUCTURE XML
' ====================================================================================================
Public gXmlStructure As XmlStructure
Public gNamespaceUri As String
Public gNamespacePrefix As String
Public gDoc As MSXML2.DOMDocument60

' ====================================================================================================
' SUB: ImportXMLFile
' DESCRIPTION: Importe un fichier XML SEPA et le charge dans Excel
' PARAMETERS: filePath - Chemin complet du fichier XML
' RETURNS: Boolean - True si succès, False sinon
' ====================================================================================================
Public Function ImportXMLFile(ByVal filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim xmlDoc As MSXML2.DOMDocument60
    Dim success As Boolean
    
    ' Initialiser
    ImportXMLFile = False
    Set xmlDoc = New MSXML2.DOMDocument60
    
    ' Configurer le parser XML
    With xmlDoc
        .async = False
        .validateOnParse = False
        .preserveWhiteSpace = True
        .resolveExternals = False
    End With
    
    ' Charger le fichier XML
    If Not xmlDoc.Load(filePath) Then
        LogError "modXMLImport.ImportXMLFile", 0, "Erreur de chargement XML: " & xmlDoc.parseError.reason
        MsgBox "Erreur lors du chargement du fichier XML:" & vbCrLf & _
               xmlDoc.parseError.reason & vbCrLf & _
               "Ligne: " & xmlDoc.parseError.line & vbCrLf & _
               "Position: " & xmlDoc.parseError.linepos, vbCritical
        Exit Function
    End If
    
    ' Vérifier que le document n'est pas vide
    If xmlDoc.DocumentElement Is Nothing Then
        LogError "modXMLImport.ImportXMLFile", 0, "Document XML vide"
        MsgBox "Le fichier XML est vide ou invalide.", vbCritical
        Exit Function
    End If
    
    ' Analyser la structure XML
    Set gDoc = xmlDoc
    success = AnalyzeXMLStructure(xmlDoc)
    
    If Not success Then
        LogError "modXMLImport.ImportXMLFile", 0, "Erreur d'analyse de la structure XML"
        Exit Function
    End If
    
    ' Mapper les données vers Excel
    success = MapXMLToExcel(xmlDoc)
    
    If Not success Then
        LogError "modXMLImport.ImportXMLFile", 0, "Erreur de mapping XML vers Excel"
        Exit Function
    End If
    
    ImportXMLFile = True
    
    MsgBox "Import XML réussi!" & vbCrLf & _
           "Nombre de transactions: " & gXmlStructure.TransactionNodes.Count, vbInformation
    
    Exit Function
    
ErrorHandler:
    ImportXMLFile = False
    LogError "modXMLImport.ImportXMLFile", Err.Number, Err.Description
    MsgBox "Erreur lors de l'import XML:" & vbCrLf & Err.Description, vbCritical
End Function

' ====================================================================================================
' FUNCTION: AnalyzeXMLStructure
' DESCRIPTION: Analyse la structure complète du fichier XML
' PARAMETERS: xmlDoc - Document XML DOM
' RETURNS: Boolean - True si succès
' ====================================================================================================
Public Function AnalyzeXMLStructure(ByVal xmlDoc As MSXML2.DOMDocument60) As Boolean
    On Error GoTo ErrorHandler
    
    Dim rootElem As MSXML2.IXMLDOMElement
    Dim nsNode As MSXML2.IXMLDOMNode
    Dim node As MSXML2.IXMLDOMNode
    
    ' Initialiser la structure
    Set gXmlStructure.Namespaces = New Collection
    Set gXmlStructure.NodeOrder = New Collection
    Set gXmlStructure.Attributes = New Collection
    Set gXmlStructure.HeaderNodes = New Collection
    Set gXmlStructure.PaymentNodes = New Collection
    Set gXmlStructure.TransactionNodes = New Collection
    
    ' Obtenir l'élément racine
    Set rootElem = xmlDoc.DocumentElement
    gXmlStructure.RootNode = rootElem.nodeName
    
    ' Extraire le namespace
    If rootElem.getAttribute("xmlns") <> "" Then
        gNamespaceUri = rootElem.getAttribute("xmlns")
    ElseIf rootElem.getAttribute("xmlns:pain.001.001.03") <> "" Then
        gNamespaceUri = rootElem.getAttribute("xmlns:pain.001.001.03")
    Else
        ' Chercher dans les attributs
        For Each nsNode In rootElem.Attributes
            If InStr(nsNode.nodeName, "xmlns") > 0 Then
                gNamespaceUri = nsNode.nodeValue
                Exit For
            End If
        Next nsNode
    End If
    
    ' Stocker le namespace
    If gNamespaceUri <> "" Then
        gXmlStructure.Namespaces.Add gNamespaceUri, "default"
    End If
    
    ' Analyser l'ordre des nœuds
    Call AnalyzeNodeOrder(rootElem, gXmlStructure.NodeOrder)
    
    ' Identifier les blocs principaux
    Call IdentifyMainBlocks(rootElem)
    
    AnalyzeXMLStructure = True
    
    Exit Function
    
ErrorHandler:
    AnalyzeXMLStructure = False
    LogError "modXMLImport.AnalyzeXMLStructure", Err.Number, Err.Description
End Function

' ====================================================================================================
' SUB: AnalyzeNodeOrder
' DESCRIPTION: Analyse récursivement l'ordre des nœuds XML
' PARAMETERS: parentNode, nodeOrderCollection
' ====================================================================================================
Private Sub AnalyzeNodeOrder(ByVal parentNode As MSXML2.IXMLDOMNode, _
                             ByVal nodeOrderCollection As Collection)
    On Error Resume Next
    
    Dim childNode As MSXML2.IXMLDOMNode
    Dim nodeInfo As String
    
    For Each childNode In parentNode.ChildNodes
        If childNode.nodeType = NODE_ELEMENT Then
            nodeInfo = childNode.nodeName
            
            ' Ajouter à la collection
            On Error Resume Next
            nodeOrderCollection.Add nodeInfo, nodeInfo
            On Error GoTo 0
            
            ' Récursivité
            Call AnalyzeNodeOrder(childNode, nodeOrderCollection)
        End If
    Next childNode
End Sub

' ====================================================================================================
' SUB: IdentifyMainBlocks
' DESCRIPTION: Identifie les blocs principaux du document SEPA
' PARAMETERS: rootElem - Élément racine du XML
' ====================================================================================================
Private Sub IdentifyMainBlocks(ByVal rootElem As MSXML2.IXMLDOMElement)
    On Error Resume Next
    
    Dim grpHdrNode As MSXML2.IXMLDOMNode
    Dim pmtInfNode As MSXML2.IXMLDOMNode
    Dim txNode As MSXML2.IXMLDOMNode
    Dim nsPrefix As String
    
    ' Déterminer le préfixe de namespace
    If gNamespaceUri <> "" Then
        nsPrefix = "ns:"
    Else
        nsPrefix = ""
    End If
    
    ' Trouver GroupHeader
    Set grpHdrNode = GetNodeByPath(rootElem, nsPrefix & "GrpHdr")
    If Not grpHdrNode Is Nothing Then
        Call StoreHeaderNodes(grpHdrNode)
    End If
    
    ' Trouver PaymentInformation
    Set pmtInfNode = GetNodeByPath(rootElem, nsPrefix & "PmtInf")
    If Not pmtInfNode Is Nothing Then
        Call StorePaymentNodes(pmtInfNode)
        
        ' Trouver les transactions (CdtTrfTxInf)
        For Each txNode In pmtInfNode.ChildNodes
            If InStr(txNode.nodeName, "CdtTrfTxInf") > 0 Then
                gXmlStructure.TransactionNodes.Add txNode
            End If
        Next txNode
    End If
End Sub

' ====================================================================================================
' SUB: StoreHeaderNodes
' DESCRIPTION: Stocke les nœuds de l'en-tête
' PARAMETERS: headerNode - Nœud GroupHeader
' ====================================================================================================
Private Sub StoreHeaderNodes(ByVal headerNode As MSXML2.IXMLDOMNode)
    On Error Resume Next
    
    Dim childNode As MSXML2.IXMLDOMNode
    
    For Each childNode In headerNode.ChildNodes
        If childNode.nodeType = NODE_ELEMENT Then
            gXmlStructure.HeaderNodes.Add childNode, childNode.nodeName
        End If
    Next childNode
End Sub

' ====================================================================================================
' SUB: StorePaymentNodes
' DESCRIPTION: Stocke les nœuds de PaymentInformation
' PARAMETERS: paymentNode - Nœud PaymentInformation
' ====================================================================================================
Private Sub StorePaymentNodes(ByVal paymentNode As MSXML2.IXMLDOMNode)
    On Error Resume Next
    
    Dim childNode As MSXML2.IXMLDOMNode
    
    For Each childNode In paymentNode.ChildNodes
        If childNode.nodeType = NODE_ELEMENT Then
            ' Exclure les transactions qui sont traitées séparément
            If InStr(childNode.nodeName, "CdtTrfTxInf") = 0 Then
                gXmlStructure.PaymentNodes.Add childNode, childNode.nodeName
            End If
        End If
    Next childNode
End Sub

' ====================================================================================================
' FUNCTION: MapXMLToExcel
' DESCRIPTION: Mappe les données XML vers les cellules Excel
' PARAMETERS: xmlDoc - Document XML DOM
' RETURNS: Boolean - True si succès
' ====================================================================================================
Public Function MapXMLToExcel(ByVal xmlDoc As MSXML2.DOMDocument60) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim nsPrefix As String
    
    ' Créer/activer la feuille Principal
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_PRINCIPAL)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_PRINCIPAL
    End If
    
    ' Effacer le contenu existant
    ws.Cells.Clear
    
    ' Créer les en-têtes
    Call CreateExcelHeaders(ws)
    
    ' Déterminer le préfixe namespace
    If gNamespaceUri <> "" Then
        nsPrefix = "ns:"
        xmlDoc.setProperty "SelectionNamespaces", "xmlns:ns='" & gNamespaceUri & "'"
    End If
    
    ' Mapper l'en-tête GroupHeader
    Call MapGroupHeader(xmlDoc, ws, nsPrefix)
    
    ' Mapper PaymentInformation
    Call MapPaymentInformation(xmlDoc, ws, nsPrefix)
    
    ' Mapper les transactions
    Call MapTransactions(xmlDoc, ws, nsPrefix)
    
    ' Ajuster la largeur des colonnes
    ws.Columns.AutoFit
    
    MapXMLToExcel = True
    
    Exit Function
    
ErrorHandler:
    MapXMLToExcel = False
    LogError "modXMLImport.MapXMLToExcel", Err.Number, Err.Description
End Function

' ====================================================================================================
' SUB: CreateExcelHeaders
' DESCRIPTION: Crée les en-têtes de la feuille Excel
' PARAMETERS: ws - Feuille Excel
' ====================================================================================================
Private Sub CreateExcelHeaders(ByVal ws As Worksheet)
    On Error Resume Next
    
    ' Section ENTÊTE
    ws.Range("A1").Value = "SECTION ENTÊTE"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Interior.Color = RGB(200, 220, 255)
    
    ws.Range("A2").Value = "MsgId"
    ws.Range("A3").Value = "CreDtTm"
    ws.Range("A4").Value = "NbOfTxs"
    ws.Range("A5").Value = "CtrlSum"
    ws.Range("A6").Value = "InitiatingParty Name"
    ws.Range("A7").Value = "Debtor Name"
    ws.Range("A8").Value = "Debtor IBAN"
    ws.Range("A9").Value = "Debtor BIC"
    ws.Range("A10").Value = "RequestedExecutionDate"
    ws.Range("A11").Value = "Currency"
    ws.Range("A12").Value = "PaymentInfoId"
    ws.Range("A13").Value = "PaymentMethod"
    ws.Range("A14").Value = "BatchBooking"
    ws.Range("A15").Value = "NbOfTxs (PI)"
    ws.Range("A16").Value = "CtrlSum (PI)"
    ws.Range("A17").Value = "Debtor Account IBAN"
    ws.Range("A18").Value = "Debtor Agent BIC"
    ws.Range("A19").Value = "ChargeBearer"
    
    ' Section DÉTAILS
    ws.Range("A21").Value = "DÉTAILS DES TRANSACTIONS"
    ws.Range("A21").Font.Bold = True
    ws.Range("A21").Interior.Color = RGB(200, 255, 200)
    
    ws.Range("B21").Value = "EndToEndId"
    ws.Range("C21").Value = "InstrId"
    ws.Range("D21").Value = "Beneficiary Name"
    ws.Range("E21").Value = "Beneficiary IBAN"
    ws.Range("F21").Value = "Beneficiary BIC"
    ws.Range("G21").Value = "Amount"
    ws.Range("H21").Value = "RemittanceInfo"
    ws.Range("I21").Value = "ExecutionDate"
    ws.Range("J21").Value = "Beneficiary Address"
    ws.Range("K21").Value = "Country"
    ws.Range("L21").Value = "City"
    
    ' Style des en-têtes de détails
    ws.Range("B21:L21").Font.Bold = True
    ws.Range("B21:L21").Interior.Color = RGB(220, 240, 220)
End Sub

' ====================================================================================================
' SUB: MapGroupHeader
' DESCRIPTION: Mappe le GroupHeader vers Excel
' PARAMETERS: xmlDoc, ws, nsPrefix
' ====================================================================================================
Private Sub MapGroupHeader(ByVal xmlDoc As MSXML2.DOMDocument60, _
                           ByVal ws As Worksheet, _
                           ByVal nsPrefix As String)
    On Error Resume Next
    
    Dim value As String
    
    ' MsgId
    value = GetNodeValue(xmlDoc, nsPrefix & "Document/" & nsPrefix & "GrpHdr/" & nsPrefix & "MsgId")
    ws.Range(RANGE_MSGID).Value = value
    
    ' CreDtTm
    value = GetNodeValue(xmlDoc, nsPrefix & "Document/" & nsPrefix & "GrpHdr/" & nsPrefix & "CreDtTm")
    ws.Range(RANGE_CREDTTM).Value = value
    
    ' NbOfTxs
    value = GetNodeValue(xmlDoc, nsPrefix & "Document/" & nsPrefix & "GrpHdr/" & nsPrefix & "NbOfTxs")
    ws.Range(RANGE_NBFTXS).Value = value
    
    ' CtrlSum
    value = GetNodeValue(xmlDoc, nsPrefix & "Document/" & nsPrefix & "GrpHdr/" & nsPrefix & "CtrlSum")
    ws.Range(RANGE_CTRLSUM).Value = value
    
    ' InitiatingParty Name
    value = GetNodeValue(xmlDoc, nsPrefix & "Document/" & nsPrefix & "GrpHdr/" & nsPrefix & "InitgPty/" & nsPrefix & "Nm")
    ws.Range(RANGE_INITIATING_PARTY_NAME).Value = value
End Sub

' ====================================================================================================
' SUB: MapPaymentInformation
' DESCRIPTION: Mappe le PaymentInformation vers Excel
' PARAMETERS: xmlDoc, ws, nsPrefix
' ====================================================================================================
Private Sub MapPaymentInformation(ByVal xmlDoc As MSXML2.DOMDocument60, _
                                  ByVal ws As Worksheet, _
                                  ByVal nsPrefix As String)
    On Error Resume Next
    
    Dim value As String
    
    ' PaymentInfoId
    value = GetNodeValue(xmlDoc, nsPrefix & "Document/" & nsPrefix & "PmtInf/" & nsPrefix & "PmtInfId")
    ws.Range(RANGE_PAYMENT_INFO_ID).Value = value
    
    ' PaymentMethod
    value = GetNodeValue(xmlDoc, nsPrefix & "Document/" & nsPrefix & "PmtInf/" & nsPrefix & "PmtMtd")
    ws.Range(RANGE_PAYMENT_METHOD).Value = value
    
    ' BatchBooking
    value = GetNodeValue(xmlDoc, nsPrefix & "Document/" & nsPrefix & "PmtInf/" & nsPrefix & "BtchBookg")
    ws.Range(RANGE_BATCH_BOOKING).Value = value
    
    ' RequestedExecutionDate
    value = GetNodeValue(xmlDoc, nsPrefix & "Document/" & nsPrefix & "PmtInf/" & nsPrefix & "ReqdExctnDt")
    ws.Range(RANGE_REQUESTED_EXECUTION_DATE).Value = value
    
    ' Debtor Name
    value = GetNodeValue(xmlDoc, nsPrefix & "Document/" & nsPrefix & "PmtInf/" & nsPrefix & "Dbtr/" & nsPrefix & "Nm")
    ws.Range(RANGE_DEBTOR_NAME).Value = value
    
    ' Debtor IBAN
    value = GetNodeValue(xmlDoc, nsPrefix & "Document/" & nsPrefix & "PmtInf/" & nsPrefix & "DbtrAcct/" & nsPrefix & "Id/" & nsPrefix & "IBAN")
    ws.Range(RANGE_DEBTOR_IBAN).Value = value
    
    ' Debtor BIC (via DbtrAgt)
    value = GetNodeValue(xmlDoc, nsPrefix & "Document/" & nsPrefix & "PmtInf/" & nsPrefix & "DbtrAgt/" & nsPrefix & "FinInstnId/" & nsPrefix & "BIC")
    ws.Range(RANGE_DEBTOR_BIC).Value = value
    
    ' DebtorAgent BIC
    ws.Range(RANGE_DEBTOR_AGENT_BIC).Value = value
    
    ' ChargeBearer
    value = GetNodeValue(xmlDoc, nsPrefix & "Document/" & nsPrefix & "PmtInf/" & nsPrefix & "ChrgBr")
    ws.Range(RANGE_CHARGE_BEARER).Value = value
    
    ' Currency (depuis CdtTrfTxInf/InstdAmt)
    value = GetNodeAttribute(xmlDoc, nsPrefix & "Document/" & nsPrefix & "PmtInf/" & nsPrefix & "CdtTrfTxInf/" & nsPrefix & "Amt/" & nsPrefix & "InstdAmt", "Ccy")
    If value = "" Then
        value = "EUR" ' Valeur par défaut
    End If
    ws.Range(RANGE_DEVISE).Value = value
End Sub

' ====================================================================================================
' SUB: MapTransactions
' DESCRIPTION: Mappe les transactions vers Excel
' PARAMETERS: xmlDoc, ws, nsPrefix
' ====================================================================================================
Private Sub MapTransactions(ByVal xmlDoc As MSXML2.DOMDocument60, _
                            ByVal ws As Worksheet, _
                            ByVal nsPrefix As String)
    On Error Resume Next
    
    Dim txNodes As MSXML2.IXMLDOMNodeList
    Dim txNode As MSXML2.IXMLDOMNode
    Dim row As Long
    Dim value As String
    
    row = TX_START_ROW
    
    ' Sélectionner tous les nœuds CdtTrfTxInf
    Set txNodes = xmlDoc.SelectNodes("//" & nsPrefix & "CdtTrfTxInf")
    
    If txNodes Is Nothing Then
        Exit Sub
    End If
    
    For Each txNode In txNodes
        ' EndToEndId
        value = GetNodeValueFromNode(txNode, nsPrefix & "PmtId/" & nsPrefix & "EndToEndId")
        ws.Cells(row, TX_COL_ENDTOENDID).Value = value
        
        ' InstrId (optionnel)
        value = GetNodeValueFromNode(txNode, nsPrefix & "PmtId/" & nsPrefix & "InstrId")
        ws.Cells(row, TX_COL_INSTRID).Value = value
        
        ' Beneficiary Name (Cdtr)
        value = GetNodeValueFromNode(txNode, nsPrefix & "Cdtr/" & nsPrefix & "Nm")
        ws.Cells(row, TX_COL_BENEFICIARY_NAME).Value = value
        
        ' Beneficiary IBAN
        value = GetNodeValueFromNode(txNode, nsPrefix & "CdtrAcct/" & nsPrefix & "Id/" & nsPrefix & "IBAN")
        ws.Cells(row, TX_COL_BENEFICIARY_IBAN).Value = value
        
        ' Beneficiary BIC
        value = GetNodeValueFromNode(txNode, nsPrefix & "CdtrAgt/" & nsPrefix & "FinInstnId/" & nsPrefix & "BIC")
        ws.Cells(row, TX_COL_BENEFICIARY_BIC).Value = value
        
        ' Amount
        value = GetNodeValueFromNode(txNode, nsPrefix & "Amt/" & nsPrefix & "InstdAmt")
        ws.Cells(row, TX_COL_AMOUNT).Value = value
        
        ' RemittanceInfo
        value = GetNodeValueFromNode(txNode, nsPrefix & "RmtInf/" & nsPrefix & "Ustrd")
        ws.Cells(row, TX_COL_REMITTANCE_INFO).Value = value
        
        ' ExecutionDate (peut être dans PmtInf/ReqdExctnDt)
        ' Déjà copié depuis PaymentInformation
        
        row = row + 1
    Next txNode
End Sub

' ====================================================================================================
' FUNCTION: GetNodeValue
' DESCRIPTION: Récupère la valeur d'un nœud par chemin XPath
' PARAMETERS: xmlDoc, xpath
' RETURNS: String - Valeur du nœud
' ====================================================================================================
Public Function GetNodeValue(ByVal xmlDoc As MSXML2.DOMDocument60, ByVal xpath As String) As String
    On Error Resume Next
    
    Dim node As MSXML2.IXMLDOMNode
    
    Set node = xmlDoc.SelectSingleNode(xpath)
    
    If Not node Is Nothing Then
        GetNodeValue = UnescapeXMLText(node.Text)
    Else
        GetNodeValue = ""
    End If
End Function

' ====================================================================================================
' FUNCTION: GetNodeValueFromNode
' DESCRIPTION: Récupère la valeur d'un nœud enfant
' PARAMETERS: parentNode, relativeXpath
' RETURNS: String - Valeur du nœud
' ====================================================================================================
Public Function GetNodeValueFromNode(ByVal parentNode As MSXML2.IXMLDOMNode, ByVal relativeXpath As String) As String
    On Error Resume Next
    
    Dim node As MSXML2.IXMLDOMNode
    
    Set node = parentNode.SelectSingleNode(relativeXpath)
    
    If Not node Is Nothing Then
        GetNodeValueFromNode = UnescapeXMLText(node.Text)
    Else
        GetNodeValueFromNode = ""
    End If
End Function

' ====================================================================================================
' FUNCTION: GetNodeAttribute
' DESCRIPTION: Récupère un attribut d'un nœud
' PARAMETERS: xmlDoc, xpath, attributeName
' RETURNS: String - Valeur de l'attribut
' ====================================================================================================
Public Function GetNodeAttribute(ByVal xmlDoc As MSXML2.DOMDocument60, _
                                 ByVal xpath As String, _
                                 ByVal attributeName As String) As String
    On Error Resume Next
    
    Dim node As MSXML2.IXMLDOMNode
    Dim attr As MSXML2.IXMLDOMAttribute
    
    Set node = xmlDoc.SelectSingleNode(xpath)
    
    If Not node Is Nothing Then
        Set attr = node.Attributes.getNamedItem(attributeName)
        If Not attr Is Nothing Then
            GetNodeAttribute = attr.nodeValue
        Else
            GetNodeAttribute = ""
        End If
    Else
        GetNodeAttribute = ""
    End If
End Function

' ====================================================================================================
' FUNCTION: GetNodeByPath
' DESCRIPTION: Récupère un nœud par chemin relatif
' PARAMETERS: rootNode, nodeName
' RETURNS: IXMLDOMNode
' ====================================================================================================
Public Function GetNodeByPath(ByVal rootNode As MSXML2.IXMLDOMNode, ByVal nodeName As String) As MSXML2.IXMLDOMNode
    On Error Resume Next
    
    Dim childNode As MSXML2.IXMLDOMNode
    
    For Each childNode In rootNode.ChildNodes
        If InStr(childNode.nodeName, nodeName) > 0 Then
            Set GetNodeByPath = childNode
            Exit Function
        End If
    Next childNode
    
    Set GetNodeByPath = Nothing
End Function
