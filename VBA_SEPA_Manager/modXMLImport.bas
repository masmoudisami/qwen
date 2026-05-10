'===============================================================================
' MODULE: modXMLImport.bas
' DESCRIPTION: Import et parsing du fichier XML SEPA (pain.001.001.03)
'===============================================================================
Option Explicit

' =============================================================================
' FONCTION: ImportXMLFile
' DESCRIPTION: Importe un fichier XML SEPA et le parse dans Excel
' =============================================================================
Public Function ImportXMLFile(ByVal FilePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim xmlDoc As Object
    Dim Success As Boolean
    
    ' Initialiser le logging
    InitializeErrorLogging
    ClearErrorLog
    
    LogInfo "Début de l'import du fichier: " & FilePath, "modXMLImport.ImportXMLFile"
    
    ' Créer le document XML
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    
    ' Configurer le parser
    With xmlDoc
        .async = False
        .validateOnParse = False
        .preserveWhiteSpace = True
        .resolveExternals = False
    End With
    
    ' Charger le fichier
    If Not xmlDoc.Load(FilePath) Then
        LogError ERR_XML_LOAD, "Impossible de charger le fichier XML: " & _
                 xmlDoc.parseError.reason & " (ligne " & xmlDoc.parseError.line & ")", _
                 "modXMLImport.ImportXMLFile", , , True
        ImportXMLFile = False
        Exit Function
    End If
    
    LogInfo "Fichier XML chargé avec succès", "modXMLImport.ImportXMLFile"
    
    ' Analyser la structure et extraire les données
    Success = ParseXMLDocument(xmlDoc)
    
    If Success Then
        LogInfo "Import terminé avec succès", "modXMLImport.ImportXMLFile"
        ImportXMLFile = True
    Else
        LogError ERR_XML_PARSE, "Erreur lors de l'analyse du document XML", _
                 "modXMLImport.ImportXMLFile", , , True
        ImportXMLFile = False
    End If
    
    Exit Function
    
ErrorHandler:
    LogError ERR_XML_LOAD, "Erreur critique lors de l'import: " & Err.Description, _
             "modXMLImport.ImportXMLFile", , , True
    ImportXMLFile = False
End Function

' =============================================================================
' FONCTION: ParseXMLDocument
' DESCRIPTION: Parse le document XML et remplit les structures de données
' =============================================================================
Private Function ParseXMLDocument(ByRef xmlDoc As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim DocNode As Object
    Dim CstmrNode As Object
    Dim GrpHdrNode As Object
    Dim PmtInfNode As Object
    Dim SEPADoc As TSEPAXMLDocument
    Dim ws As Worksheet
    
    ' Obtenir le nœud racine Document
    Set DocNode = xmlDoc.getElementsByTagName(TAG_DOCUMENT)(0)
    If DocNode Is Nothing Then
        LogError ERR_XML_MISSING_NODE, "Balise " & TAG_DOCUMENT & " introuvable", _
                 "modXMLImport.ParseXMLDocument", , , True
        ParseXMLDocument = False
        Exit Function
    End If
    
    ' Extraire les namespaces
    SEPADoc.NamespaceURI = GetNamespaceURI(DocNode, NS_ISO)
    SEPADoc.NamespaceXSI = GetNamespaceURI(DocNode, NS_XSI)
    
    LogInfo "Namespace ISO détecté: " & SEPADoc.NamespaceURI, "modXMLImport.ParseXMLDocument"
    
    ' Obtenir le nœud CstmrCdtTrfInitn
    Set CstmrNode = xmlDoc.getElementsByTagName(TAG_CSTMRCDTTRFINITN)(0)
    If CstmrNode Is Nothing Then
        LogError ERR_XML_MISSING_NODE, "Balise " & TAG_CSTMRCDTTRFINITN & " introuvable", _
                 "modXMLImport.ParseXMLDocument", , , True
        ParseXMLDocument = False
        Exit Function
    End If
    
    ' Parser Group Header
    Set GrpHdrNode = CstmrNode.getElementsByTagName(TAG_GRPHDR)(0)
    If GrpHdrNode Is Nothing Then
        LogError ERR_XML_MISSING_NODE, "Balise " & TAG_GRPHDR & " introuvable", _
                 "modXMLImport.ParseXMLDocument", , , True
        ParseXMLDocument = False
        Exit Function
    End If
    
    ParseGroupHeader GrpHdrNode, SEPADoc.GrpHdr
    
    ' Parser Payment Information
    Set PmtInfNode = CstmrNode.getElementsByTagName(TAG_PMTINF)(0)
    If PmtInfNode Is Nothing Then
        LogError ERR_XML_MISSING_NODE, "Balise " & TAG_PMTINF & " introuvable", _
                 "modXMLImport.ParseXMLDocument", , , True
        ParseXMLDocument = False
        Exit Function
    End If
    
    ParsePaymentInformation PmtInfNode, SEPADoc.PmtInf
    
    ' Écrire les données dans Excel
    Set ws = SetupPrincipalSheet()
    WriteDataToExcel ws, SEPADoc
    
    ParseXMLDocument = True
    
    Exit Function
    
ErrorHandler:
    LogError ERR_XML_PARSE, "Erreur lors du parsing: " & Err.Description, _
             "modXMLImport.ParseXMLDocument", , , True
    ParseXMLDocument = False
End Function

' =============================================================================
' SUB: ParseGroupHeader
' DESCRIPTION: Parse le nœud GrpHdr
' =============================================================================
Private Sub ParseGroupHeader(ByRef GrpHdrNode As Object, ByRef GrpHdr As TGroupHeader)
    On Error Resume Next
    
    Dim Node As Object
    Dim InitgPtyNode As Object
    
    ' MsgId
    Set Node = GrpHdrNode.getElementsByTagName(TAG_MSGID)(0)
    If Not Node Is Nothing Then GrpHdr.MsgId = Node.Text
    
    ' CreDtTm
    Set Node = GrpHdrNode.getElementsByTagName(TAG_CREDTTM)(0)
    If Not Node Is Nothing Then
        GrpHdr.CreDtTm = ParseDateFromXML(Node.Text)
    End If
    
    ' NbOfTxs
    Set Node = GrpHdrNode.getElementsByTagName(TAG_NBOFTXS)(0)
    If Not Node Is Nothing Then
        GrpHdr.NbOfTxs = CLng(Node.Text)
    End If
    
    ' CtrlSum
    Set Node = GrpHdrNode.getElementsByTagName(TAG_CTRLSUM)(0)
    If Not Node Is Nothing Then
        GrpHdr.CtrlSum = ParseAmountFromXML(Node.Text)
    End If
    
    ' InitgPty
    Set InitgPtyNode = GrpHdrNode.getElementsByTagName(TAG_INITGPTY)(0)
    If Not InitgPtyNode Is Nothing Then
        ParseParty InitgPtyNode, GrpHdr.InitgPty
    End If
    
    LogInfo "Group Header parsé: MsgId=" & GrpHdr.MsgId, "modXMLImport.ParseGroupHeader"
    
    On Error GoTo 0
End Sub

' =============================================================================
' SUB: ParsePaymentInformation
' DESCRIPTION: Parse le nœud PmtInf
' =============================================================================
Private Sub ParsePaymentInformation(ByRef PmtInfNode As Object, ByRef PmtInf As TPaymentInformation)
    On Error Resume Next
    
    Dim Node As Object
    Dim DbtrNode As Object
    Dim DbtrAcctNode As Object
    Dim DbtrAgtNode As Object
    Dim TxNodes As Object
    Dim i As Integer
    
    ' PmtInfId
    Set Node = PmtInfNode.getElementsByTagName(TAG_PMTINFID)(0)
    If Not Node Is Nothing Then PmtInf.PmtInfId = Node.Text
    
    ' PmtMtd
    Set Node = PmtInfNode.getElementsByTagName(TAG_PMTMTD)(0)
    If Not Node Is Nothing Then PmtInf.PmtMtd = Node.Text
    
    ' BtchBookg
    Set Node = PmtInfNode.getElementsByTagName(TAG_BtchBookg)(0)
    If Not Node Is Nothing Then
        PmtInf.BtchBookg = (LCase(Node.Text) = XML_TRUE)
    End If
    
    ' ReqdExctnDt
    Set Node = PmtInfNode.getElementsByTagName(TAG_REQDEXCTNDT)(0)
    If Not Node Is Nothing Then
        PmtInf.ReqdExctnDt = ParseDateFromXML(Node.Text)
    End If
    
    ' ChrgBr
    Set Node = PmtInfNode.getElementsByTagName(TAG_CHRGBR)(0)
    If Not Node Is Nothing Then PmtInf.ChrgBr = Node.Text
    
    ' Dbtr
    Set DbtrNode = PmtInfNode.getElementsByTagName(TAG_DBTR)(0)
    If Not DbtrNode Is Nothing Then
        ParseParty DbtrNode, PmtInf.Dbtr
    End If
    
    ' DbtrAcct - IBAN
    Set DbtrAcctNode = PmtInfNode.getElementsByTagName(TAG_DBTRACCT)(0)
    If Not DbtrAcctNode Is Nothing Then
        Set Node = DbtrAcctNode.getElementsByTagName(TAG_IBAN)(0)
        If Not Node Is Nothing Then
            PmtInf.DbtrAcct_IBAN = Node.Text
        End If
    End If
    
    ' DbtrAgt - BIC
    Set DbtrAgtNode = PmtInfNode.getElementsByTagName(TAG_DBTRAGT)(0)
    If Not DbtrAgtNode Is Nothing Then
        Set Node = DbtrAgtNode.getElementsByTagName(TAG_BIC)(0)
        If Not Node Is Nothing Then
            PmtInf.DbtrAgt_BIC = Node.Text
        End If
    End If
    
    ' Transactions (CdtTrfTxInf)
    Set TxNodes = PmtInfNode.getElementsByTagName(TAG_CDTTRFTXINF)
    If TxNodes.Length > 0 Then
        ReDim PmtInf.Transactions(0 To TxNodes.Length - 1)
        
        For i = 0 To TxNodes.Length - 1
            ParseTransaction TxNodes(i), PmtInf.Transactions(i)
            PmtInf.Transactions(i).RowNumber = ROW_TRANSACTION_START + i
        Next i
        
        LogInfo "Transactions parsées: " & TxNodes.Length, "modXMLImport.ParsePaymentInformation"
    End If
    
    On Error GoTo 0
End Sub

' =============================================================================
' SUB: ParseTransaction
' DESCRIPTION: Parse un nœud CdtTrfTxInf (transaction individuelle)
' =============================================================================
Private Sub ParseTransaction(ByRef TxNode As Object, ByRef Transaction As TTransaction)
    On Error Resume Next
    
    Dim Node As Object
    Dim PmtIdNode As Object
    Dim AmtNode As Object
    Dim CdtrNode As Object
    Dim CdtrAgtNode As Object
    Dim CdtrAcctNode As Object
    Dim RmtInfNode As Object
    
    ' PmtId
    Set PmtIdNode = TxNode.getElementsByTagName(TAG_PMTID)(0)
    If Not PmtIdNode Is Nothing Then
        Set Node = PmtIdNode.getElementsByTagName(TAG_ENDTOENDID)(0)
        If Not Node Is Nothing Then
            Transaction.PmtId.EndToEndId = Node.Text
        End If
        
        Set Node = PmtIdNode.getElementsByTagName(TAG_INSTRID)(0)
        If Not Node Is Nothing Then
            Transaction.PmtId.InstrId = Node.Text
        End If
    End If
    
    ' Amt - InstdAmt
    Set AmtNode = TxNode.getElementsByTagName(TAG_AMT)(0)
    If Not AmtNode Is Nothing Then
        Set Node = AmtNode.getElementsByTagName(TAG_INSTDamT)(0)
        If Not Node Is Nothing Then
            Transaction.Amt.Amount = ParseAmountFromXML(Node.Text)
            
            ' Attribut Ccy (devise)
            If Node.Attributes.length > 0 Then
                Dim Attr As Object
                For Each Attr In Node.Attributes
                    If Attr.Name = ATTR_CCY Then
                        Transaction.Amt.Ccy = Attr.Text
                        Exit For
                    End If
                Next Attr
            End If
        End If
    End If
    
    ' CdtrAgt (Banque bénéficiaire)
    Set CdtrAgtNode = TxNode.getElementsByTagName(TAG_CDTRAGT)(0)
    If Not CdtrAgtNode Is Nothing Then
        ParseFinancialInstitution CdtrAgtNode, Transaction.CdtrAgt
    End If
    
    ' Cdtr (Bénéficiaire)
    Set CdtrNode = TxNode.getElementsByTagName(TAG_CDTR)(0)
    If Not CdtrNode Is Nothing Then
        ParseParty CdtrNode, Transaction.Cdtr
    End If
    
    ' CdtrAcct - IBAN
    Set CdtrAcctNode = TxNode.getElementsByTagName(TAG_CDTRACCT)(0)
    If Not CdtrAcctNode Is Nothing Then
        Set Node = CdtrAcctNode.getElementsByTagName(TAG_IBAN)(0)
        If Not Node Is Nothing Then
            Transaction.CdtrAcct_IBAN = Node.Text
        End If
    End If
    
    ' RmtInf - Ustrd
    Set RmtInfNode = TxNode.getElementsByTagName(TAG_RMTINF)(0)
    If Not RmtInfNode Is Nothing Then
        Set Node = RmtInfNode.getElementsByTagName(TAG_USTRD)(0)
        If Not Node Is Nothing Then
            Transaction.RmtInf_Ustrd = Node.Text
        End If
    End If
    
    On Error GoTo 0
End Sub

' =============================================================================
' SUB: ParseParty
' DESCRIPTION: Parse un nœud de partie (InitgPty, Dbtr, Cdtr)
' =============================================================================
Private Sub ParseParty(ByRef PartyNode As Object, ByRef Party As TParty)
    On Error Resume Next
    
    Dim Node As Object
    Dim AddrNode As Object
    
    ' Nm
    Set Node = PartyNode.getElementsByTagName(TAG_NM)(0)
    If Not Node Is Nothing Then Party.Nm = Node.Text
    
    ' PstlAdr
    Set AddrNode = PartyNode.getElementsByTagName(TAG_PSTLADR)(0)
    If Not AddrNode Is Nothing Then
        Set Node = AddrNode.getElementsByTagName(TAG_STRTNM)(0)
        If Not Node Is Nothing Then Party.Address.StrtNm = Node.Text
        
        Set Node = AddrNode.getElementsByTagName(TAG_PSTCD)(0)
        If Not Node Is Nothing Then Party.Address.PstCd = Node.Text
        
        Set Node = AddrNode.getElementsByTagName(TAG_TWNNM)(0)
        If Not Node Is Nothing Then Party.Address.TwnNm = Node.Text
        
        Set Node = AddrNode.getElementsByTagName(TAG_CTRY)(0)
        If Not Node Is Nothing Then Party.Address.Ctry = Node.Text
    End If
    
    On Error GoTo 0
End Sub

' =============================================================================
' SUB: ParseFinancialInstitution
' DESCRIPTION: Parse un nœud FinInstnId (institution financière)
' =============================================================================
Private Sub ParseFinancialInstitution(ByRef FinInstnNode As Object, ByRef FinInstn As TFinancialInstitution)
    On Error Resume Next
    
    Dim Node As Object
    
    ' BIC
    Set Node = FinInstnNode.getElementsByTagName(TAG_BIC)(0)
    If Not Node Is Nothing Then FinInstn.BIC = Node.Text
    
    ' Nm
    Set Node = FinInstnNode.getElementsByTagName(TAG_NM)(0)
    If Not Node Is Nothing Then FinInstn.Nm = Node.Text
    
    ' PstlAdr
    Dim AddrNode As Object
    Set AddrNode = FinInstnNode.getElementsByTagName(TAG_PSTLADR)(0)
    If Not AddrNode Is Nothing Then
        Set Node = AddrNode.getElementsByTagName(TAG_STRTNM)(0)
        If Not Node Is Nothing Then FinInstn.Address.StrtNm = Node.Text
        
        Set Node = AddrNode.getElementsByTagName(TAG_PSTCD)(0)
        If Not Node Is Nothing Then FinInstn.Address.PstCd = Node.Text
        
        Set Node = AddrNode.getElementsByTagName(TAG_TWNNM)(0)
        If Not Node Is Nothing Then FinInstn.Address.TwnNm = Node.Text
        
        Set Node = AddrNode.getElementsByTagName(TAG_CTRY)(0)
        If Not Node Is Nothing Then FinInstn.Address.Ctry = Node.Text
    End If
    
    On Error GoTo 0
End Sub

' =============================================================================
' FONCTION: GetNamespaceURI
' DESCRIPTION: Extrait l'URI d'un namespace depuis un nœud
' =============================================================================
Private Function GetNamespaceURI(ByRef Node As Object, ByVal NamespaceURI As String) As String
    On Error Resume Next
    
    Dim Attr As Object
    Dim NS As String
    
    NS = ""
    
    ' Chercher dans les attributs xmlns
    For Each Attr In Node.Attributes
        If InStr(Attr.Name, "xmlns") > 0 Then
            If Attr.Text = NamespaceURI Or Attr.Name = "xmlns" Then
                NS = Attr.Text
                Exit For
            End If
        End If
    Next Attr
    
    ' Si non trouvé, utiliser la valeur par défaut
    If NS = "" Then
        NS = NamespaceURI
    End If
    
    GetNamespaceURI = NS
    
    On Error GoTo 0
End Function

' =============================================================================
' FONCTION: SetupPrincipalSheet
' DESCRIPTION: Configure la feuille Principal avec la mise en forme
' =============================================================================
Public Function SetupPrincipalSheet() As Worksheet
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim LastRow As Long
    
    ' Supprimer la feuille si elle existe déjà et la recréer
    Application.DisplayAlerts = False
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_PRINCIPAL)
    If Not ws Is Nothing Then
        ws.Delete
    End If
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = SHEET_PRINCIPAL
    Application.DisplayAlerts = True
    
    ' Configuration de la feuille
    With ws
        .Cells.Clear
        .Range("A1:Z1").Font.Bold = True
        .Range("A1:Z1").Interior.Color = RGB(79, 129, 189)
        .Range("A1:Z1").Font.Color = RGB(255, 255, 255)
        
        ' Titres des sections
        .Range("A1").Value = "SEPA XML Manager - pain.001.001.03"
        .Range("A1:C1").Merge
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Font.Size = 14
        
        ' Section ENTÊTE
        .Range("A2").Value = "MsgId:"
        .Range("A3").Value = "CreDtTm:"
        .Range("A4").Value = "NbOfTxs:"
        .Range("A5").Value = "CtrlSum:"
        
        .Range("A7").Value = "--- Initiating Party ---"
        .Range("A7:C7").Merge
        .Range("A7").Font.Bold = True
        .Range("A7").Interior.Color = RGB(200, 200, 200)
        
        .Range("A8").Value = "Nom:"
        .Range("A9").Value = "Rue:"
        .Range("A10").Value = "Code Postal:"
        .Range("A11").Value = "Ville:"
        .Range("A12").Value = "Pays:"
        
        .Range("A14").Value = "--- Debtor (Donneur d'ordre) ---"
        .Range("A14:C14").Merge
        .Range("A14").Font.Bold = True
        .Range("A14").Interior.Color = RGB(200, 200, 200)
        
        .Range("A15").Value = "Nom:"
        .Range("A16").Value = "Rue:"
        .Range("A17").Value = "Code Postal:"
        .Range("A18").Value = "Ville:"
        .Range("A19").Value = "Pays:"
        .Range("A20").Value = "IBAN:"
        .Range("A21").Value = "BIC:"
        
        .Range("A23").Value = "--- Payment Information ---"
        .Range("A23:C23").Merge
        .Range("A23").Font.Bold = True
        .Range("A23").Interior.Color = RGB(200, 200, 200)
        
        .Range("A24").Value = "PmtInfId:"
        .Range("A25").Value = "PmtMtd:"
        .Range("A26").Value = "BtchBookg:"
        .Range("A27").Value = "ReqdExctnDt:"
        .Range("A28").Value = "ChrgBr:"
        .Range("A29").Value = "Devise:"
        
        ' En-têtes de transactions
        .Cells(ROW_TRANSACTION_START - 1, COL_TX_ENDTOENDID).Value = HEADER_ENDTOENDID
        .Cells(ROW_TRANSACTION_START - 1, COL_TX_INSTRID).Value = HEADER_INSTRID
        .Cells(ROW_TRANSACTION_START - 1, COL_TX_CDTR_NM).Value = HEADER_CDTR_NM
        .Cells(ROW_TRANSACTION_START - 1, COL_TX_CDTR_STRTNM).Value = HEADER_CDTR_STRTNM
        .Cells(ROW_TRANSACTION_START - 1, COL_TX_CDTR_PSTCD).Value = HEADER_CDTR_PSTCD
        .Cells(ROW_TRANSACTION_START - 1, COL_TX_CDTR_TWNNM).Value = HEADER_CDTR_TWNNM
        .Cells(ROW_TRANSACTION_START - 1, COL_TX_CDTR_CTRY).Value = HEADER_CDTR_CTRY
        .Cells(ROW_TRANSACTION_START - 1, COL_TX_CDTR_IBAN).Value = HEADER_CDTR_IBAN
        .Cells(ROW_TRANSACTION_START - 1, COL_TX_CDTR_BIC).Value = HEADER_CDTR_BIC
        .Cells(ROW_TRANSACTION_START - 1, COL_TX_CDTR_BANKNM).Value = HEADER_CDTR_BANKNM
        .Cells(ROW_TRANSACTION_START - 1, COL_TX_AMT).Value = HEADER_AMT
        .Cells(ROW_TRANSACTION_START - 1, COL_TX_CCY).Value = HEADER_CCY
        .Cells(ROW_TRANSACTION_START - 1, COL_TX_USTRD).Value = HEADER_USTRD
        
        ' Formater la ligne d'en-tête des transactions
        With .Range(.Cells(ROW_TRANSACTION_START - 1, COL_TX_ENDTOENDID), _
                    .Cells(ROW_TRANSACTION_START - 1, COL_TX_USTRD))
            .Font.Bold = True
            .Interior.Color = RGB(79, 129, 189)
            .Font.Color = RGB(255, 255, 255)
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        
        ' Ajuster les largeurs de colonnes
        .Columns("A:A").ColumnWidth = 18
        .Columns("B:B").ColumnWidth = 35
        .Columns("C:C").ColumnWidth = 15
        .Columns("D:D").ColumnWidth = 25
        .Columns("E:E").ColumnWidth = 25
        .Columns("F:F").ColumnWidth = 12
        .Columns("G:G").ColumnWidth = 15
        .Columns("H:H").ColumnWidth = 8
        .Columns("I:I").ColumnWidth = 35
        .Columns("J:J").ColumnWidth = 15
        .Columns("K:K").ColumnWidth = 20
        .Columns("L:L").ColumnWidth = 12
        .Columns("M:M").ColumnWidth = 8
        .Columns("N:N").ColumnWidth = 30
        
        ' Figer les volets
        .Range("A" & ROW_TRANSACTION_START).Select
        ActiveWindow.FreezePanes = True
    End With
    
    Set SetupPrincipalSheet = ws
    
    On Error GoTo 0
End Function

' =============================================================================
' SUB: WriteDataToExcel
' DESCRIPTION: Écrit les données du document SEPA dans Excel
' =============================================================================
Private Sub WriteDataToExcel(ByRef ws As Worksheet, ByRef SEPADoc As TSEPAXMLDocument)
    On Error Resume Next
    
    Dim i As Long
    
    ' Group Header
    ws.Range(CELL_MSGID).Value = SEPADoc.GrpHdr.MsgId
    ws.Range(CELL_CREDTTM).Value = SEPADoc.GrpHdr.CreDtTm
    ws.Range(CELL_NBOFTXS).Value = SEPADoc.GrpHdr.NbOfTxs
    ws.Range(CELL_CTRLSUM).Value = SEPADoc.GrpHdr.CtrlSum
    
    ' Initiating Party
    ws.Range(CELL_INITGPTY_NM).Value = SEPADoc.GrpHdr.InitgPty.Nm
    ws.Range(CELL_INITGPTY_STRTNM).Value = SEPADoc.GrpHdr.InitgPty.Address.StrtNm
    ws.Range(CELL_INITGPTY_PSTCD).Value = SEPADoc.GrpHdr.InitgPty.Address.PstCd
    ws.Range(CELL_INITGPTY_TWNNM).Value = SEPADoc.GrpHdr.InitgPty.Address.TwnNm
    ws.Range(CELL_INITGPTY_CTRY).Value = SEPADoc.GrpHdr.InitgPty.Address.Ctry
    
    ' Debtor
    ws.Range(CELL_DBTR_NM).Value = SEPADoc.PmtInf.Dbtr.Nm
    ws.Range(CELL_DBTR_STRTNM).Value = SEPADoc.PmtInf.Dbtr.Address.StrtNm
    ws.Range(CELL_DBTR_PSTCD).Value = SEPADoc.PmtInf.Dbtr.Address.PstCd
    ws.Range(CELL_DBTR_TWNNM).Value = SEPADoc.PmtInf.Dbtr.Address.TwnNm
    ws.Range(CELL_DBTR_CTRY).Value = SEPADoc.PmtInf.Dbtr.Address.Ctry
    ws.Range(CELL_DBTR_IBAN).Value = SEPADoc.PmtInf.DbtrAcct_IBAN
    ws.Range(CELL_DBTR_BIC).Value = SEPADoc.PmtInf.DbtrAgt_BIC
    
    ' Payment Information
    ws.Range(CELL_PMTINFID).Value = SEPADoc.PmtInf.PmtInfId
    ws.Range(CELL_PMTMTD).Value = SEPADoc.PmtInf.PmtMtd
    ws.Range(CELL_BTCHBOOKG).Value = IIf(SEPADoc.PmtInf.BtchBookg, XML_TRUE, XML_FALSE)
    ws.Range(CELL_REQDEXCTNDT).Value = Format(SEPADoc.PmtInf.ReqdExctnDt, FORMAT_DATE_SHORT)
    ws.Range(CELL_CHRGBR).Value = SEPADoc.PmtInf.ChrgBr
    
    ' Devise (depuis la première transaction)
    If UBound(SEPADoc.PmtInf.Transactions) >= 0 Then
        ws.Range(CELL_DEVISE).Value = SEPADoc.PmtInf.Transactions(0).Amt.Ccy
    End If
    
    ' Transactions
    For i = LBound(SEPADoc.PmtInf.Transactions) To UBound(SEPADoc.PmtInf.Transactions)
        With SEPADoc.PmtInf.Transactions(i)
            ws.Cells(ROW_TRANSACTION_START + i, COL_TX_ENDTOENDID).Value = .PmtId.EndToEndId
            ws.Cells(ROW_TRANSACTION_START + i, COL_TX_INSTRID).Value = .PmtId.InstrId
            ws.Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_NM).Value = .Cdtr.Nm
            ws.Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_STRTNM).Value = .Cdtr.Address.StrtNm
            ws.Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_PSTCD).Value = .Cdtr.Address.PstCd
            ws.Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_TWNNM).Value = .Cdtr.Address.TwnNm
            ws.Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_CTRY).Value = .Cdtr.Address.Ctry
            ws.Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_IBAN).Value = .CdtrAcct_IBAN
            ws.Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_BIC).Value = .CdtrAgt.BIC
            ws.Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_BANKNM).Value = .CdtrAgt.Nm
            ws.Cells(ROW_TRANSACTION_START + i, COL_TX_AMT).Value = .Amt.Amount
            ws.Cells(ROW_TRANSACTION_START + i, COL_TX_CCY).Value = .Amt.Ccy
            ws.Cells(ROW_TRANSACTION_START + i, COL_TX_USTRD).Value = .RmtInf_Ustrd
        End With
    Next i
    
    LogInfo "Données écrites dans Excel: " & (UBound(SEPADoc.PmtInf.Transactions) + 1) & " transactions", _
            "modXMLImport.WriteDataToExcel"
    
    On Error GoTo 0
End Sub
