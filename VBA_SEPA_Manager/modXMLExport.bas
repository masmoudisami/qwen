'===============================================================================
' MODULE: modXMLExport.bas
' DESCRIPTION: Export du fichier XML SEPA conforme au modèle (pain.001.001.03)
'===============================================================================
Option Explicit

' =============================================================================
' FONCTION: ExportXMLFile
' DESCRIPTION: Exporte les données Excel vers un fichier XML SEPA
' =============================================================================
Public Function ExportXMLFile(Optional ByVal FilePath As String = "", _
                              Optional ByVal SaveAsTXT As Boolean = True) As Boolean
    On Error GoTo ErrorHandler
    
    Dim xmlDoc As Object
    Dim Success As Boolean
    Dim SEPADoc As TSEPAXMLDocument
    
    LogInfo "Début de l'export XML", "modXMLExport.ExportXMLFile"
    
    ' Valider les données avant export
    If Not ValidateAllData() Then
        LogError ERR_EXPORT_CREATE, "Validation des données échouée", _
                 "modXMLExport.ExportXMLFile", , , True
        ExportXMLFile = False
        Exit Function
    End If
    
    ' Lire les données depuis Excel
    Success = ReadDataFromExcel(SEPADoc)
    If Not Success Then
        LogError ERR_EXPORT_CREATE, "Lecture des données Excel échouée", _
                 "modXMLExport.ExportXMLFile", , , True
        ExportXMLFile = False
        Exit Function
    End If
    
    ' Créer le document XML
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    
    With xmlDoc
        .async = False
        .validateOnParse = False
        .preserveWhiteSpace = False
        .resolveExternals = False
    End With
    
    ' Générer la structure XML
    Success = BuildXMLStructure(xmlDoc, SEPADoc)
    If Not Success Then
        LogError ERR_EXPORT_WRITE, "Construction de la structure XML échouée", _
                 "modXMLExport.ExportXMLFile", , , True
        ExportXMLFile = False
        Exit Function
    End If
    
    ' Si aucun chemin fourni, demander à l'utilisateur
    If FilePath = "" Then
        FilePath = Application.GetSaveAsFilename( _
            InitialFileName:="SEPA_" & Format(Now, "yyyymmdd_hhmmss") & ".xml", _
            FileFilter:="XML Files (*.xml), *.xml")
        
        If FilePath = "False" Then
            ExportXMLFile = False
            Exit Function
        End If
    End If
    
    ' Sauvegarder le fichier XML
    If Not xmlDoc.Save(FilePath) Then
        LogError ERR_EXPORT_SAVE, "Impossible de sauvegarder le fichier XML", _
                 "modXMLExport.ExportXMLFile", , , True
        ExportXMLFile = False
        Exit Function
    End If
    
    LogInfo "Fichier XML sauvegardé: " & FilePath, "modXMLExport.ExportXMLFile"
    
    ' Exporter en TXT UTF-8 si demandé
    If SaveAsTXT Then
        Call ExportAsUTF8Text(FilePath)
    End If
    
    ExportXMLFile = True
    
    Exit Function
    
ErrorHandler:
    LogError ERR_EXPORT_CREATE, "Erreur critique lors de l'export: " & Err.Description, _
             "modXMLExport.ExportXMLFile", , , True
    ExportXMLFile = False
End Function

' =============================================================================
' FONCTION: ReadDataFromExcel
' DESCRIPTION: Lit les données depuis la feuille Excel
' =============================================================================
Private Function ReadDataFromExcel(ByRef SEPADoc As TSEPAXMLDocument) As Boolean
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim i As Long, TxCount As Long
    Dim LastRow As Long
    
    Set ws = ThisWorkbook.Worksheets(SHEET_PRINCIPAL)
    If ws Is Nothing Then
        ReadDataFromExcel = False
        Exit Function
    End If
    
    ' Group Header
    SEPADoc.GrpHdr.MsgId = SafeCStr(ws.Range(CELL_MSGID).Value)
    SEPADoc.GrpHdr.CreDtTm = SafeCDate(ws.Range(CELL_CREDTTM).Value)
    SEPADoc.GrpHdr.NbOfTxs = SafeCLng(ws.Range(CELL_NBOFTXS).Value)
    SEPADoc.GrpHdr.CtrlSum = SafeCDbl(ws.Range(CELL_CTRLSUM).Value)
    
    ' Initiating Party
    SEPADoc.GrpHdr.InitgPty.Nm = SafeCStr(ws.Range(CELL_INITGPTY_NM).Value)
    SEPADoc.GrpHdr.InitgPty.Address.StrtNm = SafeCStr(ws.Range(CELL_INITGPTY_STRTNM).Value)
    SEPADoc.GrpHdr.InitgPty.Address.PstCd = SafeCStr(ws.Range(CELL_INITGPTY_PSTCD).Value)
    SEPADoc.GrpHdr.InitgPty.Address.TwnNm = SafeCStr(ws.Range(CELL_INITGPTY_TWNNM).Value)
    SEPADoc.GrpHdr.InitgPty.Address.Ctry = SafeCStr(ws.Range(CELL_INITGPTY_CTRY).Value)
    
    ' Payment Information
    SEPADoc.PmtInf.PmtInfId = SafeCStr(ws.Range(CELL_PMTINFID).Value)
    SEPADoc.PmtInf.PmtMtd = SafeCStr(ws.Range(CELL_PMTMTD).Value)
    SEPADoc.PmtInf.BtchBookg = (LCase(SafeCStr(ws.Range(CELL_BTCHBOOKG).Value)) = XML_TRUE)
    SEPADoc.PmtInf.ReqdExctnDt = SafeCDate(ws.Range(CELL_REQDEXCTNDT).Value)
    SEPADoc.PmtInf.ChrgBr = SafeCStr(ws.Range(CELL_CHRGBR).Value)
    
    ' Debtor
    SEPADoc.PmtInf.Dbtr.Nm = SafeCStr(ws.Range(CELL_DBTR_NM).Value)
    SEPADoc.PmtInf.Dbtr.Address.StrtNm = SafeCStr(ws.Range(CELL_DBTR_STRTNM).Value)
    SEPADoc.PmtInf.Dbtr.Address.PstCd = SafeCStr(ws.Range(CELL_DBTR_PSTCD).Value)
    SEPADoc.PmtInf.Dbtr.Address.TwnNm = SafeCStr(ws.Range(CELL_DBTR_TWNNM).Value)
    SEPADoc.PmtInf.Dbtr.Address.Ctry = SafeCStr(ws.Range(CELL_DBTR_CTRY).Value)
    SEPADoc.PmtInf.DbtrAcct_IBAN = SafeCStr(ws.Range(CELL_DBTR_IBAN).Value)
    SEPADoc.PmtInf.DbtrAgt_BIC = SafeCStr(ws.Range(CELL_DBTR_BIC).Value)
    
    ' Compter les transactions
    LastRow = ws.Cells(ws.Rows.Count, COL_TX_ENDTOENDID).End(xlUp).Row
    If LastRow < ROW_TRANSACTION_START Then
        TxCount = 0
    Else
        TxCount = LastRow - ROW_TRANSACTION_START + 1
    End If
    
    ' Transactions
    If TxCount > 0 Then
        ReDim SEPADoc.PmtInf.Transactions(0 To TxCount - 1)
        
        For i = 0 To TxCount - 1
            With SEPADoc.PmtInf.Transactions(i)
                .PmtId.EndToEndId = SafeCStr(ws.Cells(ROW_TRANSACTION_START + i, COL_TX_ENDTOENDID).Value)
                .PmtId.InstrId = SafeCStr(ws.Cells(ROW_TRANSACTION_START + i, COL_TX_INSTRID).Value)
                
                .Cdtr.Nm = SafeCStr(ws.Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_NM).Value)
                .Cdtr.Address.StrtNm = SafeCStr(ws.Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_STRTNM).Value)
                .Cdtr.Address.PstCd = SafeCStr(ws.Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_PSTCD).Value)
                .Cdtr.Address.TwnNm = SafeCStr(ws.Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_TWNNM).Value)
                .Cdtr.Address.Ctry = SafeCStr(ws.Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_CTRY).Value)
                
                .CdtrAcct_IBAN = SafeCStr(ws.Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_IBAN).Value)
                
                .CdtrAgt.BIC = SafeCStr(ws.Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_BIC).Value)
                .CdtrAgt.Nm = SafeCStr(ws.Cells(ROW_TRANSACTION_START + i, COL_TX_CDTR_BANKNM).Value)
                
                .Amt.Amount = SafeCDbl(ws.Cells(ROW_TRANSACTION_START + i, COL_TX_AMT).Value)
                .Amt.Ccy = SafeCStr(ws.Cells(ROW_TRANSACTION_START + i, COL_TX_CCY).Value)
                
                .RmtInf_Ustrd = SafeCStr(ws.Cells(ROW_TRANSACTION_START + i, COL_TX_USTRD).Value)
            End With
        Next i
        
        ' Mettre à jour NbOfTxs et CtrlSum
        SEPADoc.GrpHdr.NbOfTxs = TxCount
        SEPADoc.GrpHdr.CtrlSum = CalculateCtrlSum(SEPADoc.PmtInf.Transactions)
    End If
    
    ' Namespaces par défaut (selon le modèle)
    SEPADoc.NamespaceURI = NS_ISO
    SEPADoc.NamespaceXSI = NS_XSI
    
    ReadDataFromExcel = True
    
    On Error GoTo 0
End Function

' =============================================================================
' FONCTION: BuildXMLStructure
' DESCRIPTION: Construit la structure XML complète
' =============================================================================
Private Function BuildXMLStructure(ByRef xmlDoc As Object, ByRef SEPADoc As TSEPAXMLDocument) As Boolean
    On Error GoTo ErrorHandler
    
    Dim DocNode As Object
    Dim CstmrNode As Object
    Dim GrpHdrNode As Object
    Dim PmtInfNode As Object
    Dim i As Long
    
    ' Déclaration XML
    Dim xmlDecl As Object
    Set xmlDecl = xmlDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
    xmlDoc.appendChild xmlDecl
    
    ' Document root avec namespace
    Set DocNode = xmlDoc.createElement(TAG_DOCUMENT)
    DocNode.setAttribute "xmlns", SEPADoc.NamespaceURI
    DocNode.setAttribute "xmlns:xsi", SEPADoc.NamespaceXSI
    xmlDoc.appendChild DocNode
    
    ' CstmrCdtTrfInitn
    Set CstmrNode = xmlDoc.createElement(TAG_CSTMRCDTTRFINITN)
    DocNode.appendChild CstmrNode
    
    ' GrpHdr
    Set GrpHdrNode = CreateGroupHeaderNode(xmlDoc, SEPADoc.GrpHdr)
    CstmrNode.appendChild GrpHdrNode
    
    ' PmtInf
    Set PmtInfNode = CreatePaymentInformationNode(xmlDoc, SEPADoc.PmtInf)
    CstmrNode.appendChild PmtInfNode
    
    BuildXMLStructure = True
    
    Exit Function
    
ErrorHandler:
    LogError ERR_EXPORT_WRITE, "Erreur de construction XML: " & Err.Description, _
             "modXMLExport.BuildXMLStructure", , , True
    BuildXMLStructure = False
End Function

' =============================================================================
' FONCTION: CreateGroupHeaderNode
' DESCRIPTION: Crée le nœud GrpHdr
' =============================================================================
Private Function CreateGroupHeaderNode(ByRef xmlDoc As Object, ByRef GrpHdr As TGroupHeader) As Object
    On Error Resume Next
    
    Dim GrpHdrNode As Object
    Dim Node As Object
    
    Set GrpHdrNode = xmlDoc.createElement(TAG_GRPHDR)
    
    ' MsgId
    Set Node = xmlDoc.createElement(TAG_MSGID)
    Node.Text = SanitizeStringForXML(GrpHdr.MsgId)
    GrpHdrNode.appendChild Node
    
    ' CreDtTm
    Set Node = xmlDoc.createElement(TAG_CREDTTM)
    Node.Text = FormatDateForXML(GrpHdr.CreDtTm, True)
    GrpHdrNode.appendChild Node
    
    ' NbOfTxs
    Set Node = xmlDoc.createElement(TAG_NBOFTXS)
    Node.Text = CStr(GrpHdr.NbOfTxs)
    GrpHdrNode.appendChild Node
    
    ' CtrlSum
    Set Node = xmlDoc.createElement(TAG_CTRLSUM)
    Node.Text = FormatAmountForXML(GrpHdr.CtrlSum)
    GrpHdrNode.appendChild Node
    
    ' InitgPty
    GrpHdrNode.appendChild CreatePartyNode(xmlDoc, TAG_INITGPTY, GrpHdr.InitgPty)
    
    Set CreateGroupHeaderNode = GrpHdrNode
    
    On Error GoTo 0
End Function

' =============================================================================
' FONCTION: CreatePaymentInformationNode
' DESCRIPTION: Crée le nœud PmtInf
' =============================================================================
Private Function CreatePaymentInformationNode(ByRef xmlDoc As Object, ByRef PmtInf As TPaymentInformation) As Object
    On Error Resume Next
    
    Dim PmtInfNode As Object
    Dim Node As Object
    Dim i As Long
    
    Set PmtInfNode = xmlDoc.createElement(TAG_PMTINF)
    
    ' PmtInfId
    Set Node = xmlDoc.createElement(TAG_PMTINFID)
    Node.Text = SanitizeStringForXML(PmtInf.PmtInfId)
    PmtInfNode.appendChild Node
    
    ' PmtMtd
    Set Node = xmlDoc.createElement(TAG_PMTMTD)
    Node.Text = PmtInf.PmtMtd
    PmtInfNode.appendChild Node
    
    ' BtchBookg
    Set Node = xmlDoc.createElement(TAG_BtchBookg)
    Node.Text = IIf(PmtInf.BtchBookg, XML_TRUE, XML_FALSE)
    PmtInfNode.appendChild Node
    
    ' ReqdExctnDt
    Set Node = xmlDoc.createElement(TAG_REQDEXCTNDT)
    Node.Text = FormatDateForXML(PmtInf.ReqdExctnDt, False)
    PmtInfNode.appendChild Node
    
    ' Dbtr
    PmtInfNode.appendChild CreatePartyNode(xmlDoc, TAG_DBTR, PmtInf.Dbtr)
    
    ' DbtrAcct
    Set Node = xmlDoc.createElement(TAG_DBTRACCT)
    Dim IdNode As Object
    Set IdNode = xmlDoc.createElement(TAG_ID)
    Dim IbanNode As Object
    Set IbanNode = xmlDoc.createElement(TAG_IBAN)
    IbanNode.Text = PmtInf.DbtrAcct_IBAN
    IdNode.appendChild IbanNode
    Node.appendChild IdNode
    PmtInfNode.appendChild Node
    
    ' DbtrAgt
    Set Node = xmlDoc.createElement(TAG_DBTRAGT)
    Dim FinInstnNode As Object
    Set FinInstnNode = xmlDoc.createElement(TAG_FININSTNID)
    Dim BicNode As Object
    Set BicNode = xmlDoc.createElement(TAG_BIC)
    BicNode.Text = PmtInf.DbtrAgt_BIC
    FinInstnNode.appendChild BicNode
    Node.appendChild FinInstnNode
    PmtInfNode.appendChild Node
    
    ' ChrgBr
    Set Node = xmlDoc.createElement(TAG_CHRGBR)
    Node.Text = PmtInf.ChrgBr
    PmtInfNode.appendChild Node
    
    ' Transactions (CdtTrfTxInf)
    For i = LBound(PmtInf.Transactions) To UBound(PmtInf.Transactions)
        PmtInfNode.appendChild CreateTransactionNode(xmlDoc, PmtInf.Transactions(i))
    Next i
    
    Set CreatePaymentInformationNode = PmtInfNode
    
    On Error GoTo 0
End Function

' =============================================================================
' FONCTION: CreateTransactionNode
' DESCRIPTION: Crée un nœud CdtTrfTxInf (transaction)
' =============================================================================
Private Function CreateTransactionNode(ByRef xmlDoc As Object, ByRef Transaction As TTransaction) As Object
    On Error Resume Next
    
    Dim TxNode As Object
    Dim Node As Object
    Dim SubNode As Object
    
    Set TxNode = xmlDoc.createElement(TAG_CDTTRFTXINF)
    
    ' PmtId
    Set Node = xmlDoc.createElement(TAG_PMTID)
    
    ' EndToEndId
    Set SubNode = xmlDoc.createElement(TAG_ENDTOENDID)
    SubNode.Text = SanitizeStringForXML(Transaction.PmtId.EndToEndId)
    Node.appendChild SubNode
    
    ' InstrId (optionnel)
    If Len(Trim(Transaction.PmtId.InstrId)) > 0 Then
        Set SubNode = xmlDoc.createElement(TAG_INSTRID)
        SubNode.Text = SanitizeStringForXML(Transaction.PmtId.InstrId)
        Node.appendChild SubNode
    End If
    
    TxNode.appendChild Node
    
    ' Amt - InstdAmt
    Set Node = xmlDoc.createElement(TAG_AMT)
    Set SubNode = xmlDoc.createElement(TAG_INSTDamT)
    SubNode.Text = FormatAmountForXML(Transaction.Amt.Amount)
    SubNode.setAttribute ATTR_CCY, Transaction.Amt.Ccy
    Node.appendChild SubNode
    TxNode.appendChild Node
    
    ' CdtrAgt
    TxNode.appendChild CreateFinancialInstitutionNode(xmlDoc, TAG_CDTRAGT, Transaction.CdtrAgt)
    
    ' Cdtr
    TxNode.appendChild CreatePartyNode(xmlDoc, TAG_CDTR, Transaction.Cdtr)
    
    ' CdtrAcct
    Set Node = xmlDoc.createElement(TAG_CDTRACCT)
    Set SubNode = xmlDoc.createElement(TAG_ID)
    Dim IbanNode As Object
    Set IbanNode = xmlDoc.createElement(TAG_IBAN)
    IbanNode.Text = Transaction.CdtrAcct_IBAN
    SubNode.appendChild IbanNode
    Node.appendChild SubNode
    TxNode.appendChild Node
    
    ' RmtInf - Ustrd
    Set Node = xmlDoc.createElement(TAG_RMTINF)
    Set SubNode = xmlDoc.createElement(TAG_USTRD)
    SubNode.Text = SanitizeStringForXML(Transaction.RmtInf_Ustrd)
    Node.appendChild SubNode
    TxNode.appendChild Node
    
    Set CreateTransactionNode = TxNode
    
    On Error GoTo 0
End Function

' =============================================================================
' FONCTION: CreatePartyNode
' DESCRIPTION: Crée un nœud de partie (InitgPty, Dbtr, Cdtr)
' =============================================================================
Private Function CreatePartyNode(ByRef xmlDoc As Object, ByVal TagName As String, ByRef Party As TParty) As Object
    On Error Resume Next
    
    Dim PartyNode As Object
    Dim Node As Object
    
    Set PartyNode = xmlDoc.createElement(TagName)
    
    ' Nm
    Set Node = xmlDoc.createElement(TAG_NM)
    Node.Text = SanitizeStringForXML(Party.Nm)
    PartyNode.appendChild Node
    
    ' PstlAdr
    PartyNode.appendChild CreateAddressNode(xmlDoc, Party.Address)
    
    Set CreatePartyNode = PartyNode
    
    On Error GoTo 0
End Function

' =============================================================================
' FONCTION: CreateAddressNode
' DESCRIPTION: Crée un nœud PstlAdr
' =============================================================================
Private Function CreateAddressNode(ByRef xmlDoc As Object, ByRef Addr As TAddress) As Object
    On Error Resume Next
    
    Dim AddrNode As Object
    Dim Node As Object
    
    Set AddrNode = xmlDoc.createElement(TAG_PSTLADR)
    
    ' StrtNm
    Set Node = xmlDoc.createElement(TAG_STRTNM)
    Node.Text = SanitizeStringForXML(Addr.StrtNm)
    AddrNode.appendChild Node
    
    ' PstCd
    Set Node = xmlDoc.createElement(TAG_PSTCD)
    Node.Text = SanitizeStringForXML(Addr.PstCd)
    AddrNode.appendChild Node
    
    ' TwnNm
    Set Node = xmlDoc.createElement(TAG_TWNNM)
    Node.Text = SanitizeStringForXML(Addr.TwnNm)
    AddrNode.appendChild Node
    
    ' Ctry
    Set Node = xmlDoc.createElement(TAG_CTRY)
    Node.Text = Addr.Ctry
    AddrNode.appendChild Node
    
    Set CreateAddressNode = AddrNode
    
    On Error GoTo 0
End Function

' =============================================================================
' FONCTION: CreateFinancialInstitutionNode
' DESCRIPTION: Crée un nœud FinInstnId (institution financière)
' =============================================================================
Private Function CreateFinancialInstitutionNode(ByRef xmlDoc As Object, ByVal TagName As String, _
                                                 ByRef FinInstn As TFinancialInstitution) As Object
    On Error Resume Next
    
    Dim FinInstnNode As Object
    Dim Node As Object
    Dim SubNode As Object
    
    Set FinInstnNode = xmlDoc.createElement(TagName)
    
    Set Node = xmlDoc.createElement(TAG_FININSTNID)
    
    ' BIC
    Set SubNode = xmlDoc.createElement(TAG_BIC)
    SubNode.Text = FinInstn.BIC
    Node.appendChild SubNode
    
    ' Nm
    Set SubNode = xmlDoc.createElement(TAG_NM)
    SubNode.Text = SanitizeStringForXML(FinInstn.Nm)
    Node.appendChild SubNode
    
    ' PstlAdr
    Node.appendChild CreateAddressNode(xmlDoc, FinInstn.Address)
    
    FinInstnNode.appendChild Node
    
    Set CreateFinancialInstitutionNode = FinInstnNode
    
    On Error GoTo 0
End Function

' =============================================================================
' SUB: ExportAsUTF8Text
' DESCRIPTION: Exporte le XML en fichier texte UTF-8
' =============================================================================
Private Sub ExportAsUTF8Text(ByVal XMLFilePath As String)
    On Error Resume Next
    
    Dim fNum As Integer
    Dim FileContent As String
    Dim TXTFilePath As String
    
    ' Lire le contenu du fichier XML
    fNum = FreeFile
    Open XMLFilePath For Input As #fNum
    FileContent = Input$(LOF(fNum), #fNum)
    Close #fNum
    
    ' Générer le chemin du fichier TXT
    TXTFilePath = Left(XMLFilePath, InStrRev(XMLFilePath, ".")) & "txt"
    
    ' Écrire en UTF-8
    fNum = FreeFile
    Open TXTFilePath For Output As #fNum
    Print #fNum, FileContent;
    Close #fNum
    
    LogInfo "Fichier TXT UTF-8 exporté: " & TXTFilePath, "modXMLExport.ExportAsUTF8Text"
    
    On Error GoTo 0
End Sub

' =============================================================================
' FONCTION: IndentXML
' DESCRIPTION: Formate le XML avec indentation (optionnel, pour lecture humaine)
' =============================================================================
Public Function IndentXML(ByVal XMLStr As String) As String
    On Error Resume Next
    
    Dim Result As String
    Dim Indent As String
    Dim i As Long
    Dim Char As String
    Dim InTag As Boolean
    Dim TagName As String
    
    Result = ""
    Indent = ""
    InTag = False
    
    For i = 1 To Len(XMLStr)
        Char = Mid(XMLStr, i, 1)
        
        If Char = "<" Then
            InTag = True
            If Mid(XMLStr, i + 1, 1) = "/" Then
                ' Balise fermante
                If Len(Indent) >= Len(EXPORT_INDENT) Then
                    Indent = Left(Indent, Len(Indent) - Len(EXPORT_INDENT))
                End If
            End If
            Result = Result & vbCrLf & Indent
        ElseIf Char = ">" Then
            InTag = False
            ' Vérifier si c'est une balise ouvrante sans fermeture auto
            If Mid(XMLStr, i - 1, 1) <> "/" And Mid(XMLStr, i - 1, 2) <> "/>" Then
                Indent = Indent & EXPORT_INDENT
            End If
        End If
        
        Result = Result & Char
    Next i
    
    IndentXML = Result
    
    On Error GoTo 0
End Function
