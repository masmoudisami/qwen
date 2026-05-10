' ====================================================================================================
' MODULE: modMain
' DESCRIPTION: Point d'entrée principal et initialisation de l'application SEPA XML
' ====================================================================================================
Option Explicit

' ====================================================================================================
' SUB: Auto_Open
' DESCRIPTION: S'exécute automatiquement à l'ouverture du fichier Excel
' ====================================================================================================
Public Sub Auto_Open()
    On Error Resume Next
    
    ' Initialiser l'application
    Call InitializeApplication
    
    ' Afficher un message de bienvenue
    MsgBox "SEPA XML Manager chargé avec succès!" & vbCrLf & vbCrLf & _
           "Utilisez le menu ou les macros pour:" & vbCrLf & _
           "- Importer un fichier XML SEPA" & vbCrLf & _
           "- Modifier les données dans Excel" & vbCrLf & _
           "- Exporter vers XML conforme", vbInformation, "SEPA XML Manager"
End Sub

' ====================================================================================================
' SUB: Auto_Close
' DESCRIPTION: S'exécute automatiquement à la fermeture du fichier Excel
' ====================================================================================================
Public Sub Auto_Close()
    On Error Resume Next
    
    ' Nettoyage
    Set gErrors = Nothing
    Set gXmlStructure.Namespaces = Nothing
    Set gXmlStructure.NodeOrder = Nothing
    Set gDoc = Nothing
End Sub

' ====================================================================================================
' SUB: InitializeApplication
' DESCRIPTION: Initialise l'application
' ====================================================================================================
Public Sub InitializeApplication()
    On Error GoTo ErrorHandler
    
    ' Optimiser les performances Excel
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' Initialiser les variables globales
    Call InitializeValidation
    
    ' Créer/Configurer la feuille Principal si elle n'existe pas
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_PRINCIPAL)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Call SetupWorksheet
    End If
    
    ' Restaurer les paramètres
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    LogError "modMain.InitializeApplication", Err.Number, Err.Description
End Sub

' ====================================================================================================
' SUB: Main
' DESCRIPTION: Fonction principale - point d'entrée manuel
' ====================================================================================================
Public Sub Main()
    On Error GoTo ErrorHandler
    
    Call InitializeApplication
    Call ShowMainMenu
    
    Exit Sub
    
ErrorHandler:
    LogError "modMain.Main", Err.Number, Err.Description
    MsgBox "Erreur lors de l'initialisation: " & Err.Description, vbCritical
End Sub

' ====================================================================================================
' SUB: QuickImport
' DESCRIPTION: Import rapide depuis un chemin prédéfini
' PARAMETERS: filePath - Chemin du fichier XML
' ====================================================================================================
Public Sub QuickImport(ByVal filePath As String)
    On Error GoTo ErrorHandler
    
    If Dir(filePath) = "" Then
        MsgBox "Fichier non trouvé: " & filePath, vbCritical
        Exit Sub
    End If
    
    Call InitializeApplication
    
    Dim result As Boolean
    result = ImportXMLFile(filePath)
    
    If result Then
        Call RefreshDisplay
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError "modMain.QuickImport", Err.Number, Err.Description
End Sub

' ====================================================================================================
' SUB: QuickExport
' DESCRIPTION: Export rapide vers un chemin prédéfini
' PARAMETERS: filePath - Chemin de sortie
' ====================================================================================================
Public Sub QuickExport(ByVal filePath As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_PRINCIPAL)
    
    If ws Is Nothing Then
        MsgBox "Aucune donnée à exporter.", vbCritical
        Exit Sub
    End If
    
    ' Valider
    If Not ValidateAll(ws) Then
        MsgBox "Validation échouée. Corrigez les erreurs avant d'exporter.", vbCritical
        Exit Sub
    End If
    
    ' Exporter
    Dim result As Boolean
    result = ExportToXML(filePath)
    
    Exit Sub
    
ErrorHandler:
    LogError "modMain.QuickExport", Err.Number, Err.Description
End Sub

' ====================================================================================================
' SUB: ProcessBatch
' DESCRIPTION: Traite un lot de fichiers XML
' PARAMETERS: inputFolder, outputFolder
' ====================================================================================================
Public Sub ProcessBatch(ByVal inputFolder As String, ByVal outputFolder As String)
    On Error GoTo ErrorHandler
    
    Dim fileName As String
    Dim inputFile As String
    Dim outputFile As String
    Dim count As Long
    
    count = 0
    
    ' Vérifier les dossiers
    If Right(inputFolder, 1) <> "\" Then inputFolder = inputFolder & "\"
    If Right(outputFolder, 1) <> "\" Then outputFolder = outputFolder & "\"
    
    ' Premier fichier
    fileName = Dir(inputFolder & "*.xml")
    
    Do While fileName <> ""
        inputFile = inputFolder & fileName
        outputFile = outputFolder & "processed_" & fileName
        
        ' Importer
        If ImportXMLFile(inputFile) Then
            ' Exporter
            If ExportToXML(outputFile) Then
                count = count + 1
            End If
        End If
        
        ' Fichier suivant
        fileName = Dir()
    Loop
    
    MsgBox "Traitement par lot terminé." & vbCrLf & _
           "Fichiers traités: " & count, vbInformation
    
    Exit Sub
    
ErrorHandler:
    LogError "modMain.ProcessBatch", Err.Number, Err.Description
End Sub

' ====================================================================================================
' SUB: GenerateSampleXML
' DESCRIPTION: Génère un fichier XML SEPA exemple pour test
' PARAMETERS: outputPath - Chemin de sortie
' ====================================================================================================
Public Sub GenerateSampleXML(ByVal outputPath As String)
    On Error GoTo ErrorHandler
    
    Dim xmlDoc As MSXML2.DOMDocument60
    Dim rootElem As MSXML2.IXMLDOMElement
    Dim grpHdrElem As MSXML2.IXMLDOMElement
    Dim pmtInfElem As MSXML2.IXMLDOMElement
    Dim txElem As MSXML2.IXMLDOMElement
    Dim childElem As MSXML2.IXMLDOMElement
    
    Set xmlDoc = New MSXML2.DOMDocument60
    
    With xmlDoc
        .async = False
        .validateOnParse = False
        .preserveWhiteSpace = True
    End With
    
    ' Namespace SEPA pain.001.001.03
    Const NS_SEPA As String = "urn:iso:std:iso:20022:tech:xsd:pain.001.001.03"
    
    ' Élément racine
    Set rootElem = xmlDoc.createElement("Document")
    rootElem.setAttribute "xmlns", NS_SEPA
    xmlDoc.appendChild rootElem
    
    ' GroupHeader
    Set grpHdrElem = xmlDoc.createElement("GrpHdr")
    rootElem.appendChild grpHdrElem
    
    childElem = CreateElementWithValue(xmlDoc, "MsgId", "MSG-" & Format(Now, "yyyymmddhhmmss"))
    grpHdrElem.appendChild childElem
    
    childElem = CreateElementWithValue(xmlDoc, "CreDtTm", FormatDateTimeISO(Now))
    grpHdrElem.appendChild childElem
    
    childElem = CreateElementWithValue(xmlDoc, "NbOfTxs", "2")
    grpHdrElem.appendChild childElem
    
    childElem = CreateElementWithValue(xmlDoc, "CtrlSum", "1500.00")
    grpHdrElem.appendChild childElem
    
    Dim initgPtyElem As MSXML2.IXMLDOMElement
    Set initgPtyElem = xmlDoc.createElement("InitgPty")
    grpHdrElem.appendChild initgPtyElem
    
    childElem = CreateElementWithValue(xmlDoc, "Nm", "Entreprise Exemple SARL")
    initgPtyElem.appendChild childElem
    
    ' PaymentInformation
    Set pmtInfElem = xmlDoc.createElement("PmtInf")
    rootElem.appendChild pmtInfElem
    
    childElem = CreateElementWithValue(xmlDoc, "PmtInfId", "PMT-001")
    pmtInfElem.appendChild childElem
    
    childElem = CreateElementWithValue(xmlDoc, "PmtMtd", "TRF")
    pmtInfElem.appendChild childElem
    
    childElem = CreateElementWithValue(xmlDoc, "BtchBookg", "true")
    pmtInfElem.appendChild childElem
    
    childElem = CreateElementWithValue(xmlDoc, "NbOfTxs", "2")
    pmtInfElem.appendChild childElem
    
    childElem = CreateElementWithValue(xmlDoc, "CtrlSum", "1500.00")
    pmtInfElem.appendChild childElem
    
    childElem = CreateElementWithValue(xmlDoc, "ReqdExctnDt", FormatDateISO(Date + 1))
    pmtInfElem.appendChild childElem
    
    ' Debtor
    Dim dbtrElem As MSXML2.IXMLDOMElement
    Set dbtrElem = xmlDoc.createElement("Dbtr")
    pmtInfElem.appendChild dbtrElem
    
    childElem = CreateElementWithValue(xmlDoc, "Nm", "Entreprise Exemple SARL")
    dbtrElem.appendChild childElem
    
    ' Debtor Account
    Dim dbtrAcctElem As MSXML2.IXMLDOMElement
    Set dbtrAcctElem = xmlDoc.createElement("DbtrAcct")
    pmtInfElem.appendChild dbtrAcctElem
    
    Dim idElem As MSXML2.IXMLDOMElement
    Set idElem = xmlDoc.createElement("Id")
    dbtrAcctElem.appendChild idElem
    
    childElem = CreateElementWithValue(xmlDoc, "IBAN", "FR7612345678901234567890123")
    idElem.appendChild childElem
    
    ' Debtor Agent
    Dim dbtrAgtElem As MSXML2.IXMLDOMElement
    Set dbtrAgtElem = xmlDoc.createElement("DbtrAgt")
    pmtInfElem.appendChild dbtrAgtElem
    
    Dim finInstnElem As MSXML2.IXMLDOMElement
    Set finInstnElem = xmlDoc.createElement("FinInstnId")
    dbtrAgtElem.appendChild finInstnElem
    
    childElem = CreateElementWithValue(xmlDoc, "BIC", "BNPAFRPPXXX")
    finInstnElem.appendChild childElem
    
    childElem = CreateElementWithValue(xmlDoc, "ChrgBr", "SLEV")
    pmtInfElem.appendChild childElem
    
    ' Transaction 1
    Set txElem = xmlDoc.createElement("CdtTrfTxInf")
    pmtInfElem.appendChild txElem
    
    Dim pmtIdElem As MSXML2.IXMLDOMElement
    Set pmtIdElem = xmlDoc.createElement("PmtId")
    txElem.appendChild pmtIdElem
    
    childElem = CreateElementWithValue(xmlDoc, "EndToEndId", "E2E-001")
    pmtIdElem.appendChild childElem
    
    Dim amtElem As MSXML2.IXMLDOMElement
    Set amtElem = xmlDoc.createElement("Amt")
    txElem.appendChild amtElem
    
    Dim instdAmtElem As MSXML2.IXMLDOMElement
    Set instdAmtElem = xmlDoc.createElement("InstdAmt")
    instdAmtElem.setAttribute "Ccy", "EUR"
    instdAmtElem.Text = "1000.00"
    amtElem.appendChild instdAmtElem
    
    ' Creditor Agent
    Dim cdtrAgtElem As MSXML2.IXMLDOMElement
    Set cdtrAgtElem = xmlDoc.createElement("CdtrAgt")
    txElem.appendChild cdtrAgtElem
    
    Set finInstnElem = xmlDoc.createElement("FinInstnId")
    cdtrAgtElem.appendChild finInstnElem
    
    childElem = CreateElementWithValue(xmlDoc, "BIC", "SOGEFRPPXXX")
    finInstnElem.appendChild childElem
    
    ' Creditor
    Dim cdtrElem As MSXML2.IXMLDOMElement
    Set cdtrElem = xmlDoc.createElement("Cdtr")
    txElem.appendChild cdtrElem
    
    childElem = CreateElementWithValue(xmlDoc, "Nm", "Fournisseur Alpha")
    cdtrElem.appendChild childElem
    
    ' Creditor Account
    Dim cdtrAcctElem As MSXML2.IXMLDOMElement
    Set cdtrAcctElem = xmlDoc.createElement("CdtrAcct")
    txElem.appendChild cdtrAcctElem
    
    Set idElem = xmlDoc.createElement("Id")
    cdtrAcctElem.appendChild idElem
    
    childElem = CreateElementWithValue(xmlDoc, "IBAN", "FR7612345678909876543210987")
    idElem.appendChild childElem
    
    ' Remittance Info
    Dim rmtInfElem As MSXML2.IXMLDOMElement
    Set rmtInfElem = xmlDoc.createElement("RmtInf")
    txElem.appendChild rmtInfElem
    
    childElem = CreateElementWithValue(xmlDoc, "Ustrd", "Facture F2024-001")
    rmtInfElem.appendChild childElem
    
    ' Transaction 2
    Set txElem = xmlDoc.createElement("CdtTrfTxInf")
    pmtInfElem.appendChild txElem
    
    Set pmtIdElem = xmlDoc.createElement("PmtId")
    txElem.appendChild pmtIdElem
    
    childElem = CreateElementWithValue(xmlDoc, "EndToEndId", "E2E-002")
    pmtIdElem.appendChild childElem
    
    Set amtElem = xmlDoc.createElement("Amt")
    txElem.appendChild amtElem
    
    Set instdAmtElem = xmlDoc.createElement("InstdAmt")
    instdAmtElem.setAttribute "Ccy", "EUR"
    instdAmtElem.Text = "500.00"
    amtElem.appendChild instdAmtElem
    
    Set cdtrAgtElem = xmlDoc.createElement("CdtrAgt")
    txElem.appendChild cdtrAgtElem
    
    Set finInstnElem = xmlDoc.createElement("FinInstnId")
    cdtrAgtElem.appendChild finInstnElem
    
    childElem = CreateElementWithValue(xmlDoc, "BIC", "CCFRFRPPXXX")
    finInstnElem.appendChild childElem
    
    Set cdtrElem = xmlDoc.createElement("Cdtr")
    txElem.appendChild cdtrElem
    
    childElem = CreateElementWithValue(xmlDoc, "Nm", "Société Beta SA")
    cdtrElem.appendChild childElem
    
    Set cdtrAcctElem = xmlDoc.createElement("CdtrAcct")
    txElem.appendChild cdtrAcctElem
    
    Set idElem = xmlDoc.createElement("Id")
    cdtrAcctElem.appendChild idElem
    
    childElem = CreateElementWithValue(xmlDoc, "IBAN", "FR7698765432101234567890123")
    idElem.appendChild childElem
    
    Set rmtInfElem = xmlDoc.createElement("RmtInf")
    txElem.appendChild rmtInfElem
    
    childElem = CreateElementWithValue(xmlDoc, "Ustrd", "Facture F2024-002")
    rmtInfElem.appendChild childElem
    
    ' Sauvegarder
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open outputPath For Output As #fileNum
    Print #fileNum, "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Print #fileNum, xmlDoc.xml
    Close #fileNum
    
    MsgBox "Fichier XML exemple généré:" & vbCrLf & outputPath, vbInformation
    
    Exit Sub
    
ErrorHandler:
    LogError "modMain.GenerateSampleXML", Err.Number, Err.Description
    MsgBox "Erreur lors de la génération: " & Err.Description, vbCritical
End Sub

' ====================================================================================================
' SUB: TestValidation
' DESCRIPTION: Teste les fonctions de validation
' ====================================================================================================
Public Sub TestValidation()
    On Error GoTo ErrorHandler
    
    Dim testIBAN As String
    Dim testBIC As String
    Dim result As Boolean
    
    ' Test IBAN valide
    testIBAN = "FR7612345678901234567890123"
    result = ValidateIBAN(testIBAN)
    Debug.Print "IBAN " & testIBAN & " valide: " & result
    
    ' Test BIC valide
    testBIC = "BNPAFRPPXXX"
    result = ValidateBIC(testBIC)
    Debug.Print "BIC " & testBIC & " valide: " & result
    
    ' Test format montant
    Dim amount As Double
    amount = 1234.56
    Debug.Print "Montant " & amount & " formaté: " & FormatAmount(amount)
    
    ' Test format date
    Debug.Print "Date formatée: " & FormatDateISO(Date)
    Debug.Print "DateTime formaté: " & FormatDateTimeISO(Now)
    
    MsgBox "Tests de validation terminés. Consultez la fenêtre Immediate (Ctrl+G).", vbInformation
    
    Exit Sub
    
ErrorHandler:
    LogError "modMain.TestValidation", Err.Number, Err.Description
End Sub
