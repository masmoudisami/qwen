' ====================================================================================================
' MODULE: modTypes
' DESCRIPTION: Types de données personnalisés pour l'application SEPA XML
' ====================================================================================================
Option Explicit

' ====================================================================================================
' TYPE: TransactionDetails
' DESCRIPTION: Structure contenant les informations d'une transaction individuelle
' ====================================================================================================
Public Type TransactionDetails
    EndToEndId As String
    InstrId As String
    BeneficiaryName As String
    BeneficiaryIBAN As String
    BeneficiaryBIC As String
    Amount As Double
    RemittanceInfo As String
    ExecutionDate As String
    BeneficiaryAddress As String
    BeneficiaryCountry As String
    BeneficiaryCity As String
    Currency As String
End Type

' ====================================================================================================
' TYPE: PaymentInformation
' DESCRIPTION: Structure contenant les informations du bloc PaymentInformation
' ====================================================================================================
Public Type PaymentInformation
    PaymentInfoId As String
    PaymentMethod As String
    BatchBooking As String
    NbOfTxs As Long
    CtrlSum As Double
    DebtorName As String
    DebtorIBAN As String
    DebtorBIC As String
    DebtorAgentBIC As String
    ChargeBearer As String
    RequestedExecutionDate As String
End Type

' ====================================================================================================
' TYPE: GroupHeader
' DESCRIPTION: Structure contenant les informations du bloc GroupHeader
' ====================================================================================================
Public Type GroupHeader
    MsgId As String
    CreDtTm As String
    NbOfTxs As Long
    CtrlSum As Double
    InitiatingPartyName As String
    NumberOfPayments As Long
End Type

' ====================================================================================================
' TYPE: SepaDocument
' DESCRIPTION: Structure principale contenant toutes les données du document SEPA
' ====================================================================================================
Public Type SepaDocument
    GroupHeader As GroupHeader
    PaymentInfo As PaymentInformation
    Transactions() As TransactionDetails
    NamespaceUri As String
    NamespacePrefix As String
    XmlVersion As String
    XmlDocumentId As String
End Type

' ====================================================================================================
' TYPE: ValidationError
' DESCRIPTION: Structure pour stocker les erreurs de validation
' ====================================================================================================
Public Type ValidationError
    ErrorType As String
    Description As String
    TransactionRef As String
    CellReference As String
    Level As ErrorLevel
    ErrorCode As String
End Type

' ====================================================================================================
' TYPE: XmlMapping
' DESCRIPTION: Structure pour le mapping dynamique entre XML et Excel
' ====================================================================================================
Public Type XmlMapping
    XmlNodePath As String
    ExcelCell As String
    NodeName As String
    ParentNode As String
    IsArray As Boolean
    ArrayIndex As Long
End Type

' ====================================================================================================
' TYPE: XmlStructure
' DESCRIPTION: Structure pour stocker la structure complète du XML analysé
' ====================================================================================================
Public Type XmlStructure
    RootNode As String
    Namespaces As Collection
    NodeOrder As Collection
    Attributes As Collection
    HeaderNodes As Collection
    PaymentNodes As Collection
    TransactionNodes As Collection
End Type
