'===============================================================================
' MODULE: modTypes.bas
' DESCRIPTION: Types de données personnalisés pour le gestionnaire SEPA XML
'===============================================================================
Option Explicit

' =============================================================================
' TYPE: TAddress
' DESCRIPTION: Structure d'adresse postale
' =============================================================================
Public Type TAddress
    StrtNm As String      ' Nom de la rue
    PstCd As String       ' Code postal
    TwnNm As String       ' Ville
    Ctry As String        ' Pays (code ISO 2 lettres)
End Type

' =============================================================================
' TYPE: TParty
' DESCRIPTION: Structure d'une partie (émetteur, débiteur, créditeur)
' =============================================================================
Public Type TParty
    Nm As String          ' Nom
    Address As TAddress   ' Adresse
End Type

' =============================================================================
' TYPE: TFinancialInstitution
' DESCRIPTION: Structure d'une institution financière
' =============================================================================
Public Type TFinancialInstitution
    BIC As String         ' Code BIC
    Nm As String          ' Nom de la banque
    Address As TAddress   ' Adresse
End Type

' =============================================================================
' TYPE: TPaymentIdentification
' DESCRIPTION: Structure d'identification de paiement
' =============================================================================
Public Type TPaymentIdentification
    EndToEndId As String  ' Identification de bout en bout
    InstrId As String     ' Identification d'instruction (optionnel)
End Type

' =============================================================================
' TYPE: TAmount
' DESCRIPTION: Structure de montant
' =============================================================================
Public Type TAmount
    Amount As Double      ' Montant
    Ccy As String         ' Devise (code ISO 3 lettres)
End Type

' =============================================================================
' TYPE: TTransaction
' DESCRIPTION: Structure complète d'une transaction de virement
' =============================================================================
Public Type TTransaction
    ' Identification
    PmtId As TPaymentIdentification
    
    ' Montant
    Amt As TAmount
    
    ' Créancier (Bénéficiaire)
    Cdtr As TParty
    
    ' Agent du créancier (Banque bénéficiaire)
    CdtrAgt As TFinancialInstitution
    
    ' Compte du créancier
    CdtrAcct_IBAN As String
    
    ' Informations de remise
    RmtInf_Ustrd As String
    
    ' Référence interne Excel
    RowNumber As Long
End Type

' =============================================================================
' TYPE: TGroupHeader
' DESCRIPTION: Structure de l'en-tête de groupe (GrpHdr)
' =============================================================================
Public Type TGroupHeader
    MsgId As String           ' Identification du message
    CreDtTm As Date           ' Date et heure de création
    NbOfTxs As Long           ' Nombre de transactions
    CtrlSum As Double         ' Somme de contrôle
    InitgPty As TParty        ' Partie initiante
End Type

' =============================================================================
' TYPE: TPaymentInformation
' DESCRIPTION: Structure des informations de paiement (PmtInf)
' =============================================================================
Public Type TPaymentInformation
    PmtInfId As String              ' Identification des informations de paiement
    PmtMtd As String                ' Méthode de paiement
    BtchBookg As Boolean            ' Comptabilisation par lot
    ReqdExctnDt As Date             ' Date d'exécution requise
    Dbtr As TParty                  ' Débiteur
    DbtrAcct_IBAN As String         ' IBAN du débiteur
    DbtrAgt_BIC As String           ' BIC de la banque du débiteur
    ChrgBr As String                ' Porteur de frais
    Transactions() As TTransaction  ' Tableau de transactions
End Type

' =============================================================================
' TYPE: TSEPAXMLDocument
' DESCRIPTION: Structure complète du document XML SEPA
' =============================================================================
Public Type TSEPAXMLDocument
    GrpHdr As TGroupHeader              ' En-tête de groupe
    PmtInf As TPaymentInformation       ' Informations de paiement
    NamespaceURI As String              ' URI du namespace
    NamespaceXSI As String              ' URI du namespace XSI
End Type

' =============================================================================
' TYPE: TErrorInfo
' DESCRIPTION: Structure d'information d'erreur
' =============================================================================
Public Type TErrorInfo
    ErrorCode As Long           ' Code d'erreur
    ErrorLevel As String        ' Niveau de l'erreur
    Description As String       ' Description détaillée
    TransactionRef As String    ' Référence de transaction concernée
    CellReference As String     ' Référence de cellule Excel
    Source As String            ' Source de l'erreur
    Timestamp As Date           ' Horodatage
End Type

' =============================================================================
' TYPE: TXMLElement
' DESCRIPTION: Structure pour le mapping dynamique XML-Excel
' =============================================================================
Public Type TXMLElement
    TagName As String           ' Nom de la balise
    XPath As String             ' Chemin XPath
    CellReference As String     ' Référence de cellule
    DataType As String          ' Type de données
    IsRequired As Boolean       ' Obligatoire
    DefaultValue As String      ' Valeur par défaut
    ParentTag As String         ' Balise parente
    Order As Integer            ' Ordre dans la hiérarchie
End Type

' =============================================================================
' TYPE: TXLMExportNode
' DESCRIPTION: Structure pour la génération XML à l'export
' =============================================================================
Public Type TXLMExportNode
    NodeName As String          ' Nom du nœud
    NodeValue As String         ' Valeur du nœud
    Attributes() As String      ' Tableau d'attributs (pair nom/valeur)
    Children() As TXLMExportNode ' Nœuds enfants
    Level As Integer            ' Niveau d'indentation
    HasChildren As Boolean      ' Indicateur d'enfants
End Type

' =============================================================================
' CONSTANTES DE TYPE DE DONNÉES
' =============================================================================
Public Const DATATYPE_STRING As String = "String"
Public Const DATATYPE_DATE As String = "Date"
Public Const DATATYPE_DATETIME As String = "DateTime"
Public Const DATATYPE_AMOUNT As String = "Amount"
Public Const DATATYPE_BOOLEAN As String = "Boolean"
Public Const DATATYPE_INTEGER As String = "Integer"
Public Const DATATYPE_LONG As String = "Long"
Public Const DATATYPE_DOUBLE As String = "Double"
