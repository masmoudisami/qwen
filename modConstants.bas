' ====================================================================================================
' MODULE: modConstants
' DESCRIPTION: Constantes globales pour l'application SEPA XML
' ====================================================================================================
Option Explicit

' ====================================================================================================
' NAMESPACES SEPA ISO 20022
' ====================================================================================================
Public Const NS_SEPA As String = "urn:iso:std:iso:20022:tech:xsd:pain.001.001.03"
Public Const NS_XSI As String = "http://www.w3.org/2001/XMLSchema-instance"

' ====================================================================================================
' FEUILLES EXCEL
' ====================================================================================================
Public Const SHEET_PRINCIPAL As String = "Principal"
Public Const SHEET_ERREURS As String = "Erreurs"
Public Const SHEET_LOG As String = "Log"

' ====================================================================================================
' PLAGE DE DONNÉES - SECTION ENTÊTE
' ====================================================================================================
Public Const RANGE_MSGID As String = "B2"
Public Const RANGE_CREDTTM As String = "B3"
Public Const RANGE_NBFTXS As String = "B4"
Public Const RANGE_CTRLSUM As String = "B5"
Public Const RANGE_INITIATING_PARTY_NAME As String = "B6"
Public Const RANGE_DEBTOR_NAME As String = "B7"
Public Const RANGE_DEBTOR_IBAN As String = "B8"
Public Const RANGE_DEBTOR_BIC As String = "B9"
Public Const RANGE_REQUESTED_EXECUTION_DATE As String = "B10"
Public Const RANGE_DEVISE As String = "B11"
Public Const RANGE_PAYMENT_INFO_ID As String = "B12"
Public Const RANGE_PAYMENT_METHOD As String = "B13"
Public Const RANGE_BATCH_BOOKING As String = "B14"
Public Const RANGE_NB_OF_TXS_PI As String = "B15"
Public Const RANGE_CTRL_SUM_PI As String = "B16"
Public Const RANGE_DEBTOR_ACCOUNT_IBAN As String = "B17"
Public Const RANGE_DEBTOR_AGENT_BIC As String = "B18"
Public Const RANGE_CHARGE_BEARER As String = "B19"

' ====================================================================================================
' PLAGE DE DONNÉES - SECTION DÉTAILS (TRANSACCTIONS)
' ====================================================================================================
Public Const TX_START_ROW As Long = 22
Public Const TX_COL_ENDTOENDID As Long = 2      ' B
Public Const TX_COL_INSTRID As Long = 3         ' C
Public Const TX_COL_BENEFICIARY_NAME As Long = 4 ' D
Public Const TX_COL_BENEFICIARY_IBAN As Long = 5 ' E
Public Const TX_COL_BENEFICIARY_BIC As Long = 6  ' F
Public Const TX_COL_AMOUNT As Long = 7          ' G
Public Const TX_COL_REMITTANCE_INFO As Long = 8 ' H
Public Const TX_COL_EXECUTION_DATE As Long = 9  ' I
Public Const TX_COL_BENEFICIARY_ADDRESS As Long = 10 ' J
Public Const TX_COL_BENEFICIARY_COUNTRY As Long = 11 ' K
Public Const TX_COL_BENEFICIARY_CITY As Long = 12 ' L

' ====================================================================================================
' COLONNES ERREURS
' ====================================================================================================
Public Const ERR_COL_TYPE As Long = 1
Public Const ERR_COL_DESCRIPTION As Long = 2
Public Const ERR_COL_TRANSACTION As Long = 3
Public Const ERR_COL_CELLULE As Long = 4
Public Const ERR_COL_NIVEAU As Long = 5

' ====================================================================================================
' TYPES DE VALIDATION
' ====================================================================================================
Public Enum ValidationType
    VT_IBAN = 1
    VT_BIC = 2
    VT_AMOUNT = 3
    VT_DATE = 4
    VT_MANDATORY = 5
    VT_DUPLICATE = 6
    VT_STRUCTURE = 7
    VT_NAMESPACE = 8
End Enum

' ====================================================================================================
' NIVEAUX D'ERREUR
' ====================================================================================================
Public Enum ErrorLevel
    EL_CRITICAL = 1
    EL_WARNING = 2
    EL_INFO = 3
End Enum

' ====================================================================================================
' CODES ERREUR
' ====================================================================================================
Public Const ERR_IBAN_INVALID As String = "ERR_IBAN_001"
Public Const ERR_BIC_INVALID As String = "ERR_BIC_001"
Public Const ERR_AMOUNT_INVALID As String = "ERR_AMT_001"
Public Const ERR_DATE_INVALID As String = "ERR_DAT_001"
Public Const ERR_MANDATORY_MISSING As String = "ERR_MAN_001"
Public Const ERR_DUPLICATE_FOUND As String = "ERR_DUP_001"
Public Const ERR_XML_STRUCTURE As String = "ERR_XML_001"
Public Const ERR_XML_NAMESPACE As String = "ERR_XML_002"

' ====================================================================================================
' CONFIGURATION EXPORT
' ====================================================================================================
Public Const EXPORT_INDENT As String = "  "
Public Const EXPORT_ENCODING As String = "UTF-8"
