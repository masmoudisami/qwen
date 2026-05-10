'===============================================================================
' MODULE: modConstants.bas
' DESCRIPTION: Constantes globales pour le gestionnaire SEPA XML
' VERSION: pain.001.001.03
'===============================================================================
Option Explicit

' =============================================================================
' NAMESPACES ET SCHEMAS
' =============================================================================
Public Const NS_ISO As String = "urn:iso:std:iso:20022:tech:xsd:pain.001.001.03"
Public Const NS_XSI As String = "http://www.w3.org/2001/XMLSchema-instance"
Public Const SCHEMA_LOCATION As String = NS_ISO & " pain.001.001.03.xsd"

' =============================================================================
' BALISES XML PRINCIPALES (Structure pain.001.001.03)
' =============================================================================
' Document Root
Public Const TAG_DOCUMENT As String = "Document"
Public Const TAG_CSTMRCDTTRFINITN As String = "CstmrCdtTrfInitn"

' Group Header (GrpHdr)
Public Const TAG_GRPHDR As String = "GrpHdr"
Public Const TAG_MSGID As String = "MsgId"
Public Const TAG_CREDTTM As String = "CreDtTm"
Public Const TAG_NBOFTXS As String = "NbOfTxs"
Public Const TAG_CTRLSUM As String = "CtrlSum"
Public Const TAG_INITGPTY As String = "InitgPty"
Public Const TAG_NM As String = "Nm"
Public Const TAG_PSTLADR As String = "PstlAdr"
Public Const TAG_STRTNM As String = "StrtNm"
Public Const TAG_PSTCD As String = "PstCd"
Public Const TAG_TWNNM As String = "TwnNm"
Public Const TAG_CTRY As String = "Ctry"

' Payment Information (PmtInf)
Public Const TAG_PMTINF As String = "PmtInf"
Public Const TAG_PMTINFID As String = "PmtInfId"
Public Const TAG_PMTMTD As String = "PmtMtd"
Public Const TAG_BtchBookg As String = "BtchBookg"
Public Const TAG_REQDEXCTNDT As String = "ReqdExctnDt"
Public Const TAG_DBTR As String = "Dbtr"
Public Const TAG_DBTRACCT As String = "DbtrAcct"
Public Const TAG_DBTRAGT As String = "DbtrAgt"
Public Const TAG_CHRGBR As String = "ChrgBr"
Public Const TAG_ID As String = "Id"
Public Const TAG_IBAN As String = "IBAN"
Public Const TAG_FININSTNID As String = "FinInstnId"
Public Const TAG_BIC As String = "BIC"

' Credit Transfer Transaction Information (CdtTrfTxInf)
Public Const TAG_CDTTRFTXINF As String = "CdtTrfTxInf"
Public Const TAG_PMTID As String = "PmtId"
Public Const TAG_ENDTOENDID As String = "EndToEndId"
Public Const TAG_INSTRID As String = "InstrId"
Public Const TAG_AMT As String = "Amt"
Public Const TAG_INSTDamT As String = "InstdAmt"
Public Const TAG_CDTRAGT As String = "CdtrAgt"
Public Const TAG_CDTR As String = "Cdtr"
Public Const TAG_CDTRACCT As String = "CdtrAcct"
Public Const TAG_RMTINF As String = "RmtInf"
Public Const TAG_USTRD As String = "Ustrd"

' Attributs
Public Const ATTR_CCY As String = "Ccy"

' =============================================================================
' NOMS DES FEUILLES EXCEL
' =============================================================================
Public Const SHEET_PRINCIPAL As String = "Principal"
Public Const SHEET_ERREURS As String = "Erreurs"
Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_LOG As String = "Log"

' =============================================================================
' PLAGES DE CELLULES - SECTION ENTÊTE (Feuille Principal)
' =============================================================================
' Group Header
Public Const CELL_MSGID As String = "B2"
Public Const CELL_CREDTTM As String = "B3"
Public Const CELL_NBOFTXS As String = "B4"
Public Const CELL_CTRLSUM As String = "B5"

' Initiating Party
Public Const CELL_INITGPTY_NM As String = "B7"
Public Const CELL_INITGPTY_STRTNM As String = "B8"
Public Const CELL_INITGPTY_PSTCD As String = "B9"
Public Const CELL_INITGPTY_TWNNM As String = "B10"
Public Const CELL_INITGPTY_CTRY As String = "B11"

' Debtor (Donneur d'ordre)
Public Const CELL_DBTR_NM As String = "B14"
Public Const CELL_DBTR_STRTNM As String = "B15"
Public Const CELL_DBTR_PSTCD As String = "B16"
Public Const CELL_DBTR_TWNNM As String = "B17"
Public Const CELL_DBTR_CTRY As String = "B18"
Public Const CELL_DBTR_IBAN As String = "B19"
Public Const CELL_DBTR_BIC As String = "B20"

' Payment Information
Public Const CELL_PMTINFID As String = "B23"
Public Const CELL_PMTMTD As String = "B24"
Public Const CELL_BTCHBOOKG As String = "B25"
Public Const CELL_REQDEXCTNDT As String = "B26"
Public Const CELL_CHRGBR As String = "B27"
Public Const CELL_DEVISE As String = "B28"

' =============================================================================
' PLAGES DE CELLULES - SECTION DÉTAILS (Transactions)
' =============================================================================
Public Const ROW_TRANSACTION_START As Long = 32
Public Const COL_TX_ENDTOENDID As Long = 2      ' B
Public Const COL_TX_INSTRID As Long = 3         ' C
Public Const COL_TX_CDTR_NM As Long = 4         ' D
Public Const COL_TX_CDTR_STRTNM As Long = 5     ' E
Public Const COL_TX_CDTR_PSTCD As Long = 6      ' F
Public Const COL_TX_CDTR_TWNNM As Long = 7      ' G
Public Const COL_TX_CDTR_CTRY As Long = 8       ' H
Public Const COL_TX_CDTR_IBAN As Long = 9       ' I
Public Const COL_TX_CDTR_BIC As Long = 10       ' J
Public Const COL_TX_CDTR_BANKNM As Long = 11    ' K
Public Const COL_TX_AMT As Long = 12            ' L
Public Const COL_TX_CCY As Long = 13            ' M
Public Const COL_TX_USTRD As Long = 14          ' N

' =============================================================================
' EN-TÊTES DE COLONNES
' =============================================================================
Public Const HEADER_ENDTOENDID As String = "EndToEndId"
Public Const HEADER_INSTRID As String = "InstrId"
Public Const HEADER_CDTR_NM As String = "Nom Bénéficiaire"
Public Const HEADER_CDTR_STRTNM As String = "Rue"
Public Const HEADER_CDTR_PSTCD As String = "Code Postal"
Public Const HEADER_CDTR_TWNNM As String = "Ville"
Public Const HEADER_CDTR_CTRY As String = "Pays"
Public Const HEADER_CDTR_IBAN As String = "IBAN Bénéficiaire"
Public Const HEADER_CDTR_BIC As String = "BIC Banque"
Public Const HEADER_CDTR_BANKNM As String = "Nom Banque"
Public Const HEADER_AMT As String = "Montant"
Public Const HEADER_CCY As String = "Devise"
Public Const HEADER_USTRD As String = "Libellé"

' =============================================================================
' CODES ERREUR
' =============================================================================
Public Const ERR_NONE As Long = 0
Public Const ERR_XML_LOAD As Long = 1001
Public Const ERR_XML_PARSE As Long = 1002
Public Const ERR_XML_NAMESPACE As Long = 1003
Public Const ERR_XML_MISSING_NODE As Long = 1004
Public Const ERR_XML_INVALID_STRUCTURE As Long = 1005
Public Const ERR_VALIDATION_IBAN As Long = 2001
Public Const ERR_VALIDATION_BIC As Long = 2002
Public Const ERR_VALIDATION_AMOUNT As Long = 2003
Public Const ERR_VALIDATION_DATE As Long = 2004
Public Const ERR_VALIDATION_DUPLICATE As Long = 2005
Public Const ERR_VALIDATION_MANDATORY As Long = 2006
Public Const ERR_VALIDATION_CHARS As Long = 2007
Public Const ERR_EXPORT_CREATE As Long = 3001
Public Const ERR_EXPORT_WRITE As Long = 3002
Public Const ERR_EXPORT_SAVE As Long = 3003

' =============================================================================
' NIVEAUX D'ERREUR
' =============================================================================
Public Const ERR_LEVEL_CRITICAL As String = "CRITIQUE"
Public Const ERR_LEVEL_WARNING As String = "AVERTISSEMENT"
Public Const ERR_LEVEL_INFO As String = "INFO"

' =============================================================================
' FORMATS DE DONNÉES
' =============================================================================
Public Const FORMAT_DATE_XML As String = "yyyy-mm-ddThh:mm:ss"
Public Const FORMAT_DATE_SHORT As String = "yyyy-mm-dd"
Public Const FORMAT_AMOUNT As String = "0.000"
Public Const MAX_AMOUNT_DECIMALS As Integer = 3

' =============================================================================
' LIMITES ET CONTRAINTES
' =============================================================================
Public Const MAX_NB_OF_TXS As Long = 99999
Public Const MIN_AMOUNT As Double = 0.01
Public Const MAX_AMOUNT As Double = 999999999.99
Public Const IBAN_MIN_LENGTH As Integer = 15
Public Const IBAN_MAX_LENGTH As Integer = 34
Public Const BIC_LENGTH As Integer = 8
Public Const BIC_LENGTH_WITH_BRANCH As Integer = 11

' =============================================================================
' CARACTÈRES AUTORISÉS SEPA (ISO 20022)
' =============================================================================
Public Const ALLOWED_CHARS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 .,-_+:/?&()=""'!°@£$€%§"

' =============================================================================
' MÉTHODES DE PAIEMENT
' =============================================================================
Public Const PMT_METHOD_TRF As String = "TRF"
Public Const PMT_METHOD_CHK As String = "CHK"

' =============================================================================
' FRAIS DE TRANSACTION
' =============================================================================
Public Const CHRG_BR_DEBT As String = "DEBT"
Public Const CHRG_BR_CRED As String = "CRED"
Public Const CHRG_BR_SHAR As String = "SHAR"

' =============================================================================
' BOOLEANS XML
' =============================================================================
Public Const XML_TRUE As String = "true"
Public Const XML_FALSE As String = "false"

' =============================================================================
' CONFIGURATION EXPORT
' =============================================================================
Public Const EXPORT_INDENT As String = "  "
Public Const EXPORT_ENCODING As String = "UTF-8"
Public Const EXPORT_VERSION As String = "1.0"
