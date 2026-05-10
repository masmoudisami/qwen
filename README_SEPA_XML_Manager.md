# SEPA XML Manager - Solution VBA Excel

## Vue d'ensemble

Solution complète VBA Excel pour l'import, la modification et l'export de fichiers XML SEPA ISO 20022 (pain.001.001.03).

## Architecture des Modules

### 1. modConstants.bas
Constantes globales :
- Namespaces SEPA
- Références aux feuilles Excel
- Plages de cellules
- Codes d'erreur

### 2. modTypes.bas
Types de données personnalisés :
- `TransactionDetails` : Informations d'une transaction
- `PaymentInformation` : Bloc PaymentInformation
- `GroupHeader` : Bloc GroupHeader
- `SepaDocument` : Structure complète du document
- `ValidationError` : Erreurs de validation
- `XmlStructure` : Structure XML analysée

### 3. modUtils.bas
Fonctions utilitaires :
- `FormatAmount()` / `ParseAmount()` : Formats de montants SEPA
- `FormatDateISO()` / `FormatDateTimeISO()` : Formats de dates ISO 8601
- `ValidateIBAN()` : Validation IBAN (algorithme MOD-97)
- `ValidateBIC()` : Validation code BIC/SWIFT
- `SanitizeXMLText()` / `UnescapeXMLText()` : Échappement XML
- `GenerateMsgId()` / `GetUniqueEndToEndId()` : Génération d'identifiants uniques
- `CalculateCtrlSum()` : Calcul somme de contrôle
- `CleanString()` : Nettoyage caractères SEPA

### 4. modXMLImport.bas
Import et parsing XML :
- `ImportXMLFile()` : Charge un fichier XML SEPA
- `AnalyzeXMLStructure()` : Analyse la structure XML réelle
- `MapXMLToExcel()` : Mappe les données vers Excel
- `GetNodeValue()` / `GetNodeAttribute()` : Extraction de valeurs XML
- Gestion dynamique des namespaces

### 5. modXMLExport.bas
Export XML conforme :
- `ExportToXML()` : Export vers fichier XML
- `BuildGroupHeader()` : Construit le bloc GrpHdr
- `BuildPaymentInformation()` : Construit le bloc PmtInf
- `BuildCreditTransferTransaction()` : Construit chaque transaction
- `SaveXMLFile()` : Sauvegarde avec formatage UTF-8 indenté
- `ExportToTXT()` : Export alternatif en TXT

### 6. modValidation.bas
Validation des données :
- `ValidateAll()` : Exécute toutes les validations
- `ValidateHeader()` : Valide l'en-tête
- `ValidateTransactions()` : Valide chaque transaction
- `ValidateXMLStructure()` : Valide la structure globale
- Détection des doublons EndToEndId
- Feuille "Erreurs" avec classification (CRITIQUE/AVERTISSEMENT/INFO)

### 7. modUI.bas
Interface utilisateur :
- `ShowMainMenu()` : Menu principal
- `ImportXML_Action()` : Action d'import
- `Validate_Action()` : Action de validation
- `ExportXML_Action()` : Action d'export
- `SelectFileDialog()` / `SaveFileDialog()` : Boîtes de dialogue
- `SetupWorksheet()` : Configuration de la feuille

### 8. modMain.bas
Point d'entrée :
- `Auto_Open()` : Initialisation automatique
- `Main()` : Fonction principale
- `GenerateSampleXML()` : Génère un XML exemple
- `ProcessBatch()` : Traitement par lot

## Structure des Feuilles Excel

### Feuille "Principal"

**Section ENTÊTE (lignes 1-19)**
| Cellule | Champ | Description |
|---------|-------|-------------|
| B2 | MsgId | Identifiant de message |
| B3 | CreDtTm | Date/heure de création |
| B4 | NbOfTxs | Nombre de transactions |
| B5 | CtrlSum | Somme de contrôle |
| B6 | InitiatingParty Name | Nom de l'émetteur |
| B7 | Debtor Name | Nom du donneur d'ordre |
| B8 | Debtor IBAN | IBAN du donneur d'ordre |
| B9 | Debtor BIC | BIC du donneur d'ordre |
| B10 | RequestedExecutionDate | Date d'exécution |
| B11 | Currency | Devise (EUR) |
| B12 | PaymentInfoId | ID du paiement |
| B13 | PaymentMethod | Méthode (TRF) |
| B14 | BatchBooking | Regroupement (true/false) |
| B15 | NbOfTxs (PI) | Nb transactions (PmtInf) |
| B16 | CtrlSum (PI) | Somme (PmtInf) |
| B17 | Debtor Account IBAN | IBAN compte |
| B18 | Debtor Agent BIC | BIC banque |
| B19 | ChargeBearer | Frais (SLEV) |

**Section DÉTAILS (à partir ligne 22)**
| Colonne | Champ | Description |
|---------|-------|-------------|
| B | EndToEndId | Identifiant unique |
| C | InstrId | ID instruction (optionnel) |
| D | Beneficiary Name | Nom bénéficiaire |
| E | Beneficiary IBAN | IBAN bénéficiaire |
| F | Beneficiary BIC | BIC bénéficiaire |
| G | Amount | Montant |
| H | RemittanceInfo | Libellé |
| I | ExecutionDate | Date exécution |
| J | Beneficiary Address | Adresse |
| K | Country | Pays |
| L | City | Ville |

### Feuille "Erreurs"
| Colonne | Contenu |
|---------|---------|
| A | Type d'erreur |
| B | Description |
| C | Transaction concernée |
| D | Cellule concernée |
| E | Niveau (CRITIQUE/AVERTISSEMENT/INFO) |

## Utilisation

### Importer un fichier XML
```vba
' Via le menu
Call Main

' Ou directement
Call QuickImport("C:\chemin\fichier.xml")
```

### Exporter vers XML
```vba
Call QuickExport("C:\chemin\sortie.xml")
```

### Valider les données
```vba
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Principal")
Call ValidateAll(ws)
```

### Générer un XML exemple
```vba
Call GenerateSampleXML("C:\chemin\sample.xml")
```

## Conformité SEPA

Le code respecte :
- **ISO 20022 pain.001.001.03** : Format virement SEPA
- **Namespaces XML** : Conservation exacte du modèle
- **Format des montants** : 2 décimales, point séparateur
- **Format des dates** : ISO 8601 (YYYY-MM-DD, YYYY-MM-DDThh:mm:ss)
- **Validation IBAN** : Algorithme MOD-97
- **Validation BIC** : Format SWIFT (8 ou 11 caractères)
- **Caractères autorisés** : Filtrage selon norme SEPA

## Points Clés Techniques

1. **Mapping dynamique** : Le code s'adapte à la structure réelle du XML importé
2. **Préservation des namespaces** : Attribution xmlns conservée exactement
3. **Ordre des balises** : Reproduction fidèle de la hiérarchie XML
4. **Encodage UTF-8** : Compatible avec toutes les plateformes bancaires
5. **Gestion d'erreurs robuste** : Chaque fonction inclut On Error GoTo
6. **Performance** : ScreenUpdating désactivé pendant les traitements

## Installation dans Excel

1. Ouvrir Excel
2. ALT+F11 pour ouvrir l'éditeur VBA
3. Insertion > Module pour chaque fichier .bas
4. Copier le contenu de chaque module
5. Ajouter la référence MSXML2 v6.0 :
   - Outils > Références
   - Cocher "Microsoft XML, v6.0"
6. Enregistrer en .xlsm (classeur prenant en charge les macros)

## Référence des Macros

| Macro | Description |
|-------|-------------|
| `Auto_Open` | Initialisation à l'ouverture |
| `Main` | Menu principal |
| `ImportXML_Action` | Importer XML |
| `Validate_Action` | Valider données |
| `ExportXML_Action` | Exporter XML |
| `ExportTXT_Action` | Exporter TXT |
| `Recalculate_Action` | Recalculer totaux |
| `AddTransaction_Action` | Ajouter transaction |
| `SetupWorksheet` | Configurer feuille |
| `ShowHelp` | Afficher aide |
| `TestValidation` | Tester validations |
| `GenerateSampleXML` | Générer exemple |

## Bonnes Pratiques ISO 20022

1. **EndToEndId unique** : Obligatoire et non dupliqué
2. **MsgId unique** : Identifiant de message unique par envoi
3. **Montants positifs** : Toujours > 0
4. **Dates futures** : RequestedExecutionDate >= date du jour
5. **IBAN valide** : Vérification MOD-97 obligatoire
6. **BIC valide** : Format SWIFT correct
7. **CtrlSum cohérent** : Somme des transactions exacte
8. **NbOfTxs cohérent** : Compte exact des transactions
