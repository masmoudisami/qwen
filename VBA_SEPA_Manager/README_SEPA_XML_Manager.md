# SEPA XML Manager - Documentation Complète

## Vue d'ensemble

**SEPA XML Manager** est une solution complète VBA/Excel pour la gestion des fichiers de virement SEPA au format **ISO 20022 pain.001.001.03**.

Cette solution a été développée spécifiquement pour s'adapter à la structure réelle du fichier modèle `xml-Ok.xml` fourni.

---

## Architecture du Projet

### Modules VBA (8 modules)

| Module | Fichier | Description | Lignes |
|--------|---------|-------------|--------|
| **modConstants** | modConstants.bas | Constantes globales, namespaces, balises XML | ~220 |
| **modTypes** | modTypes.bas | Types de données personnalisés (TTransaction, TSEPAXMLDocument, etc.) | ~170 |
| **modUtils** | modUtils.bas | Fonctions utilitaires (validation IBAN/BIC, formatage, etc.) | ~530 |
| **modLogging** | modLogging.bas | Gestion des logs et des erreurs | ~380 |
| **modXMLImport** | modXMLImport.bas | Import et parsing du XML SEPA | ~640 |
| **modXMLExport** | modXMLExport.bas | Export XML conforme au modèle | ~590 |
| **modValidation** | modValidation.bas | Validation des données (IBAN, BIC, doublons, etc.) | ~450 |
| **modUI** | modUI.bas | Interface utilisateur (boutons, dialogues) | ~360 |
| **modMain** | modMain.bas | Point d'entrée principal et macros | ~330 |

**Total: ~3 330 lignes de code**

---

## Structure XML Supportée

Le code est basé sur la structure exacte du fichier `xml-Ok.xml`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<Document xmlns="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03">
  <CstmrCdtTrfInitn>
    <GrpHdr>
      <MsgId>...</MsgId>
      <CreDtTm>...</CreDtTm>
      <NbOfTxs>...</NbOfTxs>
      <CtrlSum>...</CtrlSum>
      <InitgPty>
        <Nm>...</Nm>
        <PstlAdr>
          <StrtNm>...</StrtNm>
          <PstCd>...</PstCd>
          <TwnNm>...</TwnNm>
          <Ctry>...</Ctry>
        </PstlAdr>
      </InitgPty>
    </GrpHdr>
    <PmtInf>
      <PmtInfId>...</PmtInfId>
      <PmtMtd>TRF</PmtMtd>
      <BtchBookg>false</BtchBookg>
      <ReqdExctnDt>...</ReqdExctnDt>
      <Dbtr>...</Dbtr>
      <DbtrAcct><Id><IBAN>...</IBAN></Id></DbtrAcct>
      <DbtrAgt><FinInstnId><BIC>...</BIC></FinInstnId></DbtrAgt>
      <ChrgBr>DEBT</ChrgBr>
      <CdtTrfTxInf>
        <!-- Transactions -->
      </CdtTrfTxInf>
    </PmtInf>
  </CstmrCdtTrfInitn>
</Document>
```

---

## Installation dans Excel

### Étape 1: Ouvrir l'éditeur VBA
1. Ouvrez Excel
2. Appuyez sur `ALT + F11`
3. Dans le menu, cliquez sur `Insertion > Module`

### Étape 2: Importer les modules
Copiez le contenu de chaque fichier `.bas` dans un nouveau module:

1. Créer un module nommé `modConstants` → coller le contenu de modConstants.bas
2. Créer un module nommé `modTypes` → coller le contenu de modTypes.bas
3. Créer un module nommé `modUtils` → coller le contenu de modUtils.bas
4. Créer un module nommé `modLogging` → coller le contenu de modLogging.bas
5. Créer un module nommé `modXMLImport` → coller le contenu de modXMLImport.bas
6. Créer un module nommé `modXMLExport` → coller le contenu de modXMLExport.bas
7. Créer un module nommé `modValidation` → coller le contenu de modValidation.bas
8. Créer un module nommé `modUI` → coller le contenu de modUI.bas
9. Créer un module nommé `modMain` → coller le contenu de modMain.bas

### Étape 3: Ajouter le code ThisWorkbook
Dans l'éditeur VBA, double-cliquez sur `ThisWorkbook` et ajoutez:

```vba
Private Sub Workbook_Open()
    On Error Resume Next
    
    ' Initialiser le système de logging
    InitializeErrorLogging
    
    ' Créer la feuille d'erreurs si nécessaire
    Dim ws As Worksheet
    Set ws = GetWorksheetByName("Erreurs")
    If Not ws Is Nothing Then
        SetupErrorSheet ws
    End If
    
    ' Afficher un message de bienvenue
    MsgBox "Bienvenue dans SEPA XML Manager!" & vbCrLf & _
           "Format supporté: pain.001.001.03" & vbCrLf & vbCrLf & _
           "Cliquez sur 'Importer XML' pour commencer.", _
           vbInformation, "SEPA XML Manager"
End Sub
```

### Étape 4: Enregistrer le classeur
Enregistrez le fichier au format **.xlsm** (classeur Excel prenant en charge les macros)

---

## Utilisation

### Démarrage rapide

1. **Ouvrir le classeur Excel**
2. **Exécuter `InitializeApplication`** (ou laisser Workbook_Open s'exécuter)
3. **Cliquer sur "Importer XML"** pour charger un fichier existant
   - OU **exécuter `GenerateSampleData`** pour créer des données de démo

### Fonctionnalités principales

#### 1. Importer un fichier XML
- Cliquez sur le bouton **"Importer XML"**
- Sélectionnez votre fichier XML SEPA
- Les données sont chargées dans la feuille **"Principal"**

#### 2. Modifier les données
- Modifiez directement les cellules dans Excel
- Ajoutez des lignes avec **Ctrl+N** ou le bouton approprié
- Supprimez des lignes avec **Ctrl+D**

#### 3. Valider les données
- Cliquez sur **"Valider Données"**
- Les erreurs apparaissent dans la feuille **"Erreurs"**
- Corrections automatiques des totaux avec **"Recalculer Totaux"**

#### 4. Exporter vers XML
- Cliquez sur **"Exporter XML"**
- Choisissez l'emplacement de sauvegarde
- Deux fichiers sont créés: `.xml` et `.txt` (UTF-8)

---

## Feuille "Principal" - Structure

### Section ENTÊTE (lignes 2-29)

| Cellule | Champ | Description |
|---------|-------|-------------|
| B2 | MsgId | Identifiant du message |
| B3 | CreDtTm | Date/heure de création |
| B4 | NbOfTxs | Nombre de transactions (auto) |
| B5 | CtrlSum | Somme totale (auto) |
| B7-B11 | Initiating Party | Émetteur du fichier |
| B14-B21 | Debtor | Donneur d'ordre (IBAN, BIC) |
| B23-B29 | Payment Information | Infos de paiement |

### Section DÉTAILS (à partir ligne 32)

| Colonne | Champ | Description |
|---------|-------|-------------|
| B | EndToEndId | Référence unique |
| C | InstrId | Référence instruction |
| D | Nom Bénéficiaire | Nom du créditeur |
| E-H | Adresse | Rue, CP, Ville, Pays |
| I | IBAN Bénéficiaire | Compte du bénéficiaire |
| J | BIC Banque | Code banque bénéficiaire |
| K | Nom Banque | Nom de la banque |
| L | Montant | Amount à virer |
| M | Devise | Code ISO (ex: TND) |
| N | Libellé | Motif du virement |

---

## Validations implémentées

### Validation IBAN (Algorithme MOD-97)
- Vérification de la longueur (15-34 caractères)
- Vérification des 2 lettres de pays
- Calcul de la clé de contrôle MOD-97

### Validation BIC/SWIFT
- Longueur: 8 ou 11 caractères
- 4 premières lettres: code banque
- 2 lettres suivantes: code pays
- 2 caractères suivants: code lieu
- 3 derniers (optionnels): code agence

### Autres validations
- ✓ Champs obligatoires
- ✓ Formats de dates (ISO 8601)
- ✓ Formats de montants (3 décimales)
- ✓ Codes devise ISO (3 lettres)
- ✓ Codes pays ISO (2 lettres)
- ✓ Détection de doublons EndToEndId
- ✓ Caractères autorisés SEPA

---

## Raccourcis clavier

| Raccourci | Action |
|-----------|--------|
| Ctrl+I | Importer XML |
| Ctrl+E | Exporter XML |
| Ctrl+V | Valider données |
| Ctrl+T | Recalculer totaux |
| Ctrl+N | Nouvelle transaction |
| Ctrl+D | Supprimer transaction |

---

## Macros principales à connaître

| Macro | Description |
|-------|-------------|
| `Main_ImportXML` | Lance l'import d'un fichier XML |
| `Main_ExportXML` | Lance l'export vers XML |
| `Main_ValidateData` | Valide toutes les données |
| `Main_RecalculateTotals` | Recalcule NbOfTxs et CtrlSum |
| `GenerateSampleData` | Génère des données de démo |
| `TestImportFromSample` | Importe le fichier xml-Ok.xml |
| `Help_About` | Affiche les informations |

---

## Conformité ISO 20022

Le code respecte strictement:

1. **Namespace**: `urn:iso:std:iso:20022:tech:xsd:pain.001.001.03`
2. **Ordre des balises**: Identique au modèle
3. **Attributs XML**: Conservation de l'attribut `Ccy` sur `InstdAmt`
4. **Formats de données**:
   - Dates: `YYYY-MM-DDThh:mm:ss` (CreDtTm), `YYYY-MM-DD` (ReqdExctnDt)
   - Montants: 3 décimales (ex: `301.500`)
   - Booléens: `true`/`false` (minuscules)

---

## Gestion des erreurs

### Feuille "Erreurs"

Colonnes:
1. Horodatage
2. Niveau (CRITIQUE / AVERTISSEMENT / INFO)
3. Code d'erreur
4. Description
5. Source
6. Transaction concernée
7. Cellule concernée

### Codes d'erreur

| Plage | Type |
|-------|------|
| 1001-1005 | Erreurs XML |
| 2001-2007 | Erreurs de validation |
| 3001-3003 | Erreurs d'export |

---

## Bonnes pratiques

### Avant l'export
1. Toujours exécuter **"Valider Données"**
2. Corriger toutes les erreurs critiques
3. Vérifier les doublons d'EndToEndId
4. Confirmer que CtrlSum correspond au total attendu

### Pour les IBAN
- Utiliser toujours des IBAN valides (test MOD-97)
- Format: 2 lettres + 2 chiffres + reste (sans espaces)

### Pour les EndToEndId
- Doivent être uniques dans tout le fichier
- Maximum 35 caractères
- Éviter les caractères spéciaux

---

## Dépannage

### Problème: "MSXML2.DOMDocument.6.0 introuvable"
**Solution**: Installer MSXML 6.0 depuis le site Microsoft

### Problème: Erreurs de validation IBAN
**Solution**: Vérifier que l'IBAN est correct avec un validateur en ligne

### Problème: Export XML non conforme
**Solution**: 
1. Vérifier qu'aucune erreur ne figure dans la feuille Erreurs
2. S'assurer que tous les champs obligatoires sont remplis
3. Contrôler les formats (dates, montants)

---

## Limitations

- Nombre maximum de transactions: 99 999
- Montant maximum par transaction: 999 999 999.99
- Un seul bloc PmtInf par fichier (conforme au modèle)

---

## Support et évolution

Pour adapter ce code à d'autres formats SEPA:
1. Modifier les constantes dans `modConstants.bas`
2. Adapter les types dans `modTypes.bas`
3. Mettre à jour les fonctions de parsing dans `modXMLImport.bas`
4. Ajuster la génération dans `modXMLExport.bas`

---

## Licence

Ce code est fourni à titre éducatif et professionnel. Adaptez-le selon vos besoins spécifiques.

---

**Développé pour: SEPA XML Manager v1.0**  
**Format: ISO 20022 pain.001.001.03**  
**Basé sur le fichier modèle: xml-Ok.xml**
