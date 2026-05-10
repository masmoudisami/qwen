# Installation du SEPA XML Manager

## Fichiers générés

- `SEPA_XML_Manager.xlsm` : Fichier Excel avec macros (structure de base)
- `VBA_SEPA_Manager/` : Dossier contenant tous les modules VBA (.bas)
- `ImportHelper.bas` : Script d'import automatique

## Instructions d'installation

### Méthode 1 : Import manuel des modules (Recommandé)

1. Ouvrez le fichier `SEPA_XML_Manager.xlsm` dans Excel
2. Activez les macros si demandé
3. Appuyez sur `ALT + F11` pour ouvrir l'éditeur VBA
4. Dans le menu, cliquez sur `Fichier` > `Importer un fichier...`
5. Sélectionnez tous les fichiers `.bas` du dossier `VBA_SEPA_Manager/`
   - modConstants.bas
   - modTypes.bas
   - modUtils.bas
   - modLogging.bas
   - modXMLImport.bas
   - modXMLExport.bas
   - modValidation.bas
   - modUI.bas
   - modMain.bas
6. Sauvegardez le fichier Excel

### Méthode 2 : Import automatique via macro

1. Ouvrez le fichier `SEPA_XML_Manager.xlsm` dans Excel
2. Appuyez sur `ALT + F11` pour ouvrir l'éditeur VBA
3. Importez le fichier `ImportHelper.bas` (Fichier > Importer un fichier)
4. Exécutez la macro `ImportAllModules` (F5 ou Run)
5. Exécutez la macro `CreateButtons` pour créer les boutons
6. Sauvegardez le fichier Excel

### Configuration requise

- Excel 2010 ou version ultérieure
- Référence MSXML2 v6.0 activée (généralement présente par défaut)
  - Pour vérifier: Outils > Références dans l'éditeur VBA
  - Cocher "Microsoft XML, v6.0"

### Activation des macros

À l'ouverture du fichier, Excel peut afficher un avertissement de sécurité:
1. Cliquez sur "Activer le contenu" ou "Enable Content"
2. Si nécessaire, ajoutez le fichier aux emplacements approuvés:
   - Fichier > Options > Centre de gestion de la confidentialité
   - Paramètres du Centre de gestion... > Emplacements approuvés
   - Ajouter le dossier contenant le fichier

## Utilisation

### Importer un fichier XML SEPA

1. Cliquez sur le bouton "Importer XML"
2. Sélectionnez votre fichier XML (ex: xml-Ok.xml)
3. Les données sont chargées dans la feuille "Principal"

### Modifier les données

- Section ENTÊTE (lignes 2-12): Modifiez les informations globales
- Section DÉTAILS (à partir de ligne 16): Ajoutez/modifiez les transactions

### Valider les données

1. Cliquez sur "Valider Données"
2. Consultez la feuille "Erreurs" pour les éventuels problèmes

### Exporter vers XML

1. Cliquez sur "Exporter XML"
2. Choisissez l'emplacement de sauvegarde
3. Le fichier XML généré respecte la structure du modèle

## Structure du fichier Excel

### Feuille "Principal"

**Section ENTÊTE:**
- MsgId: Identifiant du message
- CreDtTm: Date de création
- NbOfTxs: Nombre de transactions (calculé auto)
- CtrlSum: Montant total (calculé auto)
- InitiatingParty: Nom de l'émetteur
- Debtor: Nom du donneur d'ordre
- IBAN Donneur: Compte de débit
- BIC: Code banque émetteur
- Date Exécution: Date de valeur
- Devise: Code devise (TND, EUR, etc.)
- PmtInfId: Identifiant de paiement
- ChrgBr: Frais (DEBT, SHAR, CRED)

**Section DÉTAILS:**
- EndToEndId: Référence de transaction
- InstrId: Identifiant d'instruction
- Nom Bénéficiaire
- IBAN Bénéficiaire
- BIC Banque
- Montant
- Devise
- Libellé
- Adresse complète

### Feuille "Erreurs"

Affiche les erreurs de validation:
- Type d'erreur
- Description détaillée
- Transaction concernée
- Cellule Excel
- Niveau (Critique, Majeur, Mineur)

### Feuille "Journal"

Historique des opérations et logs du système.

## Modules VBA

| Module | Description |
|--------|-------------|
| modConstants | Constantes et configurations |
| modTypes | Types personnalisés (structures) |
| modUtils | Fonctions utilitaires |
| modLogging | Gestion des logs et erreurs |
| modXMLImport | Import et parsing XML |
| modXMLExport | Génération XML |
| modValidation | Validations métier (IBAN, BIC, etc.) |
| modUI | Interface utilisateur |
| modMain | Point d'entrée principal |

## Support

Pour toute question ou problème, consultez le fichier README_SEPA_XML_Manager.md
