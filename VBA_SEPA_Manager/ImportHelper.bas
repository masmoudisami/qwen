Attribute VB_Name = "modImportHelper"
' =============================================================================
' SCRIPT D'IMPORT DES MODULES VBA
' Ce module aide à importer tous les modules VBA dans le projet Excel
' =============================================================================
Option Explicit

Sub ImportAllModules()
    Dim strPath As String
    Dim strFile As String
    Dim vbComp As VBComponent
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Chemin vers les fichiers .bas
    strPath = ThisWorkbook.Path & "\VBA_SEPA_Manager\"
    
    If Not fso.FolderExists(strPath) Then
        MsgBox "Le dossier VBA_SEPA_Manager n'existe pas: " & strPath, vbCritical
        Exit Sub
    End If
    
    ' Liste des fichiers à importer
    Dim files As Variant
    files = Array( _
        "modConstants.bas", _
        "modTypes.bas", _
        "modUtils.bas", _
        "modLogging.bas", _
        "modXMLImport.bas", _
        "modXMLExport.bas", _
        "modValidation.bas", _
        "modUI.bas", _
        "modMain.bas" _
    )
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim i As Integer
    For i = LBound(files) To UBound(files)
        strFile = strPath & files(i)
        
        If fso.FileExists(strFile) Then
            On Error Resume Next
            ' Supprimer le module s'il existe déjà
            Set vbComp = ThisWorkbook.VBProject.VBComponents(files(i).Replace(".bas", ""))
            If Not vbComp Is Nothing Then
                ThisWorkbook.VBProject.VBComponents.Remove vbComp
            End If
            On Error GoTo 0
            
            ' Importer le module
            ThisWorkbook.VBProject.VBComponents.Import strFile
            Debug.Print "Module importé: " & files(i)
        Else
            Debug.Print "Fichier non trouvé: " & strFile
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Import des modules terminé!" & vbCrLf & _
           "Redémarrez Excel pour activer les macros.", vbInformation
End Sub

Sub CreateButtons()
    ' Crée les boutons sur la feuille Principal
    Dim ws As Worksheet
    Dim btn As Shape
    
    Set ws = ThisWorkbook.Worksheets("Principal")
    
    ' Supprimer les anciens boutons
    For Each btn In ws.Shapes
        If btn.Type = msoButton Then btn.Delete
    Next btn
    
    ' Bouton Importer XML
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, 10, 500, 120, 30)
    With btn
        .OnAction = "ImportXML_Click"
        .OLEFormat.Object.Caption = "Importer XML"
    End With
    
    ' Bouton Exporter XML
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, 140, 500, 120, 30)
    With btn
        .OnAction = "ExportXML_Click"
        .OLEFormat.Object.Caption = "Exporter XML"
    End With
    
    ' Bouton Valider
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, 270, 500, 120, 30)
    With btn
        .OnAction = "ValidateData_Click"
        .OLEFormat.Object.Caption = "Valider Données"
    End With
    
    ' Bouton Ajouter Ligne
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, 400, 500, 120, 30)
    With btn
        .OnAction = "AddTransaction_Click"
        .OLEFormat.Object.Caption = "Ajouter Ligne"
    End With
    
    MsgBox "Boutons créés avec succès!", vbInformation
End Sub
