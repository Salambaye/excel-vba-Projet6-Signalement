Attribute VB_Name = "Module1"
'Salamata Nourou MBAYE - 19/12/2025 - Version 1.0
'Projet 6 : Signalements

' ____________Variables globales pour le fichier de sortie____________________

Dim wbLauncher As Workbook
Dim wsLauncher As Worksheet

' Déclaration des variables pour le fichier de sortie
Dim wbOutput As Workbook
Dim wsOutput As Worksheet
Dim cheminOutput As String
Dim nomFichierOutput As String
Dim derniereLigneOutput As Long
Dim derniereColonneOutput As Long

Sub Signalement()

    'Optimisation pour accélérer la macro
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    '------------------- ETAPE 1 : Déclaration des variables --------------------------------
    Dim wbTDB As Workbook
    Dim wbPilotage As Workbook

    Dim wsTDB As Worksheet
    Dim wsTableauReleves As Worksheet
    Dim wsRefQuartiers As Worksheet
    Dim wsClientsTop15 As Worksheet
    
    Dim derniereLigneTDB As Long
    Dim derniereLigneReleves As Long
    Dim derniereLigneQuartiers As Long
    Dim derniereLigneTop15 As Long
    Dim ligneDestination As Long
    Dim i As Long, j As Long
    
    Dim cheminFichierTDB As String
    Dim cheminFichierPilotage As String
    Dim fdlg As FileDialog
    Dim dossierSauvegarde As String
    
    ' Dictionnaires pour optimiser les recherches
    Dim dictReleves As Object
    Dim dictQuartiers As Object
    Dim dictTop15 As Object
    
    ' Variables temporaires
    Dim cleRecherche As String
    Dim codePostal As String
    Dim ville As String
    Dim agence As String
    Dim quartier As String
    Dim raisonSociale As String
    Dim ue As String
    
    ' Créer les dictionnaires
    Set dictReleves = CreateObject("Scripting.Dictionary")
    Set dictQuartiers = CreateObject("Scripting.Dictionary")
    Set dictTop15 = CreateObject("Scripting.Dictionary")
    
    
    ' -------------  ETAPE 2 : Sélection du premier fichier (TDB) ---------------
    MsgBox "Sélectionner le fichier TDB_INDICATEURS"
    Set fdlg = Application.FileDialog(msoFileDialogFilePicker)
    fdlg.Title = "Étape 1/2 : Choisir le fichier TDB_INDICATEURS obligatoirement"
    fdlg.Filters.Clear
    fdlg.Filters.Add "Fichiers Excel", "*.xlsx;*.xls;*.xlsm"
    fdlg.AllowMultiSelect = False
    
    If fdlg.Show <> -1 Then
        MsgBox "Sélection annulée par l'utilisateur.", vbInformation
        GoTo Fin
    End If
    
    cheminFichierTDB = fdlg.SelectedItems(1)
    
    ' ------------------  ETAPE 3 : Sélection du deuxième fichier (Pilotage) ----------------------
    MsgBox "Sélectionner le fichier Pilotage "
    Set fdlg = Application.FileDialog(msoFileDialogFilePicker)
    fdlg.Title = "Étape 2/2 : Choisir le fichier Pilotage obligatoirement"
    fdlg.Filters.Clear
    fdlg.Filters.Add "Fichiers Excel", "*.xlsx;*.xls;*.xlsm"
    fdlg.AllowMultiSelect = False
    
    If fdlg.Show <> -1 Then
        MsgBox "Sélection annulée par l'utilisateur.", vbInformation
        GoTo Fin
    End If
    
    cheminFichierPilotage = fdlg.SelectedItems(1)
     
    ' --------------- Vérification des fichiers -------------
    If Dir(cheminFichierTDB) = "" Then
        MsgBox "Le fichier TDB_INDICATEURS n'existe pas : " & cheminFichierTDB, vbCritical
        GoTo Fin
    End If
    
    If Dir(cheminFichierPilotage) = "" Then
        MsgBox "Le fichier Pilotage n'existe pas : " & cheminFichierPilotage, vbCritical
        GoTo Fin
    End If
    
    ' Vérifier que les fichiers sélectionnés soient différents
    If cheminFichierTDB = cheminFichierPilotage Then
        If MsgBox("Attention ! Vous avez sélectionné le même fichier deux fois." & vbCrLf & _
                  "Voulez-vous continuer quand même ?", vbExclamation + vbYesNo) = vbNo Then
            GoTo Fin
        End If
    End If
    
    
    ' ------------------  ETAPE 4 : Sélection du dossier de sauvegarde du fichier ---------------------
    MsgBox "Choisir le dossier dans lequel le fichier doit être enregistré"
    Set fdlgDossier = Application.FileDialog(msoFileDialogFolderPicker)
    With fdlgDossier
        .Title = "Choisir le dossier de sauvegarde du fichier"
        .AllowMultiSelect = False
        .InitialFileName = Environ("USERPROFILE") & "\DESKTOP\"
    End With

    If fdlgDossier.Show <> -1 Then
        MsgBox "Sélection du dossier annulée par l'utilisateur.", vbInformation
        Exit Sub
    End If

    dossierSauvegarde = fdlgDossier.SelectedItems(1)

    ' Vérifier que le dossier existe et est accessible
    If Dir(dossierSauvegarde, vbDirectory) = "" Then
        MsgBox "Le dossier sélectionné n'est pas accessible : " & dossierSauvegarde, vbCritical
        Exit Sub
    End If
    
    
    
    ' -------- Ouvrir les fichiers sources (UpdateLinks:=0 désactive la boîte de dialogue de mise à jour)---------------
    On Error Resume Next
    Set wbTDB = Workbooks.Open(Filename:=cheminFichierTDB, ReadOnly:=True, UpdateLinks:=0, IgnoreReadOnlyRecommended:=True)
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de l'ouverture de TDB_INDICATEURS : " & Err.Description, vbCritical
        GoTo Fin
    End If
    Err.Clear
    
    Set wbPilotage = Workbooks.Open(Filename:=cheminFichierPilotage, ReadOnly:=True, UpdateLinks:=0, IgnoreReadOnlyRecommended:=True)
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de l'ouverture de Pilotage : " & Err.Description, vbCritical
        GoTo Fin
    End If
    On Error GoTo 0
    
    ' Références aux feuilles
    On Error Resume Next
    Set wsTDB = wbTDB.Worksheets("Signalement")
    Set wsTableauReleves = wbPilotage.Worksheets("Tableau des relèves")
    Set wsRefQuartiers = wbPilotage.Worksheets("réf quartiers")
    Set wsClientsTop15 = wbPilotage.Worksheets("clients top 15")
    On Error GoTo 0
    
    ' Vérification que toutes les feuilles existent
    If wsTDB Is Nothing Then
        MsgBox "La feuille 'Signalement' n'existe pas dans TDB_INDICATEURS", vbCritical
        GoTo Fin
    End If
    If wsTableauReleves Is Nothing Then
        MsgBox "La feuille 'Tableau des relèves' n'existe pas dans Pilotage", vbCritical
        GoTo Fin
    End If
    If wsRefQuartiers Is Nothing Then
        MsgBox "La feuille 'réf quartiers' n'existe pas dans Pilotage", vbCritical
        GoTo Fin
    End If
    If wsClientsTop15 Is Nothing Then
        MsgBox "La feuille 'clients top 15' n'existe pas dans Pilotage", vbCritical
        GoTo Fin
    End If
    
    ' ------------------  ETAPE 5 : Initialisation du fichier de sortie ---------------------
    Call InitialiserLauncher
    

    
    ' ------------------  ETAPE 6 : Charger les données de référence dans les dictionnaires ---------------------
    ' Charger Tableau des relèves
    derniereLigneReleves = wsTableauReleves.Cells(wsTableauReleves.Rows.Count, "J").End(xlUp).Row
    For j = 3 To derniereLigneReleves
        ue = Trim(CStr(wsTableauReleves.Cells(j, 9).Value))
        If ue <> "" And Not dictReleves.Exists(ue) Then
            agence = Trim(CStr(wsTableauReleves.Cells(j, 8).Value))
            If Len(agence) = 1 Then agence = "0" & agence
            
            dictReleves.Add ue, Array( _
                               Trim(CStr(wsTableauReleves.Cells(j, 10).Value)), _
                            Trim(CStr(wsTableauReleves.Cells(j, 11).Value)), _
                            agence)
        End If
    Next j
    
    ' Charger réf quartiers
    derniereLigneQuartiers = wsRefQuartiers.Cells(wsRefQuartiers.Rows.Count, "A").End(xlUp).Row
    For j = 2 To derniereLigneQuartiers
        cleRecherche = Trim(CStr(wsRefQuartiers.Cells(j, 1).Value)) & "|" & _
                       Trim(CStr(wsRefQuartiers.Cells(j, 2).Value)) & "|" & _
                       Trim(CStr(wsRefQuartiers.Cells(j, 3).Value))
        If Not dictQuartiers.Exists(cleRecherche) Then
            dictQuartiers.Add cleRecherche, Trim(CStr(wsRefQuartiers.Cells(j, 4).Value))
        End If
    Next j
    
    ' Charger clients top 15
    derniereLigneTop15 = wsClientsTop15.Cells(wsClientsTop15.Rows.Count, "A").End(xlUp).Row
    For j = 2 To derniereLigneTop15
        raisonSociale = Trim(CStr(wsClientsTop15.Cells(j, 1).Value))
        If raisonSociale <> "" And Not dictTop15.Exists(raisonSociale) Then
            dictTop15.Add raisonSociale, raisonSociale
        End If
    Next j
    

    
    ' ------------------  ETAPE 7 : Copie des données dans TDB - Signalement ---------------------
    ' Déterminer la dernière ligne dans TDB Signalement
    derniereLigneTDB = wsTDB.Cells(wsTDB.Rows.Count, "E").End(xlUp).Row
    
    ' Copier les en-têtes de Signalement (ligne 4, colonnes A à N) vers launcher quotidien (colonne E)
    wsTDB.Range(wsTDB.Cells(4, 1), wsTDB.Cells(5, 14)).Copy Destination:=wsLauncher.Cells(4, 5)
    Application.CutCopyMode = False
    
    ' Ligne de destination dans launcher quotidien
    ligneDestination = 6
    
    ' Parcourir les lignes de TDB Signalement à partir de ligne 6
    For i = 6 To derniereLigneTDB
        ' Vérifier si le statut (colonne E) est "A Traiter"
        If Trim(UCase(wsTDB.Cells(i, 5).Value)) = "A TRAITER" Then
            ' Copier les données de la ligne (colonnes A à N) vers launcher quotidien (colonne E)
            wsTDB.Range(wsTDB.Cells(i, 1), wsTDB.Cells(i, 14)).Copy Destination:=wsLauncher.Cells(ligneDestination, 5)
            Application.CutCopyMode = False
            
            ' Récupérer UE pour recherche
            ue = Trim(CStr(wsTDB.Cells(i, 1).Value))
            
            ' Rechercher dans le dictionnaire des relèves
            codePostal = ""
            ville = ""
            agence = ""
            
            If dictReleves.Exists(ue) Then
                codePostal = dictReleves(ue)(0)
                ville = dictReleves(ue)(1)
                agence = dictReleves(ue)(2)
            Else
                codePostal = "#N/A"
                ville = "#N/A"
            End If
            
            ' Rechercher le quartier
            quartier = ""
            If agence <> "" Then
                cleRecherche = agence & "|" & codePostal & "|" & ville
                If dictQuartiers.Exists(cleRecherche) Then
                    quartier = dictQuartiers(cleRecherche)
                Else
                    quartier = "#N/A"
                End If
            End If
            
            ' Remplir colonnes B, C, D
            wsLauncher.Cells(ligneDestination, 2).Value = codePostal
            wsLauncher.Cells(ligneDestination, 3).Value = ville
            wsLauncher.Cells(ligneDestination, 4).Value = quartier
            
            ' Rechercher dans Top 15
            raisonSociale = Trim(CStr(wsLauncher.Cells(ligneDestination, 18).Value))
            
            If dictTop15.Exists(raisonSociale) And raisonSociale <> "" Then
                wsLauncher.Cells(ligneDestination, 1).Value = raisonSociale
            Else
                wsLauncher.Cells(ligneDestination, 1).Value = "#N/A"
            End If
            
            ligneDestination = ligneDestination + 1
        End If
    Next i
    
    ' Formatage final
    If ligneDestination > 6 Then
        With wsLauncher.Range("A1:R" & ligneDestination - 1)
            .Font.Name = "Calibri"
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
        End With
    End If
    
    Call FormaterLauncher
    
    ' _______________ETAPE 8 : Créer le fichier de sortie ___________________________
    
    nomFichierOutput = "Rapport d'analyse du pilotage signalements modèle.xlsx"
    cheminOutput = dossierSauvegarde & "\" & nomFichierOutput
    wbLauncher.SaveAs cheminOutput
    wbLauncher.Close

    Call MettreEnAvantFeuilleMacro
    
    MsgBox "Traitement terminé ! " & (ligneDestination - 6) & " lignes traitées.", vbInformation
    
    ' Ouvrir le rapport
    Dim MonApplication As Object
    Set MonApplication = CreateObject("Shell.Application")
    MonApplication.Open (cheminOutput)

Fin:
    ' Fermer les fichiers sources sans enregistrer
    If Not wbTDB Is Nothing Then wbTDB.Close SaveChanges:=False
    If Not wbPilotage Is Nothing Then wbPilotage.Close SaveChanges:=False
    
    ' Libérer la mémoire
    Set dictReleves = Nothing
    Set dictQuartiers = Nothing
    Set dictTop15 = Nothing
    
    ' Réactiver les paramètres Excel
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

End Sub

Sub InitialiserLauncher()
    ' Créer le fichier de sortie
    Set wbLauncher = Workbooks.Add
    
    ' Créer la feuille "launcher quotidien"
    Set wsLauncher = wbLauncher.Worksheets(1)
    wsLauncher.Name = "launcher quotidien"
    wsLauncher.Tab.Color = RGB(0, 113, 255)
    
End Sub

Sub FormaterLauncher()
    With wsLauncher.Range("A1:A1")
        .Value = "EXTRACTION SIGNALEMENT TSP FAIT LE : " & Format(Now, "dd/mm/yyyy")
    End With
    With wsLauncher.Range("A1:R1")
        .Font.Name = "Calibri"
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenterAcrossSelection
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' En-têtes
    With wsLauncher
        .Cells(5, 1).Value = "Top 15"
        .Cells(5, 2).Value = "Code Postal"
        .Cells(5, 3).Value = "Ville"
        .Cells(5, 4).Value = "Quartier"
    End With
    
    With wsLauncher.Range("A5:D5")
        .Font.Name = "Calibri"
        .Font.Bold = True
        .Font.Size = 11
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With wsLauncher.Range("A4:D5")
        .Interior.Color = RGB(255, 255, 0)
    End With
    
    With wsLauncher.Range("E4:R4")
        .UnMerge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    End With
    
    With wsLauncher.Range("E5:R5")
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Définir les largeurs de colonnes
    wsLauncher.Columns("A:A").ColumnWidth = 35
    wsLauncher.Columns("B:B").ColumnWidth = 19
    wsLauncher.Columns("C:C").ColumnWidth = 28
    wsLauncher.Columns("D:D").ColumnWidth = 24
    wsLauncher.Columns("E:F").ColumnWidth = 15
    wsLauncher.Columns("G:G").ColumnWidth = 42
    wsLauncher.Columns("H:J").ColumnWidth = 12
    wsLauncher.Columns("K:K").ColumnWidth = 30
    wsLauncher.Columns("L:L").ColumnWidth = 12
    wsLauncher.Columns("M:M").ColumnWidth = 30
    wsLauncher.Columns("N:N").ColumnWidth = 12
    wsLauncher.Columns("O:O").ColumnWidth = 35
    wsLauncher.Columns("P:Q").ColumnWidth = 15
    wsLauncher.Columns("R:R").ColumnWidth = 35
       
End Sub

Sub MettreEnAvantFeuilleMacro()                  'Afficher uniquement la macro et masquer les autres onglets
    Dim ws As Worksheet
    Dim feuillePrincipale As String
    feuillePrincipale = "MACRO"

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = feuillePrincipale Then
            ws.Visible = xlSheetVisible
            ws.Activate
        Else
            ws.Visible = xlSheetHidden
        End If
    Next ws
End Sub


