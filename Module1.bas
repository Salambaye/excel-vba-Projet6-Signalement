Attribute VB_Name = "Module1"
'Salamata Nourou MBAYE - 15/12/2025 - Version 1.0
'Projet 6 : Signalements

' ____________Variables globales pour le fichier de sorie____________________

Dim wbOutput As Workbook
Dim wsLauncher As Worksheet

' Déclaration des variables pour le fichier de sortie
'Dim wbRapportANOSortie As Workbook
'Dim wsLauncherSortie As Worksheet
'
'Dim cheminSortieRapportANO As String
'Dim nomFichierRapportANO  As String
'Dim derniereLigneRapportANO As Long
'Dim derniereColonneRapportANO As Long

Sub Signalement()

    'Optimisation pour accélérer la macro
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

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
    Dim ligneDestination As Long
    Dim i As Long, j As Long
    
    Dim cheminFichierTDB As String
    Dim cheminFichierPilotage As String
    Dim fdlg As FileDialog
    Dim dossierSauvegarde As String
    Dim fdlgDossier As FileDialog
    Dim cheminOutput As String
    
    Dim cleRecherche As String
    Dim codePostal As String
    Dim ville As String
    Dim agence As String
    Dim quartier As String
    Dim raisonSociale As String
    Dim top15 As Variant
    
    
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
        Exit Sub
    End If
    
    cheminFichierPilotage = fdlg.SelectedItems(1)
     
    ' --------------- Vérification des fichiers -------------
    If Dir(cheminFichierTDB) = "" Then
        MsgBox "Le fichier TDB_INDICATEURS n'existe pas : " & cheminFichierTDB, vbCritical
        GoTo Fin
    End If
    
    If Dir(cheminFichierPilotage) = "" Then
        MsgBox "Le fichier Pilotage n'existe pas : " & cheminFichierPilotage, vbCritical
        Exit Sub
    End If
    
    ' Vérifier que les fichiers sélectionnés soient différents
    If cheminFichierTDB = cheminFichierPilotage Then
        If MsgBox("Attention ! Vous avez sélectionné le même fichier deux fois." & vbCrLf & _
                  "Voulez-vous continuer quand même ?", vbExclamation + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Ouvrir les fichiers sources
    On Error Resume Next
    Set wbTDB = Workbooks.Open(cheminFichierTDB, ReadOnly:=True)
    On Error GoTo 0
  
    On Error Resume Next
    Set wbPilotage = Workbooks.Open(cheminFichierPilotage, ReadOnly:=True)
    On Error GoTo 0

'    ' ------------------  ETAPE 4 : Sélection du dossier de sauvegarde du fichier ---------------------
'    MsgBox "Choisir le dossier dans lequel le fichier doit être enregistré"
'    Set fdlgDossier = Application.FileDialog(msoFileDialogFolderPicker)
'    With fdlgDossier
'        .Title = "Choisir le dossier de sauvegarde du fichier"
'        .AllowMultiSelect = False
'        .InitialFileName = Environ("USERPROFILE") & "\DESKTOP\"
'    End With
'
'    If fdlgDossier.Show <> -1 Then
'        MsgBox "Sélection du dossier annulée par l'utilisateur.", vbInformation
'        Exit Sub
'    End If
'
'    dossierSauvegarde = fdlgDossier.SelectedItems(1)
'
'    ' Vérifier que le dossier existe et est accessible
'    If Dir(dossierSauvegarde, vbDirectory) = "" Then
'        MsgBox "Le dossier sélectionné n'est pas accessible : " & dossierSauvegarde, vbCritical
'        Exit Sub
'    End If
    
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
    
    
    ' ------------------  ETAPE 6 : Copie des données dans TDB - Signalement ---------------------
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
        
            ' Récupérer les informations pour remplir colonnes A-D
    ' Rechercher dans Tableau des relèves
    derniereLigneReleves = wsTableauReleves.Cells(wsTableauReleves.Rows.Count, "J").End(xlUp).Row
    codePostal = ""
    ville = ""
    agence = ""
            
    ' Clef commune : UE
    For j = 3 To derniereLigneReleves
        If wsTDB.Cells(i, 1).Value = wsTableauReleves.Cells(j, 9).Value Then
                
            codePostal = wsTableauReleves.Cells(j, 10).Value ' Colonne J
            ville = wsTableauReleves.Cells(j, 11).Value ' Colonne K
            agence = wsTableauReleves.Cells(j, 8).Value ' Colonne H
            
            If Len(Trim(agence)) = 1 Then
                agence = "0" & agence
            End If
                
        End If
        '                Exit For
    Next j
            
    ' Rechercher le quartier dans réf quartiers
    quartier = ""
    cleRecherche = agence & "|" & codePostal & "|" & ville

    derniereLigneQuartiers = wsRefQuartiers.Cells(wsRefQuartiers.Rows.Count, "A").End(xlUp).Row
    For j = 2 To derniereLigneQuartiers
        Dim cleQuartier As String
        cleQuartier = wsRefQuartiers.Cells(j, 1).Value & "|" & _
                      wsRefQuartiers.Cells(j, 2).Value & "|" & _
                      wsRefQuartiers.Cells(j, 3).Value

        If cleQuartier = cleRecherche Then
            quartier = wsRefQuartiers.Cells(j, 4).Value ' Colonne D
            Exit For
        End If
    Next j
            
'     Remplir colonnes B, C, D
    wsLauncher.Cells(ligneDestination, 2).Value = codePostal
    wsLauncher.Cells(ligneDestination, 3).Value = ville
    wsLauncher.Cells(ligneDestination, 4).Value = quartier
            
    ' Rechercher dans Top 15 (colonne A = Top 15)
    ' Raison sociale est en colonne R (18) de launcher quotidien
    raisonSociale = wsLauncher.Cells(ligneDestination, 18).Value

    ' RECHERCHEV dans clients top 15
    On Error Resume Next
    top15 = Application.WorksheetFunction.VLookup(raisonSociale, wsClientsTop15.Range("A:B"), 1, False)
    On Error GoTo 0

    If Not IsError(top15) And top15 <> "" Then
        wsLauncher.Cells(ligneDestination, 1).Value = top15

        ' Colorer la ligne en rouge
'        wsLauncher.Range(wsLauncher.Cells(ligneDestination, 1), _
'                         wsLauncher.Cells(ligneDestination, 18)).Interior.Color = RGB(255, 0, 0)
'        wsLauncher.Range(wsLauncher.Cells(ligneDestination, 1), _
'                         wsLauncher.Cells(ligneDestination, 18)).Font.Color = RGB(255, 255, 255)
    End If

            ligneDestination = ligneDestination + 1
        End If
    Next i
    
    With wsLauncher.Range("A1:R" & ligneDestination)
        .Font.Name = "Calibri"
    End With
    
    Call FormaterLauncher

Fin:
    ' Fermer les fichiers sources sans enregistrer
    '    If Not wbTDB Is Nothing Then wbTDB.Close SaveChanges:=False
    '    If Not wbPilotage Is Nothing Then wbPilotage.Close SaveChanges:=False
    
    ' Réactiver les paramètres Excel
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True


End Sub

Sub InitialiserLauncher()
    ' Créer le fichier de sortie
    Set wbOutput = Workbooks.Add
    
    ' Créer la feuille "launcher quotidien"
    Set wsLauncher = wbOutput.Worksheets(1)
    wsLauncher.Name = "launcher quotidien"
    wsLauncher.Tab.Color = RGB(0, 113, 255)      'RGB(27, 235, 151)
    
    '    Call FormaterLauncher
    
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
        '        .Interior.Color = RGB(0, 51, 102)
        '        .Font.Color = RGB(255, 255, 255)
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
        '        .Borders.Weight = xlMedium
    End With
    
    With wsLauncher.Range("A4:D5")
        .Interior.Color = RGB(255, 255, 0)
        '        .Font.Color = RGB(255, 255, 255)
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
        '        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
    End With

    
    ' Définir les largeurs de colonnes
    wsLauncher.Columns("A:A").ColumnWidth = 35   ' Top 15
    wsLauncher.Columns("B:B").ColumnWidth = 19   ' Code postal
    wsLauncher.Columns("C:C").ColumnWidth = 28   ' Ville
    wsLauncher.Columns("D:D").ColumnWidth = 24   ' Quartier
    wsLauncher.Columns("E:F").ColumnWidth = 15   ' Code UEX, Code Agence
    wsLauncher.Columns("G:G").ColumnWidth = 42   ' Dénomination
    wsLauncher.Columns("H:J").ColumnWidth = 12   ' Numéro, Statut et Code Observation
    wsLauncher.Columns("K:K").ColumnWidth = 30   ' Libellé observation
    wsLauncher.Columns("L:L").ColumnWidth = 12   ' Code motif de non résolution
    wsLauncher.Columns("M:M").ColumnWidth = 30   ' Libellé motif de non résolution
    wsLauncher.Columns("N:N").ColumnWidth = 12   ' Initiales
    wsLauncher.Columns("O:O").ColumnWidth = 35   ' Identité
    wsLauncher.Columns("P:Q").ColumnWidth = 15   ' Date de passage planifiée
    wsLauncher.Columns("R:R").ColumnWidth = 35   ' Raison sociale
       
End Sub


