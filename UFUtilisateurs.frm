VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFUtilisateurs 
   Caption         =   "UserForm1"
   ClientHeight    =   8868.001
   ClientLeft      =   -10956
   ClientTop       =   504
   ClientWidth     =   23424
   OleObjectBlob   =   "UFUtilisateurs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFUtilisateurs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================
' UFUtilisateurs - CODE COMPLET (V3)
' ============================================================
' NOUVEAUTES V2 :
'  - Renfort : passage de 2 CheckBox independantes a 3 OptionButton
'    exclusifs (Aucun / PRESS / COFIT)
'  - Bloc "Planning Semaine" : grille editable
'    Entree / Sortie / Pause Debut / Pause Fin / Commentaire
'    pour les 7 jours de la semaine selectionnee dans UFGenerer
'    (variable Public BOOM.g_LundiCible)
'  - Bouton "Mettre a jour" -> reecrit PLANNING, CONSOLIDATION,
'    la feuille du projet du collaborateur, et la feuille ROTATION
'    (nb renforts + date de derniere MAJ)
'
' NOUVEAUTES V3 :
'  - Feuille Utilisateurs reorganisee sur 22 colonnes (voir en-tete
'    du module BOOM.bas pour le detail complet des colonnes)
'  - Cadre "Conge" : OUI/NON passent en OptionButton exclusifs
'    (remplacent chkConge) + nouvelle liste deroulante "Type de Conge"
'  - Cadre "TeleTravail" : OUI/NON passent en OptionButton exclusifs
'    (remplacent chkTT)
'  - NOUVEAU cadre "Maladie" : OUI/NON (exclusif) + Date d'arret +
'    Date de reprise
'  - NOUVEAU cadre "Contrat" : CDI/CDD (exclusif) + Date d'expiration
'    (n'est utile/saisie que pour un CDD ; ecrite dans la colonne
'    "Date de sortie" de la feuille Utilisateurs)
'  - Matricule / N de telephone / Date d'embauche restent des colonnes
'    de la feuille Utilisateurs mais ne sont pas saisies depuis ce
'    formulaire (non presentes sur la maquette) : leur valeur est
'    conservee telle quelle lors de l'enregistrement d'une fiche
'    existante.
'
' ------------------------------------------------------------
' CONTROLES A CREER / RENOMMER DANS L'EDITEUR DE FORMULAIRE
' (le fichier .frx n'etant pas modifiable depuis ce chat, ces
'  controles doivent etre ajoutes/renommes manuellement dans
'  l'editeur VBA, en suivant le visuel de la maquette) :
'
'   Cadre "Conge" :
'     optCongeOui / optCongeNon  (OptionButton, GroupName="GrpConge")
'     -> remplacent chkConge (a supprimer)
'     cboTypeConge   (ComboBox, associe au Label29 "Type de Conge")
'
'   Cadre "TeleTravail" :
'     optTTOui / optTTNon  (OptionButton, GroupName="GrpTT")
'     -> remplacent chkTT (a supprimer)
'
'   Cadre "Maladie" (nouveau) :
'     optMaladieOui / optMaladieNon  (OptionButton, GroupName="GrpMaladie")
'     txtDateArret    (TextBox)
'     txtDateReprise  (TextBox)
'
'   Cadre "Contrat" (nouveau) :
'     optCDI / optCDD  (OptionButton, GroupName="GrpContrat")
'     txtDateExpiration (TextBox)
'
'   Cadre "RENFORT" :
'     optRenfortAucun  (OptionButton, GroupName="GrpRenfort")
'     optRenfortPress  (OptionButton, GroupName="GrpRenfort", Caption "Renfort PRESS")
'     optRenfortCofit  (OptionButton, GroupName="GrpRenfort", Caption "RENFORT COFIT")
'     -> remplacent chkRenfortP / chkRenfortI (a supprimer)
'
'   Cadre "Planning Semaine" (7 lignes Lundi..Dimanche x 5 colonnes) :
'     txtEntree1 .. txtEntree7   (TextBox)
'     txtSortie1 .. txtSortie7   (TextBox)
'     txtPauseD1 .. txtPauseD7   (TextBox)
'     txtPauseF1 .. txtPauseF7   (TextBox)
'     txtComment1 .. txtComment7 (TextBox)
'     cmdMettreAJour             (CommandButton, Caption "Mettre a jour")
'     (index 1=Lundi, 2=Mardi, 3=Mercredi, 4=Jeudi, 5=Vendredi,
'            6=Samedi, 7=Dimanche)
'
' ------------------------------------------------------------
' STRUCTURE FEUILLE "Utilisateurs" (22 colonnes) :
'   1=NOM COMPLET | 2=Activite | 3=Ville | 4=ZONES
'   5=Conge | 6=Conge D | 7=Conge F | 8=Type de Conge
'   9=TRANSPORT
'   10=TT | 11=TT D | 12=TT F
'   13=RENFORT PRESS | 14=RENFORT ITALY
'   15=Matricule | 16=N de telephone
'   17=Type de contrat (CDI/CDD) | 18=Date d'embauche | 19=Date de sortie
'   20=Maladie | 21=Date D'Arret | 22=DATE DE REPRISE
' ============================================================

Option Explicit

Private m_ligneSelectionnee As Long
Private m_modeAjout As Boolean

' Valeurs "passives" (colonnes 15/16/18) non saisies sur ce formulaire :
' conservees telles quelles lors de la sauvegarde d'une fiche existante.
Private m_matricule As String
Private m_telephone As String
Private m_dateEmbauche As Variant

' ============================================================
Private Sub UserForm_Initialize()
    ' Remplir la liste des projets
    cboProjet.Clear
    Dim Projets As Variant
    Projets = Array("AFEDIM", "ACCESSIBILITE", "CM Leasing", "GLF", "EBRA", _
                    "EBRA PRESSE", "GOOGLE LEADS", "TLV", "TELEVENTE", "FACTO", "DAC")
    Dim p As Variant
    For Each p In Projets
        cboProjet.AddItem p
    Next p

    ' Transport
    cboTransport.Clear
    cboTransport.AddItem "OUI"
    cboTransport.AddItem "NON"
    cboTransport.Text = "NON"

    ' Type de conge
    cboTypeConge.Clear
    cboTypeConge.AddItem "Conge Paye"
    cboTypeConge.AddItem "RTT"
    cboTypeConge.AddItem "Sans Solde"
    cboTypeConge.AddItem "Conge Exceptionnel"
    cboTypeConge.AddItem "Autre"

    m_ligneSelectionnee = 0
    m_modeAjout = False
    m_matricule = ""
    m_telephone = ""
    m_dateEmbauche = ""
    ChargerListe
    VerrouillerFormulaire True
    ViderPlanningSemaine
    cmdMettreAJour.Enabled = False
End Sub

' --- Charger la liste des collaborateurs ---
Private Sub ChargerListe()
    lstCollabs.Clear
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Utilisateurs")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).Value) <> "" Then
            lstCollabs.AddItem ws.Cells(i, 1).Value & "  [" & ws.Cells(i, 2).Value & "]"
            lstCollabs.List(lstCollabs.ListCount - 1, 0) = ws.Cells(i, 1).Value & "  [" & ws.Cells(i, 2).Value & "]"
        End If
    Next i
End Sub

' --- Selection dans la liste ---
Private Sub lstCollabs_Click()
    If lstCollabs.ListIndex < 0 Then Exit Sub
    m_modeAjout = False
    m_ligneSelectionnee = lstCollabs.ListIndex + 2  ' +2 car ligne 1 = entete

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Utilisateurs")

    txtNom.Text = CStr(ws.Cells(m_ligneSelectionnee, 1).Value)
    cboProjet.Text = CStr(ws.Cells(m_ligneSelectionnee, 2).Value)
    txtVille.Text = CStr(ws.Cells(m_ligneSelectionnee, 3).Value)
    txtZone.Text = CStr(ws.Cells(m_ligneSelectionnee, 4).Value)

    ' Conge (col 5-8)
    Dim congeOui As Boolean
    congeOui = (UCase(Trim(ws.Cells(m_ligneSelectionnee, 5).Value)) = "OUI")
    optCongeOui.Value = congeOui
    optCongeNon.Value = Not congeOui
    txtCongeD.Text = IIf(IsDate(ws.Cells(m_ligneSelectionnee, 6).Value), _
                           Format(ws.Cells(m_ligneSelectionnee, 6).Value, "dd/mm/yyyy"), "")
    txtCongeF.Text = IIf(IsDate(ws.Cells(m_ligneSelectionnee, 7).Value), _
                           Format(ws.Cells(m_ligneSelectionnee, 7).Value, "dd/mm/yyyy"), "")
    cboTypeConge.Text = CStr(ws.Cells(m_ligneSelectionnee, 8).Value)

    ' Transport (col 9)
    cboTransport.Text = CStr(ws.Cells(m_ligneSelectionnee, 9).Value)

    ' TT (col 10-12)
    Dim ttOui As Boolean
    ttOui = (UCase(Trim(ws.Cells(m_ligneSelectionnee, 10).Value)) = "OUI")
    optTTOui.Value = ttOui
    optTTNon.Value = Not ttOui
    txtTTD.Text = IIf(IsDate(ws.Cells(m_ligneSelectionnee, 11).Value), _
                           Format(ws.Cells(m_ligneSelectionnee, 11).Value, "dd/mm/yyyy"), "")
    txtTTF.Text = IIf(IsDate(ws.Cells(m_ligneSelectionnee, 12).Value), _
                           Format(ws.Cells(m_ligneSelectionnee, 12).Value, "dd/mm/yyyy"), "")

    ' Renforts (col 13-14) -> 3 OptionButton exclusifs
    Dim rp As String, ri As String
    rp = UCase(Trim(ws.Cells(m_ligneSelectionnee, 13).Value))
    ri = UCase(Trim(ws.Cells(m_ligneSelectionnee, 14).Value))
    If rp = "OUI" Then
        optRenfortPress.Value = True
    ElseIf ri = "OUI" Then
        optRenfortCofit.Value = True
    Else
        optRenfortAucun.Value = True
    End If

    ' Matricule / telephone / date embauche (col 15/16/18) : non
    ' saisis ici, conserves tels quels en memoire pour la sauvegarde.
    m_matricule = CStr(ws.Cells(m_ligneSelectionnee, 15).Value)
    m_telephone = CStr(ws.Cells(m_ligneSelectionnee, 16).Value)
    m_dateEmbauche = ws.Cells(m_ligneSelectionnee, 18).Value

    ' Contrat (col 17 + 19)
    Dim typeContrat As String
    typeContrat = UCase(Trim(ws.Cells(m_ligneSelectionnee, 17).Value))
    If typeContrat = "CDD" Then
        optCDD.Value = True
    Else
        optCDI.Value = True
    End If
    txtDateExpiration.Text = IIf(IsDate(ws.Cells(m_ligneSelectionnee, 19).Value), _
                           Format(ws.Cells(m_ligneSelectionnee, 19).Value, "dd/mm/yyyy"), "")

    ' Maladie (col 20-22)
    Dim maladieOui As Boolean
    maladieOui = (UCase(Trim(ws.Cells(m_ligneSelectionnee, 20).Value)) = "OUI")
    optMaladieOui.Value = maladieOui
    optMaladieNon.Value = Not maladieOui
    txtDateArret.Text = IIf(IsDate(ws.Cells(m_ligneSelectionnee, 21).Value), _
                           Format(ws.Cells(m_ligneSelectionnee, 21).Value, "dd/mm/yyyy"), "")
    txtDateReprise.Text = IIf(IsDate(ws.Cells(m_ligneSelectionnee, 22).Value), _
                           Format(ws.Cells(m_ligneSelectionnee, 22).Value, "dd/mm/yyyy"), "")

    VerrouillerFormulaire False
    AppliquerEtatsConditionnels

    ' Charger le planning de la semaine active pour ce collaborateur
    ChargerPlanningSemaine txtNom.Text
    cmdMettreAJour.Enabled = True
End Sub

' --- Nouveau collaborateur ---
Private Sub cmdNouveau_Click()
    m_modeAjout = True
    m_ligneSelectionnee = 0
    ViderFormulaire
    VerrouillerFormulaire False
    AppliquerEtatsConditionnels
    ViderPlanningSemaine
    cmdMettreAJour.Enabled = False
    txtNom.SetFocus
End Sub

' --- Supprimer ---
Private Sub cmdSupprimer_Click()
    If lstCollabs.ListIndex < 0 Then
        MsgBox "Selectionnez d'abord un collaborateur.", vbExclamation: Exit Sub
    End If
    Dim nomSel As String: nomSel = txtNom.Text
    Dim rep As Integer
    rep = MsgBox("Supprimer " & nomSel & " ?", vbYesNo + vbWarning, "Confirmation")
    If rep = vbNo Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Utilisateurs")
    ws.Rows(m_ligneSelectionnee).Delete

    ViderFormulaire
    VerrouillerFormulaire True
    ViderPlanningSemaine
    cmdMettreAJour.Enabled = False
    m_ligneSelectionnee = 0
    ChargerListe
    MsgBox nomSel & " supprime.", vbInformation
End Sub

' Ecrit la fiche collaborateur (colonnes 1-22) a la ligne "lr" de la
' feuille Utilisateurs, apres validation des champs. Renvoie False (et
' affiche un message) si une validation echoue, sans rien ecrire.
' Utilisee a la fois par cmdSauver_Click et cmdMettreAJour_Click (pour
' que "Mettre a jour" prenne aussi en compte les modifications des
' Informations collaborateurs, pas seulement le planning).
Private Function EnregistrerFicheUtilisateur(lr As Long) As Boolean
    EnregistrerFicheUtilisateur = False

    ' Validations
    If Trim(txtNom.Text) = "" Then
        MsgBox "Le nom est obligatoire.", vbExclamation: txtNom.SetFocus: Exit Function
    End If
    If Trim(cboProjet.Text) = "" Then
        MsgBox "Le projet est obligatoire.", vbExclamation: cboProjet.SetFocus: Exit Function
    End If
    If optCongeOui.Value Then
        If Not IsDate(txtCongeD.Text) Or Not IsDate(txtCongeF.Text) Then
            MsgBox "Dates de conge invalides (format jj/mm/aaaa).", vbExclamation: Exit Function
        End If
    End If
    If optTTOui.Value Then
        If Not IsDate(txtTTD.Text) Or Not IsDate(txtTTF.Text) Then
            MsgBox "Dates TT invalides (format jj/mm/aaaa).", vbExclamation: Exit Function
        End If
    End If
    If optMaladieOui.Value Then
        If Not IsDate(txtDateArret.Text) Or Not IsDate(txtDateReprise.Text) Then
            MsgBox "Dates de maladie invalides (format jj/mm/aaaa).", vbExclamation: Exit Function
        End If
    End If
    If optCDD.Value And Trim(txtDateExpiration.Text) <> "" Then
        If Not IsDate(txtDateExpiration.Text) Then
            MsgBox "Date d'expiration invalide (format jj/mm/aaaa).", vbExclamation: Exit Function
        End If
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Utilisateurs")

    ws.Cells(lr, 1).Value = Trim(txtNom.Text)
    ws.Cells(lr, 2).Value = Trim(cboProjet.Text)
    ws.Cells(lr, 3).Value = Trim(txtVille.Text)
    ws.Cells(lr, 4).Value = Trim(txtZone.Text)

    ' Conge (col 5-8)
    ws.Cells(lr, 5).Value = IIf(optCongeOui.Value, "OUI", "NON")
    If optCongeOui.Value And IsDate(txtCongeD.Text) Then
        ws.Cells(lr, 6).Value = CDate(txtCongeD.Text)
        ws.Cells(lr, 6).NumberFormat = "dd/mm/yyyy"
    Else
        ws.Cells(lr, 6).Value = ""
    End If
    If optCongeOui.Value And IsDate(txtCongeF.Text) Then
        ws.Cells(lr, 7).Value = CDate(txtCongeF.Text)
        ws.Cells(lr, 7).NumberFormat = "dd/mm/yyyy"
    Else
        ws.Cells(lr, 7).Value = ""
    End If
    ws.Cells(lr, 8).Value = IIf(optCongeOui.Value, Trim(cboTypeConge.Text), "")

    ' Transport (col 9)
    ws.Cells(lr, 9).Value = cboTransport.Text

    ' TT (col 10-12)
    ws.Cells(lr, 10).Value = IIf(optTTOui.Value, "OUI", "NON")
    If optTTOui.Value And IsDate(txtTTD.Text) Then
        ws.Cells(lr, 11).Value = CDate(txtTTD.Text)
        ws.Cells(lr, 11).NumberFormat = "dd/mm/yyyy"
    Else
        ws.Cells(lr, 11).Value = ""
    End If
    If optTTOui.Value And IsDate(txtTTF.Text) Then
        ws.Cells(lr, 12).Value = CDate(txtTTF.Text)
        ws.Cells(lr, 12).NumberFormat = "dd/mm/yyyy"
    Else
        ws.Cells(lr, 12).Value = ""
    End If

    ' Renforts (col 13-14) -> reflete le choix exclusif des OptionButton
    ws.Cells(lr, 13).Value = IIf(optRenfortPress.Value, "OUI", "NON")
    ws.Cells(lr, 14).Value = IIf(optRenfortCofit.Value, "OUI", "NON")

    ' Matricule / telephone / date embauche (col 15/16/18) : conserves
    ' tels quels (non saisis sur ce formulaire). Pour une nouvelle
    ' fiche, restent vides jusqu'a saisie directe dans la feuille.
    ws.Cells(lr, 15).Value = m_matricule
    ws.Cells(lr, 16).Value = m_telephone
    ws.Cells(lr, 18).Value = m_dateEmbauche
    If IsDate(m_dateEmbauche) Then ws.Cells(lr, 18).NumberFormat = "dd/mm/yyyy"

    ' Contrat (col 17 + 19)
    ws.Cells(lr, 17).Value = IIf(optCDD.Value, "CDD", "CDI")
    If optCDD.Value And IsDate(txtDateExpiration.Text) Then
        ws.Cells(lr, 19).Value = CDate(txtDateExpiration.Text)
        ws.Cells(lr, 19).NumberFormat = "dd/mm/yyyy"
    Else
        ws.Cells(lr, 19).Value = ""
    End If

    ' Maladie (col 20-22)
    ' Regle : Maladie = OUI equivaut a un statut "OFF" pour la
    ' periode d'arret (voir aussi cmdMettreAJour_Click, qui force les
    ' jours du planning concernes a "OFF").
    ws.Cells(lr, 20).Value = IIf(optMaladieOui.Value, "OUI", "NON")
    If optMaladieOui.Value And IsDate(txtDateArret.Text) Then
        ws.Cells(lr, 21).Value = CDate(txtDateArret.Text)
        ws.Cells(lr, 21).NumberFormat = "dd/mm/yyyy"
    Else
        ws.Cells(lr, 21).Value = ""
    End If
    If optMaladieOui.Value And IsDate(txtDateReprise.Text) Then
        ws.Cells(lr, 22).Value = CDate(txtDateReprise.Text)
        ws.Cells(lr, 22).NumberFormat = "dd/mm/yyyy"
    Else
        ws.Cells(lr, 22).Value = ""
    End If

    EnregistrerFicheUtilisateur = True
End Function

' --- Sauvegarder (fiche collaborateur) ---
Private Sub cmdSauver_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Utilisateurs")
    Dim lr As Long

    If m_modeAjout Then
        ' Nouvelle ligne apres la derniere
        lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Else
        lr = m_ligneSelectionnee
    End If

    If Not EnregistrerFicheUtilisateur(lr) Then Exit Sub

    ChargerListe
    VerrouillerFormulaire True

    Dim etaitAjout As Boolean: etaitAjout = m_modeAjout
    m_modeAjout = False
    m_ligneSelectionnee = lr

    MsgBox "Collaborateur " & IIf(etaitAjout, "ajoute", "modifie") & " avec succes !", vbInformation
End Sub

Private Sub cmdAnnuler_Click()
    If m_ligneSelectionnee > 0 Then
        lstCollabs_Click
    Else
        ViderFormulaire
        ViderPlanningSemaine
        cmdMettreAJour.Enabled = False
    End If
    m_modeAjout = False
    VerrouillerFormulaire True
End Sub

Private Sub cmdFermer_Click()
    Unload Me
End Sub

' --- Helpers fiche collaborateur ---
Private Sub ViderFormulaire()
    txtNom.Text = "": cboProjet.Text = "": txtVille.Text = "": txtZone.Text = ""
    optCongeNon.Value = True: txtCongeD.Text = "": txtCongeF.Text = "": cboTypeConge.Text = ""
    cboTransport.Text = "NON"
    optTTNon.Value = True: txtTTD.Text = "": txtTTF.Text = ""
    optRenfortAucun.Value = True
    optCDI.Value = True: txtDateExpiration.Text = ""
    optMaladieNon.Value = True: txtDateArret.Text = "": txtDateReprise.Text = ""
    m_matricule = "": m_telephone = "": m_dateEmbauche = ""
End Sub

Private Sub VerrouillerFormulaire(bVerrouille As Boolean)
    txtNom.Enabled = Not bVerrouille
    cboProjet.Enabled = Not bVerrouille
    txtVille.Enabled = Not bVerrouille
    txtZone.Enabled = Not bVerrouille
    optCongeOui.Enabled = Not bVerrouille
    optCongeNon.Enabled = Not bVerrouille
    txtCongeD.Enabled = Not bVerrouille
    txtCongeF.Enabled = Not bVerrouille
    cboTypeConge.Enabled = Not bVerrouille
    cboTransport.Enabled = Not bVerrouille
    optTTOui.Enabled = Not bVerrouille
    optTTNon.Enabled = Not bVerrouille
    txtTTD.Enabled = Not bVerrouille
    txtTTF.Enabled = Not bVerrouille
    optRenfortAucun.Enabled = Not bVerrouille
    optRenfortPress.Enabled = Not bVerrouille
    optRenfortCofit.Enabled = Not bVerrouille
    optCDI.Enabled = Not bVerrouille
    optCDD.Enabled = Not bVerrouille
    txtDateExpiration.Enabled = Not bVerrouille
    optMaladieOui.Enabled = Not bVerrouille
    optMaladieNon.Enabled = Not bVerrouille
    txtDateArret.Enabled = Not bVerrouille
    txtDateReprise.Enabled = Not bVerrouille
    cmdSauver.Enabled = Not bVerrouille
    cmdAnnuler.Enabled = Not bVerrouille
End Sub

' Reapplique les activations/desactivations dependantes des choix
' OUI/NON (conge, TT, maladie) et CDI/CDD, par exemple juste apres
' avoir deverrouille le formulaire pour edition.
Private Sub AppliquerEtatsConditionnels()
    txtCongeD.Enabled = optCongeOui.Value
    txtCongeF.Enabled = optCongeOui.Value
    cboTypeConge.Enabled = optCongeOui.Value
    txtTTD.Enabled = optTTOui.Value
    txtTTF.Enabled = optTTOui.Value
    txtDateArret.Enabled = optMaladieOui.Value
    txtDateReprise.Enabled = optMaladieOui.Value
    txtDateExpiration.Enabled = optCDD.Value
End Sub

' --- Active/desactive les champs dependants selon les OptionButton ---
Private Sub optCongeOui_Click()
    txtCongeD.Enabled = True
    txtCongeF.Enabled = True
    cboTypeConge.Enabled = True
End Sub

Private Sub optCongeNon_Click()
    txtCongeD.Enabled = False: txtCongeD.Text = ""
    txtCongeF.Enabled = False: txtCongeF.Text = ""
    cboTypeConge.Enabled = False: cboTypeConge.Text = ""
End Sub

Private Sub optTTOui_Click()
    txtTTD.Enabled = True
    txtTTF.Enabled = True
End Sub

Private Sub optTTNon_Click()
    txtTTD.Enabled = False: txtTTD.Text = ""
    txtTTF.Enabled = False: txtTTF.Text = ""
End Sub

Private Sub optMaladieOui_Click()
    txtDateArret.Enabled = True
    txtDateReprise.Enabled = True
End Sub

Private Sub optMaladieNon_Click()
    txtDateArret.Enabled = False: txtDateArret.Text = ""
    txtDateReprise.Enabled = False: txtDateReprise.Text = ""
End Sub

Private Sub optCDI_Click()
    txtDateExpiration.Enabled = False
    txtDateExpiration.Text = ""
End Sub

Private Sub optCDD_Click()
    txtDateExpiration.Enabled = True
End Sub

' Ouverture rapide depuis le ruban / macro
Public Sub Ouvrir()
    UFUtilisateurs.Show
End Sub

' ============================================================
' BLOC "PLANNING SEMAINE"
' ============================================================

Private Function NomJour(j As Integer) As String
    Select Case j
        Case 1: NomJour = "Lundi"
        Case 2: NomJour = "Mardi"
        Case 3: NomJour = "Mercredi"
        Case 4: NomJour = "Jeudi"
        Case 5: NomJour = "Vendredi"
        Case 6: NomJour = "Samedi"
        Case 7: NomJour = "Dimanche"
        Case Else: NomJour = ""
    End Select
End Function

Private Sub ViderPlanningSemaine()
    Dim j As Integer
    For j = 1 To 7
        Me.Controls("txtEntree" & j).Text = ""
        Me.Controls("txtSortie" & j).Text = ""
        Me.Controls("txtPauseD" & j).Text = ""
        Me.Controls("txtPauseF" & j).Text = ""
        Me.Controls("txtComment" & j).Text = ""
    Next j
End Sub

' Charge, pour le collaborateur donne, le planning de la semaine
' actuellement ciblee (BOOM.g_LundiCible, definie via UFGenerer)
Private Sub ChargerPlanningSemaine(nomComplet As String)
    ViderPlanningSemaine
    If Trim(nomComplet) = "" Then Exit Sub
    If Not BOOM.FeuilleExiste("PLANNING") Then Exit Sub
    If Not BOOM.FeuilleExiste("CONSOLIDATION") Then Exit Sub

    Dim wsP As Worksheet, wsC As Worksheet
    Set wsP = ThisWorkbook.Sheets("PLANNING")
    Set wsC = ThisWorkbook.Sheets("CONSOLIDATION")

    Dim lundiCible As Date: lundiCible = BOOM.LundiSemaine()
    Dim semCible As Integer
    semCible = Application.WorksheetFunction.WeekNum(lundiCible, 2)

    ' --- Entree / Sortie depuis PLANNING (1 ligne / collab / semaine) ---
    Dim lastP As Long: lastP = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    Dim lrP As Long: lrP = 0
    Dim i As Long
    For i = 2 To lastP
        If CStr(wsP.Cells(i, 1).Value) = CStr(semCible) And wsP.Cells(i, 5).Value = nomComplet Then
            lrP = i: Exit For
        End If
    Next i

    Dim j As Integer
    If lrP > 0 Then
        For j = 1 To 7
            Dim colE As Integer: colE = 10 + (j * 2)
            Dim colS As Integer: colS = colE + 1
            Dim vE As String: vE = FormatValeurPlanning(wsP.Cells(lrP, colE).Value)
            Dim vS As String: vS = FormatValeurPlanning(wsP.Cells(lrP, colS).Value)
            If Left(vE, 3) = "TT " Then vE = Mid(vE, 4)
            If Left(vS, 3) = "TT " Then vS = Mid(vS, 4)
            Me.Controls("txtEntree" & j).Text = vE
            Me.Controls("txtSortie" & j).Text = vS
        Next j
    End If

    ' --- Pause / Commentaire depuis CONSOLIDATION (1 ligne / collab / jour) ---
    Dim lastC As Long: lastC = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
    For j = 1 To 7
        Dim dJour As Date: dJour = lundiCible + (j - 1)
        Dim k As Long
        For k = 2 To lastC
            If wsC.Cells(k, 1).Value = nomComplet And IsDate(wsC.Cells(k, 2).Value) Then
                If CLng(CDate(wsC.Cells(k, 2).Value)) = CLng(dJour) Then
                    Me.Controls("txtPauseD" & j).Text = FormatValeurPlanning(wsC.Cells(k, 5).Value)
                    Me.Controls("txtPauseF" & j).Text = FormatValeurPlanning(wsC.Cells(k, 6).Value)
                    Me.Controls("txtComment" & j).Text = CStr(wsC.Cells(k, 15).Value)
                    Exit For
                End If
            End If
        Next k
    Next j
End Sub

' Formate une valeur de cellule PLANNING/CONSOLIDATION pour l'affichage
' dans les TextBox : si la cellule contient une heure Excel (fraction
' de jour, ex 0,375), on la convertit en "hh:mm". Sinon (texte du
' type "OFF", "CONGE", ou une chaine deja au format hh:mm), on la
' laisse telle quelle. Corrige l'affichage de type "0,375" au lieu
' de "09:00" dans le planning.
Private Function FormatValeurPlanning(v As Variant) As String
    On Error GoTo Secours
    If IsEmpty(v) Or IsNull(v) Then
        FormatValeurPlanning = ""
        Exit Function
    End If
    If VarType(v) = vbString Then
        FormatValeurPlanning = CStr(v)
        Exit Function
    End If
    If IsNumeric(v) Then
        Dim d As Double: d = CDbl(v)
        If d = 0 Then
            FormatValeurPlanning = ""
        Else
            FormatValeurPlanning = Format(d, "hh:mm")
        End If
        Exit Function
    End If
    FormatValeurPlanning = CStr(v)
    Exit Function
Secours:
    ' En cas de type inattendu (erreur, tableau, etc.), on ne bloque pas
    ' le chargement du planning : on affiche la valeur brute si possible.
    On Error Resume Next
    FormatValeurPlanning = CStr(v)
    On Error GoTo 0
End Function

Private Function EstHeureValide(h As String) As Boolean
    If h = "" Then EstHeureValide = True: Exit Function
    If InStr(h, ":") = 0 Then EstHeureValide = False: Exit Function
    Dim p() As String
    p = Split(h, ":")
    If UBound(p) <> 1 Then EstHeureValide = False: Exit Function
    If Not IsNumeric(p(0)) Or Not IsNumeric(p(1)) Then EstHeureValide = False: Exit Function
    EstHeureValide = (CInt(p(0)) >= 0 And CInt(p(0)) <= 23 And CInt(p(1)) >= 0 And CInt(p(1)) <= 59)
End Function

Private Sub ColorerCellulePlanning(cel As Range)
    cel.HorizontalAlignment = xlCenter
    Select Case True
        Case cel.Value = "OFF"
            cel.Interior.Color = RGB(255, 199, 206): cel.Font.Color = RGB(192, 0, 0): cel.Font.Bold = True
        Case cel.Value = "CONGE"
            cel.Interior.Color = RGB(255, 230, 153): cel.Font.Color = RGB(156, 87, 0): cel.Font.Bold = True
        Case Left(CStr(cel.Value), 2) = "TT"
            cel.Interior.Color = RGB(220, 190, 255): cel.Font.Color = RGB(70, 0, 130): cel.Font.Bold = False
        Case Else
            cel.Interior.ColorIndex = xlNone: cel.Font.Color = RGB(0, 0, 0): cel.Font.Bold = False
    End Select
End Sub

' Met a jour la cellule du jour dans la feuille horizontale du projet
' (ex : AFEDIM, TLV...). Ne fait rien si le collaborateur n'y figure pas.
Private Sub MettreAJourCelluleProjet(ws As Worksheet, nomComplet As String, j As Integer, _
                                      entree As String, sortie As String, pd As String, pf As String)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long, lr As Long: lr = 0
    For r = 4 To lastRow   ' les donnees commencent apres l'entete (lignes 1-3)
        If ws.Cells(r, 1).Value = nomComplet Then lr = r: Exit For
    Next r
    If lr = 0 Then Exit Sub

    Dim cel As Range: Set cel = ws.Cells(lr, 3 + j)
    cel.Value = BOOM.FormatCelluleJour(entree, sortie, pd, pf)
    cel.WrapText = True
    cel.HorizontalAlignment = xlCenter
    cel.VerticalAlignment = xlCenter

    Select Case True
        Case entree = "OFF"
            cel.Interior.Color = RGB(255, 199, 206): cel.Font.Bold = True: cel.Font.Color = RGB(192, 0, 0)
        Case entree = "CONGE"
            cel.Interior.Color = RGB(255, 192, 0): cel.Font.Bold = True: cel.Font.Color = RGB(0, 0, 0)
        Case Else
            cel.Font.Color = RGB(0, 0, 0): cel.Font.Bold = False
            cel.Interior.Color = IIf(lr Mod 2 = 0, RGB(235, 241, 255), RGB(255, 255, 255))
    End Select
End Sub

' Met a jour (ou cree) la ligne ROTATION du collaborateur : date de
' derniere MAJ manuelle, semaine, et compteur de renforts si un
' commentaire "RENFORT" a ete saisi sur un jour de la semaine.
Private Sub MajLigneRotation(nomComplet As String, projet As String, semCible As Integer, aRenfort As Boolean)
    If Not BOOM.FeuilleExiste("ROTATION") Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("ROTATION")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long, lr As Long: lr = 0
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = nomComplet And ws.Cells(i, 2).Value = projet Then lr = i: Exit For
    Next i
    If lr = 0 Then
        lr = lastRow + 1
        ws.Cells(lr, 1).Value = nomComplet
        ws.Cells(lr, 2).Value = projet
        ws.Cells(lr, 3).Value = 0
        ws.Cells(lr, 6).Value = 0
    End If
    ws.Cells(lr, 4).Value = Format(Now, "dd/mm/yyyy hh:mm")
    ws.Cells(lr, 5).Value = semCible
    If aRenfort Then
        Dim v As Variant: v = ws.Cells(lr, 6).Value
        ws.Cells(lr, 6).Value = IIf(IsNumeric(v), CInt(v) + 1, 1)
    End If
End Sub

' Indique si le jour donne tombe dans la periode d'arret maladie
' actuellement affichee sur le formulaire (Maladie = OUI + dates
' d'arret/reprise valides). Utilisee par cmdMettreAJour_Click pour
' forcer ces jours a "OFF" dans le planning (regle : Maladie OUI = OFF).
Private Function EstJourMaladie(dJourCheck As Date) As Boolean
    EstJourMaladie = False
    If Not optMaladieOui.Value Then Exit Function
    If Not IsDate(txtDateArret.Text) Or Not IsDate(txtDateReprise.Text) Then Exit Function
    Dim dA As Date: dA = CDate(txtDateArret.Text)
    Dim dR As Date: dR = CDate(txtDateReprise.Text)
    EstJourMaladie = (dJourCheck >= dA And dJourCheck <= dR)
End Function

' --- Bouton "Mettre a jour" ---
Private Sub cmdMettreAJour_Click()
    If lstCollabs.ListIndex < 0 Then
        MsgBox "Selectionnez d'abord un collaborateur.", vbExclamation: Exit Sub
    End If
    If Not BOOM.FeuilleExiste("PLANNING") Or Not BOOM.FeuilleExiste("CONSOLIDATION") Then
        MsgBox "Les feuilles PLANNING / CONSOLIDATION sont introuvables." & Chr(10) & _
               "Lancez d'abord 'Generer Planning' depuis UFGenerer.", vbCritical
        Exit Sub
    End If

    ' Enregistre d'abord les eventuelles modifications des Informations
    ' collaborateur (Ville, Zone, Conge, TT, Renfort, Contrat, Maladie...)
    ' afin que "Mettre a jour" reflete aussi ces changements, pas
    ' seulement la grille Planning Semaine.
    If Not EnregistrerFicheUtilisateur(m_ligneSelectionnee) Then Exit Sub
    Dim idxListeAvant As Long: idxListeAvant = m_ligneSelectionnee - 2
    ChargerListe
    If idxListeAvant >= 0 And idxListeAvant < lstCollabs.ListCount Then lstCollabs.ListIndex = idxListeAvant

    Dim nomComplet As String: nomComplet = Trim(txtNom.Text)
    Dim nomProjet As String: nomProjet = Trim(cboProjet.Text)

    Dim wsP As Worksheet, wsC As Worksheet, wsProjet As Worksheet
    Set wsP = ThisWorkbook.Sheets("PLANNING")
    Set wsC = ThisWorkbook.Sheets("CONSOLIDATION")
    If BOOM.FeuilleExiste(nomProjet) Then Set wsProjet = ThisWorkbook.Sheets(nomProjet)

    Dim lundiCible As Date: lundiCible = BOOM.LundiSemaine()
    Dim semCible As Integer
    semCible = Application.WorksheetFunction.WeekNum(lundiCible, 2)

    ' Retrouver la ligne PLANNING pour ce collaborateur + cette semaine
    Dim lastP As Long: lastP = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    Dim lrP As Long: lrP = 0
    Dim i As Long
    For i = 2 To lastP
        If CStr(wsP.Cells(i, 1).Value) = CStr(semCible) And wsP.Cells(i, 5).Value = nomComplet Then
            lrP = i: Exit For
        End If
    Next i
    If lrP = 0 Then
        MsgBox "Aucun planning genere pour ce collaborateur sur la semaine " & semCible & "." & Chr(10) & _
               "Lancez d'abord 'Generer Planning' depuis UFGenerer.", vbExclamation
        Exit Sub
    End If

    ' Validation de toutes les heures saisies avant toute ecriture
    ' (les jours couverts par un arret maladie sont ignores : ils
    ' seront de toute facon forces a "OFF" plus bas)
    Dim j As Integer
    For j = 1 To 7
        If Not EstJourMaladie(lundiCible + (j - 1)) Then
            Dim eTest As String: eTest = UCase(Trim(Me.Controls("txtEntree" & j).Text))
            Dim sTest As String: sTest = UCase(Trim(Me.Controls("txtSortie" & j).Text))
            If eTest <> "" And eTest <> "OFF" And eTest <> "CONGE" And Not EstHeureValide(eTest) Then
                MsgBox "Heure d'entree invalide (" & NomJour(j) & ") : " & eTest, vbExclamation
                Exit Sub
            End If
            If sTest <> "" And sTest <> "OFF" And sTest <> "CONGE" And Not EstHeureValide(sTest) Then
                MsgBox "Heure de sortie invalide (" & NomJour(j) & ") : " & sTest, vbExclamation
                Exit Sub
            End If
        End If
    Next j

    Dim aRenfortCetteSemaine As Boolean: aRenfortCetteSemaine = False

    For j = 1 To 7
        Dim entree As String: entree = UCase(Trim(Me.Controls("txtEntree" & j).Text))
        Dim sortie As String: sortie = UCase(Trim(Me.Controls("txtSortie" & j).Text))
        Dim pd As String: pd = Trim(Me.Controls("txtPauseD" & j).Text)
        Dim pf As String: pf = Trim(Me.Controls("txtPauseF" & j).Text)
        Dim commentaire As String: commentaire = Trim(Me.Controls("txtComment" & j).Text)

        Dim dJour As Date: dJour = lundiCible + (j - 1)
        Dim colE As Integer: colE = 10 + (j * 2)
        Dim colS As Integer: colS = colE + 1

        ' Regle : Maladie = OUI equivaut a "OFF" pour les jours de la
        ' periode d'arret, quelle que soit la saisie du planning.
        If EstJourMaladie(dJour) Then
            entree = "OFF"
            sortie = "OFF"
        End If

        ' --- PLANNING ---
        wsP.Cells(lrP, colE).Value = entree
        wsP.Cells(lrP, colS).Value = sortie
        ColorerCellulePlanning wsP.Cells(lrP, colE)
        ColorerCellulePlanning wsP.Cells(lrP, colS)

        ' --- CONSOLIDATION (creer la ligne du jour si elle n'existe pas) ---
        Dim lastC As Long: lastC = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
        Dim lrC As Long: lrC = 0
        Dim k As Long
        For k = 2 To lastC
            If wsC.Cells(k, 1).Value = nomComplet And IsDate(wsC.Cells(k, 2).Value) Then
                If CLng(CDate(wsC.Cells(k, 2).Value)) = CLng(dJour) Then lrC = k: Exit For
            End If
        Next k
        If lrC = 0 Then
            lrC = lastC + 1
            wsC.Cells(lrC, 1).Value = nomComplet
            wsC.Cells(lrC, 2).Value = dJour
            wsC.Cells(lrC, 2).NumberFormat = "dd/mm/yyyy"
            wsC.Cells(lrC, 7).Value = semCible
            wsC.Cells(lrC, 11).Value = txtVille.Text
            wsC.Cells(lrC, 12).Value = txtZone.Text
        End If

        Dim pauseOK As Boolean: pauseOK = (entree <> "OFF" And entree <> "CONGE" And entree <> "")
        wsC.Cells(lrC, 3).Value = entree
        wsC.Cells(lrC, 4).Value = sortie
        wsC.Cells(lrC, 5).Value = IIf(pauseOK, pd, "")
        wsC.Cells(lrC, 6).Value = IIf(pauseOK, pf, "")

        Dim activite As String
        Select Case True
            Case entree = "CONGE": activite = "CONGE"
            Case entree = "OFF" Or entree = "": activite = "OFF"
            Case InStr(UCase(commentaire), "RENFORT") > 0
                activite = "RENFORT": aRenfortCetteSemaine = True
            Case Else: activite = nomProjet
        End Select
        wsC.Cells(lrC, 8).Value = activite
        wsC.Cells(lrC, 15).Value = commentaire

        ' --- Feuille projet (affichage horizontal) ---
        If Not wsProjet Is Nothing Then
            MettreAJourCelluleProjet wsProjet, nomComplet, j, entree, sortie, pd, pf
        End If
    Next j

    ' Recalcule les cumuls NB HEURE / NB JOUR dans CONSOLIDATION
    BOOM.CalculerCumulsSemaine

    ' Met a jour ROTATION (tracabilite + nb renforts)
    MajLigneRotation nomComplet, nomProjet, semCible, aRenfortCetteSemaine

    MsgBox "Planning de la semaine " & semCible & " mis a jour pour " & nomComplet & ".", vbInformation
End Sub
