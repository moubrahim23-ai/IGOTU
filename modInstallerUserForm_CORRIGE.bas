Attribute VB_Name = "modInstallerUserForm"
Option Explicit

Private Const TYPE_COMPOSANT_USERFORM As Long = 3
Private Const STYLE_LISTE_DEROULANTE As Long = 2
Private Const ALIGNEMENT_CENTRE As Long = 2
Private Const ALIGNEMENT_GAUCHE As Long = 1
Private Const BORDURE_SIMPLE As Long = 1
Private Const EFFET_PLAT As Long = 0
Private Const FICHIER_CODE_USERFORM As String = "UserForm1_Code.txt"
Private Const VERSION_INSTALLATEUR As String = "1.1 - LTSC 2021"

Private EtapeInstallation As String

Public Sub InstallerUserFormComplet()
    Dim projet As Object
    Dim composantExistant As Object
    Dim nouveauComposant As Object
    Dim cheminCode As String
    Dim cheminSauvegarde As String
    Dim ancienSupprime As Boolean
    Dim nouveauCree As Boolean
    Dim reponse As VbMsgBoxResult
    Dim numeroErreur As Long
    Dim descriptionErreur As String

    EtapeInstallation = "acces au projet VBA"
    On Error GoTo AccesRefuse
    Set projet = ThisWorkbook.VBProject
    On Error GoTo GestionErreur

    If Not ClasseurMacrosActif Then
        MsgBox "Enregistrez d'abord ce classeur au format .xlsm, puis relancez " & _
            "InstallerUserFormComplet.", vbExclamation, "Classeur avec macros requis"
        Exit Sub
    End If

    EtapeInstallation = "recherche du code du formulaire"
    cheminCode = TrouverFichierCode
    If Len(cheminCode) = 0 Then Exit Sub

    EtapeInstallation = "verification du formulaire existant"
    Set composantExistant = TrouverComposant(projet, "UserForm1")
    If Not composantExistant Is Nothing Then
        reponse = MsgBox("UserForm1 existe deja." & vbCrLf & vbCrLf & _
            "L'installateur peut l'exporter comme sauvegarde, puis le remplacer " & _
            "par la version organisee." & vbCrLf & vbCrLf & _
            "Continuer ?", vbYesNo + vbQuestion + vbDefaultButton2, _
            "Remplacer UserForm1")
        If reponse <> vbYes Then Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    EtapeInstallation = "fermeture du formulaire existant"
    FermerInstanceUserForm "UserForm1"

    If Not composantExistant Is Nothing Then
        EtapeInstallation = "sauvegarde du formulaire existant"
        cheminSauvegarde = ThisWorkbook.Path & Application.PathSeparator & _
            "Sauvegarde_UserForm1_" & Format$(Now, "yyyymmdd_hhnnss") & ".frm"
        composantExistant.Export cheminSauvegarde
        projet.VBComponents.Remove composantExistant
        ancienSupprime = True
    End If

    EtapeInstallation = "creation du composant UserForm1"
    Set nouveauComposant = projet.VBComponents.Add(TYPE_COMPOSANT_USERFORM)
    nouveauComposant.Name = "UserForm1"
    nouveauCree = True

    EtapeInstallation = "construction de l'interface"
    ConstruireInterface nouveauComposant.Designer
    EtapeInstallation = "injection du code du formulaire"
    InjecterCode nouveauComposant, cheminCode

    EtapeInstallation = "enregistrement du classeur"
    ThisWorkbook.Save

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "UserForm1 a ete cree avec une mise en page organisee." & vbCrLf & _
        "Installateur " & VERSION_INSTALLATEUR & vbCrLf & vbCrLf & _
        "- saisie a gauche" & vbCrLf & _
        "- file d'attente en temps reel a droite" & vbCrLf & _
        "- synthese et actions regroupees" & vbCrLf & _
        "- bouton Dashboard integre" & vbCrLf & vbCrLf & _
        IIf(Len(cheminSauvegarde) > 0, _
            "L'ancien formulaire a ete sauvegarde dans :" & vbCrLf & cheminSauvegarde, _
            "Vous pouvez maintenant lancer OuvrirFormulaireSaisie."), _
        vbInformation, "Installation terminee"
    Exit Sub

AccesRefuse:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Excel bloque l'acces au projet VBA." & vbCrLf & vbCrLf & _
        "Activez : Fichier > Options > Centre de gestion de la confidentialite > " & _
        "Parametres du Centre > Parametres des macros > Acces approuve au modele " & _
        "d'objet du projet VBA." & vbCrLf & vbCrLf & _
        "Fermez puis rouvrez Excel avant de relancer InstallerUserFormComplet.", _
        vbExclamation, "Autorisation VBA requise"
    Exit Sub

GestionErreur:
    numeroErreur = Err.Number
    descriptionErreur = Err.Description
    On Error Resume Next
    If nouveauCree Then
        Set nouveauComposant = TrouverComposant(projet, "UserForm1")
        If Not nouveauComposant Is Nothing Then projet.VBComponents.Remove nouveauComposant
    End If
    If ancienSupprime And Len(cheminSauvegarde) > 0 Then
        If Len(Dir$(cheminSauvegarde)) > 0 Then projet.VBComponents.Import cheminSauvegarde
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0

    MsgBox "Le UserForm n'a pas pu etre installe : " & numeroErreur & " - " & _
        descriptionErreur & vbCrLf & _
        "Etape : " & EtapeInstallation & vbCrLf & vbCrLf & _
        "Aucune ancienne version n'a ete perdue ; une sauvegarde .frm est conservee " & _
        "lorsqu'un formulaire existait.", vbCritical, "Installation interrompue"
End Sub

Private Function ClasseurMacrosActif() As Boolean
    Dim extension As String

    extension = LCase$(Mid$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") + 1))
    ClasseurMacrosActif = (extension = "xlsm" Or extension = "xlsb")
End Function

Private Function TrouverFichierCode() As String
    Dim chemin As String

    chemin = ThisWorkbook.Path & Application.PathSeparator & FICHIER_CODE_USERFORM
    If Len(Dir$(chemin)) > 0 Then
        TrouverFichierCode = chemin
        Exit Function
    End If

    With Application.FileDialog(3)
        .Title = "Selectionnez " & FICHIER_CODE_USERFORM
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Code du UserForm", "*.txt"
        If .Show = -1 Then TrouverFichierCode = .SelectedItems(1)
    End With

    If Len(TrouverFichierCode) = 0 Then
        MsgBox "Installation annulee : " & FICHIER_CODE_USERFORM & _
            " n'a pas ete selectionne.", vbInformation, "Fichier requis"
    End If
End Function

Private Function TrouverComposant(ByVal projet As Object, ByVal nomComposant As String) As Object
    Dim composant As Object

    For Each composant In projet.VBComponents
        If StrComp(composant.Name, nomComposant, vbTextCompare) = 0 Then
            Set TrouverComposant = composant
            Exit Function
        End If
    Next composant
End Function

Private Sub FermerInstanceUserForm(ByVal nomFormulaire As String)
    Dim formulaire As Object

    On Error Resume Next
    For Each formulaire In VBA.UserForms
        If StrComp(formulaire.Name, nomFormulaire, vbTextCompare) = 0 Then Unload formulaire
    Next formulaire
    On Error GoTo 0
End Sub

Private Sub InjecterCode(ByVal composant As Object, ByVal cheminCode As String)
    With composant.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromFile cheminCode
    End With
End Sub

Private Sub ConstruireInterface(ByVal formulaire As Object)
    Dim fraContexte As Object
    Dim fraSaisie As Object
    Dim fraAttente As Object
    Dim fraResume As Object
    Dim fraActions As Object
    Dim controle As Object

    EtapeInstallation = "configuration generale du formulaire"
    ' Configuration tolérante aux différences entre versions de MSForms.
    DefinirProprieteOptionnelle formulaire, "Caption", "Saisie des realisations"
    DefinirProprieteOptionnelle formulaire, "Width", 930
    DefinirProprieteOptionnelle formulaire, "Height", 590
    DefinirProprieteOptionnelle formulaire, "BackColor", RGB(241, 245, 249)
    DefinirProprieteOptionnelle formulaire, "StartUpPosition", 1

    ' ShowModal n'est volontairement pas modifie ici : selon la version de
    ' MSForms, cette propriete de conception peut provoquer l'erreur 438.
    ' Le formulaire est deja ouvert en mode non modal par OuvrirFormulaireSaisie.

    EtapeInstallation = "creation du bandeau"
    Set controle = AjouterLabel(formulaire, "lblTitre", _
        "SAISIE DES REALISATIONS", 0, 0, 918, 38, 18, True, RGB(255, 255, 255), RGB(23, 54, 93))
    controle.TextAlign = ALIGNEMENT_CENTRE

    Set controle = AjouterLabel(formulaire, "lblSousTitre", _
        "Ajout temporaire  |  validation CSV  |  suivi en temps reel", _
        0, 38, 918, 17, 9, False, RGB(23, 54, 93), RGB(217, 234, 247))
    controle.TextAlign = ALIGNEMENT_CENTRE

    EtapeInstallation = "creation du contexte agent"
    Set fraContexte = AjouterCadre(formulaire, "fraContexte", "Contexte agent", 12, 62, 285, 86)
    AjouterLibelleChamp fraContexte, "lblAgent", "Agent", 10, 21, 70
    Set controle = AjouterZoneTexte(fraContexte, "txtPersoID", 82, 17, 190, 22, True, False)
    controle.TabStop = False

    AjouterLibelleChamp fraContexte, "lblDateHeure", "Horodatage", 10, 52, 70
    Set controle = AjouterZoneTexte(fraContexte, "txtDateTime", 82, 48, 190, 22, True, False)
    controle.TabStop = False

    EtapeInstallation = "creation des champs de saisie"
    Set fraSaisie = AjouterCadre(formulaire, "fraSaisie", "Nouvelle realisation", 12, 156, 285, 344)

    AjouterLibelleChamp fraSaisie, "lblNumeroDossier", "No dossier *", 10, 23, 78
    Set controle = AjouterZoneTexte(fraSaisie, "txtNumeroDossier", 92, 19, 180, 22, False, False)
    controle.TabIndex = 0
    DefinirProprieteOptionnelle controle, "ControlTipText", "Numero de dossier obligatoire"

    AjouterLibelleChamp fraSaisie, "lblActivite", "Activite *", 10, 55, 78
    Set controle = AjouterListeDeroulante(fraSaisie, "cboActivite", 92, 51, 180, 22)
    controle.TabIndex = 1

    AjouterLibelleChamp fraSaisie, "lblProduit", "Produit *", 10, 87, 78
    Set controle = AjouterListeDeroulante(fraSaisie, "cboProduit", 92, 83, 180, 22)
    controle.TabIndex = 2

    AjouterLibelleChamp fraSaisie, "lblTypeDossier", "Type dossier", 10, 119, 78
    Set controle = AjouterListeDeroulante(fraSaisie, "cboTypeDossier", 92, 115, 180, 22)
    controle.TabIndex = 3

    AjouterLibelleChamp fraSaisie, "lblDecision", "Decision *", 10, 151, 78
    Set controle = AjouterListeDeroulante(fraSaisie, "cboDecision", 92, 147, 180, 22)
    controle.TabIndex = 4

    AjouterLibelleChamp fraSaisie, "lblDispo", "Dispo", 10, 183, 78
    Set controle = AjouterListeDeroulante(fraSaisie, "cboDispo", 92, 179, 180, 22)
    controle.TabIndex = 5

    AjouterLibelleChamp fraSaisie, "lblCommentaire", "Commentaire *", 10, 215, 100
    Set controle = AjouterZoneTexte(fraSaisie, "txtCommentaire", 10, 234, 262, 58, False, True)
    controle.TabIndex = 6
    DefinirProprieteOptionnelle controle, "MaxLength", 500
    DefinirProprieteOptionnelle controle, "ControlTipText", _
        "Commentaire obligatoire - 500 caracteres maximum"

    Set controle = AjouterBouton(fraSaisie, "btnValider", "Ajouter a la file d'attente", _
        10, 300, 262, 32, RGB(47, 117, 181), RGB(255, 255, 255))
    controle.TabIndex = 7
    DefinirProprieteOptionnelle controle, "Default", True
    DefinirProprieteOptionnelle controle, "ControlTipText", _
        "Ajoute uniquement dans DATA_Temp ; le CSV n'est pas modifie"

    EtapeInstallation = "creation de la file d'attente"
    Set fraAttente = AjouterCadre(formulaire, "fraAttente", _
        "File d'attente - non enregistree dans le CSV", 307, 62, 610, 298)
    Set controle = AjouterListe(fraAttente, "ListBox1", 10, 22, 588, 264)
    controle.ColumnCount = 10
    controle.ColumnWidths = "58;88;66;70;52;62;36;78;50;0"
    controle.TabIndex = 8
    DefinirProprieteOptionnelle controle, "ControlTipText", _
        "Selectionnez une ligne pour voir son commentaire dans la zone d'etat"

    EtapeInstallation = "creation de la synthese"
    Set fraResume = AjouterCadre(formulaire, "fraResume", "Synthese agent", 307, 370, 245, 130)
    Set controle = AjouterListe(fraResume, "ListBox2", 10, 22, 223, 96)
    controle.ColumnCount = 2
    controle.ColumnWidths = "125;75"
    controle.TabStop = False

    EtapeInstallation = "creation des boutons d'action"
    Set fraActions = AjouterCadre(formulaire, "fraActions", "Actions", 562, 370, 355, 130)
    Set controle = AjouterBouton(fraActions, "btnSupprimer", "Supprimer la selection", _
        10, 24, 160, 34, RGB(192, 0, 0), RGB(255, 255, 255))
    controle.TabIndex = 9
    DefinirProprieteOptionnelle controle, "ControlTipText", _
        "Supprime uniquement la ligne temporaire selectionnee"

    Set controle = AjouterBouton(fraActions, "btnEnregistrer", "Enregistrer dans le CSV", _
        178, 24, 164, 34, RGB(112, 173, 71), RGB(255, 255, 255))
    controle.TabIndex = 10
    DefinirProprieteOptionnelle controle, "ControlTipText", _
        "Valide les lignes en attente et optimise les doublons"

    Set controle = AjouterBouton(fraActions, "btnDashboard", "Ouvrir le Dashboard", _
        10, 70, 160, 34, RGB(47, 117, 181), RGB(255, 255, 255))
    controle.TabIndex = 11

    Set controle = AjouterBouton(fraActions, "btnFermer", "Fermer", _
        178, 70, 164, 34, RGB(89, 89, 89), RGB(255, 255, 255))
    controle.TabIndex = 12
    DefinirProprieteOptionnelle controle, "Cancel", True

    EtapeInstallation = "creation de la zone d'etat"
    Set controle = AjouterLabel(formulaire, "Label1", _
        "Pret. Ajoutez une realisation ; elle restera temporaire jusqu'a Enregistrer.", _
        307, 510, 610, 42, 9, False, RGB(31, 41, 55), RGB(248, 250, 252))
    DefinirProprieteOptionnelle controle, "BorderStyle", BORDURE_SIMPLE
    DefinirProprieteOptionnelle controle, "TextAlign", ALIGNEMENT_GAUCHE
    DefinirProprieteOptionnelle controle, "WordWrap", True
    DefinirProprieteOptionnelle controle, "SpecialEffect", EFFET_PLAT
End Sub

Private Function AjouterCadre(ByVal conteneur As Object, ByVal nom As String, _
        ByVal legende As String, ByVal gauche As Single, ByVal haut As Single, _
        ByVal largeur As Single, ByVal hauteur As Single) As Object

    Dim controle As Object

    Set controle = conteneur.Controls.Add("Forms.Frame.1", nom, True)
    With controle
        .Caption = legende
        .Left = gauche
        .Top = haut
        .Width = largeur
        .Height = hauteur
        .BackColor = RGB(248, 250, 252)
        .ForeColor = RGB(23, 54, 93)
        DefinirProprieteOptionnelle controle, "SpecialEffect", EFFET_PLAT
        DefinirProprieteOptionnelle controle, "Font.Name", "Segoe UI"
        DefinirProprieteOptionnelle controle, "Font.Size", 10
        DefinirProprieteOptionnelle controle, "Font.Bold", True
    End With
    Set AjouterCadre = controle
End Function

Private Function AjouterLabel(ByVal conteneur As Object, ByVal nom As String, _
        ByVal legende As String, ByVal gauche As Single, ByVal haut As Single, _
        ByVal largeur As Single, ByVal hauteur As Single, ByVal taillePolice As Single, _
        ByVal gras As Boolean, ByVal couleurTexte As Long, ByVal couleurFond As Long) As Object

    Dim controle As Object

    Set controle = conteneur.Controls.Add("Forms.Label.1", nom, True)
    With controle
        .Caption = legende
        .Left = gauche
        .Top = haut
        .Width = largeur
        .Height = hauteur
        DefinirProprieteOptionnelle controle, "BackStyle", 1
        DefinirProprieteOptionnelle controle, "BackColor", couleurFond
        DefinirProprieteOptionnelle controle, "ForeColor", couleurTexte
        DefinirProprieteOptionnelle controle, "Font.Name", "Segoe UI"
        DefinirProprieteOptionnelle controle, "Font.Size", taillePolice
        DefinirProprieteOptionnelle controle, "Font.Bold", gras
        DefinirProprieteOptionnelle controle, "TextAlign", ALIGNEMENT_GAUCHE
    End With
    Set AjouterLabel = controle
End Function

Private Sub AjouterLibelleChamp(ByVal conteneur As Object, ByVal nom As String, _
        ByVal legende As String, ByVal gauche As Single, ByVal haut As Single, _
        ByVal largeur As Single)

    Dim controle As Object

    Set controle = AjouterLabel(conteneur, nom, legende, gauche, haut, largeur, 18, _
        9, False, RGB(31, 41, 55), RGB(248, 250, 252))
    controle.BackStyle = 0
End Sub

Private Function AjouterZoneTexte(ByVal conteneur As Object, ByVal nom As String, _
        ByVal gauche As Single, ByVal haut As Single, ByVal largeur As Single, _
        ByVal hauteur As Single, ByVal verrouille As Boolean, _
        ByVal multiligne As Boolean) As Object

    Dim controle As Object

    Set controle = conteneur.Controls.Add("Forms.TextBox.1", nom, True)
    With controle
        .Left = gauche
        .Top = haut
        .Width = largeur
        .Height = hauteur
        DefinirProprieteOptionnelle controle, "Locked", verrouille
        DefinirProprieteOptionnelle controle, "MultiLine", multiligne
        DefinirProprieteOptionnelle controle, "WordWrap", multiligne
        DefinirProprieteOptionnelle controle, "EnterKeyBehavior", multiligne
        DefinirProprieteOptionnelle controle, "BackColor", IIf(verrouille, RGB(226, 232, 240), RGB(255, 255, 255))
        DefinirProprieteOptionnelle controle, "ForeColor", RGB(31, 41, 55)
        DefinirProprieteOptionnelle controle, "BorderStyle", BORDURE_SIMPLE
        DefinirProprieteOptionnelle controle, "SpecialEffect", EFFET_PLAT
        DefinirProprieteOptionnelle controle, "Font.Name", "Segoe UI"
        DefinirProprieteOptionnelle controle, "Font.Size", 9
    End With
    If multiligne Then DefinirProprieteOptionnelle controle, "ScrollBars", 2
    Set AjouterZoneTexte = controle
End Function

Private Function AjouterListeDeroulante(ByVal conteneur As Object, ByVal nom As String, _
        ByVal gauche As Single, ByVal haut As Single, ByVal largeur As Single, _
        ByVal hauteur As Single) As Object

    Dim controle As Object

    Set controle = conteneur.Controls.Add("Forms.ComboBox.1", nom, True)
    With controle
        .Left = gauche
        .Top = haut
        .Width = largeur
        .Height = hauteur
        DefinirProprieteOptionnelle controle, "Style", STYLE_LISTE_DEROULANTE
        DefinirProprieteOptionnelle controle, "BackColor", RGB(255, 255, 255)
        DefinirProprieteOptionnelle controle, "ForeColor", RGB(31, 41, 55)
        DefinirProprieteOptionnelle controle, "SpecialEffect", EFFET_PLAT
        DefinirProprieteOptionnelle controle, "Font.Name", "Segoe UI"
        DefinirProprieteOptionnelle controle, "Font.Size", 9
        DefinirProprieteOptionnelle controle, "ListRows", 12
    End With
    Set AjouterListeDeroulante = controle
End Function

Private Function AjouterListe(ByVal conteneur As Object, ByVal nom As String, _
        ByVal gauche As Single, ByVal haut As Single, ByVal largeur As Single, _
        ByVal hauteur As Single) As Object

    Dim controle As Object

    Set controle = conteneur.Controls.Add("Forms.ListBox.1", nom, True)
    With controle
        .Left = gauche
        .Top = haut
        .Width = largeur
        .Height = hauteur
        DefinirProprieteOptionnelle controle, "BackColor", RGB(255, 255, 255)
        DefinirProprieteOptionnelle controle, "ForeColor", RGB(31, 41, 55)
        DefinirProprieteOptionnelle controle, "BorderStyle", BORDURE_SIMPLE
        DefinirProprieteOptionnelle controle, "SpecialEffect", EFFET_PLAT
        DefinirProprieteOptionnelle controle, "Font.Name", "Segoe UI"
        DefinirProprieteOptionnelle controle, "Font.Size", 8
    End With
    DefinirProprieteOptionnelle controle, "IntegralHeight", False
    Set AjouterListe = controle
End Function

Private Function AjouterBouton(ByVal conteneur As Object, ByVal nom As String, _
        ByVal legende As String, ByVal gauche As Single, ByVal haut As Single, _
        ByVal largeur As Single, ByVal hauteur As Single, ByVal couleurFond As Long, _
        ByVal couleurTexte As Long) As Object

    Dim controle As Object

    Set controle = conteneur.Controls.Add("Forms.CommandButton.1", nom, True)
    With controle
        .Caption = legende
        .Left = gauche
        .Top = haut
        .Width = largeur
        .Height = hauteur
        DefinirProprieteOptionnelle controle, "BackColor", couleurFond
        DefinirProprieteOptionnelle controle, "ForeColor", couleurTexte
        DefinirProprieteOptionnelle controle, "Font.Name", "Segoe UI"
        DefinirProprieteOptionnelle controle, "Font.Size", 9
        DefinirProprieteOptionnelle controle, "Font.Bold", True
    End With
    DefinirProprieteOptionnelle controle, "TakeFocusOnClick", True
    Set AjouterBouton = controle
End Function

Private Sub DefinirProprieteOptionnelle(ByVal cible As Object, ByVal nomPropriete As String, _
        ByVal valeur As Variant)

    ' Certaines propriétés de MSForms varient selon la version d'Office.
    ' On ignore volontairement une propriété absente/non supportée.
    On Error Resume Next
    Err.Clear
    CallByName cible, nomPropriete, VbLet, valeur
    Err.Clear
    On Error GoTo 0
End Sub
