Attribute VB_Name = "BOOM"

' ============================================================
' GÉNÉRATEUR AUTOMATIQUE DE PLANNING HEBDOMADAIRE - V7
' ============================================================
' NOUVEAUTÉS V7 :
'  Feuille Utilisateurs colonnes (RÉORGANISÉES) :
'    1=NOM COMPLET      | 2=Activité        | 3=Ville
'    4=ZONES            | 5=Congé           | 6=Congé D
'    7=Congé F           | 8=Type de Congé   | 9=TRANSPORT
'    10=TT               | 11=TT D           | 12=TT F
'    13=RENFORT PRESS    | 14=RENFORT ITALY  | 15=Matricule
'    16=N de téléphone   | 17=Type de contrat (CDI/CDD)
'    18=Date d'embauche  | 19=Date de sortie (= date d'expiration
'                           du contrat pour un CDD)
'    20=Maladie          | 21=Date D'Arret   | 22=DATE DE REPRISE
'
'  - NOUVEAU : bloc "Maladie" (Oui/Non + date d'arrêt / de reprise)
'  - NOUVEAU : bloc "Contrat" (CDI/CDD + date d'expiration)
'  - NOUVEAU : "Type de Congé" (liste déroulante) associé au congé
'  - TT (Télétravail) : affiché dans les plannings et consolidation
'    avec fond violet clair. Google Leads Dimanche = TT par défaut.
'    TOUS LES AUTRES PROJETS : TT uniquement si renseigné dans Utilisateurs.
'  - RENFORT : mission temporaire sans modifier le planning de base.
'    Feuille BESOINS â†’ macro TraiterRenforts() propose les meilleurs
'    candidats disponibles selon critères d'équité.
'  - CONSOLIDATION : +2 colonnes NB HEURE et NB JOUR (cumulés semaine)
'  - Feuille PLANNING : trié par Date puis Nom
'    colonnes Nom | Date | Entrée | Sortie | No Semaine | Activité
'
'  NOTE IMPORTANTE : les colonnes NOM / PRENOM / POINT DE REPERE ont
'  été retirées de la feuille Utilisateurs lors de la réorganisation.
'  Le Nom/Prénom sont désormais déduits de NOM COMPLET (premier mot =
'  Prénom, une approximation : adapter SepareNomPrenom si votre
'  convention de saisie est différente). PointRepere n'a plus de
'  source dans Utilisateurs et reste vide (adapter si nécessaire).
' ============================================================

Option Explicit

' ============================================================
' TYPE COLLABORATEUR
' ============================================================
Type Collaborateur
    nomComplet      As String
    nom             As String
    Prenom          As String
    Matricule       As String
    projet          As String
    Ville           As String
    zone            As String
    PointRepere     As String
    Telephone       As String
    DateEmbauche    As String
    IndexRotation   As Integer
    EnConge         As Boolean
    CongeDebut      As Date
    CongeFin        As Date
    EnTT            As Boolean
    TTDebut         As Date
    TTFin           As Date
    RenforcPress    As Boolean
    RenforcItaly    As Boolean
    TypeConge       As String
    TypeContrat     As String      ' "CDI" ou "CDD"
    DateSortie      As String      ' = date d'expiration du contrat (CDD)
    EnMaladie       As Boolean
    DateArret       As Date
    DateReprise     As Date
End Type

Dim jours(1 To 7) As String
Public g_LundiCible As Date   ' Lundi de la semaine cible (défini par UFGenerer ou auto)

' ============================================================
' POINT D'ENTRÉE PRINCIPAL
' ============================================================
Sub GenererPlanning()
    jours(1) = "Lundi"
    jours(2) = "Mardi"
    jours(3) = "Mercredi"
    jours(4) = "Jeudi"
    jours(5) = "Vendredi"
    jours(6) = "Samedi"
    jours(7) = "Dimanche"

    ' Si g_LundiCible n'est pas définie (appel direct sans UserForm), prendre semaine courante
    If g_LundiCible = 0 Or g_LundiCible = CDate("01/01/1900") Then
        g_LundiCible = LundiSemaineAuto()
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrHandler

    If Not VerifierFeuillesExistantes() Then
        MsgBox "Erreur : certaines feuilles requises sont manquantes.", vbCritical
        GoTo Cleanup
    End If

    InitialiserFeuilleRotation
    InitialiserFeuilleConsolidation
    InitialiserFeuillePlanning
    EffacerAnciensPlannings

    Dim collaborateurs() As Collaborateur
    Dim nbCollab As Integer
    nbCollab = LireCollaborateurs(collaborateurs)

    If nbCollab = 0 Then
        MsgBox "Aucun collaborateur trouve dans la feuille Utilisateurs."
        GoTo Cleanup
    End If

    GenererPlanningAFEDIM collaborateurs, nbCollab
    GenererPlanningACCESSIBILITE collaborateurs, nbCollab
    GenererPlanningCMLEASING collaborateurs, nbCollab
    GenererPlanningGLF collaborateurs, nbCollab
    GenererPlanningEBRA collaborateurs, nbCollab
    GenererPlanningGOOGLELEADS collaborateurs, nbCollab
    GenererPlanningTLV collaborateurs, nbCollab
    GenererPlanningFACTO collaborateurs, nbCollab
    GenererPlanningDAC collaborateurs, nbCollab
    MettreAJourRotation collaborateurs, nbCollab
    TraiterRenforts collaborateurs, nbCollab
    AfficherRenfortsDansPlanning collaborateurs, nbCollab

    Dim semAff As Integer
    semAff = Application.WorksheetFunction.WeekNum(LundiSemaine(), 2)
    MsgBox "Planning genere avec succes !" & Chr(10) & _
           "Semaine " & semAff & " - Du " & Format(LundiSemaine(), "dd/mm/yyyy") & _
           " au " & Format(LundiSemaine() + 6, "dd/mm/yyyy"), _
           vbInformation, "Generation Planning"

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical, "Erreur"
    Resume Cleanup
End Sub

' ============================================================
' VÉRIFICATION / UTILITAIRES
' ============================================================
Function VerifierFeuillesExistantes() As Boolean
    Dim req() As String
    req = Split("Utilisateurs,AFEDIM,ACCESSIBILITE,CM Leasing,GLF,EBRA,GOOGLE LEADS,TLV,FACTO,DAC,CONSOLIDATION,PLANNING,BESOINS", ",")
    Dim i As Integer
    For i = 0 To UBound(req)
        If Not FeuilleExiste(req(i)) Then
            MsgBox "Feuille manquante : [" & req(i) & "]", vbCritical
            VerifierFeuillesExistantes = False
            Exit Function
        End If
    Next i
    VerifierFeuillesExistantes = True
End Function

Function FeuilleExiste(nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nom)
    On Error GoTo 0
    FeuilleExiste = Not (ws Is Nothing)
End Function

Function AjouterMinutes(Heure As String, minutes As Integer) As String
    If Heure = "" Then AjouterMinutes = "": Exit Function
    Dim p() As String
    p = Split(Heure, ":")
    Dim t As Integer
    t = CInt(p(0)) * 60 + CInt(p(1)) + minutes
    If t < 0 Then t = 0
    AjouterMinutes = Format(t \ 60, "00") & ":" & Format(t Mod 60, "00")
End Function

' Convertit "HH:MM" en nombre de minutes depuis minuit
Function HeureEnMinutes(Heure As String) As Integer
    Heure = Trim(Heure)
    If Heure = "" Or Heure = "OFF" Then HeureEnMinutes = 0: Exit Function

    Dim p() As String
    p = Split(Heure, ":")

    If UBound(p) < 1 Then HeureEnMinutes = 0: Exit Function

    Dim h As String, m As String
    h = Right("00" & Trim(p(0)), 2)
    m = Left(Trim(p(1)), 2)

    If IsNumeric(h) And IsNumeric(m) Then
        HeureEnMinutes = CInt(h) * 60 + CInt(m)
    Else
        HeureEnMinutes = 0
    End If
End Function

' Calcule les heures nettes travaillées (sortie-entrée-pause) en décimal
Function HeuresNettes(entree As String, sortie As String, pd As String, pf As String) As Double
    If entree = "" Or entree = "OFF" Or sortie = "" Or sortie = "OFF" Then
        HeuresNettes = 0: Exit Function
    End If
    Dim tTotal As Integer
    tTotal = HeureEnMinutes(sortie) - HeureEnMinutes(entree)
    Dim tPause As Integer
    tPause = 0
    If pd <> "" And pf <> "" Then
        tPause = HeureEnMinutes(pf) - HeureEnMinutes(pd)
    End If
    If tTotal < 0 Then tTotal = 0
    HeuresNettes = (tTotal - tPause) / 60
End Function

' Calcule le lundi de la semaine courante (utilisé si aucune date cible définie)
Function LundiSemaineAuto() As Date
    Dim d As Date
    d = Date
    Dim wd As Integer
    wd = Weekday(d, vbMonday)
    LundiSemaineAuto = d - (wd - 1)
End Function

Function LundiSemaine() As Date
    If g_LundiCible = 0 Or g_LundiCible = CDate("01/01/1900") Then
        LundiSemaine = LundiSemaineAuto()
    Else
        LundiSemaine = g_LundiCible
    End If
End Function

Function DateDuJour(j As Integer) As Date
    DateDuJour = LundiSemaine() + (j - 1)
End Function

Function EstEnConge(c As Collaborateur, d As Date) As Boolean
    If Not c.EnConge Then EstEnConge = False: Exit Function
    EstEnConge = (d >= c.CongeDebut And d <= c.CongeFin)
End Function

Function EstEnTT(c As Collaborateur, d As Date) As Boolean
    If Not c.EnTT Then EstEnTT = False: Exit Function
    EstEnTT = (d >= c.TTDebut And d <= c.TTFin)
End Function

Function NomJourToIndex(NomJour As String) As Integer
    Select Case UCase(Trim(NomJour))
        Case "LUNDI":    NomJourToIndex = 1
        Case "MARDI":    NomJourToIndex = 2
        Case "MERCREDI": NomJourToIndex = 3
        Case "JEUDI":    NomJourToIndex = 4
        Case "VENDREDI": NomJourToIndex = 5
        Case "SAMEDI":   NomJourToIndex = 6
        Case "DIMANCHE": NomJourToIndex = 7
        Case Else:       NomJourToIndex = 0
    End Select
End Function

Function FormatCelluleJour(debut As String, fin As String, pd As String, pf As String) As String
    If debut = "OFF" Or debut = "" Then
        FormatCelluleJour = "OFF": Exit Function
    End If
    Dim s As String
    s = debut & " - " & fin
    If pd <> "" Then s = s & Chr(10) & "Pause: " & pd & "-" & pf
    FormatCelluleJour = s
End Function

' ============================================================
' EN-TÃŠTE HORIZONTALE
' ============================================================
Sub EcrireEnTeteHorizontale(ws As Worksheet, projet As String)
    ws.Cells(1, 1).Value = "PLANNING HEBDOMADAIRE - " & UCase(projet)
    With ws.Cells(1, 1)
        .Font.Bold = True: .Font.Size = 14
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 11)).Merge

    Dim semNum As Integer
    semNum = Application.WorksheetFunction.WeekNum(LundiSemaine(), 2)
    ws.Cells(2, 1).Value = "Semaine " & semNum & _
                            "  |  Du " & Format(LundiSemaine(), "dd/mm/yyyy") & _
                            " au " & Format(LundiSemaine() + 6, "dd/mm/yyyy") & _
                            "  |  Generee le " & Format(Date, "dd/mm/yyyy")
    ws.Cells(2, 1).Font.Italic = True
    ws.Range(ws.Cells(2, 1), ws.Cells(2, 11)).Merge

    ws.Cells(3, 1).Value = "Collaborateur"
    ws.Cells(3, 2).Value = "Ville"
    ws.Cells(3, 3).Value = "Zone"
    Dim j As Integer
    For j = 1 To 7
        ws.Cells(3, 3 + j).Value = jours(j)
    Next j
    ws.Cells(3, 11).Value = "NB HEURES"   ' colonne cumul hebdo

    With ws.Rows(3)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 14
    Dim c As Integer
    For c = 4 To 10
        ws.Columns(c).ColumnWidth = 22
    Next c
    ws.Columns("K").ColumnWidth = 12   ' NB HEURES
    ws.Rows(3).RowHeight = 20
End Sub

' ============================================================
' ÉCRIRE UNE LIGNE HORIZONTALE
' ============================================================
Sub EcrireLigneHorizontale(ws As Worksheet, ligne As Integer, nom As String, _
                            Ville As String, zone As String, cellules() As String, _
                            Optional nbHeures As Double = -1)
    ws.Cells(ligne, 1).Value = nom
    ws.Cells(ligne, 2).Value = Ville
    ws.Cells(ligne, 3).Value = zone

    Dim j As Integer
    For j = 1 To 7
        Dim cel As Range
        Set cel = ws.Cells(ligne, 3 + j)
        cel.Value = cellules(j)
        cel.HorizontalAlignment = xlCenter
        cel.VerticalAlignment = xlCenter
        cel.WrapText = True

        If cellules(j) = "OFF" Then
            cel.Interior.Color = RGB(255, 199, 206)
            cel.Font.Bold = True
            cel.Font.Color = RGB(192, 0, 0)
        ElseIf cellules(j) = "CONGE" Then
            cel.Interior.Color = RGB(255, 192, 0)
            cel.Font.Bold = True
            cel.Font.Color = RGB(0, 0, 0)
        ElseIf Left(cellules(j), 2) = "TT" Then
            cel.Interior.Color = RGB(220, 190, 255)
            cel.Font.Bold = False
            cel.Font.Color = RGB(70, 0, 130)
        ElseIf InStr(cellules(j), "[RENFORT]") > 0 Then
            cel.Interior.Color = RGB(169, 208, 142)   ' vert renfort
            cel.Font.Bold = True
            cel.Font.Color = RGB(0, 97, 0)
        ElseIf InStr(cellules(j), "[SHIFT Réduit]") > 0 Then
            cel.Interior.Color = RGB(255, 220, 140)   ' orange clair shift réduit
            cel.Font.Bold = True
            cel.Font.Color = RGB(150, 75, 0)
        Else
            cel.Font.Color = RGB(0, 0, 0)
            If ligne Mod 2 = 0 Then
                cel.Interior.Color = RGB(235, 241, 255)
            Else
                cel.Interior.Color = RGB(255, 255, 255)
            End If
        End If
    Next j

    ' Colonne NB HEURES (col 11)
    Dim celH As Range
    Set celH = ws.Cells(ligne, 11)
    If nbHeures >= 0 Then
        celH.Value = Round(nbHeures, 2)
        celH.NumberFormat = "0.00"
        celH.HorizontalAlignment = xlCenter
        celH.Font.Bold = True
        celH.Interior.Color = RGB(197, 224, 180)   ' vert clair
        celH.Font.Color = RGB(0, 97, 0)
    End If

    If ligne Mod 2 = 0 Then
        ws.Cells(ligne, 1).Interior.Color = RGB(235, 241, 255)
        ws.Cells(ligne, 2).Interior.Color = RGB(235, 241, 255)
        ws.Cells(ligne, 3).Interior.Color = RGB(235, 241, 255)
    End If
    ws.Rows(ligne).RowHeight = 40
End Sub

Sub AppliquerBorduresH(ws As Worksheet, ligneDebut As Integer, ligneFin As Integer)
    If ligneFin < ligneDebut Then Exit Sub
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(ligneDebut, 1), ws.Cells(ligneFin, 11))
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(189, 189, 189)
    End With
    With rng.Borders(xlEdgeLeft): .Weight = xlMedium: .Color = RGB(68, 114, 196): End With
    With rng.Borders(xlEdgeRight): .Weight = xlMedium: .Color = RGB(68, 114, 196): End With
    With rng.Borders(xlEdgeTop): .Weight = xlMedium: .Color = RGB(68, 114, 196): End With
    With rng.Borders(xlEdgeBottom): .Weight = xlMedium: .Color = RGB(68, 114, 196): End With
End Sub

' ============================================================
' CONSOLIDATION
' ============================================================
Sub InitialiserFeuilleConsolidation()
    Dim ws As Worksheet
    If Not FeuilleExiste("CONSOLIDATION") Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "CONSOLIDATION"
    Else
        Set ws = ThisWorkbook.Sheets("CONSOLIDATION")
    End If
    ws.Cells.Clear

    Dim headers As Variant
    headers = Array("Nom", "Date", "Entree", "Sortie", "Pause D", "Pause F", _
                    "No Semaine", "Activite", "Conge D", "Conge F", "Ville", "Zone", _
                    "NB HEURE", "NB JOUR", "Commentaire")
    Dim c As Integer
    For c = 0 To UBound(headers)
        ws.Cells(1, c + 1).Value = headers(c)
    Next c
    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 10
    ws.Columns("E").ColumnWidth = 10
    ws.Columns("F").ColumnWidth = 10
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 16
    ws.Columns("I").ColumnWidth = 14
    ws.Columns("J").ColumnWidth = 14
    ws.Columns("K").ColumnWidth = 14
    ws.Columns("L").ColumnWidth = 14
    ws.Columns("M").ColumnWidth = 12
    ws.Columns("N").ColumnWidth = 12
    ws.Columns("O").ColumnWidth = 30   ' Commentaire
End Sub

Sub AjouterLigneConsolidation(collab As Collaborateur, d As Date, _
                               entree As String, sortie As String, _
                               pd As String, pf As String, _
                               activite As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("CONSOLIDATION")
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    Dim sem As Integer
    sem = Application.WorksheetFunction.WeekNum(d, 2)

    ws.Cells(lr, 1).Value = collab.nomComplet
    ws.Cells(lr, 2).Value = d
    ws.Cells(lr, 2).NumberFormat = "dd/mm/yyyy"
    ws.Cells(lr, 3).Value = entree
    ws.Cells(lr, 4).Value = sortie
    ws.Cells(lr, 5).Value = pd
    ws.Cells(lr, 6).Value = pf
    ws.Cells(lr, 7).Value = sem
    ws.Cells(lr, 8).Value = activite

    If collab.EnConge Then
        ws.Cells(lr, 9).Value = collab.CongeDebut
        ws.Cells(lr, 9).NumberFormat = "dd/mm/yyyy"
        ws.Cells(lr, 10).Value = collab.CongeFin
        ws.Cells(lr, 10).NumberFormat = "dd/mm/yyyy"
    End If

    ws.Cells(lr, 11).Value = collab.Ville
    ws.Cells(lr, 12).Value = collab.zone

    Select Case activite
        Case "OFF":    ws.Rows(lr).Interior.Color = RGB(255, 199, 206)
        Case "CONGE":  ws.Rows(lr).Interior.Color = RGB(255, 230, 153)
        Case "TT":     ws.Rows(lr).Interior.Color = RGB(230, 210, 255)
        Case Else
            If lr Mod 2 = 0 Then
                ws.Rows(lr).Interior.Color = RGB(235, 241, 255)
            Else
                ws.Rows(lr).Interior.Color = RGB(255, 255, 255)
            End If
    End Select
End Sub

Sub CalculerCumulsSemaine()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("CONSOLIDATION")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim noms() As String
    Dim sems() As Integer
    Dim nbH() As Double
    Dim nbJ() As Integer
    Dim nbGroupe As Integer
    nbGroupe = 0

    Dim i As Long
    For i = 2 To lastRow
        Dim nom As String: nom = ws.Cells(i, 1).Value
        Dim sem As Integer: sem = CInt(ws.Cells(i, 7).Value)
        Dim activite As String: activite = ws.Cells(i, 8).Value
        Dim entree As String: entree = ws.Cells(i, 3).Value
        Dim sortie As String: sortie = ws.Cells(i, 4).Value
        Dim pdStr As String: pdStr = ws.Cells(i, 5).Value
        Dim pfStr As String: pfStr = ws.Cells(i, 6).Value

        Dim gIdx As Integer: gIdx = -1
        Dim g As Integer
        For g = 1 To nbGroupe
            If noms(g) = nom And sems(g) = sem Then
                gIdx = g: Exit For
            End If
        Next g

        If gIdx = -1 Then
            nbGroupe = nbGroupe + 1
            ReDim Preserve noms(1 To nbGroupe)
            ReDim Preserve sems(1 To nbGroupe)
            ReDim Preserve nbH(1 To nbGroupe)
            ReDim Preserve nbJ(1 To nbGroupe)
            noms(nbGroupe) = nom
            sems(nbGroupe) = sem
            nbH(nbGroupe) = 0
            nbJ(nbGroupe) = 0
            gIdx = nbGroupe
        End If

        If activite <> "OFF" And activite <> "CONGE" Then
            nbJ(gIdx) = nbJ(gIdx) + 1
            Dim hNet As Double
            hNet = HeuresNettes(entree, sortie, pdStr, pfStr)
            nbH(gIdx) = nbH(gIdx) + hNet
        End If
    Next i

    For i = 2 To lastRow
        Dim nomL As String: nomL = ws.Cells(i, 1).Value
        Dim semL As Integer: semL = CInt(ws.Cells(i, 7).Value)
        For g = 1 To nbGroupe
            If noms(g) = nomL And sems(g) = semL Then
                ws.Cells(i, 13).Value = Round(nbH(g), 2)
                ws.Cells(i, 14).Value = nbJ(g)
                Exit For
            End If
        Next g
    Next i
End Sub

Sub InitialiserFeuillePlanning()
    Dim ws As Worksheet
    If Not FeuilleExiste("PLANNING") Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "PLANNING"
    Else
        Set ws = ThisWorkbook.Sheets("PLANNING")
    End If
    ws.Cells.Clear

    Dim headers As Variant
    headers = Array( _
        "Semaine", "Matricule", "NOM", "PRENOM", "NOM COMPLET", _
        "Date d'embauche", "Activité", "N de téléphone", "Ville", _
        "POINT DE REPERE", "ZONES", _
        "LUN. Entrée", "LUN. Sortie", _
        "MAR. Entrée", "MAR. Sortie", _
        "MER. Entrée", "MER. Sortie", _
        "JEU. Entrée", "JEU. Sortie", _
        "VEN. Entrée", "VEN. Sortie", _
        "SAM. Entrée", "SAM. Sortie", _
        "DIM. Entrée", "DIM. Sortie")

    Dim c As Integer
    For c = 0 To UBound(headers)
        ws.Cells(1, c + 1).Value = headers(c)
    Next c

    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With

    ' Largeurs colonnes
    ws.Columns(1).ColumnWidth = 10    ' Semaine
    ws.Columns(2).ColumnWidth = 14    ' Matricule
    ws.Columns(3).ColumnWidth = 18    ' NOM
    ws.Columns(4).ColumnWidth = 18    ' PRENOM
    ws.Columns(5).ColumnWidth = 28    ' NOM COMPLET
    ws.Columns(6).ColumnWidth = 16    ' Date embauche
    ws.Columns(7).ColumnWidth = 16    ' Activité
    ws.Columns(8).ColumnWidth = 16    ' Téléphone
    ws.Columns(9).ColumnWidth = 14    ' Ville
    ws.Columns(10).ColumnWidth = 18   ' Point de repère
    ws.Columns(11).ColumnWidth = 12   ' Zones
    Dim d As Integer
    For d = 12 To 25
        ws.Columns(d).ColumnWidth = 11  ' Entrée/Sortie par jour
    Next d

    ' Geler la première ligne
    ws.Activate
    ws.Rows(2).Select
    ActiveWindow.FreezePanes = True
End Sub

Sub AjouterLignePlanning(nom As String, d As Date, entree As String, sortie As String, activite As String)
    ' Appel legacy sans collab â†’ ignoré (remplacé par AjouterLignePlanningCollab)
End Sub

Sub AjouterLignePlanningCollab(c As Collaborateur, d As Date, entree As String, sortie As String, activite As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PLANNING")

    Dim sem As Integer
    sem = Application.WorksheetFunction.WeekNum(d, 2)

    ' Numéro de jour de la semaine (1=Lun â€¦ 7=Dim)
    Dim wd As Integer
    wd = Weekday(d, vbMonday)   ' 1=Lun, 7=Dim

    ' Colonnes Entrée/Sortie pour ce jour
    ' Lun=12/13, Mar=14/15, Mer=16/17, Jeu=18/19, Ven=20/21, Sam=22/23, Dim=24/25
    Dim colEntree As Integer: colEntree = 10 + (wd * 2)     ' 12,14,16,18,20,22,24
    Dim colSortie As Integer: colSortie = colEntree + 1      ' 13,15,17,19,21,23,25

    ' Chercher si la ligne collab+semaine existe déjà
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim lr As Long: lr = 0
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = sem And ws.Cells(i, 5).Value = c.nomComplet Then
            lr = i: Exit For
        End If
    Next i

    ' Créer la ligne si elle n'existe pas
    If lr = 0 Then
        lr = lastRow + 1
        ws.Cells(lr, 1).Value = sem
        ws.Cells(lr, 2).Value = c.Matricule
        ws.Cells(lr, 3).Value = c.nom
        ws.Cells(lr, 4).Value = c.Prenom
        ws.Cells(lr, 5).Value = c.nomComplet
        ws.Cells(lr, 6).Value = c.DateEmbauche
        ws.Cells(lr, 7).Value = c.projet
        ws.Cells(lr, 8).Value = c.Telephone
        ws.Cells(lr, 9).Value = c.Ville
        ws.Cells(lr, 10).Value = c.PointRepere
        ws.Cells(lr, 11).Value = c.zone

        ' Couleur de ligne alternée
        If lr Mod 2 = 0 Then
            ws.Rows(lr).Interior.Color = RGB(235, 241, 255)
        Else
            ws.Rows(lr).Interior.Color = RGB(255, 255, 255)
        End If
        ws.Rows(lr).RowHeight = 18
    End If

    ' Remplir les cellules Entrée/Sortie du jour
    Dim celE As Range: Set celE = ws.Cells(lr, colEntree)
    Dim celS As Range: Set celS = ws.Cells(lr, colSortie)

    ' Valeur à afficher
    Dim valE As String, valS As String
    Select Case True
        Case activite = "CONGE"
            valE = "CONGE": valS = "CONGE"
        Case entree = "OFF" Or activite = "OFF"
            valE = "OFF": valS = "OFF"
        Case Left(activite, 2) = "TT"
            valE = "TT " & entree: valS = "TT " & sortie
        Case activite = "RENFORT"
            valE = entree: valS = sortie
        Case Else
            valE = entree: valS = sortie
    End Select

    celE.Value = valE
    celS.Value = valS
    celE.HorizontalAlignment = xlCenter
    celS.HorizontalAlignment = xlCenter

    ' Colorier les cellules du jour selon statut
    Select Case True
        Case valE = "OFF"
            celE.Interior.Color = RGB(255, 199, 206): celE.Font.Color = RGB(192, 0, 0): celE.Font.Bold = True
            celS.Interior.Color = RGB(255, 199, 206): celS.Font.Color = RGB(192, 0, 0): celS.Font.Bold = True
        Case valE = "CONGE"
            celE.Interior.Color = RGB(255, 230, 153): celE.Font.Color = RGB(156, 87, 0): celE.Font.Bold = True
            celS.Interior.Color = RGB(255, 230, 153): celS.Font.Color = RGB(156, 87, 0): celS.Font.Bold = True
        Case Left(valE, 2) = "TT"
            celE.Interior.Color = RGB(220, 190, 255): celE.Font.Color = RGB(70, 0, 130)
            celS.Interior.Color = RGB(220, 190, 255): celS.Font.Color = RGB(70, 0, 130)
        Case Else
            celE.Interior.ColorIndex = xlNone: celE.Font.Color = RGB(0, 0, 0)
            celS.Interior.ColorIndex = xlNone: celS.Font.Color = RGB(0, 0, 0)
    End Select
End Sub

' ============================================================
' TRI FEUILLE PLANNING PAR SEMAINE PUIS NOM COMPLET
' ============================================================
Sub TrierFeuillePlanning()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PLANNING")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 3 Then Exit Sub

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 25))

    With ws.Sort
        .SortFields.Clear
        ' Tri primaire : Semaine (col 1)
        .SortFields.Add Key:=ws.Range("A2:A" & lastRow), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ' Tri secondaire : NOM COMPLET (col 5)
        .SortFields.Add Key:=ws.Range("E2:E" & lastRow), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange rng
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ' Réappliquer couleurs alternées après tri (les cellules jour gardent leur couleur)
    Dim i As Long
    For i = 2 To lastRow
        ' Couleur de fond ligne : alterner blanc/bleu clair sauf si déjà colorié par statut
        ' On recolorie uniquement les colonnes fixes (1-11)
        Dim bgColor As Long
        bgColor = IIf(i Mod 2 = 0, RGB(235, 241, 255), RGB(255, 255, 255))
        Dim col As Integer
        For col = 1 To 11
            ws.Cells(i, col).Interior.Color = bgColor
        Next col
    Next i
End Sub

' ============================================================
' ROTATION / GESTION
' ============================================================
Sub InitialiserFeuilleRotation()
    Dim ws As Worksheet
    If Not FeuilleExiste("ROTATION") Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "ROTATION"
    Else
        Set ws = ThisWorkbook.Sheets("ROTATION")
    End If
    If ws.Cells(1, 1).Value = "" Then
        ws.Cells(1, 1).Value = "Collaborateur"
        ws.Cells(1, 2).Value = "Projet"
        ws.Cells(1, 3).Value = "Index Rotation"
        ws.Cells(1, 4).Value = "Derniere MAJ"
        ws.Cells(1, 5).Value = "Semaine"
        ws.Cells(1, 6).Value = "Nb Renforts"
        With ws.Rows(1)
            .Font.Bold = True
            .Interior.Color = RGB(31, 73, 125)
            .Font.Color = RGB(255, 255, 255)
        End With
    End If
End Sub

Sub EffacerAnciensPlannings()
    Dim feuilles() As String
    feuilles = Split("AFEDIM,ACCESSIBILITE,CM Leasing,GLF,EBRA,GOOGLE LEADS,TLV,FACTO,DAC", ",")
    Dim i As Integer
    For i = 0 To UBound(feuilles)
        ThisWorkbook.Sheets(feuilles(i)).Cells.Clear
    Next i
    Dim wsC As Worksheet
    Set wsC = ThisWorkbook.Sheets("CONSOLIDATION")
    If wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row > 1 Then
        wsC.Range(wsC.Cells(2, 1), wsC.Cells(wsC.Rows.Count, 15)).Clear
    End If
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets("PLANNING")
    If wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row > 1 Then
        wsP.Range(wsP.Cells(2, 1), wsP.Cells(wsP.Rows.Count, 25)).Clear
    End If
End Sub

' ============================================================
' LECTURE COLLABORATEURS
' Structure feuille Utilisateurs (colonnes RÉORGANISÉES - V7) :
'   1=NOM COMPLET | 2=Activité | 3=Ville | 4=ZONES
'   5=Congé | 6=Congé D | 7=Congé F | 8=Type de Congé
'   9=TRANSPORT (ignoré)
'   10=TT | 11=TT D | 12=TT F
'   13=RENFORT PRESS | 14=RENFORT ITALY
'   15=Matricule | 16=N de téléphone
'   17=Type de contrat (CDI/CDD) | 18=Date d'embauche | 19=Date de sortie
'   20=Maladie | 21=Date D'Arret | 22=DATE DE REPRISE
' NOTE : Si vos colonnes sont dans un ordre différent, ajustez
'        les numéros ci-dessous en conséquence.
' ============================================================
Function LireCollaborateurs(ByRef collabs() As Collaborateur) As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Utilisateurs")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then LireCollaborateurs = 0: Exit Function
    Dim nb As Integer
    nb = lastRow - 1
    ReDim collabs(1 To nb)
    Dim i As Integer
    For i = 1 To nb
        collabs(i).nomComplet = Trim(ws.Cells(i + 1, 1).Value)
        collabs(i).projet = Trim(ws.Cells(i + 1, 2).Value)
        collabs(i).Ville = Trim(ws.Cells(i + 1, 3).Value)
        collabs(i).zone = Trim(ws.Cells(i + 1, 4).Value)
        collabs(i).IndexRotation = LireIndexRotation(collabs(i).nomComplet, collabs(i).projet)

        ' Congé (col 5-8)
        Dim cv As String: cv = UCase(Trim(ws.Cells(i + 1, 5).Value))
        collabs(i).EnConge = (cv = "OUI" Or cv = "O" Or cv = "YES")
        If collabs(i).EnConge Then
            Dim rawD As Variant: rawD = ws.Cells(i + 1, 6).Value
            Dim rawF As Variant: rawF = ws.Cells(i + 1, 7).Value
            If IsDate(rawD) Then collabs(i).CongeDebut = CDate(rawD) Else collabs(i).CongeDebut = Date
            If IsDate(rawF) Then collabs(i).CongeFin = CDate(rawF) Else collabs(i).CongeFin = Date
        End If
        collabs(i).TypeConge = Trim(ws.Cells(i + 1, 8).Value)

        ' col 9 = TRANSPORT â†’ ignoré

        ' TT (col 10-12)
        Dim tv As String: tv = UCase(Trim(ws.Cells(i + 1, 10).Value))
        collabs(i).EnTT = (tv = "OUI" Or tv = "O" Or tv = "YES")
        If collabs(i).EnTT Then
            Dim rawTD As Variant: rawTD = ws.Cells(i + 1, 11).Value
            Dim rawTF As Variant: rawTF = ws.Cells(i + 1, 12).Value
            If IsDate(rawTD) Then collabs(i).TTDebut = CDate(rawTD) Else collabs(i).TTDebut = Date
            If IsDate(rawTF) Then collabs(i).TTFin = CDate(rawTF) Else collabs(i).TTFin = Date
        End If

        ' Renfort (col 13-14)
        Dim rpv As String: rpv = UCase(Trim(ws.Cells(i + 1, 13).Value))
        collabs(i).RenforcPress = (rpv = "OUI" Or rpv = "O" Or rpv = "YES")
        Dim riv As String: riv = UCase(Trim(ws.Cells(i + 1, 14).Value))
        collabs(i).RenforcItaly = (riv = "OUI" Or riv = "O" Or riv = "YES")

        ' Matricule / téléphone (col 15-16)
        collabs(i).Matricule = Trim(ws.Cells(i + 1, 15).Value)
        collabs(i).Telephone = Trim(ws.Cells(i + 1, 16).Value)

        ' Contrat (col 17-19)
        collabs(i).TypeContrat = Trim(ws.Cells(i + 1, 17).Value)
        Dim rawE As Variant: rawE = ws.Cells(i + 1, 18).Value
        collabs(i).DateEmbauche = IIf(IsDate(rawE), Format(CDate(rawE), "dd/mm/yyyy"), CStr(rawE))
        Dim rawS As Variant: rawS = ws.Cells(i + 1, 19).Value
        collabs(i).DateSortie = IIf(IsDate(rawS), Format(CDate(rawS), "dd/mm/yyyy"), CStr(rawS))

        ' Maladie (col 20-22)
        Dim mv As String: mv = UCase(Trim(ws.Cells(i + 1, 20).Value))
        collabs(i).EnMaladie = (mv = "OUI" Or mv = "O" Or mv = "YES")
        If collabs(i).EnMaladie Then
            Dim rawArret As Variant: rawArret = ws.Cells(i + 1, 21).Value
            Dim rawRepr As Variant: rawRepr = ws.Cells(i + 1, 22).Value
            If IsDate(rawArret) Then collabs(i).DateArret = CDate(rawArret) Else collabs(i).DateArret = Date
            If IsDate(rawRepr) Then collabs(i).DateReprise = CDate(rawRepr) Else collabs(i).DateReprise = Date
        End If

        ' NOM / PRENOM / POINT DE REPERE ne sont plus des colonnes de la
        ' feuille Utilisateurs (retirées lors de la réorganisation).
        ' Approximation : premier mot de NOM COMPLET = Prénom, reste = Nom.
        ' Adapter SepareNomPrenom si votre convention de saisie diffère.
        SepareNomPrenom collabs(i).nomComplet, collabs(i).nom, collabs(i).Prenom
        collabs(i).PointRepere = ""
    Next i
    LireCollaborateurs = nb
End Function

' Sépare grossièrement NOM COMPLET en Nom / Prénom (1er mot = prénom).
' Utilisé uniquement pour alimenter la feuille PLANNING (colonnes
' NOM / PRENOM) puisque ces colonnes n'existent plus dans Utilisateurs.
Private Sub SepareNomPrenom(nomComplet As String, ByRef nom As String, ByRef prenom As String)
    Dim p As Long: p = InStr(nomComplet, " ")
    If p = 0 Then
        prenom = nomComplet
        nom = ""
    Else
        prenom = Trim(Left(nomComplet, p - 1))
        nom = Trim(Mid(nomComplet, p + 1))
    End If
End Sub

Function LireIndexRotation(nom As String, projet As String) As Integer
    If Not FeuilleExiste("ROTATION") Then LireIndexRotation = 0: Exit Function
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ROTATION")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = nom And ws.Cells(i, 2).Value = projet Then
            LireIndexRotation = CInt(ws.Cells(i, 3).Value)
            Exit Function
        End If
    Next i
    LireIndexRotation = 0
End Function

Function LireNbRenforts(nom As String, projet As String) As Integer
    If Not FeuilleExiste("ROTATION") Then LireNbRenforts = 0: Exit Function
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ROTATION")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = nom And ws.Cells(i, 2).Value = projet Then
            Dim v As Variant: v = ws.Cells(i, 6).Value
            LireNbRenforts = IIf(IsNumeric(v), CInt(v), 0)
            Exit Function
        End If
    Next i
    LireNbRenforts = 0
End Function

Sub IncrementerNbRenforts(nom As String, projet As String)
    If Not FeuilleExiste("ROTATION") Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ROTATION")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        ' FIX 2 : si projet fourni, matcher les deux ; sinon matcher uniquement le nom
        Dim match As Boolean
        If projet <> "" Then
            match = (ws.Cells(i, 1).Value = nom And ws.Cells(i, 2).Value = projet)
        Else
            match = (ws.Cells(i, 1).Value = nom)
        End If
        If match Then
            Dim v As Variant: v = ws.Cells(i, 6).Value
            ws.Cells(i, 6).Value = IIf(IsNumeric(v), CInt(v) + 1, 1)
            Exit Sub
        End If
    Next i
End Sub

Function TrouverLigneRotation(ws As Worksheet, nom As String, projet As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = nom And ws.Cells(i, 2).Value = projet Then
            TrouverLigneRotation = i: Exit Function
        End If
    Next i
    TrouverLigneRotation = 0
End Function

' ============================================================
' HELPERS COMMUNS
' ============================================================

' FIX 1 : TT Dimanche par défaut UNIQUEMENT pour GOOGLE LEADS
' Pour tous les autres projets : TT s'applique SEULEMENT si
' la colonne TT de la feuille Utilisateurs est à OUI avec dates.
' AppliquerCongesEtTT est utilisé par tous les projets SAUF Google Leads
' (qui gère son propre Dimanche TT dans GenererPlanningGOOGLELEADS).
Sub AppliquerCongesEtTT(ByRef cellules() As String, _
                         ByRef entTab() As String, _
                         ByRef sorTab() As String, _
                         ByRef pdTab() As String, _
                         ByRef pfTab() As String, _
                         c As Collaborateur)
    Dim j As Integer
    For j = 1 To 7
        Dim d As Date: d = DateDuJour(j)
        If cellules(j) <> "OFF" Then
            If EstEnConge(c, d) Then
                ' Congé prioritaire sur tout
                cellules(j) = "CONGE"
                entTab(j) = "": sorTab(j) = "": pdTab(j) = "": pfTab(j) = ""
            ElseIf EstEnTT(c, d) Then
                ' TT uniquement si explicitement renseigné dans Utilisateurs
                If Left(cellules(j), 2) <> "TT" Then
                    cellules(j) = "TT " & cellules(j)
                End If
                ' Les horaires entTab/sorTab/pdTab/pfTab restent inchangés
            End If
        End If
    Next j
End Sub

Sub EcrireLigneAvecConsolidation(ws As Worksheet, ligne As Integer, _
                                  c As Collaborateur, cellules() As String, _
                                  entrees() As String, sorties() As String, _
                                  pDs() As String, pFs() As String)
    ' Calculer NB HEURES semaine pour affichage dans planning
    Dim totalHeures As Double: totalHeures = 0
    Dim j As Integer
    For j = 1 To 7
        If cellules(j) <> "OFF" And cellules(j) <> "CONGE" Then
            Dim hNet As Double
            hNet = HeuresNettes(entrees(j), sorties(j), pDs(j), pFs(j))
            totalHeures = totalHeures + hNet
        End If
    Next j

    EcrireLigneHorizontale ws, ligne, c.nomComplet, c.Ville, c.zone, cellules, totalHeures

    For j = 1 To 7
        Dim d As Date: d = DateDuJour(j)
        Dim activite As String
        Dim entStr As String, sorStr As String, pdStr As String, pfStr As String
        Dim planEntree As String, planSortie As String

        Select Case True
            Case cellules(j) = "CONGE"
                activite = "CONGE"
                entStr = "": sorStr = "": pdStr = "": pfStr = ""
                planEntree = "OFF": planSortie = "OFF"
            Case cellules(j) = "OFF"
                activite = "OFF"
                entStr = "": sorStr = "": pdStr = "": pfStr = ""
                planEntree = "OFF": planSortie = "OFF"
            Case Left(cellules(j), 2) = "TT"
                activite = "TT"
                entStr = entrees(j): sorStr = sorties(j)
                pdStr = pDs(j): pfStr = pFs(j)
                planEntree = entrees(j): planSortie = sorties(j)
            Case InStr(cellules(j), "[RENFORT]") > 0
                activite = "RENFORT"
                entStr = entrees(j): sorStr = sorties(j)
                pdStr = pDs(j): pfStr = pFs(j)
                planEntree = entrees(j): planSortie = sorties(j)
            Case Else
                activite = c.projet
                entStr = entrees(j): sorStr = sorties(j)
                pdStr = pDs(j): pfStr = pFs(j)
                planEntree = entrees(j): planSortie = sorties(j)
        End Select

        AjouterLigneConsolidation c, d, entStr, sorStr, pdStr, pfStr, activite
        AjouterLignePlanningCollab c, d, planEntree, planSortie, activite
    Next j
End Sub

' ============================================================
' APPLIQUER LES RENFORTS SUR LES FEUILLES PROJET
' Appelé après TraiterRenforts â€” lit la feuille BESOINS et
' annote la cellule du bon jour dans la feuille projet du collab renfort
' ============================================================
Sub AfficherRenfortsDansPlanning(collabs() As Collaborateur, nb As Integer)
    If Not FeuilleExiste("BESOINS") Then Exit Sub
    Dim wsB As Worksheet
    Set wsB = ThisWorkbook.Sheets("BESOINS")
    Dim lastRow As Long
    lastRow = wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim r As Long
    For r = 2 To lastRow
        Dim Statut As String: Statut = CStr(wsB.Cells(r, 9).Value)
        If Left(Statut, 2) <> "OK" And Left(Statut, 7) <> "PARTIEL" Then GoTo NextR

        Dim proposes As String: proposes = CStr(wsB.Cells(r, 7).Value)
        If proposes = "" Or proposes = "Aucun candidat disponible" Then GoTo NextR

        Dim jourBesoin As String: jourBesoin = Trim(wsB.Cells(r, 3).Value)
        Dim hdebut As String: hdebut = Trim(wsB.Cells(r, 4).Value)
        Dim hfin As String: hfin = Trim(wsB.Cells(r, 5).Value)
        Dim projetBesoin As String: projetBesoin = Trim(wsB.Cells(r, 1).Value)

        Dim jIdx As Integer: jIdx = NomJourToIndex(jourBesoin)
        If jIdx = 0 Then GoTo NextR

        ' Pour chaque agent proposé
        Dim agents() As String
        agents = Split(proposes, " | ")
        Dim a As Integer
        For a = 0 To UBound(agents)
            Dim nomAgent As String: nomAgent = Trim(agents(a))
            If nomAgent = "" Then GoTo NextAgent

            ' Trouver le projet du collab pour savoir dans quelle feuille chercher
            Dim projetCollab As String: projetCollab = ""
            Dim ci As Integer
            For ci = 1 To nb
                If collabs(ci).nomComplet = nomAgent Then
                    projetCollab = collabs(ci).projet
                    Exit For
                End If
            Next ci
            If projetCollab = "" Then GoTo NextAgent

            ' Normaliser nom feuille
            Dim nomFeuille As String
            Select Case UCase(projetCollab)
                Case "AFEDIM":        nomFeuille = "AFEDIM"
                Case "ACCESSIBILITE": nomFeuille = "ACCESSIBILITE"
                Case "CM LEASING":    nomFeuille = "CM Leasing"
                Case "GLF":           nomFeuille = "GLF"
                Case "EBRA", "EBRA PRESSE": nomFeuille = "EBRA"
                Case "GOOGLE LEADS":  nomFeuille = "GOOGLE LEADS"
                Case "TLV", "TELEVENTE": nomFeuille = "TLV"
                Case "FACTO":         nomFeuille = "FACTO"
                Case "DAC":           nomFeuille = "DAC"
                Case Else:            GoTo NextAgent
            End Select
            If Not FeuilleExiste(nomFeuille) Then GoTo NextAgent

            Dim ws As Worksheet
            Set ws = ThisWorkbook.Sheets(nomFeuille)

            ' Chercher la ligne du collab (colonne A, à partir de la ligne 4)
            Dim ligneCollab As Long: ligneCollab = 0
            Dim lr As Long
            For lr = 4 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
                If ws.Cells(lr, 1).Value = nomAgent Then
                    ligneCollab = lr: Exit For
                End If
            Next lr
            If ligneCollab = 0 Then GoTo NextAgent

            ' Colonne du jour (col 3+jIdx)
            Dim colJour As Integer: colJour = 3 + jIdx
            Dim cel As Range
            Set cel = ws.Cells(ligneCollab, colJour)

            ' Annoter la cellule avec le renfort
            Dim valActuelle As String: valActuelle = CStr(cel.Value)
            Dim mention As String
            mention = "[RENFORT] " & projetBesoin & Chr(10) & hdebut & "-" & hfin
            If InStr(valActuelle, "[RENFORT]") = 0 Then
                cel.Value = valActuelle & Chr(10) & mention
            End If
            cel.Interior.Color = RGB(169, 208, 142)
            cel.Font.Bold = True
            cel.Font.Color = RGB(0, 97, 0)
            cel.WrapText = True

NextAgent:
        Next a
NextR:
    Next r
End Sub

' ============================================================
' PROJETS FIXES : AFEDIM / ACCESSIBILITE / CM LEASING
' ============================================================
Sub GenererPlanningFixe(nomFeuille As String, collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nomFeuille)
    EcrireEnTeteHorizontale ws, nomFeuille

    Dim ligne As Integer: ligne = 4
    Dim i As Integer
    For i = 1 To nb
        If UCase(Trim(collabs(i).projet)) = UCase(nomFeuille) Then
            Dim cellules(1 To 7) As String
            Dim entrees(1 To 7) As String
            Dim sorties(1 To 7) As String
            Dim pDs(1 To 7) As String
            Dim pFs(1 To 7) As String
            Dim j As Integer

            For j = 1 To 7
                Select Case j
                    Case 1, 2, 3, 4
                        cellules(j) = FormatCelluleJour("08:00", "18:00", "13:00", "14:00")
                        entrees(j) = "08:00": sorties(j) = "18:00"
                        pDs(j) = "13:00": pFs(j) = "14:00"
                    Case 5
                        cellules(j) = FormatCelluleJour("08:00", "17:00", "13:00", "14:00")
                        entrees(j) = "08:00": sorties(j) = "17:00"
                        pDs(j) = "13:00": pFs(j) = "14:00"
                    Case Else
                        cellules(j) = "OFF"
                        entrees(j) = "": sorties(j) = "": pDs(j) = "": pFs(j) = ""
                End Select
            Next j

            AppliquerCongesEtTT cellules, entrees, sorties, pDs, pFs, collabs(i)
            EcrireLigneAvecConsolidation ws, ligne, collabs(i), cellules, entrees, sorties, pDs, pFs
            ligne = ligne + 1
        End If
    Next i

    If ligne > 4 Then
        ws.Cells(ligne + 1, 1).Value = "Total : 44h | Pause fixe 13:00-14:00 | TT = fond violet"
        ws.Cells(ligne + 1, 1).Font.Italic = True
        ws.Cells(ligne + 1, 1).Font.Color = RGB(31, 73, 125)
    End If
    AppliquerBorduresH ws, 4, ligne - 1
End Sub

Sub GenererPlanningAFEDIM(collabs() As Collaborateur, nb As Integer)
    GenererPlanningFixe "AFEDIM", collabs, nb
End Sub
Sub GenererPlanningACCESSIBILITE(collabs() As Collaborateur, nb As Integer)
    GenererPlanningFixe "ACCESSIBILITE", collabs, nb
End Sub
Sub GenererPlanningCMLEASING(collabs() As Collaborateur, nb As Integer)
    GenererPlanningFixe "CM Leasing", collabs, nb
End Sub

' ============================================================
' GLF
' ============================================================
Sub GenererPlanningGLF(collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("GLF")
    EcrireEnTeteHorizontale ws, "GLF"

    Dim vaguesPause(1 To 5) As String
    vaguesPause(1) = "12:00": vaguesPause(2) = "12:30": vaguesPause(3) = "13:00"
    vaguesPause(4) = "13:30": vaguesPause(5) = "14:00"

    Dim glfIdx() As Integer
    Dim nbGLF As Integer: nbGLF = 0
    Dim i As Integer
    For i = 1 To nb
        If UCase(Trim(collabs(i).projet)) = "GLF" Then
            nbGLF = nbGLF + 1
            ReDim Preserve glfIdx(1 To nbGLF)
            glfIdx(nbGLF) = i
        End If
    Next i
    If nbGLF = 0 Then Exit Sub

    Dim ligne As Integer: ligne = 4
    Dim k As Integer
    For k = 1 To nbGLF
        Dim idx As Integer: idx = glfIdx(k)
        Dim groupeBase As Integer: groupeBase = (k - 1) Mod 5
        Dim vagueIdx As Integer
        vagueIdx = ((groupeBase + collabs(idx).IndexRotation) Mod 5) + 1
        Dim pauseH As String: pauseH = vaguesPause(vagueIdx)
        Dim pauseF As String: pauseF = AjouterMinutes(pauseH, 60)

        Dim cellules(1 To 7) As String
        Dim entrees(1 To 7) As String
        Dim sorties(1 To 7) As String
        Dim pDs(1 To 7) As String
        Dim pFs(1 To 7) As String
        Dim j As Integer

        For j = 1 To 7
            Select Case j
                Case 1, 2, 3, 4
                    cellules(j) = FormatCelluleJour("08:00", "18:00", pauseH, pauseF)
                    entrees(j) = "08:00": sorties(j) = "18:00"
                    pDs(j) = pauseH: pFs(j) = pauseF
                Case 5
                    cellules(j) = FormatCelluleJour("08:00", "17:00", pauseH, pauseF)
                    entrees(j) = "08:00": sorties(j) = "17:00"
                    pDs(j) = pauseH: pFs(j) = pauseF
                Case Else
                    cellules(j) = "OFF"
                    entrees(j) = "": sorties(j) = "": pDs(j) = "": pFs(j) = ""
            End Select
        Next j

        AppliquerCongesEtTT cellules, entrees, sorties, pDs, pFs, collabs(idx)
        EcrireLigneAvecConsolidation ws, ligne, collabs(idx), cellules, entrees, sorties, pDs, pFs
        ligne = ligne + 1
    Next k

    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "LÉGENDE VAGUES GLF (groupes ~5, rotation hebdo)"
    ws.Cells(ligne, 1).Font.Bold = True: ws.Cells(ligne, 1).Font.Color = RGB(31, 73, 125)
    Dim v As Integer
    For v = 1 To 5
        ws.Cells(ligne + v, 1).Value = "Vague " & v & " : " & vaguesPause(v) & "-" & AjouterMinutes(vaguesPause(v), 60)
    Next v
    AppliquerBorduresH ws, 4, ligne - 2
End Sub

' ============================================================
' EBRA
' ============================================================
Sub GenererPlanningEBRA(collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("EBRA")
    EcrireEnTeteHorizontale ws, "EBRA"

    Dim vaguesPause(1 To 5) As String
    vaguesPause(1) = "11:00": vaguesPause(2) = "11:30": vaguesPause(3) = "12:00"
    vaguesPause(4) = "12:30": vaguesPause(5) = "13:00"

    Dim ebraIdx() As Integer
    Dim nbEBRA As Integer: nbEBRA = 0
    Dim i As Integer
    For i = 1 To nb
        If UCase(Trim(collabs(i).projet)) = "EBRA" Or UCase(Trim(collabs(i).projet)) = "EBRA PRESSE" Then
            nbEBRA = nbEBRA + 1
            ReDim Preserve ebraIdx(1 To nbEBRA)
            ebraIdx(nbEBRA) = i
        End If
    Next i
    If nbEBRA = 0 Then Exit Sub

    Dim ligne As Integer: ligne = 4
    Dim k As Integer
    For k = 1 To nbEBRA
        Dim idx As Integer: idx = ebraIdx(k)
        Dim groupeBase As Integer: groupeBase = ((k - 1) \ 10) Mod 5
        Dim vagueIdx As Integer
        vagueIdx = ((groupeBase + collabs(idx).IndexRotation) Mod 5) + 1
        Dim pauseH As String: pauseH = vaguesPause(vagueIdx)
        Dim pauseF As String: pauseF = AjouterMinutes(pauseH, 60)

        Dim cellules(1 To 7) As String
        Dim entrees(1 To 7) As String
        Dim sorties(1 To 7) As String
        Dim pDs(1 To 7) As String
        Dim pFs(1 To 7) As String
        Dim j As Integer

        For j = 1 To 7
            Select Case j
                Case 1 To 5
                    cellules(j) = FormatCelluleJour("07:00", "16:00", pauseH, pauseF)
                    entrees(j) = "07:00": sorties(j) = "16:00"
                    pDs(j) = pauseH: pFs(j) = pauseF
                Case 6
                    cellules(j) = FormatCelluleJour("07:00", "11:00", "", "")
                    entrees(j) = "07:00": sorties(j) = "11:00"
                    pDs(j) = "": pFs(j) = ""
                Case 7
                    cellules(j) = "OFF"
                    entrees(j) = "": sorties(j) = "": pDs(j) = "": pFs(j) = ""
            End Select
        Next j

        AppliquerCongesEtTT cellules, entrees, sorties, pDs, pFs, collabs(idx)
        EcrireLigneAvecConsolidation ws, ligne, collabs(idx), cellules, entrees, sorties, pDs, pFs
        ligne = ligne + 1
    Next k

    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "LÉGENDE VAGUES EBRA (groupes ~10, rotation hebdo)"
    ws.Cells(ligne, 1).Font.Bold = True: ws.Cells(ligne, 1).Font.Color = RGB(31, 73, 125)
    Dim v As Integer
    For v = 1 To 5
        ws.Cells(ligne + v, 1).Value = "Vague " & v & " : " & vaguesPause(v) & "-" & AjouterMinutes(vaguesPause(v), 60)
    Next v
    ws.Cells(ligne + 6, 1).Value = "Sam 07:00-11:00 sans pause | Dim OFF"
    ws.Cells(ligne + 6, 1).Font.Italic = True
    AppliquerBorduresH ws, 4, ligne - 2
End Sub

' ============================================================
' GOOGLE LEADS
' FIX 1 : Dimanche = TT par défaut UNIQUEMENT pour ce projet.
' SHIFT Réduit : chaque agent a 1 jour/semaine avec sortie -2h
'   (ex: shift 11-20 â†’ 11-18 ce jour-là)
'   Priorité weekend (Dim > Sam) puis semaine en rotation équitable.
'   Le shift réduit s'applique sur un jour TRAVAILLÉ (pas OFF/CONGE).
' ============================================================
Sub GenererPlanningGOOGLELEADS(collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("GOOGLE LEADS")
    EcrireEnTeteHorizontale ws, "GOOGLE LEADS"

    ' 5 shifts : entrée / sortie normale / sortie réduite (-2h)
    ' RÃˆGLE : shift 11H EST aussi considéré "shift réduit" (11-18 au lieu 11-20)
    ' â†’ chaque agent a au maximum 1 shift réduit par semaine (toutes causes confondues)
    Dim entrees(1 To 5) As String
    Dim sorties(1 To 5) As String
    Dim sortiesReduit(1 To 5) As String
    Dim estShiftReduitParDefaut(1 To 5) As Boolean  ' True si le shift est "réduit" par nature
    entrees(1) = "07:00": sorties(1) = "17:00": sortiesReduit(1) = "16:00": estShiftReduitParDefaut(1) = False
    entrees(2) = "08:00": sorties(2) = "18:00": sortiesReduit(2) = "17:00": estShiftReduitParDefaut(2) = False
    entrees(3) = "09:00": sorties(3) = "19:00": sortiesReduit(3) = "18:00": estShiftReduitParDefaut(3) = False
    entrees(4) = "10:00": sorties(4) = "20:00": sortiesReduit(4) = "19:00": estShiftReduitParDefaut(4) = False
    entrees(5) = "07:00": sorties(5) = "17:00": sortiesReduit(5) = "16:00": estShiftReduitParDefaut(1) = False
    ' Le shift 11H (index 5) est intrinsèquement un "shift réduit" :
    ' un agent affecté au shift 5 a DÉJÃ€ son quota de 1 shift réduit
    ' â†’ il ne peut pas recevoir un 2ème jour en horaire raccourci

    Dim glIdx() As Integer
    Dim nbGL As Integer: nbGL = 0
    Dim i As Integer
    For i = 1 To nb
        If UCase(Trim(collabs(i).projet)) = "GOOGLE LEADS" Then
            nbGL = nbGL + 1
            ReDim Preserve glIdx(1 To nbGL)
            glIdx(nbGL) = i
        End If
    Next i
    If nbGL = 0 Then Exit Sub

    ' Pré-comptage congés par jour
    Dim congeParJour(1 To 7) As Integer
    Dim j As Integer
    For j = 1 To 7: congeParJour(j) = 0: Next j
    Dim k As Integer
    For k = 1 To nbGL
        Dim idx As Integer: idx = glIdx(k)
        For j = 1 To 7
            If EstEnConge(collabs(idx), DateDuJour(j)) Then
                congeParJour(j) = congeParJour(j) + 1
            End If
        Next j
    Next k

    ' Quotas OFF par jour
    Dim quotaOFF(1 To 7) As Integer
    For j = 1 To 7
        Select Case j
            Case 7:    quotaOFF(j) = Application.WorksheetFunction.Max(0, 7 - congeParJour(j))
            Case 6:    quotaOFF(j) = Application.WorksheetFunction.Max(0, 6 - congeParJour(j))
            Case Else: quotaOFF(j) = Application.WorksheetFunction.Max(0, 5 - congeParJour(j))
        End Select
    Next j

    Dim offPlanifieParJour(1 To 7) As Integer
    For j = 1 To 7: offPlanifieParJour(j) = 0: Next j

    ' Compteur de shifts réduits planifiés par jour (équité)
    ' Priorité : Dim(7) > Sam(6) > Ven(5) > Jeu(4) > Mer(3) > Mar(2) > Lun(1)
    Dim reduitParJour(1 To 7) As Integer
    For j = 1 To 7: reduitParJour(j) = 0: Next j

    Dim ligne As Integer: ligne = 4

    For k = 1 To nbGL
        idx = glIdx(k)
        Dim shiftIdx As Integer
        shiftIdx = ((k - 1 + collabs(idx).IndexRotation) Mod 5) + 1
        Dim pd As String: pd = AjouterMinutes(entrees(shiftIdx), 300)
        Dim pf As String: pf = AjouterMinutes(pd, 60)

        ' Pause réduite : même heure début pause, fin = sortie réduite si avant fin pause normale
        Dim pDReduit As String: pDReduit = pd
        Dim pFReduit As String
        ' Si la pause normale déborde après la sortie réduite â†’ pas de pause
        If HeureEnMinutes(pd) >= HeureEnMinutes(sortiesReduit(shiftIdx)) Then
            pDReduit = "": pFReduit = ""
        Else
            pFReduit = pf
        End If

        Dim joursConge(1 To 7) As Boolean
        For j = 1 To 7
            joursConge(j) = EstEnConge(collabs(idx), DateDuJour(j))
        Next j

        Dim off1 As Integer, off2 As Integer
        Call CalculerJoursOFF_GL_V5(offPlanifieParJour, quotaOFF, joursConge, off1, off2)
        If off1 > 0 Then offPlanifieParJour(off1) = offPlanifieParJour(off1) + 1
        If off2 > 0 Then offPlanifieParJour(off2) = offPlanifieParJour(off2) + 1

        ' Choisir le jour du shift réduit :
        ' Priorité Dim > Sam > Ven > Jeu > Mer > Mar > Lun
        ' RÃˆGLE : si l'agent est sur le shift 5 (11H), il a déjà son shift réduit
        '         â†’ jourReduit reste 0 (aucun jour supplémentaire en horaire raccourci)
        Dim prioriteReduit(1 To 7) As Integer
        prioriteReduit(1) = 7: prioriteReduit(2) = 6: prioriteReduit(3) = 5
        prioriteReduit(4) = 4: prioriteReduit(5) = 3: prioriteReduit(6) = 2
        prioriteReduit(7) = 1
        Dim jourReduit As Integer: jourReduit = 0

        ' Si shift 11H â†’ quota déjà consommé, pas de shift réduit supplémentaire
        If Not estShiftReduitParDefaut(shiftIdx) Then
            Dim pr As Integer
            For pr = 1 To 7
                Dim dpr As Integer: dpr = prioriteReduit(pr)
                If Not joursConge(dpr) And dpr <> off1 And dpr <> off2 Then
                    Dim limiteReduit As Integer
                    limiteReduit = Application.WorksheetFunction.Max(1, Int(nbGL / 7) + 1)
                    If reduitParJour(dpr) < limiteReduit Then
                        jourReduit = dpr
                        Exit For
                    End If
                End If
            Next pr
            ' Fallback
            If jourReduit = 0 Then
                For pr = 1 To 7
                    dpr = prioriteReduit(pr)
                    If Not joursConge(dpr) And dpr <> off1 And dpr <> off2 Then
                        jourReduit = dpr: Exit For
                    End If
                Next pr
            End If
        End If
        If jourReduit > 0 Then reduitParJour(jourReduit) = reduitParJour(jourReduit) + 1

        ' Construire les cellules
        Dim cellules(1 To 7) As String
        Dim entTab(1 To 7) As String
        Dim sorTab(1 To 7) As String
        Dim pdTab(1 To 7) As String
        Dim pfTab(1 To 7) As String

        For j = 1 To 7
            If joursConge(j) Then
                cellules(j) = "CONGE"
                entTab(j) = "": sorTab(j) = "": pdTab(j) = "": pfTab(j) = ""
            ElseIf j = off1 Or j = off2 Then
                cellules(j) = "OFF"
                entTab(j) = "": sorTab(j) = "": pdTab(j) = "": pfTab(j) = ""
            ElseIf j = jourReduit Or estShiftReduitParDefaut(shiftIdx) Then
                ' Shift réduit :
                '   - soit jour explicitement choisi (jourReduit)
                '   - soit shift 11H qui est réduit tous les jours travaillés
                Dim sorReduit As String: sorReduit = sortiesReduit(shiftIdx)
                Dim contenuR As String
                If pDReduit <> "" Then
                    contenuR = entrees(shiftIdx) & " - " & sorReduit & Chr(10) & _
                               "Pause: " & pDReduit & "-" & pFReduit & Chr(10) & "[SHIFT Réduit]"
                Else
                    contenuR = entrees(shiftIdx) & " - " & sorReduit & Chr(10) & "[SHIFT Réduit]"
                End If
                If j = 7 Then
                    cellules(j) = "TT " & contenuR
                Else
                    cellules(j) = contenuR
                End If
                entTab(j) = entrees(shiftIdx): sorTab(j) = sorReduit
                pdTab(j) = pDReduit: pfTab(j) = pFReduit
            Else
                ' Shift normal
                Dim contenu As String
                contenu = FormatCelluleJour(entrees(shiftIdx), sorties(shiftIdx), pd, pf)
                If j = 7 Then
                    cellules(j) = "TT " & contenu
                Else
                    cellules(j) = contenu
                End If
                entTab(j) = entrees(shiftIdx): sorTab(j) = sorties(shiftIdx)
                pdTab(j) = pd: pfTab(j) = pf
            End If
        Next j

        ' Appliquer TT perso (peut écraser d'autres jours si plage TT définie)
        For j = 1 To 7
            If cellules(j) <> "CONGE" And cellules(j) <> "OFF" Then
                Dim dj As Date: dj = DateDuJour(j)
                If EstEnTT(collabs(idx), dj) And j <> 7 Then
                    cellules(j) = "TT " & FormatCelluleJour(entTab(j), sorTab(j), pdTab(j), pfTab(j))
                End If
            End If
        Next j

        EcrireLigneAvecConsolidation ws, ligne, collabs(idx), cellules, entTab, sorTab, pdTab, pfTab
        ligne = ligne + 1
    Next k

    ' Légende
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "SHIFTS GOOGLE LEADS | Dimanche = TT par défaut | 1 shift réduit/semaine (-2h, priorité weekend)"
    ws.Cells(ligne, 1).Font.Bold = True: ws.Cells(ligne, 1).Font.Color = RGB(31, 73, 125)
    Dim s As Integer
    For s = 1 To 5
        Dim pauseGL As String: pauseGL = AjouterMinutes(entrees(s), 300)
        ws.Cells(ligne + s, 1).Value = "Shift " & s & " : " & entrees(s) & "-" & sorties(s) & _
                                       " (réduit: " & sortiesReduit(s) & ")" & _
                                       "  Pause: " & pauseGL & "-" & AjouterMinutes(pauseGL, 60)
    Next s
    ws.Cells(ligne + 6, 1).Value = "~5 OFF/jour | Shift réduit = fond orange | TT = fond violet"
    ws.Cells(ligne + 6, 1).Font.Italic = True
    AppliquerBorduresH ws, 4, ligne - 2
End Sub

Sub CalculerJoursOFF_GL_V5(offPlanifie() As Integer, quota() As Integer, _
                             joursConge() As Boolean, _
                             ByRef off1 As Integer, ByRef off2 As Integer)
    Dim priorite(1 To 7) As Integer
    priorite(1) = 7: priorite(2) = 6: priorite(3) = 5: priorite(4) = 4
    priorite(5) = 3: priorite(6) = 2: priorite(7) = 1

    Dim p As Integer
    off1 = 0: off2 = 0

    For p = 1 To 7
        Dim d As Integer: d = priorite(p)
        If Not joursConge(d) And offPlanifie(d) < quota(d) Then
            off1 = d: Exit For
        End If
    Next p

    For p = 1 To 7
        d = priorite(p)
        If d <> off1 And Not joursConge(d) And offPlanifie(d) < quota(d) Then
            off2 = d: Exit For
        End If
    Next p

    If off1 = 0 Then off1 = 7
    If off2 = 0 Then
        For p = 1 To 7
            If priorite(p) <> off1 And Not joursConge(priorite(p)) Then
                off2 = priorite(p): Exit For
            End If
        Next p
    End If
    If off2 = 0 Then off2 = 6
End Sub

' ============================================================
' TLV
' ============================================================
' RÈGLES TLV (4 agents) :
'
' SYLLA SOKHNA SAFIETOU & DIOP MAMADOU MOUSTAPHA DOKY
'   Planning fixe : 08:00-17:00 du Lundi au Vendredi, OFF Samedi et Dimanche
'
' ABDELAOUI KHADIJA & AZIANE YASSINE
'   Chaque agent a 2 jours OFF : Dimanche (fixe) + Lundi ou Samedi (rotation)
'   L un a OFF Lundi  => travaille Sam 08:00-14:00
'   L autre a OFF Sam => travaille Lundi 08:00-17:00
'   Ils alternent chaque semaine via IndexRotation
'   Mardi-Mercredi : 08:00-18:00 | Jeudi-Vendredi : 08:00-17:00
' ============================================================
Sub GenererPlanningTLV(collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TLV")
    EcrireEnTeteHorizontale ws, "TLV"

    Const NOM_SYLLA      As String = "SYLLA SOKHNA SAFIETOU"
    Const NOM_DIOP       As String = "DIOP MAMADOU MOUSTAPHA DOKY"
    Const NOM_ABDELAOUI  As String = "ABDELAOUI KHADIJA"
    Const NOM_AZIANE     As String = "AZIANE YASSINE"

    Dim tlvIdx() As Integer
    Dim nbTLV As Integer: nbTLV = 0
    Dim i As Integer
    For i = 1 To nb
        If UCase(Trim(collabs(i).projet)) = "TLV" Or UCase(Trim(collabs(i).projet)) = "TELEVENTE" Then
            nbTLV = nbTLV + 1
            ReDim Preserve tlvIdx(1 To nbTLV)
            tlvIdx(nbTLV) = i
        End If
    Next i
    If nbTLV = 0 Then Exit Sub

    Dim ligne As Integer: ligne = 4
    Dim k As Integer

    For k = 1 To nbTLV
        Dim idx As Integer: idx = tlvIdx(k)
        Dim nomAgent As String: nomAgent = UCase(Trim(collabs(idx).nomComplet))

        Dim cellules(1 To 7) As String
        Dim entTab(1 To 7) As String
        Dim sorTab(1 To 7) As String
        Dim pdTab(1 To 7) As String
        Dim pfTab(1 To 7) As String
        Dim j As Integer

        ' -------------------------------------------------------
        ' Groupe 1 : SYLLA et DIOP — planning fixe 08:00-17:00 Lun-Sam
        ' -------------------------------------------------------
        If nomAgent = UCase(NOM_SYLLA) Or nomAgent = UCase(NOM_DIOP) Then
            For j = 1 To 7
                If j = 6 Or j = 7 Then  ' Samedi et Dimanche OFF
                    cellules(j) = "OFF"
                    entTab(j) = "": sorTab(j) = "": pdTab(j) = "": pfTab(j) = ""
                Else  ' Lundi a Vendredi 08:00-17:00
                    cellules(j) = FormatCelluleJour("08:00", "17:00", "13:00", "14:00")
                    entTab(j) = "08:00": sorTab(j) = "17:00"
                    pdTab(j) = "13:00": pfTab(j) = "14:00"
                End If
            Next j

        ' -------------------------------------------------------
        ' Groupe 2 : ABDELAOUI et AZIANE — rotation hebdomadaire
        ' Chaque agent a 2 jours OFF : Dimanche fixe + 1 jour rotatif
        ' Un agent a OFF Lundi, l'autre a OFF Samedi — ils alternent chaque semaine
        '
        ' ABDELAOUI (IndexRotation pair)   : OFF Lundi | Sam 08:00-14:00
        ' ABDELAOUI (IndexRotation impair) : OFF Samedi | Lundi 08:00-17:00
        ' AZIANE    (IndexRotation pair)   : OFF Samedi | Lundi 08:00-17:00
        ' AZIANE    (IndexRotation impair) : OFF Lundi  | Sam 08:00-14:00
        '
        ' Mardi-Mercredi : 08:00-18:00 | Jeudi-Vendredi : 08:00-17:00 | Dim : OFF fixe
        ' -------------------------------------------------------
        ElseIf nomAgent = UCase(NOM_ABDELAOUI) Or nomAgent = UCase(NOM_AZIANE) Then
            ' Calcul du jour OFF rotatif :
            ' ABDELAOUI pair => OFF Lundi ; AZIANE pair => OFF Samedi (et vice versa impair)
            Dim offLundi As Boolean
            If nomAgent = UCase(NOM_ABDELAOUI) Then
                offLundi = ((collabs(idx).IndexRotation Mod 2) = 0)
            Else  ' AZIANE
                offLundi = ((collabs(idx).IndexRotation Mod 2) = 1)
            End If
            ' offLundi=True  => OFF Lundi + Sam 08:00-14:00
            ' offLundi=False => OFF Samedi + Lundi 08:00-17:00

            For j = 1 To 7
                Select Case j
                    Case 1  ' Lundi
                        If offLundi Then
                            cellules(j) = "OFF"
                            entTab(j) = "": sorTab(j) = "": pdTab(j) = "": pfTab(j) = ""
                        Else
                            cellules(j) = FormatCelluleJour("08:00", "17:00", "13:00", "14:00")
                            entTab(j) = "08:00": sorTab(j) = "17:00"
                            pdTab(j) = "13:00": pfTab(j) = "14:00"
                        End If
                    Case 2  ' Mardi 08:00-18:00
                        cellules(j) = FormatCelluleJour("08:00", "18:00", "13:00", "14:00")
                        entTab(j) = "08:00": sorTab(j) = "18:00"
                        pdTab(j) = "13:00": pfTab(j) = "14:00"
                    Case 3  ' Mercredi 08:00-18:00
                        cellules(j) = FormatCelluleJour("08:00", "18:00", "13:00", "14:00")
                        entTab(j) = "08:00": sorTab(j) = "18:00"
                        pdTab(j) = "13:00": pfTab(j) = "14:00"
                    Case 4  ' Jeudi 08:00-17:00
                        cellules(j) = FormatCelluleJour("08:00", "17:00", "13:00", "14:00")
                        entTab(j) = "08:00": sorTab(j) = "17:00"
                        pdTab(j) = "13:00": pfTab(j) = "14:00"
                    Case 5  ' Vendredi 08:00-17:00
                        cellules(j) = FormatCelluleJour("08:00", "17:00", "13:00", "14:00")
                        entTab(j) = "08:00": sorTab(j) = "17:00"
                        pdTab(j) = "13:00": pfTab(j) = "14:00"
                    Case 6  ' Samedi
                        If offLundi Then
                            ' Travaille Samedi 08:00-14:00
                            cellules(j) = FormatCelluleJour("08:00", "14:00", "", "")
                            entTab(j) = "08:00": sorTab(j) = "14:00"
                            pdTab(j) = "": pfTab(j) = ""
                        Else
                            ' OFF Samedi
                            cellules(j) = "OFF"
                            entTab(j) = "": sorTab(j) = "": pdTab(j) = "": pfTab(j) = ""
                        End If
                    Case 7  ' Dimanche OFF fixe
                        cellules(j) = "OFF"
                        entTab(j) = "": sorTab(j) = "": pdTab(j) = "": pfTab(j) = ""
                End Select
            Next j

        ' -------------------------------------------------------
        ' Fallback : agent TLV non identifié — planning générique
        ' -------------------------------------------------------
        Else
            For j = 1 To 7
                If j = 7 Then
                    cellules(j) = "OFF"
                    entTab(j) = "": sorTab(j) = "": pdTab(j) = "": pfTab(j) = ""
                Else
                    cellules(j) = FormatCelluleJour("08:00", "17:00", "13:00", "14:00")
                    entTab(j) = "08:00": sorTab(j) = "17:00"
                    pdTab(j) = "13:00": pfTab(j) = "14:00"
                End If
            Next j
        End If

        AppliquerCongesEtTT cellules, entTab, sorTab, pdTab, pfTab, collabs(idx)
        EcrireLigneAvecConsolidation ws, ligne, collabs(idx), cellules, entTab, sorTab, pdTab, pfTab
        ligne = ligne + 1
    Next k

    If ligne > 4 Then
        ligne = ligne + 1
        ws.Cells(ligne, 1).Value = "SHIFTS TLV | SYLLA & DIOP : 08:00-17:00 Lun-Ven (Sam+Dim OFF) | ABDELAOUI & AZIANE : rotation Lun/Sam + Mar-Mer 08-18 + Jeu-Ven 08-17"
        ws.Cells(ligne, 1).Font.Bold = True: ws.Cells(ligne, 1).Font.Color = RGB(31, 73, 125)
        ws.Cells(ligne + 1, 1).Value = "Rotation ABDELAOUI/AZIANE : IndexRotation pair = Sam(08-14)+Lun travail | impair = Sam+Lun OFF"
        ws.Cells(ligne + 1, 1).Font.Italic = True
    End If
    AppliquerBorduresH ws, 4, ligne - 2
End Sub

' ============================================================
' FACTO / DAC
' ============================================================
Sub GenererPlanningFactoDAC(nomFeuille As String, collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nomFeuille)
    EcrireEnTeteHorizontale ws, nomFeuille

    Dim entrees(1 To 2) As String: Dim sorties(1 To 2) As String
    entrees(1) = "07:00": sorties(1) = "17:00"
    entrees(2) = "08:00": sorties(2) = "18:00"

    Dim fdIdx() As Integer
    Dim nbFD As Integer: nbFD = 0
    Dim i As Integer
    For i = 1 To nb
        If UCase(Trim(collabs(i).projet)) = UCase(nomFeuille) Then
            nbFD = nbFD + 1
            ReDim Preserve fdIdx(1 To nbFD)
            fdIdx(nbFD) = i
        End If
    Next i
    If nbFD = 0 Then Exit Sub

    Dim ligne As Integer: ligne = 4
    Dim k As Integer
    For k = 1 To nbFD
        Dim idx As Integer: idx = fdIdx(k)
        Dim shiftIdx As Integer
        shiftIdx = ((k - 1 + collabs(idx).IndexRotation) Mod 2) + 1

        Dim finNormale As String: finNormale = sorties(shiftIdx)
        Dim finVen As String: finVen = AjouterMinutes(finNormale, -60)
        Dim pd As String: pd = AjouterMinutes(entrees(shiftIdx), 300)
        Dim pf As String: pf = AjouterMinutes(pd, 60)

        Dim cellules(1 To 7) As String
        Dim entTab(1 To 7) As String
        Dim sorTab(1 To 7) As String
        Dim pdTab(1 To 7) As String
        Dim pfTab(1 To 7) As String
        Dim j As Integer

        For j = 1 To 7
            Select Case j
                Case 1, 2, 3, 4
                    cellules(j) = FormatCelluleJour(entrees(shiftIdx), finNormale, pd, pf)
                    entTab(j) = entrees(shiftIdx): sorTab(j) = finNormale
                    pdTab(j) = pd: pfTab(j) = pf
                Case 5
                    cellules(j) = FormatCelluleJour(entrees(shiftIdx), finVen, pd, pf)
                    entTab(j) = entrees(shiftIdx): sorTab(j) = finVen
                    pdTab(j) = pd: pfTab(j) = pf
                Case Else
                    cellules(j) = "OFF"
                    entTab(j) = "": sorTab(j) = "": pdTab(j) = "": pfTab(j) = ""
            End Select
        Next j

        AppliquerCongesEtTT cellules, entTab, sorTab, pdTab, pfTab, collabs(idx)
        EcrireLigneAvecConsolidation ws, ligne, collabs(idx), cellules, entTab, sorTab, pdTab, pfTab
        ligne = ligne + 1
    Next k

    If ligne > 4 Then
        ligne = ligne + 1
        ws.Cells(ligne, 1).Value = "SHIFTS " & UCase(nomFeuille) & " | Shift 1: 07-17 | Shift 2: 08-18 | Ven -1h"
        ws.Cells(ligne, 1).Font.Bold = True: ws.Cells(ligne, 1).Font.Color = RGB(31, 73, 125)
    End If
    AppliquerBorduresH ws, 4, ligne - 2
End Sub

Sub GenererPlanningFACTO(collabs() As Collaborateur, nb As Integer)
    GenererPlanningFactoDAC "FACTO", collabs, nb
End Sub
Sub GenererPlanningDAC(collabs() As Collaborateur, nb As Integer)
    GenererPlanningFactoDAC "DAC", collabs, nb
End Sub

' ============================================================
' MISE Ã€ JOUR ROTATION
' FIX 3 : appel de TrierFeuillePlanning après CalculerCumulsSemaine
' ============================================================
Sub MettreAJourRotation(collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ROTATION")
    Dim sem As Integer
    sem = Application.WorksheetFunction.WeekNum(Date, 2)
    Dim i As Integer
    For i = 1 To nb
        Dim lr As Long
        lr = TrouverLigneRotation(ws, collabs(i).nomComplet, collabs(i).projet)
        If lr = 0 Then
            lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
            ws.Cells(lr, 1).Value = collabs(i).nomComplet
            ws.Cells(lr, 2).Value = collabs(i).projet
            ws.Cells(lr, 3).Value = 1
            ws.Cells(lr, 6).Value = 0
        Else
            ws.Cells(lr, 3).Value = CInt(ws.Cells(lr, 3).Value) + 1
        End If
        ws.Cells(lr, 4).Value = Date
        ws.Cells(lr, 5).Value = sem
    Next i

    ' Calculer les cumulés après toutes les générations
    CalculerCumulsSemaine

    ' FIX 3 : Trier la feuille PLANNING par Date puis Nom
    TrierFeuillePlanning

    ws.Columns("A:F").AutoFit
End Sub

' ============================================================
' RENFORTS - TRAITEMENT AUTOMATIQUE
'
' Feuille BESOINS colonnes :
'   A=Projet | B=Semaine | C=Jour | D=Heure début | E=Heure fin | F=Nb agents
'
' FIX 2 : La disponibilité est vérifiée dans la feuille Utilisateurs
' (comme avant). Le bug était que IncrementerNbRenforts était appelé
' avec un projet vide "". Maintenant on stocke aussi le projet du
' candidat et on le passe correctement à IncrementerNbRenforts.
'
' De plus, la boucle de recherche dans Utilisateurs cherche maintenant
' par Nom (col 1) sans dépendre d'une feuille PLANNING intermédiaire â€”
' on reconstruit la disponibilité depuis les tableaux collabs() déjà
' chargés en mémoire, ce qui est plus fiable.
' ============================================================
Sub TraiterRenforts(collabs() As Collaborateur, nb As Integer)
    If Not FeuilleExiste("BESOINS") Then Exit Sub

    Dim wsB As Worksheet
    Set wsB = ThisWorkbook.Sheets("BESOINS")
    Dim lastRow As Long
    lastRow = wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' En-têtes résultats si besoin
    If wsB.Cells(1, 7).Value = "" Then
        wsB.Cells(1, 7).Value = "Agents proposes"
        wsB.Cells(1, 8).Value = "Nb disponibles"
        wsB.Cells(1, 9).Value = "Statut"
        With wsB.Range(wsB.Cells(1, 7), wsB.Cells(1, 9))
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
        End With
    End If

    ' FIX 2 : On utilise la feuille PLANNING (déjà remplie et triée)
    ' pour vérifier les horaires réels du collab ce jour-là
    Dim wsPlan As Worksheet
    Set wsPlan = ThisWorkbook.Sheets("PLANNING")
    Dim lastPlan As Long
    lastPlan = wsPlan.Cells(wsPlan.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        Dim projetBesoin As String: projetBesoin = UCase(Trim(wsB.Cells(r, 1).Value))
        Dim semBesoin As Integer
        If IsNumeric(wsB.Cells(r, 2).Value) Then semBesoin = CInt(wsB.Cells(r, 2).Value) Else semBesoin = 0
        Dim jourBesoin As String: jourBesoin = Trim(wsB.Cells(r, 3).Value)
        Dim hdebut As String: hdebut = Trim(wsB.Cells(r, 4).Value)
        Dim hfin As String: hfin = Trim(wsB.Cells(r, 5).Value)
        Dim nbAgents As Integer
        If IsNumeric(wsB.Cells(r, 6).Value) Then nbAgents = CInt(wsB.Cells(r, 6).Value) Else nbAgents = 1

        If projetBesoin = "" Or jourBesoin = "" Or hdebut = "" Then
            wsB.Cells(r, 7).Value = "Ligne incomplete"
            wsB.Cells(r, 9).Value = "ERREUR"
            GoTo NextBesoin
        End If

        ' Type de renfort
        Dim typeRenfort As String
        If InStr(projetBesoin, "PRESS") > 0 Or projetBesoin = "EBRA PRESS" Or projetBesoin = "EBRA PRESSE" Then
            typeRenfort = "PRESS"
        ElseIf InStr(projetBesoin, "ITALY") > 0 Or projetBesoin = "COFIT" Then
            typeRenfort = "ITALY"
        Else
            typeRenfort = "PRESS"
        End If

        ' Date du jour concerné
        Dim jIdx As Integer: jIdx = NomJourToIndex(jourBesoin)
        If jIdx = 0 Then
            wsB.Cells(r, 7).Value = "Jour invalide : " & jourBesoin
            wsB.Cells(r, 9).Value = "ERREUR"
            GoTo NextBesoin
        End If
        Dim dateBesoin As Date: dateBesoin = DateDuJour(jIdx)

        ' Tableaux candidats â€” FIX 2 : on stocke aussi le projet
        Dim candidats() As String
        Dim candidatsProjets() As String
        Dim candidatsScore() As Integer
        Dim nbCandidats As Integer: nbCandidats = 0

        Dim i As Integer
        For i = 1 To nb
            ' Critère 1 : éligible renfort du bon type
            Dim eligible As Boolean
            If typeRenfort = "PRESS" Then
                eligible = collabs(i).RenforcPress
            Else
                eligible = collabs(i).RenforcItaly
            End If
            If Not eligible Then GoTo NextCollab

            ' Critère 2 : pas en congé ce jour
            If EstEnConge(collabs(i), dateBesoin) Then GoTo NextCollab

            ' FIX 2 : Vérifier disponibilité dans la feuille PLANNING
            ' (qui contient les horaires réels après génération)
            Dim dispo As Boolean: dispo = False
            Dim planEntree As String: planEntree = ""
            Dim planSortie As String: planSortie = ""
            Dim p As Long
            ' Nouveau format PLANNING : 1 ligne par collab par semaine
            ' Col 1=Semaine | Col 5=NOM COMPLET
            ' Colonnes Entrée/Sortie par jour : Lun=12/13, Mar=14/15 ... Dim=24/25
            Dim wdBesoin As Integer
            wdBesoin = Weekday(dateBesoin, vbMonday)  ' 1=Lun â€¦ 7=Dim
            Dim colPlanE As Integer: colPlanE = 10 + (wdBesoin * 2)
            Dim colPlanS As Integer: colPlanS = colPlanE + 1
            Dim semBesoinCalc As Integer
            semBesoinCalc = Application.WorksheetFunction.WeekNum(dateBesoin, 2)
            For p = 2 To lastPlan
                If wsPlan.Cells(p, 5).Value = collabs(i).nomComplet And _
                   CStr(wsPlan.Cells(p, 1).Value) = CStr(semBesoinCalc) Then
                    planEntree = CStr(wsPlan.Cells(p, colPlanE).Value)
                    planSortie = CStr(wsPlan.Cells(p, colPlanS).Value)
                    ' Nettoyer préfixe TT si présent
                    If Left(planEntree, 3) = "TT " Then planEntree = Mid(planEntree, 4)
                    If Left(planSortie, 3) = "TT " Then planSortie = Mid(planSortie, 4)
                    If planEntree <> "OFF" And planEntree <> "CONGE" And planEntree <> "" Then
                        dispo = True
                    End If
                    Exit For
                End If
            Next p
            If Not dispo Then GoTo NextCollab

            ' Critère 4 : créneau besoin dans les heures de travail du collab
            Dim hDebutM As Integer: hDebutM = HeureEnMinutes(hdebut)
            Dim hFinM As Integer: hFinM = HeureEnMinutes(hfin)
            Dim planEntM As Integer: planEntM = HeureEnMinutes(planEntree)
            Dim planSorM As Integer: planSorM = HeureEnMinutes(planSortie)
            If hDebutM < planEntM Or hFinM > planSorM Then GoTo NextCollab

            ' Candidat valide
            nbCandidats = nbCandidats + 1
            ReDim Preserve candidats(1 To nbCandidats)
            ReDim Preserve candidatsProjets(1 To nbCandidats)   ' FIX 2
            ReDim Preserve candidatsScore(1 To nbCandidats)
            candidats(nbCandidats) = collabs(i).nomComplet
            candidatsProjets(nbCandidats) = collabs(i).projet   ' FIX 2 : stocker le projet
            candidatsScore(nbCandidats) = LireNbRenforts(collabs(i).nomComplet, collabs(i).projet)

NextCollab:
        Next i

        ' Tri bubble sort par score croissant
        If nbCandidats > 1 Then
            Dim a As Integer, b As Integer
            For a = 1 To nbCandidats - 1
                For b = a + 1 To nbCandidats
                    If candidatsScore(b) < candidatsScore(a) Then
                        Dim tmpS As String: tmpS = candidats(a)
                        candidats(a) = candidats(b): candidats(b) = tmpS
                        Dim tmpP As String: tmpP = candidatsProjets(a)  ' FIX 2
                        candidatsProjets(a) = candidatsProjets(b): candidatsProjets(b) = tmpP
                        Dim tmpI As Integer: tmpI = candidatsScore(a)
                        candidatsScore(a) = candidatsScore(b): candidatsScore(b) = tmpI
                    End If
                Next b
            Next a
        End If

        ' Sélectionner les N meilleurs
        Dim proposes As String: proposes = ""
        Dim selectionnes As Integer: selectionnes = 0
        Dim c As Integer
        For c = 1 To nbCandidats
            If selectionnes >= nbAgents Then Exit For
            If proposes <> "" Then proposes = proposes & " | "
            proposes = proposes & candidats(c)
            selectionnes = selectionnes + 1
            ' FIX 2 : passer le bon projet à IncrementerNbRenforts
            IncrementerNbRenforts candidats(c), candidatsProjets(c)
        Next c

        wsB.Cells(r, 7).Value = IIf(proposes = "", "Aucun candidat disponible", proposes)
        wsB.Cells(r, 8).Value = nbCandidats
        If selectionnes >= nbAgents Then
            wsB.Cells(r, 9).Value = "OK - " & selectionnes & "/" & nbAgents
            wsB.Cells(r, 9).Interior.Color = RGB(198, 239, 206)
            wsB.Cells(r, 9).Font.Color = RGB(0, 97, 0)
        Else
            wsB.Cells(r, 9).Value = "PARTIEL - " & selectionnes & "/" & nbAgents
            wsB.Cells(r, 9).Interior.Color = RGB(255, 235, 156)
            wsB.Cells(r, 9).Font.Color = RGB(156, 87, 0)
        End If

NextBesoin:
    Next r

    wsB.Columns("A:I").AutoFit
End Sub

' ============================================================
' RESET ROTATIONS
' ============================================================
Sub ResetRotations()
    If MsgBox("Reinitialiser toutes les rotations ?", vbYesNo + vbWarning) = vbNo Then Exit Sub
    If FeuilleExiste("ROTATION") Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("ROTATION")
        If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > 1 Then
            ws.Range(ws.Cells(2, 1), ws.Cells(ws.Rows.Count, 6)).Clear
        End If
        MsgBox "Rotations reinitialisees.", vbInformation
    End If
End Sub

' ============================================================
' LANCEURS USERFORMS
' Point d'entrée principal : OuvrirUFMain
' Assigner OuvrirUFMain à un bouton du ruban ou d'une feuille
' ============================================================
Sub OuvrirUFMain()
    UFMain.Show
End Sub

Sub OuvrirUFGenerer()
    UFGenerer.Show
End Sub

Sub OuvrirUFUtilisateurs()
    UFUtilisateurs.Show
End Sub

Sub OuvrirUFBesoins()
    UFBesoins.Show
End Sub

' ============================================================
' ACTUALISER LES STATISTIQUES ETAT
' Assigner ce sub au bouton "Actualiser les statistiques"
' sur la feuille ETAT ou dans le ruban
' ============================================================
Sub ActualiserStatistiques()
    ModuleTableaux.GenererFeuille
End Sub



