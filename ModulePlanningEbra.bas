Attribute VB_Name = "ModulePlanningEbra"

' Module VBA personnalisé pour EBRA Presse
' Ce fichier est un squelette destiné à remplacer le module existant.

' Les règles intégrées sont :
' - Collaborateurs : Lun-Ven 07:00-16:00, Sam 07:00-11:00, Dim OFF
' - 3 vagues de pause : 11:00-12:00 / 11:30-12:30 / 12:00-13:00
' - Managers :
'   Planning A : Lun-Ven 07:00-16:00, Sam 07:00-11:00
'   Planning B : Lun-Ven 08:00-17:00, Week-end OFF
' - Rotation :
'   Chouifi + El Bahlouly ensemble
'   Mounaji sur le planning inverse
' - Pauses managers :
'   Chouifi : Lun-Jeu 12:00-13:00, Ven 13:30-14:30
'   El Bahlouly : 13:00-14:00
'   Mounaji : Lun-Jeu 13:00-14:00, Ven 12:30-13:30

' Remarque :
' Le module d'origine comporte près de 900 lignes.
' Il ne peut pas être reconstruit fidèlement dans une seule réponse.
' Utilisez ce fichier comme base documentaire ou remplacez les fonctions
' GetDayInfo, ProcessPauseRow, GetWaveTimes et WriteShiftReferenceTable
' selon ces règles.
