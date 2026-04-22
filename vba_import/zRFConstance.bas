Attribute VB_Name = "zRFConstance"
Option Explicit

' =============================================
' FEUILLES
' =============================================
Public Const SHEET_MAIN As String = "GMC"
Public Const SHEET_TITRES As String = "titres"
Public Const SHEET_MENU_DEROULANT As String = "MENU DEROULANT"

' =============================================
' MOT DE PASSE / SECURITE
' =============================================
Public Const MDP_DEV As String = "11223344"

' =============================================
' SESSION
' =============================================
Public Const CELL_NOM_SESSION As String = "A1"
Public Const CELL_DATE_SESSION As String = "B1"

' =============================================
' STRUCTURE
' =============================================
Public Const ROW_START As Long = 8
Public Const ROW_RECHERCHE As Long = 4
Public Const ROW_HEADER As Long = 7
Public Const ROW_TITRES As Long = 1

Public Const COL_FIRST As String = "A"
Public Const COL_LAST As String = "AE"
Public Const COL_LAST_RECHERCHE As String = "AD"

Public Const NB_COL_RECHERCHE As Long = 30
Public Const NB_COL_UI As Long = 30
Public Const NB_COL_TABLE As Long = 26
Public Const NB_COL_BLOC_MAIN As Long = 13
Public Const FIRST_COL_BLOC_ASSOC As Long = 14

Public Const PLAGE_RECHERCHE As String = "A4:AD4"

' =============================================
' COLONNES RF - BLOC PRINCIPAL
' =============================================
Public Const COL_TRANCHE_MAIN As String = "A"
Public Const COL_SYSTEME_MAIN As String = "B"
Public Const COL_NUMERO_MAIN As String = "C"
Public Const COL_BIGRAMME_MAIN As String = "D"
Public Const COL_COMP1_MAIN As String = "E"
Public Const COL_COMP2_MAIN As String = "F"
Public Const COL_LIBELLE_MAIN As String = "G"
Public Const COL_LOCAL_MAIN As String = "H"
Public Const COL_QUALITE_MAIN As String = "I"
Public Const COL_CARROYAGE_MAIN As String = "J"
Public Const COL_FOLIO_MAIN As String = "K"
Public Const COL_CONTROLE_MAIN As String = "L"
Public Const COL_SITUATION_MAIN As String = "M"

' =============================================
' COLONNES RF - BLOC ASSOCIE
' =============================================
Public Const COL_TRANCHE_ASSOC As String = "N"
Public Const COL_SYSTEME_ASSOC As String = "O"
Public Const COL_NUMERO_ASSOC As String = "P"
Public Const COL_BIGRAMME_ASSOC As String = "Q"
Public Const COL_COMP1_ASSOC As String = "R"
Public Const COL_COMP2_ASSOC As String = "S"
Public Const COL_LIBELLE_ASSOC As String = "T"
Public Const COL_LOCAL_ASSOC As String = "U"
Public Const COL_QUALITE_ASSOC As String = "V"
Public Const COL_CARROYAGE_ASSOC As String = "W"
Public Const COL_FOLIO_ASSOC As String = "X"
Public Const COL_CONTROLE_ASSOC As String = "Y"
Public Const COL_SITUATION_ASSOC As String = "Z"

' =============================================
' SUIVI AGENT / CONTROLE
' =============================================
Public Const COL_DATE As String = "AA"
Public Const COL_NOM As String = "AB"
Public Const COL_CONF As String = "AC"
Public Const COL_OBS As String = "AD"
Public Const COL_AIDE As String = "AE"

' =============================================
' FACTORISATION - EDITION / PROTECTION
' =============================================
Public Const PLAGE_EDITABLE_MAIN As String = "G:Z"
Public Const PLAGE_EDITABLE_SUIVI As String = "AC:AD"
Public Const PLAGE_EDITABLE_AIDE As String = "AE:AE"

Public Const COL_VALIDATION_CONF As String = "AC"

' =============================================
' FACTORISATION - COLLAGE
' =============================================
Public Const PLAGE_COLLER_RECHERCHE As String = "A4:AD4"
Public Const PLAGE_COLLER_EDITABLE As String = "G:Z"
Public Const PLAGE_COLLER_SUIVI As String = "AC:AD"

' =============================================
' VALEURS CONFORMITE
' =============================================
Public Const VAL_CONF_1 As String = "conforme"
Public Const VAL_CONF_2 As String = "non conforme"
Public Const VAL_CONF_3 As String = "examen"

Public Const MSG_VALEURS_CONF As String = "Valeurs autorisées : conforme, non conforme ou examen."

' =============================================
' BOUTON RECHERCHE RF PRINCIPAL
' =============================================
Public Const COL_RF_CONCAT As String = "AE"
Public Const COL_RECHERCHE_EXACTE As String = COL_RF_CONCAT

' =============================================
' BOUTON PREMIERE LIGNE VIDE / COLONNE MASQUEE
' =============================================
Public Const COL_PREMIERE_LIGNE_VIDE As String = "G"
Public Const COLONNES_MASQUEES As String = "N:Z"

' =============================================
' BOUTON ZOOM
' =============================================
Public Const ZOOM_ECRAN_PRINCIPAL As Long = 74
Public Const ZOOM_ECRAN_SECONDAIRE As Long = 97

' =============================================
' USERFORM UF-AGENT
' =============================================
Public Const COL_AGENTS As String = "C"
Public Const ROW_AGENTS_START As Long = 2
Public Const ROW_AGENTS_END As Long = 44

' =============================================
' COMPORTEMENT
' =============================================
Public Const HAS_MAIN_EDIT_ZONE As Boolean = True
Public Const HAS_AIDE_ZONE As Boolean = True
Public Const HAS_DOC_DOUBLECLICK As Boolean = False
Public Const HAS_RECHERCHE_EXACTE As Boolean = True
Public Const HAS_COLONNES_MASQUEES As Boolean = True
Public Const HAS_PREMIERE_LIGNE_VIDE As Boolean = True

' =============================================
' COULEURS
' =============================================
Public Const COLOR_RECHERCHE_FOND As Long = 16510410
Public Const COLOR_RECHERCHE_ACTIVE As Long = 15781618
Public Const COLOR_PLACEHOLDER As Long = 9868950
Public Const COLOR_TEXTE_NOIR As Long = 0
Public Const COLOR_BORDURE_BLEUE As Long = 10318348
Public Const COLOR_BORDURE_VIOLETTE As Long = 9644960
Public Const COLOR_ERREUR_ROUGE As Long = 9869055

' =============================================
' COULEURS CONFORMITE
' =============================================
Public Const COLOR_CONF_CONFORME     As Long = 10675893  ' #B5E6A2 vert
Public Const COLOR_CONF_NON_CONFORME As Long = 14524132  ' #E49EDD rose
Public Const COLOR_CONF_EXAMEN       As Long = 7531262   ' #FEEA72 jaune

' =============================================
' SEUILS / SECURITE
' =============================================
Public Const MAX_SELECTION_CHANGE As Long = 500
Public Const SEUIL_CONFIRMATION_MASSE As Long = 2000
Public Const SEUIL_BLOCAGE_MASSE As Long = 15000

' =====================================================
' LISTES DE CONTROLE AU COLLAGE (BDD-RF)
' =====================================================
Public Const HAS_VALIDATION_LISTES_COLLAGE As Boolean = True
Public Const SHEET_LISTES_RF As String = "_DataProvider"

Public Const PLAGE_LISTE_TRANCHE As String = "D4:D13"
Public Const PLAGE_LISTE_POSE As String = "F4:F5"
Public Const PLAGE_LISTE_SITUATION As String = "H4:H6"
Public Const PLAGE_LISTE_QUALITE As String = "L4:L7"

Public Const PLAGE_COLONNES_LISTE_TRANCHE As String = "A:A,N:N"
Public Const PLAGE_COLONNES_LISTE_POSE As String = "L:L,Y:Y"
Public Const PLAGE_COLONNES_LISTE_SITUATION As String = "M:M,Z:Z"
Public Const PLAGE_COLONNES_LISTE_QUALITE As String = "I:I,V:V"

Public Const MSG_VALEURS_TRANCHE As String = "Valeurs autorisées : 0 à 9."
Public Const MSG_VALEURS_POSE As String = "Valeurs autorisées : O ou N."
Public Const MSG_VALEURS_SITUATION As String = "Valeurs autorisées : EX, HX ou NC."
Public Const MSG_VALEURS_QUALITE As String = "Valeurs autorisées : NQS, QS, IPS ou NC."

' =============================================
' REFERENTIELS RF (BDD-RF)
' =============================================
Public Const SHEET_LOCAUX As String = "local"
Public Const SHEET_CARROYAGES As String = "carroyage"
Public Const COL_REFERENTIEL As String = "A"
Public Const ROW_REFERENTIEL_START As Long = 2

Public Const MSG_VALEURS_LOCAL As String = "Local non autorisé."
Public Const MSG_VALEURS_CARROYAGE As String = "Carroyage non autorisé."



