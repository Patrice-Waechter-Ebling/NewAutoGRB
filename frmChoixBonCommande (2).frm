VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChoixBonCommande 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choix des pièces à commander"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   Icon            =   "frmChoixBonCommande.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   10380
   Begin VB.CommandButton cmdDeselectAll 
      Caption         =   "Aucun"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Tous"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame fraFournisseur 
      BackColor       =   &H00000000&
      Caption         =   "Fournisseurs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   10095
      Begin VB.CommandButton cmdAnnulerFRS 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   7680
         TabIndex        =   2
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdOKFRS 
         Caption         =   "OK"
         Height          =   375
         Left            =   8880
         TabIndex        =   3
         Top             =   1920
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvwFournisseur 
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fournisseur"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pers. Ress."
            Object.Width           =   1984
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   1746
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Par"
            Object.Width           =   805
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Valide"
            Object.Width           =   1746
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Prix listé"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Escompte"
            Object.Width           =   1561
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Prix net"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Prix spécial"
            Object.Width           =   1720
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Quoter"
            Object.Width           =   1191
         EndProperty
      End
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   7920
      TabIndex        =   7
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdCommander 
      Caption         =   "Commander"
      Height          =   375
      Left            =   9120
      TabIndex        =   8
      Top             =   6000
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwPiece 
      Height          =   5055
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Qté"
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No. Item"
         Object.Width           =   3228
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   6720
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Manufacturier"
         Object.Width           =   2037
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Fournisseur"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Stock"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmChoixBonCommande"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Index des colonnes de lvwPiece
Private Const I_COL_QTE As Integer = 0
Private Const I_COL_NO_ITEM As Integer = 1
Private Const I_COL_DESCRIPTION As Integer = 2
Private Const I_COL_MANUFACTURIER As Integer = 3
Private Const I_COL_FOURNISSEUR As Integer = 4
Private Const I_COL_QTE_STOCK As Integer = 5

'Index des colonnes de lvwFournisseur
Private Const I_COL_FRS_FRS As Integer = 0
Private Const I_COL_FRS_PERS_RESS As Integer = 1
Private Const I_COL_FRS_DATE As Integer = 2
Private Const I_COL_FRS_ENTRER_PAR As Integer = 3
Private Const I_COL_FRS_VALIDE As Integer = 4
Private Const I_COL_FRS_PRIX_LIST As Integer = 5
Private Const I_COL_FRS_ESCOMPTE As Integer = 6
Private Const I_COL_FRS_PRIX_NET As Integer = 7
Private Const I_COL_FRS_PRIX_SP As Integer = 8
Private Const I_COL_FRS_QUOTER As Integer = 9

'Énumération servant à savoir si le form est en anglais ou en francais
Private Enum enumLangage
 FRANCAIS = 0
 ANGLAIS = 1
End Enum

Private m_sNoProjet As String
Private m_frmSource As Form
Private m_sType As String

Private m_sIDAchat As String
Private m_iIndexAchat As Integer

Private m_collPiece As Collection
Private m_collNoLigne As Collection

Private m_eLangage As enumLangage

Private m_collNoLigneFRS As Collection
Private m_collPrixList As Collection
Private m_collPrixOrigine As Collection
Private m_collPrixNet As Collection
Private m_collEscompte As Collection
Private m_collPrixSP As Collection

Public Sub AfficherAchat(ByVal sIDAchat As String, ByVal iIndexAchat As Integer, ByVal eType As enumCatalogue)

 On Error GoTo Oups

 m_sIDAchat = sIDAchat

 m_iIndexAchat = iIndexAchat

 Set m_frmSource = frmAchat

 cmdSelectAll.Visible = True

 If eType = ELECTRIQUE Then
 m_sType = "E"
 Else
 m_sType = "M"
 End If

 Call lvwPiece.ColumnHeaders.Remove(I_COL_QTE_STOCK)

  Call RemplirListViewPieceAchat

  Call Me.Show(vbModal)

  Exit Sub

Oups:

  wOups "frmChoixBonCommande", "AfficherAchat", Err, Err.number, Err.Description
End Sub

Public Sub Afficher(ByVal sNoProjet As String, ByVal frmSource As Form, ByVal iLangage As Integer)

 On Error GoTo Oups

 'Méthode pour afficher le form
 m_sNoProjet = sNoProjet

 m_eLangage = iLangage

 Set m_frmSource = frmSource

 If frmSource.Name = "FrmProjSoumElec" Then
 m_sType = "E"
 Else
 m_sType = "M"
 End If
 
 Call RemplirListViewPieces
 
 Call Me.Show(vbModal)

  Exit Sub

Oups:

  wOups "frmChoixBonCommande", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewPieceAchat()

 On Error GoTo Oups

 'Remplis les pièces de l'achat avec la BD
 Dim rstAchat As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim itmAchat As ListItem
 Dim lColor As Long
 
 Call lvwPiece.ListItems.Clear
 
 Set rstFRS = New ADODB.Recordset
 Set rstAchat = New ADODB.Recordset
 
 Call rstAchat.Open("SELECT * FROM GrbAchat_Pieces WHERE IDAchat = '" & m_sIDAchat & "' AND IndexAchat = " & m_iIndexAchat & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
 
 Do While Not rstAchat.EOF
 If rstAchat.Fields("Recu") = True Then
  lColor = COLOR_GRIS 'Gris
  Else
  If rstAchat.Fields("Commandé") = True Then
  lColor = COLOR_ORANGE 'COLOR_ORANGE
  Else
  lColor = COLOR_NOIR
  End If
  End If

Set itmAchat = lvwPiece.ListItems.Add
 
 'Quantité
1 If Not IsNull(rstAchat.Fields("Qté")) Then
 itmAchat.Text = rstAchat.Fields("Qté")
 Else
 itmAchat.Text = vbNullString
 End If

 itmAchat.ForeColor = lColor
 
 itmAchat.Tag = rstAchat.Fields("DateRéception")

 'Numéro d'item
 If Not IsNull(rstAchat.Fields("PIECE")) Then
 itmAchat.SubItems(I_COL_NO_ITEM) = rstAchat.Fields("PIECE")
 Else
 itmAchat.SubItems(I_COL_NO_ITEM) = vbNullString
End If

 itmAchat.ListSubItems(I_COL_NO_ITEM).ForeColor = lColor

 itmAchat.ListSubItems(I_COL_NO_ITEM).Tag = rstAchat.Fields("NuméroLigne")
 
 'Description en francais
 If Not IsNull(rstAchat.Fields("DESC_FR")) Then
 itmAchat.SubItems(I_COL_DESCRIPTION) = rstAchat.Fields("DESC_FR")
 Else
 itmAchat.SubItems(I_COL_DESCRIPTION) = vbNullString
1  End If

 itmAchat.ListSubItems(I_COL_DESCRIPTION).ForeColor = lColor
 
 'On met la description en anglais dans le tag de la description en francais
 If Not IsNull(rstAchat.Fields("Desc_EN")) Then
 itmAchat.ListSubItems(I_COL_DESCRIPTION).Tag = rstAchat.Fields("Desc_EN")
 Else
 itmAchat.ListSubItems(I_COL_DESCRIPTION).Tag = vbNullString
 End If
 
 'Fabricant
 If Not IsNull(rstAchat.Fields("Manufact")) Then
 itmAchat.SubItems(I_COL_MANUFACTURIER) = rstAchat.Fields("Manufact")
 Else
 itmAchat.SubItems(I_COL_MANUFACTURIER) = vbNullString
 End If

 itmAchat.ListSubItems(I_COL_MANUFACTURIER).ForeColor = lColor

itmAchat.ListSubItems(I_COL_MANUFACTURIER).Tag = rstAchat.Fields("NoRetour")
 
 'Fournisseur
 If Not IsNull(rstAchat.Fields("IDFRS")) Then
 If rstAchat.Fields("IDFRS") <> 0 Then
 Call rstFRS.Open("SELECT NomFournisseur FROM GrbFournisseur WHERE IDFRS = " & rstAchat.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'On affiche le nom dans la colonne
 itmAchat.SubItems(I_COL_FOURNISSEUR) = rstFRS.Fields("NomFournisseur")
 
 'On affiche l'Id dans le tag
 itmAchat.ListSubItems(I_COL_FOURNISSEUR).Tag = rstAchat.Fields("IDFRS")
 
 Call rstFRS.Close
 Else
 itmAchat.SubItems(I_COL_FOURNISSEUR) = " "
End If
 Else
 itmAchat.SubItems(I_COL_FOURNISSEUR) = vbNullString
 End If

 itmAchat.ListSubItems(I_COL_FOURNISSEUR).ForeColor = lColor

 Call rstAchat.MoveNext
Loop
 
Call rstAchat.Close
Set rstAchat = Nothing

Set rstFRS = Nothing

Exit Sub

Oups:

3  wOups "frmChoixBonCommande", "RemplirListViewAchat", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewPieces()

 On Error GoTo Oups

 'Rempli le ListView selon le no. du projet
 Dim rstPieces As ADODB.Recordset
 Dim rstSection As ADODB.Recordset
 Dim rstInventaire As ADODB.Recordset
 Dim rstFRS As Recordset
 Dim itmPieces As ListItem
 Dim iCompteur As Integer
 Dim bPremierEnr As Boolean
 Dim iOrdreSection As Integer
 Dim sSousSection As String
 Dim sSection As String
  Dim lCouleur As Long
 
  bPremierEnr = True
 
  If m_eLangage = ANGLAIS Then
  sSection = "NomSectionEN"
  Else
  sSection = "NomSectionFR"
  End If
 
  lvwPiece.Sorted = False

10 Set rstFRS = New ADODB.Recordset
Set rstPieces = New ADODB.Recordset
Set rstSection = New ADODB.Recordset
Set rstInventaire = New ADODB.Recordset

 'Ouverture du recordset
Call rstPieces.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & m_sNoProjet & "' AND Type = '" & m_sType & "' AND PieceExtraChargeable = False AND PieceExtraNonChargeable = False ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
 
Do While Not rstPieces.EOF
 Set itmPieces = lvwPiece.ListItems.Add
 
 'Si c'est le premier enregistrement, il faut ajouter la section et la sous-section
 If bPremierEnr = True Then
 sSousSection = rstPieces.Fields("SousSection")
 iOrdreSection = rstPieces.Fields("OrdreSection")
 
 'Pour avoir le nom de la section
 'Si c'est un projet électrique
 If m_sType = "E" Then
 Call rstSection.Open("SELECT " & sSection & " FROM GrbSoumProjSectionElec WHERE IDSection = " & rstPieces.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstSection.Open("SELECT " & sSection & " FROM GrbSoumProjSectionMec WHERE IDSection = " & rstPieces.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
 End If

 'Ajout du nom de la section
 If Not IsNull(rstSection.Fields(sSection)) Then
 itmPieces.SubItems(I_COL_NO_ITEM) = rstSection.Fields(sSection)
 Else
 itmPieces.SubItems(I_COL_NO_ITEM) = vbNullString
1  End If
 
 itmPieces.ListSubItems(I_COL_NO_ITEM).Bold = True
 
 Call rstSection.Close
 
 Set itmPieces = lvwPiece.ListItems.Add
 
 'Ajout du nom de la sous-section
 If sSousSection = "PAS DE SOUS-SECTION" Then
 itmPieces.SubItems(I_COL_DESCRIPTION) = vbNullString
 Else
 itmPieces.SubItems(I_COL_DESCRIPTION) = sSousSection
 End If
 
 itmPieces.Tag = "PAS UNE SECTION"
 
 itmPieces.ListSubItems(I_COL_DESCRIPTION).Bold = True
 
 Set itmPieces = lvwPiece.ListItems.Add
 
 bPremierEnr = False
Else
 'Si c'est pas le premier enregistrement, il faut vérifier avec l'ancienne section
 If iOrdreSection <> rstPieces.Fields("OrdreSection") Then
 iOrdreSection = rstPieces.Fields("OrdreSection")
 
 If m_sType = "E" Then
 Call rstSection.Open("SELECT " & sSection & " FROM GrbSoumProjSectionElec WHERE IDSection = " & rstPieces.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstSection.Open("SELECT " & sSection & " FROM GrbSoumProjSectionMec WHERE IDSection = " & rstPieces.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
 End If
 
 If Not IsNull(rstSection.Fields(sSection)) Then
 itmPieces.SubItems(I_COL_NO_ITEM) = rstSection.Fields(sSection)
 Else
 itmPieces.SubItems(I_COL_NO_ITEM) = vbNullString
 End If
 
 itmPieces.ListSubItems(I_COL_NO_ITEM).Bold = True
 
 Call rstSection.Close
 
 Set itmPieces = lvwPiece.ListItems.Add
 
 sSousSection = rstPieces.Fields("SousSection")
 
 If sSousSection = "PAS DE SOUS-SECTION" Then
 itmPieces.SubItems(I_COL_DESCRIPTION) = vbNullString
 Else
 itmPieces.SubItems(I_COL_DESCRIPTION) = rstPieces.Fields("SousSection")
 End If
 
 itmPieces.Tag = "PAS UNE SECTION"
 
 itmPieces.ListSubItems(I_COL_DESCRIPTION).Bold = True
 
 Set itmPieces = lvwPiece.ListItems.Add
 Else
 'il faut vérifier avec l'ancienne sous-section
 If sSousSection <> rstPieces.Fields("SousSection") Then
 sSousSection = rstPieces.Fields("SousSection")
 
 If sSousSection = "PAS DE SOUS-SECTION" Then
4 itmPieces.SubItems(I_COL_DESCRIPTION) = vbNullString
4 Else
4 itmPieces.SubItems(I_COL_DESCRIPTION) = sSousSection
4 End If
 
4 itmPieces.ListSubItems(I_COL_DESCRIPTION).Bold = True

4 itmPieces.Tag = "PAS UNE SECTION"
 
4 Set itmPieces = lvwPiece.ListItems.Add
4 End If
4 End If
4 End If
 
4 If rstPieces.Fields("Commandé") = True Then
4  lCouleur = COLOR_ORANGE
4  Else
4  If rstPieces.Fields("Recu") = True Then
4  lCouleur = COLOR_GRIS
4  Else
4  lCouleur = COLOR_NOIR
4  End If
4  End If
 
 'Quantité
50 If Not IsNull(rstPieces.Fields("Qté")) Then
itmPieces.Text = rstPieces.Fields("Qté")
 Else
 itmPieces.Text = vbNullString
 End If

 itmPieces.ForeColor = lCouleur
 
 itmPieces.Tag = "PAS UNE SECTION"
 
 'Numéro d'item
 If Not IsNull(rstPieces.Fields("NumItem")) Then
 itmPieces.SubItems(I_COL_NO_ITEM) = rstPieces.Fields("NumItem")
 Else
 itmPieces.SubItems(I_COL_NO_ITEM) = vbNullString
 End If

5  itmPieces.ListSubItems(I_COL_NO_ITEM).ForeColor = lCouleur

5  itmPieces.ListSubItems(I_COL_NO_ITEM).Tag = rstPieces.Fields("NuméroLigne")
 
5  If m_eLangage = FRANCAIS Then
 'Description en francais
5  If Not IsNull(rstPieces.Fields("Desc_FR")) Then
5  itmPieces.SubItems(I_COL_DESCRIPTION) = rstPieces.Fields("Desc_FR")
5  Else
5  itmPieces.SubItems(I_COL_DESCRIPTION) = vbNullString
5  End If
60 Else
 'Description en anglais
  If Not IsNull(rstPieces.Fields("Desc_EN")) Then
  itmPieces.SubItems(I_COL_DESCRIPTION) = rstPieces.Fields("Desc_EN")
  Else
  itmPieces.SubItems(I_COL_DESCRIPTION) = vbNullString
  End If
  End If
 
  itmPieces.ListSubItems(I_COL_DESCRIPTION).ForeColor = lCouleur
 
 'Fabricant
  If Not IsNull(rstPieces.Fields("Manufact")) Then
  itmPieces.SubItems(I_COL_MANUFACTURIER) = rstPieces.Fields("Manufact")
  Else
  itmPieces.SubItems(I_COL_MANUFACTURIER) = vbNullString
6  End If
 
6  itmPieces.ListSubItems(I_COL_MANUFACTURIER).ForeColor = lCouleur
 
 'Fournisseur
6  If Not IsNull(rstPieces.Fields("IDFRS")) And rstPieces.Fields("IDFRS") > 0 Then
6  If itmPieces.SubItems(I_COL_NO_ITEM) <> "Texte" Then
6  Call rstFRS.Open("SELECT NomFournisseur FROM GrbFournisseur WHERE IDFRS = " & rstPieces.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'On affiche le nom dans la colonne
6  itmPieces.SubItems(I_COL_FOURNISSEUR) = rstFRS.Fields("NomFournisseur")
 
6  Call rstFRS.Close
6  End If
70 Else
  itmPieces.SubItems(I_COL_FOURNISSEUR) = vbNullString
  End If
 
  itmPieces.ListSubItems(I_COL_FOURNISSEUR).ForeColor = lCouleur
 
  If m_sType = "E" Then
  Call rstInventaire.Open("SELECT * FROM GrbInventaireElec WHERE NoItem = '" & Replace(rstPieces.Fields("NumItem"), "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
  Else
  Call rstInventaire.Open("SELECT * FROM GrbInventaireMec WHERE NoItem = '" & Replace(rstPieces.Fields("NumItem"), "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
  End If

  If Not rstInventaire.EOF Then
  itmPieces.SubItems(I_COL_QTE_STOCK) = rstInventaire.Fields("QuantitéStock")
  End If

   Call rstInventaire.Close
 
   Call rstPieces.MoveNext
7  Loop
 
7  Call rstPieces.Close
7  Set rstPieces = Nothing

7  Set rstFRS = Nothing
7  Set rstInventaire = Nothing
7  Set rstSection = Nothing

80 Exit Sub

Oups:

80 wOups "frmChoixBonCommande", "RemplirListViewPieces", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixBonCommande", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdCommander_Click()

 On Error GoTo Oups

 Dim bChecked As Boolean
 Dim iCompteur As Integer
 Dim sNoBC As String
 Dim rstProjet As ADODB.Recordset
 
 bChecked = False
 
 For iCompteur = 1 To lvwPiece.ListItems.count
 If lvwPiece.ListItems(iCompteur).Checked = True Then
 bChecked = True
 
 Exit For
 End If
  Next
 
  If bChecked = True Then
  Set m_collPiece = New Collection
  Set m_collNoLigne = New Collection

  If m_frmSource.Name <> "frmAchat" Then
  Call ModifierFournisseurBD
  End If
 
  For iCompteur = 1 To lvwPiece.ListItems.count
 If lvwPiece.ListItems(iCompteur).Checked = True Then
 Call m_collPiece.Add(lvwPiece.ListItems(iCompteur).SubItems(I_COL_NO_ITEM))
 Call m_collNoLigne.Add(lvwPiece.ListItems(iCompteur).ListSubItems(I_COL_NO_ITEM).Tag)
 End If
 Next
 
 If m_frmSource.Name <> "frmAchat" Then
 Set rstProjet = New ADODB.Recordset

 If m_sType = "E" Then
 Call rstProjet.Open("SELECT ProchaineCommande FROM GrbProjetElec WHERE IDProjet = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstProjet.Open("SELECT ProchaineCommande FROM GrbProjetMec WHERE IDProjet = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If

 If Not IsNull(rstProjet.Fields("ProchaineCommande")) Then
 sNoBC = m_sNoProjet & "-" & Right$("00" & rstProjet.Fields("ProchaineCommande"), 3)
 Else
 sNoBC = m_sNoProjet
 End If

 Call rstProjet.Close
 Set rstProjet = Nothing
1  Else
 sNoBC = m_sIDAchat & "-" & Right$("00" & m_iIndexAchat, 3)
 End If

 If sNoBC <> vbNullString Then
 If m_frmSource.Name = "FrmProjSoumElec" Then
 Call frmBonCommande.AfficherFormProjetAchat(m_sNoProjet, sNoBC, m_collPiece, m_collNoLigne, I_PROJET_ELEC, m_eLangage)
 Else
 If m_frmSource.Name = "FrmProjSoumMec" Then
 Call frmBonCommande.AfficherFormProjetAchat(m_sNoProjet, sNoBC, m_collPiece, m_collNoLigne, I_PROJET_MEC, m_eLangage)
 Else
 If m_sType = "E" Then
 Call frmBonCommande.AfficherFormProjetAchat(m_sIDAchat & "-" & Right$("00" & m_iIndexAchat, 3), sNoBC, m_collPiece, m_collNoLigne, I_ACHAT_ELEC, 0)
 Else
 
 Call frmBonCommande.AfficherFormProjetAchat(m_sIDAchat & "-" & Right$("00" & m_iIndexAchat, 3), sNoBC, m_collPiece, m_collNoLigne, I_ACHAT_MEC, 0)
 End If
 End If
 End If

 Call Unload(Me)
 End If
2  Else
 Call MsgBox("Aucune pièce n'est sélectionnée!", vbOKOnly, "Erreur")
30 End If

Exit Sub

Oups:

wOups "frmChoixBonCommande", "cmdCommander_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdSelectAll_Click()

 On Error GoTo Oups

 Dim iCompteur As Integer

 For iCompteur = 1 To lvwPiece.ListItems.count
 If lvwPiece.ListItems(iCompteur).SubItems(I_COL_NO_ITEM) <> "Texte" And lvwPiece.ListItems(iCompteur).SubItems(I_COL_NO_ITEM) <> "Text" Then
 If m_frmSource.Name <> "frmAchat" Then
 If lvwPiece.ListItems(iCompteur).Tag <> vbNullString Then
 If lvwPiece.ListItems(iCompteur).SubItems(I_COL_NO_ITEM) <> vbNullString Then
 If CDbl(lvwPiece.ListItems(iCompteur).Text) > 0 Then
 lvwPiece.ListItems(iCompteur).Checked = True
 End If
 End If
  End If
  Else
  If CDbl(lvwPiece.ListItems(iCompteur).Text) > 0 Then
  lvwPiece.ListItems(iCompteur).Checked = True
  End If
  End If
  End If
  Next

10 Exit Sub

Oups:

wOups "frmChoixBonCommande", "cmdSelectAll_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDeSelectAll_Click()

 On Error GoTo Oups

 Dim iCompteur As Integer

 For iCompteur = 1 To lvwPiece.ListItems.count
 lvwPiece.ListItems(iCompteur).Checked = False
 Next

 Exit Sub

Oups:

 wOups "frmChoixBonCommande", "cmdDeselectAll_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Set m_collNoLigneFRS = New Collection
 Set m_collEscompte = New Collection
 Set m_collPrixList = New Collection
 Set m_collPrixNet = New Collection
 Set m_collPrixOrigine = New Collection
 Set m_collPrixSP = New Collection

 Exit Sub

Oups:

 wOups "frmChoixBonCommande", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub lvwPiece_DblClick()

 On Error GoTo Oups

 
 If m_frmSource.Name <> "frmAchat" Then
 'Si ce n'est pas une section
 If lvwPiece.SelectedItem.Tag <> "" Then
 'Si ce n'est pas une sous-section
 If lvwPiece.SelectedItem.SubItems(I_COL_NO_ITEM) <> "" Then
 Call ChangerFournisseur
 End If
 End If
 End If

 Exit Sub

Oups:

 wOups "frmChoixBonCommande", "lvwPiece_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwPiece_ItemCheck(ByVal Item As MSComctlLib.ListItem)

 On Error GoTo Oups
 
 If m_frmSource.Name <> "frmAchat" Then
 'Si c'est du texte
 If Item.SubItems(I_COL_NO_ITEM) = "Texte" Or Item.Tag = vbNullString Or Item.SubItems(I_COL_NO_ITEM) = vbNullString Then
 'On enlève le check
 Item.Checked = False
 Else
 If CDbl(Replace(Item.Text, "*", "")) <= 0 Then
 'On enlève le check
 Item.Checked = False
 End If
 End If
 End If

  Exit Sub

Oups:

  wOups "frmChoixBonCommande", "lvwPiece_ItemCheck", Err, Err.number, Err.Description
End Sub

Private Sub ChangerFournisseur()

 On Error GoTo Oups

 Call AfficherListeFournisseurs

 If lvwfournisseur.ListItems.count = 0 Then
 Call MsgBox("Il n'y a aucun fournisseur pour cette pièce!", vbOKOnly, "Erreur")
 
 Exit Sub
 Else
 frafournisseur.Visible = True
 End If

 Exit Sub

Oups:

 wOups "frmChoixBonCommande", "ChangerFournisseur", Err, Err.number, Err.Description
End Sub

Private Sub AfficherListeFournisseurs()

 On Error GoTo Oups

 'Méthode qui sert à afficher la liste des fournisseurs
 'Affiche le frame seulement s'il y a des items dans le ListView
 Call RemplirListViewFournisseur
 
 If lvwfournisseur.ListItems.count > 1 Then
 frafournisseur.Visible = True
 Call lvwfournisseur.SetFocus
 End If

 Exit Sub

Oups:

 wOups "frmChoixBonCommande", "AfficherListeFournisseurs", Err, Err.number, Err.Description
End Sub

Private Sub lvwFournisseur_DblClick()

 On Error GoTo Oups

 Call ChoisirFournisseur

 Exit Sub

Oups:

 wOups "frmChoixBonCommande", "lvwFournisseur_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub ChoisirFournisseur()

 On Error GoTo Oups

 Dim itmBC As ListItem
 Dim itmFRS As ListItem

 Set itmBC = lvwPiece.SelectedItem
 Set itmFRS = lvwfournisseur.SelectedItem

 itmBC.SubItems(I_COL_FOURNISSEUR) = itmFRS.Text

 itmBC.ListSubItems(I_COL_FOURNISSEUR).Tag = itmFRS.Tag

 Call m_collNoLigneFRS.Add(itmBC.ListSubItems(I_COL_NO_ITEM).Tag)
 Call m_collEscompte.Add(itmFRS.SubItems(I_COL_FRS_ESCOMPTE))
 Call m_collPrixList.Add(itmFRS.SubItems(I_COL_FRS_PRIX_LIST))
 Call m_collPrixNet.Add(itmFRS.SubItems(I_COL_FRS_PRIX_NET))
  Call m_collPrixOrigine.Add(itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag)
  Call m_collPrixSP.Add(itmFRS.SubItems(I_COL_FRS_PRIX_SP))

 'On cache le listview
  frafournisseur.Visible = False

  Exit Sub

Oups:

  wOups "frmChoixBonCommande", "ChoisirFournisseur", Err, Err.number, Err.Description
End Sub

Private Sub cmdOKFRS_Click()

 On Error GoTo Oups

 Call ChoisirFournisseur

 Exit Sub

Oups:

 wOups "frmChoixBonCommande", "cmdOKFRS_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulerFRS_Click()

 On Error GoTo Oups

 frafournisseur.Visible = False

 Exit Sub

Oups:

 wOups "frmChoixBonCommande", "cmdAnnulerFRS_Click", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewFournisseur()

 On Error GoTo Oups

 'Rempli le listview des distributeur pour une pièce choisie
 Dim rstPieceFRS As ADODB.Recordset
 Dim rstContact As ADODB.Recordset
 Dim rstFRS As Recordset
 Dim iCompteur As Integer
 Dim itmFRS As ListItem
 Dim iNoClient As Integer
 Dim sDevise As String
 
 'vide le lister
 Call lvwfournisseur.ListItems.Clear

 Set rstPieceFRS = New ADODB.Recordset
 Set rstContact = New ADODB.Recordset
  Set rstFRS = New ADODB.Recordset

  Call rstFRS.Open("SELECT IDFRS FROM GrbFournisseur WHERE NomFournisseur = 'FOURNI PAR LE CLIENT'", g_connData, adOpenDynamic, adLockOptimistic)
 
  iNoClient = rstFRS.Fields("IDFRS")

  Call rstFRS.Close
  Set rstFRS = Nothing
 
  Call rstPieceFRS.Open("SELECT GrbPiecesFRS.*, GrbFournisseur.NomFournisseur FROM GrbPiecesFRS INNER JOIN GrbFournisseur ON GrbPiecesFRS.IDFRS = GrbFournisseur.IDFRS WHERE PIECE = '" & Replace(lvwPiece.SelectedItem.SubItems(I_COL_NO_ITEM), "'", "''") & "' AND Type = '" & m_sType & "' ORDER BY PrixReel", g_connData, adOpenDynamic, adLockOptimistic)
 
 'tant il y a des fournisseur de la piece , ajoute dans lister
  Do While Not rstPieceFRS.EOF
 'on change la couleur de l'enregistrement selon la devise monétaire.
 'CAN = rouge, USA ou ESP = bleu
  If rstPieceFRS.Fields("DeviseMonétaire") = "CAN" Then
 sDevise = "CAN"
1 Else
 If rstPieceFRS.Fields("DeviseMonétaire") = "USA" Then
 sDevise = "USA"
 Else
 sDevise = "SPA"
 End If
 End If
 
 Set itmFRS = lvwfournisseur.ListItems.Add
 
 itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = " "
 itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = " "
 itmFRS.SubItems(I_COL_FRS_PRIX_NET) = " "
itmFRS.SubItems(I_COL_FRS_PRIX_SP) = " "
 
 'Nom du FRS
 itmFRS.Text = rstPieceFRS.Fields("NomFournisseur")
 
 itmFRS.Tag = rstPieceFRS.Fields("IDFRS")
 
 'Personne ressource
 If Trim(rstPieceFRS.Fields("PERS_RESS")) <> vbNullString Then
 Call rstContact.Open("SELECT NomContact FROM GrbContact WHERE IDContact = " & rstPieceFRS.Fields("PERS_RESS"), g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstContact.EOF Then
 itmFRS.SubItems(I_COL_FRS_PERS_RESS) = rstContact.Fields("NomContact")
1  End If

 Call rstContact.Close
 End If
 
 'Date
 If Not IsNull(rstPieceFRS.Fields("Date")) Then
 itmFRS.SubItems(I_COL_FRS_DATE) = rstPieceFRS.Fields("Date")
 Else
 itmFRS.SubItems(I_COL_FRS_DATE) = vbNullString
 End If
 
 'Entrer par
 If Not IsNull(rstPieceFRS.Fields("Entrer_Par")) Then
 itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = rstPieceFRS.Fields("Entrer_Par")
 Else
 itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = vbNullString
 End If
  
 'Valide
If Not IsNull(rstPieceFRS.Fields("Valide")) Then
 itmFRS.SubItems(I_COL_FRS_VALIDE) = rstPieceFRS.Fields("Valide")
Else
 itmFRS.SubItems(I_COL_FRS_VALIDE) = vbNullString
End If
 
 'Prix listé
 If rstPieceFRS.Fields("PRIX_LIST") <> vbNullString Then
 If sDevise = "USA" Then
 itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = Conversion(CStr(Round(CDbl(rstPieceFRS.Fields("PRIX_LIST")), 4)), MODE_ARGENT, 4)
 Else
 If sDevise = "SPA" Then
 itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = Conversion(CStr(Round(CDbl(rstPieceFRS.Fields("PRIX_LIST")), 4)), MODE_ARGENT, 4)
 Else
 itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = Conversion(Round(Replace(rstPieceFRS.Fields("PRIX_LIST"), ".", ","), 4), MODE_ARGENT, 4)
 End If
 End If
 End If
 
 'Escompte
 If rstPieceFRS.Fields("ESCOMPTE") <> vbNullString Then
 itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = Conversion(Replace(Replace(rstPieceFRS.Fields("ESCOMPTE"), "_", vbNullString), ".", ",") * 100, MODE_POURCENT)
 End If
 
 'Prix net
 If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
 If sDevise = "USA" Then
 itmFRS.SubItems(I_COL_FRS_PRIX_NET) = Conversion(CStr(Round(CDbl(rstPieceFRS.Fields("PRIX_NET")), 4)), MODE_ARGENT, 4)
 Else
 If sDevise = "SPA" Then
 itmFRS.SubItems(I_COL_FRS_PRIX_NET) = Conversion(CStr(Round(CDbl(rstPieceFRS.Fields("PRIX_NET")), 4)), MODE_ARGENT, 4)
 Else
 itmFRS.SubItems(I_COL_FRS_PRIX_NET) = Conversion(Round(Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ","), 4), MODE_ARGENT, 4)
 End If
 End If
4 End If
 
 'Prix spécial
4 If rstPieceFRS.Fields("PRIX_SP") <> vbNullString Then
4 If sDevise = "USA" Then
4 itmFRS.SubItems(I_COL_FRS_PRIX_SP) = Conversion(CStr(Round(CDbl(rstPieceFRS.Fields("PRIX_SP")), 4)), MODE_ARGENT, 4)
4 Else
4 If sDevise = "SPA" Then
4 itmFRS.SubItems(I_COL_FRS_PRIX_SP) = Conversion(Round(CDbl(rstPieceFRS.Fields("PRIX_SP")), 4), MODE_ARGENT, 4)
4 Else
4 itmFRS.SubItems(I_COL_FRS_PRIX_SP) = Conversion(Round(rstPieceFRS.Fields("PRIX_SP"), 4), MODE_ARGENT, 4)
4 End If
4 End If
4  End If
 
 'Quoter
4  If rstPieceFRS.Fields("QUOTER") = True Then
4  itmFRS.SubItems(I_COL_FRS_QUOTER) = "Oui"
4  Else
4  itmFRS.SubItems(I_COL_FRS_QUOTER) = "Non"
4  End If
 
 'Pour garder en mémoire le prix d'origine, je le mets dans le
 'tag de la colonne Prix Listé
4  If itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = vbNullString Then
4  itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = " "
50 End If
 
5 If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
 itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ",")
 Else
 itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_SP"), ".", ",")
 End If

 'Pour avoir le no d'enregistrement de PiecesFRS

 If itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = vbNullString Then
 itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = " "
 End If

 itmFRS.ListSubItems(I_COL_FRS_ENTRER_PAR).Tag = rstPieceFRS.Fields("NoEnreg")
 
 Call rstPieceFRS.MoveNext
5  Loop
 
 'ferme la table
5  Call rstPieceFRS.Close
5  Set rstPieceFRS = Nothing

5  Exit Sub

Oups:

5  wOups "frmChoixBonCommande", "RemplirListViewFournisseur", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTotalRecordsetElec(ByVal sNoProjSoum As String)

 On Error GoTo Oups

 'Méthode pour calculer le prix
 Dim rstProjet As ADODB.Recordset
 Dim rstPiece As ADODB.Recordset
 Dim rstPunch As ADODB.Recordset
 Dim dblTotalDessin As Double
 Dim dblTotalFabrication As Double
 Dim dblTotalAssemblage As Double
 Dim dblTotalProgInterface As Double
 Dim dblTotalProgAutomate As Double
 Dim dblTotalProgRobot As Double
 Dim dblTotalVision As Double
  Dim dblTotalTest As Double
  Dim dblTotalInstallation As Double
  Dim dblTotalMiseService As Double
  Dim dblTotalFormation As Double
  Dim dblTotalGestion As Double
  Dim dblTotalShipping As Double
  Dim dblHebergement As Double
  Dim dblRepas As Double
10 Dim dblTransport As Double
Dim dblUniteMobile As Double
Dim dblPrixEmballage As Double
Dim dblTotalResteTemps As Double
Dim dblPrixPieces As Double
Dim dblPrixTotal As Double
Dim dblCommission As Double
Dim dblTotalTemps As Double
Dim dblProfit As Double
Dim dblTotalManuel As Double
Dim dblTotalPieceImprevue As Double
Dim dblGrandTotal As Double
1  Dim sDateDebut As String
Dim sDateFin As String
 Dim sTotal As String

Set rstProjet = New ADODB.Recordset

 Call rstProjet.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

If Not rstProjet.EOF Then
 Set rstPunch = New ADODB.Recordset

 'Total des temps
1  sDateDebut = "TIMESERIAL(Left(GrbPunch.HeureDébut,2),RIGHT(GrbPunch.HeureDébut,2),0)"

 sDateFin = "TIMESERIAL(Left(GrbPunch.HeureFin,2),RIGHT(GrbPunch.HeureFin,2),0)"

 sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"


 Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GrbPunch WHERE NoProjet = '" & sNoProjSoum & "' And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)

 dblTotalDessin = 0
 dblTotalFabrication = 0
 dblTotalAssemblage = 0
 dblTotalProgInterface = 0
 dblTotalProgAutomate = 0
 dblTotalProgRobot = 0
 dblTotalVision = 0
 dblTotalTest = 0
 dblTotalInstallation = 0
dblTotalMiseService = 0
 dblTotalFormation = 0
dblTotalGestion = 0
 dblTotalShipping = 0

Do While Not rstPunch.EOF
 If Not IsNull(rstPunch.Fields("Total")) Then
 Select Case rstPunch.Fields("Type")
 Case "Dessin": dblTotalDessin = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxDessin"))
 Case "Fabrication": dblTotalFabrication = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxFabrication"))
 Case "Assemblage": dblTotalAssemblage = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxAssemblage"))
 Case "ProgInterface": dblTotalProgInterface = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxProgInterface"))
 Case "ProgAutomate": dblTotalProgAutomate = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxProgAutomate"))
 Case "ProgRobot": dblTotalProgRobot = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxProgRobot"))
 Case "Vision": dblTotalVision = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxVision"))
 Case "Test": dblTotalTest = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxTest"))
 Case "Installation": dblTotalInstallation = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxInstallation"))
 Case "MiseService": dblTotalMiseService = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxMiseService"))
 Case "Formation": dblTotalFormation = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxFormation"))
 Case "Gestion": dblTotalGestion = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxGestion"))
 Case "Shipping": dblTotalShipping = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxShipping"))
 End Select
 End If
 
 Call rstPunch.MoveNext
Loop

 Call rstPunch.Close
Set rstPunch = Nothing

 dblTotalTemps = dblTotalDessin + _
 dblTotalFabrication + _
 dblTotalAssemblage + _
 dblTotalProgInterface + _
 dblTotalProgAutomate + _
 dblTotalProgRobot + _
 dblTotalVision + _
 dblTotalTest + _
 dblTotalInstallation + _
 dblTotalMiseService + _
 dblTotalFormation + _
 dblTotalGestion + _
 dblTotalShipping
 
 Set rstPiece = New ADODB.Recordset

 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoProjSoum & "' AND Type = 'E'", g_connData, adOpenDynamic, adLockOptimistic)

 'Pour chaque élément du recordset
Do While Not rstPiece.EOF
4 If Trim(rstPiece.Fields("Prix_total")) <> vbNullString Then
 'On additionne le prix total
4 dblPrixPieces = dblPrixPieces + CDbl(rstPiece.Fields("Prix_total")) - CDbl(rstPiece.Fields("Profit_Argent"))

 'On additionne le profit
4 dblProfit = dblProfit + CDbl(rstPiece.Fields("Profit_Argent"))
4 End If

4 Call rstPiece.MoveNext
4 Loop

4 Call rstPiece.Close
4 Set rstPiece = Nothing

4 dblHebergement = 0
4 dblRepas = 0
4 dblTransport = 0
4  dblUniteMobile = 0

 'Correction d'un bug de Type Incompatible
4  If IsNumeric(rstProjet.Fields("PrixEmballage")) Then
4  dblPrixEmballage = CDbl(rstProjet.Fields("PrixEmballage"))
4  Else
4  dblPrixEmballage = 0
4  End If
 
4  dblTotalResteTemps = dblHebergement + dblRepas + dblTransport + dblUniteMobile + dblPrixEmballage

4  If IsNumeric(rstProjet.Fields("total_manuel")) Then
50 dblTotalManuel = CDbl(rstProjet.Fields("total_manuel"))
5 Else
 dblTotalManuel = 0
 End If

 dblTotalPieceImprevue = (dblPrixPieces + dblProfit) * (1 + CDbl(rstProjet.Fields("Imprevue")))

 dblPrixTotal = dblTotalTemps + dblTotalManuel + dblTotalPieceImprevue + dblTotalResteTemps

 'Le calcul de la commission est sur les manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps
 dblCommission = dblPrixTotal * CDbl(rstProjet.Fields("Commission"))

 'Le prix total est le calcul des manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps + total de la commission
 dblGrandTotal = dblPrixTotal + dblCommission

 'Format monétaires avec 2 chiffres après la virgule
 rstProjet.Fields("total_commission") = dblCommission
 rstProjet.Fields("Total_manuel") = dblTotalManuel
 rstProjet.Fields("Total_temps") = dblTotalTemps
 rstProjet.Fields("total_imprevue") = dblTotalPieceImprevue - (dblPrixPieces + dblProfit)
5  rstProjet.Fields("total_piece") = dblPrixPieces
5  rstProjet.Fields("Total_Commission") = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
5  rstProjet.Fields("PrixTotal") = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
5  rstProjet.Fields("Total_Profit") = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)

5  Call rstProjet.Update
5  Else
5  Call MsgBox("Le projet " & sNoProjSoum & " est inexistant!", vbOKOnly, "Erreur")
5  End If

60 Call rstProjet.Close
60 Set rstProjet = Nothing

  Exit Sub

Oups:

  wOups "frmChoixBonCommande", "CalculerTotalRecordset", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTotalRecordsetMec(ByVal sNoProjSoum As String)

 On Error GoTo Oups

 'Méthode pour calculer le prix
 Dim rstProjet As ADODB.Recordset
 Dim rstPiece As ADODB.Recordset
 Dim rstPunch As ADODB.Recordset
 Dim dblPrixPieces As Double
 Dim dblPrixTotal As Double
 Dim dblCommission As Double
 Dim dblTotalTemps As Double
 Dim dblProfit As Double
 Dim dblTotalManuel As Double
 Dim dblTotalImprevue As Double
  Dim dblGrandTotal As Double
  Dim dblTotalDessin As Double
  Dim dblTotalCoupe As Double
  Dim dblTotalMachinage As Double
  Dim dblTotalSoudure As Double
  Dim dblTotalAssemblage As Double
  Dim dblTotalPeinture As Double
  Dim dblTotalTest As Double
10 Dim dblTotalInstallation As Double
Dim dblTotalFormation As Double
Dim dblTotalGestion As Double
Dim dblTotalShipping As Double
Dim dblPrixEmballage As Double
Dim dblTotalResteTemps As Double
Dim sDateDebut As String
Dim sDateFin As String
Dim sTotal As String

Set rstProjet = New ADODB.Recordset

Call rstProjet.Open("SELECT * FROM GrbProjetMec WHERE IDProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

If Not rstProjet.EOF Then
Set rstPiece = New ADODB.Recordset

 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoProjSoum & "' AND Type = 'M'", g_connData, adOpenDynamic, adLockOptimistic)

 'Pour chaque élément du recordset
 Do While Not rstPiece.EOF
 If Trim(rstPiece.Fields("Prix_total")) <> vbNullString Then
 'On additionne le prix total
 dblPrixPieces = dblPrixPieces + rstPiece.Fields("Prix_total") - rstPiece.Fields("Profit_Argent")
 
 'On additionne le profit
 dblProfit = dblProfit + rstPiece.Fields("Profit_Argent")
 End If

1  Call rstPiece.MoveNext
 Loop
 
 'Total des temps
 sDateDebut = "TIMESERIAL(Left(GrbPunch.HeureDébut,2),RIGHT(GrbPunch.HeureDébut,2),0)"

 sDateFin = "TIMESERIAL(Left(GrbPunch.HeureFin,2),RIGHT(GrbPunch.HeureFin,2),0)"

 sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"

 Set rstPunch = New ADODB.Recordset

 Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GrbPunch WHERE NoProjet = '" & sNoProjSoum & "' And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)
 
 dblTotalDessin = 0
 dblTotalCoupe = 0
 dblTotalMachinage = 0
 dblTotalSoudure = 0
 dblTotalAssemblage = 0
 dblTotalPeinture = 0
dblTotalTest = 0
 dblTotalInstallation = 0
dblTotalFormation = 0
 dblTotalGestion = 0
dblTotalShipping = 0

 Do While Not rstPunch.EOF
 If Not IsNull(rstPunch.Fields("Total")) Then
 Select Case rstPunch.Fields("Type")
 Case "Dessin": dblTotalDessin = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxDessin"))
 Case "Coupe": dblTotalCoupe = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxCoupe"))
 Case "Machinage": dblTotalMachinage = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxMachinage"))
 Case "Soudure": dblTotalSoudure = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxSoudure"))
 Case "Assemblage": dblTotalAssemblage = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxAssemblage"))
 Case "Peinture": dblTotalPeinture = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxPeinture"))
 Case "Test": dblTotalTest = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxTest"))
 Case "Installation": dblTotalInstallation = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxInstallation"))
 Case "Formation": dblTotalFormation = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxFormation"))
 Case "Gestion": dblTotalGestion = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxGestion"))
 Case "Shipping": dblTotalShipping = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxShipping"))
 End Select
 End If

 Call rstPunch.MoveNext
 Loop

Call rstPunch.Close
 Set rstPunch = Nothing
 
dblTotalTemps = dblTotalDessin + _
 dblTotalCoupe + _
 dblTotalMachinage + _
 dblTotalSoudure + _
 dblTotalAssemblage + _
 dblTotalPeinture + _
 dblTotalTest + _
 dblTotalInstallation + _
 dblTotalFormation + _
 dblTotalGestion + _
 dblTotalShipping
 
 'Correction d'un bug de Type Incompatible
 If IsNumeric(rstProjet.Fields("PrixEmballage")) Then
 dblPrixEmballage = CDbl(rstProjet.Fields("PrixEmballage"))
 Else
 dblPrixEmballage = 0
4 End If
 
4 dblTotalResteTemps = dblPrixEmballage
  
4 If IsNumeric(rstProjet.Fields("total_manuel")) Then
4 dblTotalManuel = CDbl(rstProjet.Fields("total_manuel"))
4 Else
4 dblTotalManuel = 0
4 End If
 
4 dblTotalImprevue = Round((dblPrixPieces + dblProfit) * CDbl(rstProjet.Fields("Imprevue")), 2)
 
4 dblPrixTotal = dblPrixPieces + dblProfit + dblTotalTemps + dblTotalImprevue + dblTotalManuel + dblTotalResteTemps
 
 'Le calcul de la commission est sur les manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps
4 dblCommission = Round(dblPrixTotal * CDbl(rstProjet.Fields("Commission")), 2)
 
 'Le prix total est le calcul des manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps + total de la commission
4 dblGrandTotal = dblPrixTotal + dblCommission

 'Format monétaires avec 2 chiffres après la virgule
4  rstProjet.Fields("Total_Piece") = Conversion(CStr(Round(dblPrixPieces, 2)), MODE_ARGENT)
4  rstProjet.Fields("Total_Temps") = Conversion(CStr(Round(dblTotalTemps, 2)), MODE_ARGENT)
4  rstProjet.Fields("PrixTotal") = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
4  rstProjet.Fields("total_Imprevue") = Conversion(CStr(Round(dblTotalImprevue, 2)), MODE_ARGENT)
4  rstProjet.Fields("total_commission") = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
4  rstProjet.Fields("total_profit") = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)

4  Call rstProjet.Update

4  Call rstPiece.Close
50 Set rstPiece = Nothing
50 Else
 Call MsgBox("Le projet " & sNoProjSoum & " est inexistant!", vbOKOnly, "Erreur")
 End If

 Call rstProjet.Close
 Set rstProjet = Nothing

 Exit Sub

Oups:

 wOups "frmChoixBonCommande", "CalculerTotalRecordset", Err, Err.number, Err.Description
End Sub

Private Sub ModifierFournisseurBD()

 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 Dim rstProjet As ADODB.Recordset
 Dim sProfit As String
 Dim iCompteur As Integer
 Dim bModif As Boolean
 Dim iCompteurColl As Integer
 Dim iIndexColl As Integer
 Dim sLiaison As String

 Set rstProjet = New ADODB.Recordset

 If m_sType = "E" Then
  Call rstProjet.Open("SELECT Profit, LiaisonChargeable FROM GrbProjetElec WHERE IDProjet = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
  Else
  Call rstProjet.Open("SELECT Profit, LiaisonChargeable FROM GrbProjetMec WHERE IDProjet = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
  End If

  sProfit = rstProjet.Fields("Profit")

  If Not IsNull(rstProjet.Fields("LiaisonChargeable")) Then
  sLiaison = rstProjet.Fields("LiaisonChargeable")
  Else
sLiaison = ""
End If

Call rstProjet.Close
Set rstProjet = Nothing

Set rstPiece = New ADODB.Recordset

For iCompteur = 1 To lvwPiece.ListItems.count
 If lvwPiece.ListItems(iCompteur).Checked = True Then
 If lvwPiece.ListItems(iCompteur).ListSubItems(I_COL_FOURNISSEUR).Tag <> "" Then
 bModif = True

 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & m_sNoProjet & "' AND NuméroLigne = " & lvwPiece.ListItems(iCompteur).ListSubItems(I_COL_NO_ITEM).Tag, g_connData, adOpenDynamic, adLockOptimistic)

 For iCompteurColl = 1 To m_collNoLigneFRS.count
 If m_collNoLigneFRS(iCompteurColl) = lvwPiece.ListItems(iCompteur).ListSubItems(I_COL_NO_ITEM).Tag Then
 iIndexColl = iCompteurColl

 Exit For
 End If
 Next

 rstPiece.Fields("IDFRS") = lvwPiece.ListItems(iCompteur).ListSubItems(I_COL_FOURNISSEUR).Tag

 'Prix listé
 If Trim$(m_collPrixList(iIndexColl)) = vbNullString Then
 rstPiece.Fields("PRIX_LIST") = Conversion("0", MODE_PAS_FORMAT, 4)
1  Else
 rstPiece.Fields("PRIX_LIST") = Conversion(m_collPrixList(iIndexColl), MODE_PAS_FORMAT, 4)
 rstPiece.Fields("PrixOrigine") = Conversion(m_collPrixOrigine(iIndexColl), MODE_PAS_FORMAT, 4)
 End If
 
 'S'il y a un prix net, on le met l'escompte et le prix net sinon, on prend le prix
 'spécial pour mettre dans le prix net
 If Trim$(m_collPrixNet(iIndexColl)) <> vbNullString Then
 rstPiece.Fields("ESCOMPTE") = Conversion(m_collEscompte(iIndexColl), MODE_PAS_FORMAT)
 rstPiece.Fields("PRIX_NET") = Conversion(m_collPrixNet(iIndexColl), MODE_PAS_FORMAT, 4)
 Else
 If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) <> vbNullString Then
 rstPiece.Fields("PRIX_NET") = Conversion(m_collPrixSP(iIndexColl), MODE_PAS_FORMAT, 4)
 Else
 rstPiece.Fields("ESCOMPTE") = Conversion("0", MODE_PAS_FORMAT)
 rstPiece.Fields("PRIX_NET") = Conversion("0", MODE_PAS_FORMAT, 4)
 End If
 End If
 
 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
 rstPiece.Fields("Prix_Total") = Conversion(CStr(Round(Replace(rstPiece.Fields("Qté"), "*", vbNullString) * rstPiece.Fields("PRIX_NET") * CSng(sProfit), 2)), MODE_PAS_FORMAT)
 
 'Pour le profit, c'est le prix total - (prix net * quantité)
 rstPiece.Fields("Profit_argent") = Conversion(CStr(Round(rstPiece.Fields("Prix_Total") - (rstPiece.Fields("PRIX_NET") * Replace(rstPiece.Fields("Qté"), "*", vbNullString)), 2)), MODE_PAS_FORMAT)

 Call rstPiece.Update

 Call rstPiece.Close

 If sLiaison <> "" Then
 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & Left$(m_sNoProjet, Len(m_sNoProjet) - 2) & sLiaison & "' AND Provenance = '" & Right$(m_sNoProjet, 2) & "' AND NumItem = '" & lvwPiece.ListItems(iCompteur).SubItems(I_COL_NO_ITEM) & "' AND Qté = '" & lvwPiece.ListItems(iCompteur).Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

 For iCompteurColl = 1 To m_collNoLigneFRS.count
 If m_collNoLigneFRS(iCompteurColl) = lvwPiece.ListItems(iCompteur).ListSubItems(I_COL_NO_ITEM).Tag Then
 iIndexColl = iCompteurColl

 Exit For
 End If
 Next

 rstPiece.Fields("IDFRS") = lvwPiece.ListItems(iCompteur).ListSubItems(I_COL_FOURNISSEUR).Tag

 'Prix listé
 If Trim$(m_collPrixList(iIndexColl)) = vbNullString Then
 rstPiece.Fields("PRIX_LIST") = Conversion("0", MODE_PAS_FORMAT, 4)
 Else
 rstPiece.Fields("PRIX_LIST") = Conversion(m_collPrixList(iIndexColl), MODE_PAS_FORMAT, 4)
 rstPiece.Fields("PrixOrigine") = Conversion(m_collPrixOrigine(iIndexColl), MODE_PAS_FORMAT, 4)
 End If
 
 'S'il y a un prix net, on le met l'escompte et le prix net sinon, on prend le prix
 'spécial pour mettre dans le prix net
 If Trim$(m_collPrixNet(iIndexColl)) <> vbNullString Then
 rstPiece.Fields("ESCOMPTE") = Conversion(m_collEscompte(iIndexColl), MODE_PAS_FORMAT)
 rstPiece.Fields("PRIX_NET") = Conversion(m_collPrixNet(iIndexColl), MODE_PAS_FORMAT, 4)
 Else
 If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) <> vbNullString Then
 rstPiece.Fields("PRIX_NET") = Conversion(m_collPrixSP(iIndexColl), MODE_PAS_FORMAT, 4)
 Else
 rstPiece.Fields("ESCOMPTE") = Conversion("0", MODE_PAS_FORMAT)
4 rstPiece.Fields("PRIX_NET") = Conversion("0", MODE_PAS_FORMAT, 4)
4 End If
4 End If
 
 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
4 rstPiece.Fields("Prix_Total") = Conversion(CStr(Round(Replace(rstPiece.Fields("Qté"), "*", vbNullString) * rstPiece.Fields("PRIX_NET") * CSng(sProfit), 2)), MODE_PAS_FORMAT)
 
 'Pour le profit, c'est le prix total - (prix net * quantité)
4 rstPiece.Fields("Profit_argent") = Conversion(CStr(Round(rstPiece.Fields("Prix_Total") - (rstPiece.Fields("PRIX_NET") * Replace(rstPiece.Fields("Qté"), "*", vbNullString)), 2)), MODE_PAS_FORMAT)

4 Call rstPiece.Update

4 Call rstPiece.Close
4 End If
4 End If
4 End If
4 Next

4  Set rstPiece = Nothing

4  If bModif = True Then
4  If m_sType = "E" Then
4  Call CalculerTotalRecordsetElec(m_sNoProjet)

4  If sLiaison <> "" Then
4  Call CalculerTotalRecordsetElec(Left$(m_sNoProjet, Len(m_sNoProjet) - 2) & sLiaison)
4  End If

4  FrmProjSoumElec.m_bModifFournisseurBC = True
50 Else
Call CalculerTotalRecordsetMec(m_sNoProjet)

 If sLiaison <> "" Then
 Call CalculerTotalRecordsetMec(Left$(m_sNoProjet, Len(m_sNoProjet) - 2) & sLiaison)
 End If

 FrmProjSoumMec.m_bModifFournisseurBC = True
 End If
 Else
 If m_sType = "E" Then
 FrmProjSoumElec.m_bModifFournisseurBC = False
 Else
 FrmProjSoumMec.m_bModifFournisseurBC = False
5  End If
5  End If

5  Exit Sub

Oups:

5  wOups "frmChoixBonCommande", "ModifierFournisseurBD", Err, Err.number, Err.Description
End Sub
