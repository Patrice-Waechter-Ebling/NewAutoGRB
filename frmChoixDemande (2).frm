VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChoixDemande 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demande de prix"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   8760
   Begin MSComctlLib.ListView lvwCategorie 
      Height          =   4335
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7646
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Catégorie"
         Object.Width           =   12303
      EndProperty
   End
   Begin MSComctlLib.ListView lvwPiece 
      Height          =   3975
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7011
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Quantité"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Pièce"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description française"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description anglaise"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Fabricant"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdDeselectAll 
      Caption         =   "Aucun"
      Height          =   375
      Left            =   960
      TabIndex        =   20
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Tous"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   5160
      Width           =   735
   End
   Begin VB.ComboBox cmbTri 
      Height          =   315
      ItemData        =   "frmChoixDemande.frx":0000
      Left            =   4440
      List            =   "frmChoixDemande.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdRechercher 
      Caption         =   "Rechercher"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   300
      Width           =   1095
   End
   Begin VB.TextBox txtRechercher 
      Height          =   285
      Left            =   5640
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdRafraichir 
      Caption         =   "Rafraichir"
      Height          =   315
      Left            =   7680
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdTri 
      Caption         =   "Trier"
      Height          =   315
      Left            =   6600
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   7560
      TabIndex        =   31
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox txtCommentaire 
      Height          =   495
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   5880
      Width           =   4695
   End
   Begin VB.TextBox txtNoGRB 
      Height          =   285
      Left            =   5040
      TabIndex        =   27
      Top             =   5520
      Width           =   1215
   End
   Begin MSMask.MaskEdBox mskDateRequise 
      Height          =   255
      Left            =   5040
      TabIndex        =   23
      Top             =   5160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdLangage 
      Caption         =   "En anglais"
      Height          =   375
      Left            =   2280
      TabIndex        =   21
      Top             =   5160
      Width           =   1335
   End
   Begin VB.ComboBox cmbCategorie 
      Height          =   315
      Left            =   4440
      TabIndex        =   2
      Text            =   "cmbCategorie"
      Top             =   360
      Width           =   4215
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Height          =   375
      Left            =   6360
      TabIndex        =   30
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   7560
      TabIndex        =   25
      Top             =   5160
      Width           =   1095
   End
   Begin VB.ComboBox cmbManufacturier 
      Height          =   315
      Left            =   2640
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtNoPiece 
      Height          =   285
      Left            =   2640
      MaxLength       =   37
      TabIndex        =   11
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtDescription 
      Height          =   525
      Left            =   5040
      MaxLength       =   61
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdAjouter 
      Caption         =   "Ajouter"
      Height          =   375
      Left            =   7560
      TabIndex        =   14
      Top             =   1320
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwManufacturier 
      Height          =   4335
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7646
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Manufacturier"
         Object.Width           =   8784
      EndProperty
   End
   Begin MSComctlLib.ListView lvwNouvellesPieces 
      Height          =   3255
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5741
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Quantité"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Pièce"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Manufacturier"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Catégorie"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwFournisseur 
      Height          =   4335
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7646
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nom Fournisseur"
         Object.Width           =   8784
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Langage de la demande"
         Object.Width           =   3387
      EndProperty
   End
   Begin VB.Label lblFormatDate 
      BackStyle       =   0  'Transparent
      Caption         =   "AA-MM-JJ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   24
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Description :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblManufacturier 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturier :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblNoPiece 
      BackStyle       =   0  'Transparent
      Caption         =   "No de pièce :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblCommentaire 
      BackStyle       =   0  'Transparent
      Caption         =   "Commentaire"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   28
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label lblNoGRB 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "# GRB :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   26
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label lblDateRequise 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date Requise :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblCategorie 
      BackStyle       =   0  'Transparent
      Caption         =   "Categorie de pièce :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmChoixDemande"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Index de cmbTri
Private Const I_CMB_PIECE As Integer = 0
Private Const I_CMB_DESCRIPTION_FR As Integer = 1
Private Const I_CMB_DESCRIPTION_EN As Integer = 2
Private Const I_CMB_FABRICANT As Integer = 3

'Index des colonnes de lvwPiece
Private Const I_COL_QUANTITE As Integer = 0
Private Const I_COL_PIECE As Integer = 1
Private Const I_COL_DESC_FR As Integer = 2
Private Const I_COL_DESC_EN As Integer = 3
Private Const I_COL_FABRICANT As Integer = 4
 
'Index des colonnes de lvwFournisseur
Private Const I_COL_NOM_FRS As Integer = 0
Private Const I_COL_LANGAGE As Integer = 1
 
'Index des colonnes de lvwNouvellesPieces
Private Const I_COL_QTE As Integer = 0
Private Const I_COL_NO_PIECE As Integer = 1
Private Const I_COL_DESCRIPTION As Integer = 2
Private Const I_COL_MANUFACT As Integer = 3
Private Const I_COL_CATEGORIE As Integer = 4
 
'Caption du bouton cmdLangage
Private Const S_DEMANDE_FRANCAIS As String = "En français"
Private Const S_DEMANDE_ANGLAIS As String = "En anglais"
 
'Texte de la colonne Langage de la demande de lvwFournisseur
Private Const S_FRANCAIS As String = "Français"
Private Const S_ANGLAIS As String = "Anglais"

'Pour savoir si l'appel de la demande a été fait à partir d'une soumission
'ou d'un projet
Private Enum enumType
 TYPE_PROJET = 0
 TYPE_SOUMISSION = 1
End Enum

'Pour savoir quel ListView est affiché
Private Enum enumMode
 Fournisseur = 0
 PIECE = 1
 Categorie = 2
 NOUVELLE_PIECE = 3
 Manufacturier = 4
End Enum

Public Enum enumModeDemande
 MODE_PIECE = 0 'Pour une pièce
 MODE_FOURNISSEUR = 1 'Pour toutes les pièces d'un fournisseur
 MODE_CATEGORIE = 2 'Pour catégorie
 MODE_NOUVELLE = 3
End Enum

Private m_eMode As enumMode

Private m_eDemande As enumModeDemande

'Contient la valeur électrique ou mécanique
Private m_eCatalogue As enumCatalogue

'Pour conserver en mémoire les pièces choisies
Private m_collPiece As Collection
Private m_collQuantite As Collection
Private m_collDescriptionFR As Collection
Private m_collDescriptionEN As Collection
Private m_collCategorie As Collection
Private m_collManufacturier As Collection

'Pour savoir si les fournisseurs sont affichés après avoir choisit des pièces
Private m_bPiece As Boolean

'Pour savoir si la catégorie a changé ou non,
'sert également pour mettre dans m_collCategorie
Private m_sCategorie As String

Private m_sLangage As String

'Pour savoir si la demande a été fait à partir des Projets / Soumissions
Private m_bProjSoum As Boolean

'Pour savoir si la demande a été fait à partir des achats
Private m_bAchat As Boolean

'Contient le numéro du projet si la demande à été fait à partir d'un Projet
Private m_sNoProjSoum As String

'Contient le numéro de l'achat si la demande a été fait à partir des achats
Private m_sNoAchat As String

'Contient l'index de l'achat si la demande a été fait à partir des achats
Private m_iIndexAchat As Integer

'Pour savoir à quel index la rechercher est rendu
Private m_iIndexRecherche As Integer

'Pour savoir par quoi trier le ListView
Private m_sTri As String

'Pour savoir si c'est une soumission ou un projet
Private m_eType As enumType

Public m_bAnnulerContact As Boolean
Public m_sContact As String

Public Sub Afficher(ByVal eCatalogue As enumCatalogue, ByVal eDemande As enumModeDemande)

 On Error GoTo Oups

 m_eCatalogue = eCatalogue
 m_eDemande = eDemande
 m_bProjSoum = False
 
 Select Case eDemande
 Case MODE_FOURNISSEUR:
 Call RemplirListViewFournisseur(False)
 
 Call AfficherControles(Fournisseur)
 
 Case MODE_PIECE:
 Call RemplirComboCategorie
 
 Call AfficherControles(PIECE)
 
 Case MODE_CATEGORIE:
  Call RemplirListViewCatalogue
 
  Call AfficherControles(Categorie)
 
  Case MODE_NOUVELLE:
  Call RemplirComboCategorie
 
  Call AfficherControles(NOUVELLE_PIECE)

  If m_eDemande = MODE_NOUVELLE Then
  Call RemplirComboManufacturiers
  End If
10 End Select
 
Call Show(vbModal)

Exit Sub

Oups:

wOups "frmChoixDemande", "Afficher", Err, Err.number, Err.Description
End Sub

Public Sub AfficherProjetSoumission(ByVal sNoProjSoum As String, ByVal eCatalogue As enumCatalogue, ByVal eDemande As enumModeDemande, ByVal iType As Integer)

 On Error GoTo Oups

 m_eCatalogue = eCatalogue
 m_eDemande = eDemande
 m_sNoProjSoum = sNoProjSoum
 m_eType = iType

 txtNoGRB.Text = sNoProjSoum
 
 Call RemplirListViewPieceProjetSoumission
 
 Call AfficherControles(PIECE)
 
 cmbTri.Visible = False
 cmdTri.Visible = False
 cmdRafraichir.Visible = False
 
  cmbCategorie.Visible = False
  lblCategorie.Visible = False
 
  m_bProjSoum = True
 
  Call Me.Show(vbModal)

  Exit Sub

Oups:

  wOups "frmChoixDemande", "AfficherProjet", Err, Err.number, Err.Description
End Sub

Public Sub AfficherAchat(ByVal sNoAchat As String, ByVal eCatalogue As enumCatalogue, ByVal eDemande As enumModeDemande)

 On Error GoTo Oups

 m_eCatalogue = eCatalogue
 m_eDemande = eDemande

 txtNoGRB.Text = sNoAchat

 m_sNoAchat = Left$(sNoAchat, 9)
 m_iIndexAchat = CInt(Right$(sNoAchat, 3))
 
 Call RemplirListViewPieceAchat
 
 Call AfficherControles(PIECE)
 
 cmbTri.Visible = False
 cmdTri.Visible = False
 cmdRafraichir.Visible = False
 
  cmbCategorie.Visible = False
  lblCategorie.Visible = False
 
  m_bAchat = True
 
  Call Show(vbModal)

  Exit Sub

Oups:

  wOups "frmChoixDemande", "AfficherProjet", Err, Err.number, Err.Description
End Sub

Private Sub AfficherControles(ByVal eMode As enumMode)

 On Error GoTo Oups

 Dim bCategorie As Boolean
 Dim bLvwPiece As Boolean
 Dim bLvwFournisseur As Boolean
 Dim bLvwCategorie As Boolean
 Dim bLvwManufacturier As Boolean
 Dim bLvwNouvelle As Boolean
 Dim bCmdOK As Boolean
 Dim bCmdImprimer As Boolean
 Dim bCmdLangage As Boolean
 Dim bNoGRB As Boolean
  Dim bDate As Boolean
  Dim bCommentaire As Boolean
  Dim bNoPiece As Boolean
  Dim bManufact As Boolean
  Dim bDescription As Boolean
  Dim bCmdAjouter As Boolean
  Dim iHeight As Integer
  Dim bRechercher As Boolean
10 Dim bTri As Boolean
Dim bSelectAll As Boolean
Dim bDeselectAll As Boolean
 
m_eMode = eMode
 
Select Case eMode
 Case Fournisseur:
 bLvwFournisseur = True
 bCmdImprimer = True
 bCmdLangage = True
 bNoGRB = True
 bDate = True
 bCommentaire = True
 bRechercher = True
 bSelectAll = True
 bDeselectAll = True
 
 iHeight = 6855
 
 Case PIECE:
 bCategorie = True
 bLvwPiece = True
 bCmdOK = True
 bTri = True
1  bSelectAll = True
 bDeselectAll = True
 
 iHeight = 6150

 Case Manufacturier:
 bLvwManufacturier = True
 bCmdOK = True
 
 bSelectAll = True
 bDeselectAll = True
 
 iHeight = 6150
 
 Case Categorie:
 bLvwCategorie = True
 bCmdOK = True
 bSelectAll = True
 bDeselectAll = True
 
 iHeight = 6150
 
 Case NOUVELLE_PIECE:
 bLvwNouvelle = True
 bNoPiece = True
 bManufact = True
 bDescription = True
 bCategorie = True
 bCmdOK = True
 bCmdAjouter = True
 
 iHeight = 6150
30 End Select
 
Me.Height = iHeight
 
lblCategorie.Visible = bCategorie
cmbCategorie.Visible = bCategorie
 
lvwPiece.Visible = bLvwPiece
lvwfournisseur.Visible = bLvwFournisseur
lvwCategorie.Visible = bLvwCategorie
lvwNouvellesPieces.Visible = bLvwNouvelle
lvwManufacturier.Visible = bLvwManufacturier
 
cmdSelectAll.Visible = bSelectAll
cmdDeselectAll.Visible = bDeselectAll
 
cmdOk.Visible = bCmdOK
 
3  lblNoPiece.Visible = bNoPiece
txtNoPiece.Visible = bNoPiece
 
3  lblManufacturier.Visible = bManufact
cmbManufacturier.Visible = bManufact
 
3  lblDescription.Visible = bDescription
txtDescription.Visible = bDescription
 
3  Cmdajouter.Visible = bCmdAjouter
 
 cmdImprimer.Visible = bCmdImprimer
40 cmdLangage.Visible = bCmdLangage

lblNoGRB.Visible = bNoGRB
4 txtNoGRB.Visible = bNoGRB
 
4 lblDateRequise.Visible = bDate
4 mskDateRequise.Visible = bDate
4 lblFormatDate.Visible = bDate
 
4 lblCommentaire.Visible = bCommentaire
4 txtcommentaire.Visible = bCommentaire
 
4 txtRechercher.Visible = bRechercher
4 cmdRechercher.Visible = bRechercher
 
4 cmbTri.Visible = bTri
4 cmdRafraichir.Visible = bTri
4  cmdTri.Visible = bTri

4  Exit Sub

Oups:

4  wOups "frmChoixDemande", "AfficherControles", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewManufacturier()
 
 On Error GoTo Oups

 Dim rstManufact As ADODB.Recordset
 Dim sWhere As String
 Dim iCompteur As Integer

 Call lvwManufacturier.ListItems.Clear

 lvwManufacturier.Sorted = True
 lvwManufacturier.SortKey = 0

 sWhere = "CATEGORIE In ("

 For iCompteur = 1 To m_collCategorie.count
 If iCompteur <> m_collCategorie.count Then
 sWhere = sWhere & "'" & Replace(m_collCategorie(iCompteur), "'", "''") & "',"
  Else
  sWhere = sWhere & "'" & Replace(m_collCategorie(iCompteur), "'", "''") & "')"
  End If
  Next

  Set rstManufact = New ADODB.Recordset

  If m_eCatalogue = ELECTRIQUE Then
  Call rstManufact.Open("SELECT DISTINCT FABRICANT FROM GrbCatalogueElec WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
  Else
Call rstManufact.Open("SELECT DISTINCT FABRICANT FROM GrbCatalogueMec WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
End If

Do While Not rstManufact.EOF
 Call lvwManufacturier.ListItems.Add(, , rstManufact.Fields("FABRICANT"))
 
 Call rstManufact.MoveNext
Loop

Call rstManufact.Close
Set rstManufact = Nothing

Exit Sub

Oups:

wOups "frmChoixDemande", "RemplirListViewManufacturier", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewFournisseur(ByVal bPiece As Boolean)

 On Error GoTo Oups

 Dim rstFRS As ADODB.Recordset
 Dim itmFRS As ListItem
 Dim sWhere As String
 Dim iCompteur As Integer
 
 m_bPiece = bPiece
 
 Call lvwfournisseur.ListItems.Clear
 
 Set rstFRS = New ADODB.Recordset
 
 If bPiece = False Then
 Call rstFRS.Open("SELECT NomFournisseur, IDFRS FROM GrbFournisseur WHERE Supprimé = False ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
 Else
  sWhere = "PIECE In ("

  For iCompteur = 1 To m_collPiece.count
  If iCompteur <> m_collPiece.count Then
  sWhere = sWhere & "'" & Replace(m_collPiece(iCompteur), "'", "''") & "',"
  Else
  sWhere = sWhere & "'" & Replace(m_collPiece(iCompteur), "'", "''") & "')"
  End If
  Next
 
Call rstFRS.Open("SELECT DISTINCT GrbFournisseur.NomFournisseur, GrbFournisseur.IDFRS FROM GrbPiecesFRS INNER JOIN GrbFournisseur ON GrbPiecesFRS.IDFRS = GrbFournisseur.IDFRS WHERE " & sWhere & " ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
End If
 
Do While Not rstFRS.EOF
 Set itmFRS = lvwfournisseur.ListItems.Add
 
 itmFRS.Text = rstFRS.Fields("NomFournisseur")
 
 itmFRS.Tag = rstFRS.Fields("IDFRS")
 
 itmFRS.SubItems(I_COL_LANGAGE) = S_FRANCAIS
 
 Call rstFRS.MoveNext
Loop
 
Call rstFRS.Close
Set rstFRS = Nothing

Exit Sub

Oups:

1  wOups "frmChoixDemande", "RemplirListViewFournisseur", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboCategorie()

 On Error GoTo Oups

 Dim rstCategorie As ADODB.Recordset
 
 Call cmbCategorie.Clear
 
 Set rstCategorie = New ADODB.Recordset
 
 If m_eCatalogue = ELECTRIQUE Then
 Call rstCategorie.Open("SELECT DISTINCT CATEGORIE FROM GrbCatalogueElec ORDER BY Categorie", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstCategorie.Open("SELECT DISTINCT CATEGORIE FROM GrbCatalogueMec ORDER BY Categorie", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 
 Do While Not rstCategorie.EOF
 If Not IsNull(rstCategorie.Fields("CATEGORIE")) Then
  Call cmbCategorie.AddItem(rstCategorie.Fields("CATEGORIE"))
  End If
 
  Call rstCategorie.MoveNext
  Loop
 
  Call rstCategorie.Close
  Set rstCategorie = Nothing
 
  If cmbCategorie.ListCount > 0 Then
  cmbCategorie.ListIndex = 0
10 End If

Exit Sub

Oups:

wOups "frmChoixDemande", "RemplirComboCategorie", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewCatalogue()

 On Error GoTo Oups

 Dim rstCategorie As ADODB.Recordset
 
 Call lvwCategorie.ListItems.Clear
 
 Set rstCategorie = New ADODB.Recordset
 
 If m_eCatalogue = ELECTRIQUE Then
 Call rstCategorie.Open("SELECT DISTINCT CATEGORIE FROM GrbCatalogueElec ORDER BY CATEGORIE", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstCategorie.Open("SELECT DISTINCT CATEGORIE FROM GrbCatalogueMec ORDER BY CATEGORIE", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 
 Do While Not rstCategorie.EOF
 Call lvwCategorie.ListItems.Add(, , rstCategorie.Fields("CATEGORIE"))
 
  Call rstCategorie.MoveNext
  Loop
 
  Call rstCategorie.Close
  Set rstCategorie = Nothing

  Exit Sub

Oups:

  wOups "frmChoixDemande", "RemplirListViewCatalogue", Err, Err.number, Err.Description
End Sub

Private Function TrouverIndexPiece(ByVal sPiece As String, ByVal sDescriptionFR As String, ByVal sDescriptionEN As String, ByVal sFabricant As String, ByVal iIndexActuel As Integer)

 On Error GoTo Oups

 Dim iIndex As Integer
 Dim sTri As String
 Dim bDebut As Boolean
 
 sTri = UCase(m_sTri)
 
 sPiece = UCase(sPiece)
 sDescriptionFR = UCase(sDescriptionFR)
 sDescriptionEN = UCase(sDescriptionEN)
 sFabricant = UCase(sFabricant)
 
 If sTri <> vbNullString Then
 bDebut = False
 
 'Selon le tri
  Select Case cmbTri.ListIndex
 'Si c'est trier par PIECE
 Case I_CMB_PIECE:
 'Si la PIECE contient la recherche
  If InStr(1, sPiece, sTri) > 0 Then
 'On met la variable à true pour l'ajouter au début
  bDebut = True
  End If
 
 'Si c'est trier par DESCR_FR
 Case I_CMB_DESCRIPTION_FR:
 'Si la description contient la recherche
  If InStr(1, sDescriptionFR, sTri) > 0 Then
 'On met la variable à true pour l'ajouter au début
  bDebut = True
  End If
 
 'Si c'est trier par DESCR_EN
 Case I_CMB_DESCRIPTION_EN:
 'Si la description contient la recherche
  If InStr(1, sDescriptionEN, sTri) > 0 Then
 'On met la variable à true pour l'ajouter au début
 bDebut = True
 End If
 
 'Si c'est la colonne Manufacturier
 Case I_CMB_FABRICANT:
 'Si le manufacturier contient la recherche
 If InStr(1, sFabricant, sTri) > 0 Then
 'On met la variable à true pour l'ajouter au début
 bDebut = True
 End If
 End Select
 
 If bDebut = True Then
 iIndex = iIndexActuel + 1
 Else
 iIndex = 0
 End If
Else
iIndex = 0
End If
 
 TrouverIndexPiece = iIndex

Exit Function

Oups:

 wOups "frmChoixDemande", "TrouverIndexPiece", Err, Err.number, Err.Description
End Function

Private Sub RemplirListViewPiece()

 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 Dim itmPiece As ListItem
 Dim sCategorie As String
 Dim sOrderBy As String
 Dim iIndex As Integer
 Dim iCompteur As Integer
 
 sCategorie = Replace(cmbCategorie.Text, "'", "''")
 
 Call lvwPiece.ListItems.Clear
 
 'Pour savoir par quoi trier le recordset
 Select Case cmbTri.ListIndex
 Case I_CMB_PIECE: sOrderBy = "PIECE"
 Case I_CMB_DESCRIPTION_FR: sOrderBy = "DESC_FR"
  Case I_CMB_DESCRIPTION_EN: sOrderBy = "DESC_EN"
  Case I_CMB_FABRICANT: sOrderBy = "FABRICANT"
  End Select
 
  Set rstPiece = New ADODB.Recordset
 
  If m_eCatalogue = ELECTRIQUE Then
  Call rstPiece.Open("SELECT * FROM GrbCatalogueElec WHERE CATEGORIE = '" & sCategorie & "' ORDER BY " & sOrderBy, g_connData, adOpenDynamic, adLockOptimistic)
  Else
  Call rstPiece.Open("SELECT * FROM GrbCatalogueMec WHERE CATEGORIE = '" & sCategorie & "' ORDER BY " & sOrderBy, g_connData, adOpenDynamic, adLockOptimistic)
10 End If
 
Do While Not rstPiece.EOF
 If Not IsNull(rstPiece.Fields("DESC_FR")) And Not IsNull(rstPiece.Fields("DESC_EN")) Then
 iIndex = TrouverIndexPiece(rstPiece.Fields("PIECE"), rstPiece.Fields("DESC_FR"), rstPiece.Fields("DESC_EN"), rstPiece.Fields("FABRICANT"), iIndex)
 Else
 If Not IsNull(rstPiece.Fields("DESC_FR")) Then
 iIndex = TrouverIndexPiece(rstPiece.Fields("PIECE"), rstPiece.Fields("DESC_FR"), vbNullString, rstPiece.Fields("FABRICANT"), iIndex)
 Else
 If Not IsNull(rstPiece.Fields("DESC_EN")) Then
 iIndex = TrouverIndexPiece(rstPiece.Fields("PIECE"), vbNullString, rstPiece.Fields("DESC_EN"), rstPiece.Fields("FABRICANT"), iIndex)
 Else
 iIndex = TrouverIndexPiece(rstPiece.Fields("PIECE"), vbNullString, vbNullString, rstPiece.Fields("FABRICANT"), iIndex)
 End If
 End If
 End If

 If iIndex = 0 Then
 Set itmPiece = lvwPiece.ListItems.Add
 Else
 Set itmPiece = lvwPiece.ListItems.Add(iIndex)
1  End If
 
 For iCompteur = 1 To m_collPiece.count
 If m_collCategorie(iCompteur) = cmbCategorie.Text Then
 If m_collPiece(iCompteur) = rstPiece.Fields("PIECE") Then
 If m_collDescriptionFR(iCompteur) = rstPiece.Fields("DESC_FR") Then
 If m_collDescriptionEN(iCompteur) = rstPiece.Fields("DESC_EN") Then
 If m_collManufacturier(iCompteur) = rstPiece.Fields("FABRICANT") Then
 itmPiece.Checked = True

 Exit For
 End If
 End If
 End If
 End If
 End If
 Next
 
itmPiece.SubItems(I_COL_PIECE) = rstPiece.Fields("PIECE")

 If Not IsNull(rstPiece.Fields("DESC_FR")) Then
 itmPiece.SubItems(I_COL_DESC_FR) = rstPiece.Fields("DESC_FR")
 Else
 itmPiece.SubItems(I_COL_DESC_FR) = vbNullString
 End If

If Not IsNull(rstPiece.Fields("DESC_EN")) Then
itmPiece.SubItems(I_COL_DESC_EN) = rstPiece.Fields("DESC_EN")
 Else
 itmPiece.SubItems(I_COL_DESC_EN) = vbNullString
 End If

 itmPiece.SubItems(I_COL_FABRICANT) = rstPiece.Fields("FABRICANT")
 
 Call rstPiece.MoveNext
Loop
 
Call rstPiece.Close
Set rstPiece = Nothing

Exit Sub

Oups:

wOups "frmChoixDemande", "RemplirListViewPiece", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewPieceProjetSoumission()

 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 Dim itmPiece As ListItem
 
 Call lvwPiece.ListItems.Clear
 
 lvwPiece.Sorted = False
 
 Set rstPiece = New ADODB.Recordset
 
 'Si c'est un projet
 If m_eType = TYPE_PROJET Then
 Call rstPiece.Open("SELECT Qté, NumItem, Desc_FR, Desc_EN, Manufact, IDFRS, PieceExtraChargeable, PieceExtraNonChargeable FROM GrbProjet_Pieces WHERE (IDProjet = '" & m_sNoProjSoum & "') AND (IDFRS = 0 AND NumItem <> 'Texte') ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 'Si c'est une soumission
 Call rstPiece.Open("SELECT Qté, NumItem, Desc_FR, Desc_En, Manufact, IDFRS, PieceExtraChargeable, PieceExtraNonChargeable FROM GrbSoumission_Pieces WHERE (IDSoumission = '" & m_sNoProjSoum & "') AND (IDFRS = 0 AND NumItem <> 'Texte') ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 
  Do While Not rstPiece.EOF
  If rstPiece.Fields("PieceExtraChargeable") = False And rstPiece.Fields("PieceExtraNonChargeable") = False Then
  Set itmPiece = lvwPiece.ListItems.Add
 
  itmPiece.Text = rstPiece.Fields("Qté")
 
  itmPiece.SubItems(I_COL_PIECE) = rstPiece.Fields("NumItem")
 
  If Not IsNull(rstPiece.Fields("Desc_FR")) Then
  itmPiece.SubItems(I_COL_DESC_FR) = rstPiece.Fields("Desc_FR")
  Else
 itmPiece.SubItems(I_COL_DESC_FR) = vbNullString
End If

 If Not IsNull(rstPiece.Fields("Desc_En")) Then
 itmPiece.SubItems(I_COL_DESC_EN) = rstPiece.Fields("Desc_En")
 Else
 itmPiece.SubItems(I_COL_DESC_EN) = vbNullString
 End If
 
 itmPiece.SubItems(I_COL_FABRICANT) = rstPiece.Fields("Manufact")
 End If
 
 Call rstPiece.MoveNext
Loop
 
Call rstPiece.Close
1  Set rstPiece = Nothing

Exit Sub

Oups:

 wOups "frmChoixDemande", "RemplirListViewPieceProjet", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewPieceAchat()

 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 Dim itmPiece As ListItem
 
 Call lvwPiece.ListItems.Clear
 
 Set rstPiece = New ADODB.Recordset
 
 Call rstPiece.Open("SELECT Qté, PIECE, Desc_FR, Desc_EN, Manufact, IDFRS FROM GrbAchat_Pieces WHERE IDAchat = '" & m_sNoAchat & "' AND IndexAchat = " & m_iIndexAchat & " AND IDFRS = 0 ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
 
 Do While Not rstPiece.EOF
 Set itmPiece = lvwPiece.ListItems.Add
 
 itmPiece.Text = rstPiece.Fields("Qté")
 
 itmPiece.SubItems(I_COL_PIECE) = rstPiece.Fields("PIECE")
 
 If Not IsNull(rstPiece.Fields("Desc_FR")) Then
  itmPiece.SubItems(I_COL_DESC_FR) = rstPiece.Fields("Desc_FR")
  Else
  itmPiece.SubItems(I_COL_DESC_FR) = vbNullString
  End If

  If Not IsNull(rstPiece.Fields("Desc_En")) Then
  itmPiece.SubItems(I_COL_DESC_EN) = rstPiece.Fields("Desc_En")
  Else
  itmPiece.SubItems(I_COL_DESC_EN) = vbNullString
End If
 
1 itmPiece.SubItems(I_COL_FABRICANT) = rstPiece.Fields("Manufact")
 
 Call rstPiece.MoveNext
Loop
 
Call rstPiece.Close
Set rstPiece = Nothing

Exit Sub

Oups:

wOups "frmChoixDemande", "RemplirListViewPieceProjet", Err, Err.number, Err.Description
End Sub

Private Sub cmbCategorie_Click()

 On Error GoTo Oups

 Call AjouterPieceCollection

 m_sCategorie = cmbCategorie.Text
 
 Call RemplirListViewPiece
 
 Call CocherCases

 Exit Sub

Oups:

 wOups "frmChoixDemande", "cmbCategorie_Click", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboManufacturiers()

 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 
 Call cmbManufacturier.Clear
 
 Set rstPiece = New ADODB.Recordset
 
 If m_eCatalogue = ELECTRIQUE Then
 Call rstPiece.Open("SELECT DISTINCT FABRICANT FROM GrbCatalogueElec", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstPiece.Open("SELECT DISTINCT FABRICANT FROM GrbCatalogueMec", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 
 Do While Not rstPiece.EOF
 If Not IsNull(rstPiece.Fields("FABRICANT")) Then
  Call cmbManufacturier.AddItem(rstPiece.Fields("FABRICANT"))
  End If
 
  Call rstPiece.MoveNext
  Loop
 
  Call rstPiece.Close
  Set rstPiece = Nothing

  Exit Sub

Oups:

  wOups "frmChoixDemande", "RemplirComboManufacturiers", Err, Err.number, Err.Description
End Sub

Private Sub CocherCases()

 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim iCompteur2 As Integer
 
 For iCompteur = 1 To m_collCategorie.count
 If m_collCategorie(iCompteur) = cmbCategorie.Text Then
 For iCompteur2 = 1 To lvwPiece.ListItems.count
 If lvwPiece.ListItems(iCompteur2).SubItems(I_COL_PIECE) = m_collPiece(iCompteur) Then
 lvwPiece.ListItems(iCompteur2).Checked = True
 End If
 Next iCompteur2
 End If
  Next iCompteur

  Exit Sub

Oups:

  wOups "frmChoixDemande", "CocherCases", Err, Err.number, Err.Description
End Sub

Private Sub cmbManufacturier_KeyPress(KeyAscii As Integer)

 On Error GoTo Oups

 If KeyAscii <= 122 And KeyAscii >=   Then
 KeyAscii = KeyAscii - 32
 End If

 Exit Sub

Oups:

 wOups "frmChoixDemande", "cmbManufacturier_KeyPress", Err, Err.number, Err.Description
End Sub

Private Sub Cmdajouter_Click()

 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 Dim iCompteur As Integer
 Dim bTrouver As Boolean
 Dim itmPiece As ListItem
 Dim sQuantite As String
 
 If txtNoPiece.Text = vbNullString Or cmbManufacturier.Text = vbNullString Or txtDescription.Text = vbNullString Then
 Call MsgBox("Vous devez absolument remplir tous les champs!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
 
 If InStr(1, txtNoPiece.Text, "'") > 0 Then
  Call MsgBox("Numéro invalide! Le numéro ne doit pas contenir d'appostrophes!", vbOKOnly, "Erreur")
 
  Exit Sub
  End If
 
  Set rstPiece = New ADODB.Recordset
 
  If m_eCatalogue = ELECTRIQUE Then
  Call rstPiece.Open("SELECT * FROM GrbCatalogueElec WHERE PIECE = '" & Replace(txtNoPiece.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
  Else
  Call rstPiece.Open("SELECT * FROM GrbCatalogueMec WHERE PIECE = '" & Replace(txtNoPiece.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
10 End If
 
If rstPiece.EOF = False Then
 bTrouver = True
End If
 
Call rstPiece.Close
Set rstPiece = Nothing
 
If bTrouver = True Then
 Call MsgBox("Le numéro de pièce existe déjà!", vbOKOnly, "Erreur")
 
 Exit Sub
End If
 
sQuantite = InputBox("Quelle est la quantité?")
 
1  If sQuantite <> vbNullString Then
 If Not IsNumeric(sQuantite) Then
 Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
Else
 sQuantite = "1"
1  End If
 
 Set itmPiece = lvwNouvellesPieces.ListItems.Add
 
 itmPiece.Text = sQuantite
itmPiece.SubItems(I_COL_NO_PIECE) = txtNoPiece.Text
itmPiece.SubItems(I_COL_DESCRIPTION) = txtDescription.Text
itmPiece.SubItems(I_COL_MANUFACT) = cmbManufacturier.Text
itmPiece.SubItems(I_COL_CATEGORIE) = cmbCategorie.Text

Exit Sub

Oups:

wOups "frmChoixDemande", "cmdAjouter_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixDemande", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdImprimer_Click()

 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim itmFRS As ListItem
 
 If lvwfournisseur.ListItems.count > 0 Then
 If VerifierSiChecked = True Then
 For iCompteur = 1 To lvwfournisseur.ListItems.count
 If lvwfournisseur.ListItems(iCompteur).Checked = True Then
 Set itmFRS = lvwfournisseur.ListItems(iCompteur)
 
 m_sLangage = itmFRS.SubItems(I_COL_LANGAGE)

 Call frmChoixContactFRS.Afficher(itmFRS.Tag)
 
 If m_bAnnulerContact = False Then
  If m_eDemande = MODE_NOUVELLE Then
  Call EnregistrerDemandePrixNouvellesPieces
  Else
  Call EnregistrerDemandePrix(itmFRS.Tag)
  End If
 
  Call ImprimerDemandePrix(itmFRS.Tag)
  End If
  End If
 Next
 
If m_eDemande = MODE_NOUVELLE Then
 Call EnregistrerPieces
 End If
 End If
End If

Exit Sub

Oups:

wOups "frmChoixDemande", "cmdImprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub EnregistrerPieces()

 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 Dim rstPiecesFRS As ADODB.Recordset
 Dim iCompteurPieces As Integer
 Dim iCompteurFRS As Integer
 
 Set rstPiece = New ADODB.Recordset
 Set rstPiecesFRS = New ADODB.Recordset

 For iCompteurPieces = 1 To m_collPiece.count
 If m_eCatalogue = ELECTRIQUE Then
 Call rstPiece.Open("SELECT * FROM GrbCatalogueElec", g_connData, adOpenDynamic, adLockOptimistic)
 Else
  Call rstPiece.Open("SELECT * FROM GrbCatalogueMec", g_connData, adOpenDynamic, adLockOptimistic)
  End If
 
  Call rstPiece.AddNew
 
  rstPiece.Fields("PIECE") = m_collPiece(iCompteurPieces)
  rstPiece.Fields("PIECE_GRB") = m_collPiece(iCompteurPieces) & "GRB"
  rstPiece.Fields("DESC_FR") = m_collDescriptionFR(iCompteurPieces)
  rstPiece.Fields("FABRICANT") = m_collManufacturier(iCompteurPieces)
  rstPiece.Fields("CATEGORIE") = m_collCategorie(iCompteurPieces)
 
Call rstPiece.Update
 
1 Call rstPiecesFRS.Open("SELECT * FROM GrbPiecesFRS", g_connData, adOpenDynamic, adLockOptimistic)
 
 For iCompteurFRS = 1 To lvwfournisseur.ListItems.count
 If lvwfournisseur.ListItems(iCompteurFRS).Checked = True Then
 Call rstPiecesFRS.AddNew
 
 rstPiecesFRS.Fields("IDFRS") = lvwfournisseur.ListItems(iCompteurFRS).Tag
 rstPiecesFRS.Fields("PIECE") = m_collPiece(iCompteurPieces)
 rstPiecesFRS.Fields("DATE") = ConvertDate(Date)
 rstPiecesFRS.Fields("ENTRER_PAR") = g_sInitiale
 rstPiecesFRS.Fields("PRIX_SP") = vbNullString
 rstPiecesFRS.Fields("PERS_RESS") = vbNullString
 rstPiecesFRS.Fields("PRIX_LIST") = "0"
 rstPiecesFRS.Fields("ESCOMPTE") = "0"
 rstPiecesFRS.Fields("PRIX_NET") = "0"
 rstPiecesFRS.Fields("DeviseMonétaire") = "CAN"
 rstPiecesFRS.Fields("PrixReel") = "0"
 rstPiecesFRS.Fields("Type") = "M"
 
 Call rstPiecesFRS.Update
 End If
1  Next
 
 Call rstPiecesFRS.Close
 
 Call rstPiece.Close
Next

Set rstPiece = Nothing
Set rstPiecesFRS = Nothing
 
Exit Sub

Oups:

wOups "frmChoixDemande", "EnregistrerPieces", Err, Err.number, Err.Description
End Sub

Private Sub cmdLangage_Click()

 On Error GoTo Oups

 If cmdLangage.Caption = S_DEMANDE_FRANCAIS Then
 lvwfournisseur.SelectedItem.SubItems(I_COL_LANGAGE) = S_FRANCAIS
 
 cmdLangage.Caption = S_DEMANDE_ANGLAIS
 Else
 lvwfournisseur.SelectedItem.SubItems(I_COL_LANGAGE) = S_ANGLAIS
 
 cmdLangage.Caption = S_DEMANDE_FRANCAIS
 End If
 
 Call lvwfournisseur.SetFocus

 Exit Sub

Oups:

 wOups "frmChoixDemande", "cmdLangage_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdOK_Click()

 On Error GoTo Oups

 If m_eMode = PIECE Then
 Call AjouterPieceCollection
 
 Call AfficherFournisseur
 Else
 If m_eMode = NOUVELLE_PIECE Then
 Call AjouterNouvellePieceCollection
 
 Call RemplirListViewFournisseur(False)
 
 Call AfficherControles(Fournisseur)
 Else
 If m_eMode = Manufacturier Then
  Call AjouterManufacturierCollection

  Call RemplirListViewFournisseur(False)

  Call AfficherControles(Fournisseur)
  Else
  Call AjouterCategorieCollection
 
  If VerifierSiChecked = True Then
  Call RemplirListViewManufacturier
 
  Call AfficherControles(Manufacturier)
 End If
End If
 End If
End If

Exit Sub

Oups:

wOups "frmChoixDemande", "cmdOK_Click", Err, Err.number, Err.Description
End Sub

Private Sub AjouterNouvellePieceCollection()

 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim itmPiece As ListItem
 
 For iCompteur = 1 To lvwNouvellesPieces.ListItems.count
 Set itmPiece = lvwNouvellesPieces.ListItems(iCompteur)
 
 Call m_collQuantite.Add(itmPiece.Text)
 Call m_collPiece.Add(itmPiece.SubItems(I_COL_NO_PIECE))
 Call m_collDescriptionFR.Add(itmPiece.SubItems(I_COL_DESCRIPTION))
 Call m_collManufacturier.Add(itmPiece.SubItems(I_COL_MANUFACT))
 Call m_collCategorie.Add(itmPiece.SubItems(I_COL_CATEGORIE))
 Next

  Exit Sub

Oups:

  wOups "frmChoixDemande", "AjouterNouvellePieceCollection", Err, Err.number, Err.Description
End Sub

Private Sub AjouterManufacturierCollection()

 On Error GoTo Oups

 Dim iCompteur As Integer

 For iCompteur = 1 To lvwManufacturier.ListItems.count
 If lvwManufacturier.ListItems(iCompteur).Checked = True Then
 Call m_collManufacturier.Add(lvwManufacturier.ListItems(iCompteur).Text)
 End If
 Next

 Exit Sub

Oups:

 wOups "frmChoixDemande", "AjouterManufacturierCollection", Err, Err.number, Err.Description
End Sub

Private Sub AjouterCategorieCollection()

 On Error GoTo Oups
 
 Dim iCompteur As Integer
 
 For iCompteur = 1 To lvwCategorie.ListItems.count
 If lvwCategorie.ListItems(iCompteur).Checked = True Then
 Call m_collCategorie.Add(lvwCategorie.ListItems(iCompteur).Text)
 End If
 Next

 Exit Sub

Oups:

 wOups "frmChoixDemande", "AjouterCategorieCollection", Err, Err.number, Err.Description
End Sub

Private Sub cmdRechercher_Click()

 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim bTrouver As Boolean
 
 'Si le texte du bouton est rechercher
 If cmdRechercher.Caption = "Rechercher" Then
 'Pour chaque élément du listview
 For iCompteur = 1 To lvwfournisseur.ListItems.count
 'si le nom du fournisseur contient le texte à rechercher
 If InStr(1, UCase(lvwfournisseur.ListItems(iCompteur).Text), UCase(txtRechercher.Text)) > 0 Then
 bTrouver = True
 
 lvwfournisseur.ListItems(iCompteur).Selected = True
 
 Call lvwfournisseur.ListItems(iCompteur).EnsureVisible
 
 Call lvwfournisseur.SetFocus
 
 m_iIndexRecherche = iCompteur
 
  Exit For
  End If
  Next
 
  If bTrouver = True Then
  cmdRechercher.Caption = "Suivant"
  Else
  Call MsgBox("Aucun fournisseur trouvé!", vbOKOnly)
  End If
10 Else
 'Pour chaque élément restant du listview
1 For iCompteur = m_iIndexRecherche + 1 To lvwfournisseur.ListItems.count
 'Si le nom du fournisseur contient le texte à rechercher
 If InStr(1, UCase(lvwfournisseur.ListItems(iCompteur).Text), UCase(txtRechercher.Text)) > 0 Then
 bTrouver = True
 
 lvwfournisseur.ListItems(iCompteur).Selected = True
 
 Call lvwfournisseur.ListItems(iCompteur).EnsureVisible
 
 Call lvwfournisseur.SetFocus
 
 m_iIndexRecherche = iCompteur
 
 Exit For
 End If
 Next
 
 If bTrouver = False Then
 Call MsgBox("Aucun fournisseur trouvé!", vbOKOnly)
 
 cmdRechercher.Caption = "Rechercher"
 End If
End If

 Exit Sub

Oups:

wOups "frmChoixDemande", "cmdRechercher_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdSelectAll_Click()

 On Error GoTo Oups

 Dim lvwSource As ListView
 Dim iCompteur As Integer

 Select Case m_eMode
 Case PIECE: Set lvwSource = lvwPiece
 Case Categorie: Set lvwSource = lvwCategorie
 Case Fournisseur: Set lvwSource = lvwfournisseur
 Case Manufacturier: Set lvwSource = lvwManufacturier
 End Select

 For iCompteur = 1 To lvwSource.ListItems.count
 lvwSource.ListItems(iCompteur).Checked = True
 Next

 Exit Sub

Oups:

 wOups "frmChoixDemande", "cmdSelectAll_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDeSelectAll_Click()

 On Error GoTo Oups

 Dim lvwSource As ListView
 Dim iCompteur As Integer

 Select Case m_eMode
 Case PIECE: Set lvwSource = lvwPiece
 Case Categorie: Set lvwSource = lvwCategorie
 Case Fournisseur: Set lvwSource = lvwfournisseur
 Case Manufacturier: Set lvwSource = lvwManufacturier
 End Select

 For iCompteur = 1 To lvwSource.ListItems.count
 lvwSource.ListItems(iCompteur).Checked = False
 Next

 Exit Sub

Oups:

 wOups "frmChoixDemande", "cmdSelectAll_Click", Err, Err.number, Err.Description
End Sub


Private Sub cmdTri_Click()

 On Error GoTo Oups

 m_sTri = InputBox("Quel est la pièce à trier?")
 
 If m_sTri <> vbNullString Then
 lvwCategorie.Sorted = False
 lvwPiece.Sorted = False
 lvwfournisseur.Sorted = False
 lvwNouvellesPieces.Sorted = False

 Call AjouterPieceCollection

 Call RemplirListViewPiece
 End If

 Exit Sub

Oups:

  wOups "frmChoixDemande", "cmdTri_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRafraichir_Click()

 On Error GoTo Oups

 If m_sTri <> vbNullString Then
 m_sTri = vbNullString
 
 Call RemplirListViewPiece
 End If

 Exit Sub

Oups:

 wOups "frmChoixDemande", "cmdRafraichir_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Set m_collPiece = New Collection
 Set m_collQuantite = New Collection
 Set m_collDescriptionFR = New Collection
 Set m_collDescriptionEN = New Collection
 Set m_collCategorie = New Collection
 Set m_collManufacturier = New Collection
 
 cmbTri.ListIndex = I_CMB_PIECE

 Exit Sub

Oups:

 wOups "frmChoixDemande", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
 
 On Error GoTo Oups
 
 m_sTri = vbNullString

 Exit Sub

Oups:

 wOups "frmChoixDemande", "Form_Unload", Err, Err.number, Err.Description
End Sub

Private Sub lvwCategorie_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

 On Error GoTo Oups

 lvwCategorie.Sorted = True
 
 If lvwCategorie.SortOrder = lvwAscending Then
 lvwCategorie.SortOrder = lvwDescending
 Else
 lvwCategorie.SortOrder = lvwAscending
 End If
 
 lvwCategorie.SortKey = ColumnHeader.Index - 1

 Exit Sub

Oups:

 wOups "frmChoixDemande", "lvwCategorie_ColumnClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwFournisseur_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

 On Error GoTo Oups

 lvwfournisseur.Sorted = True
 
 If lvwfournisseur.SortOrder = lvwAscending Then
 lvwfournisseur.SortOrder = lvwDescending
 Else
 lvwfournisseur.SortOrder = lvwAscending
 End If
 
 lvwfournisseur.SortKey = ColumnHeader.Index - 1

 Exit Sub

Oups:

 wOups "frmChoixDemande", "lvwFournisseur_ColumnClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwFournisseur_ItemClick(ByVal Item As MSComctlLib.ListItem)

 On Error GoTo Oups

 If Item.SubItems(I_COL_LANGAGE) = S_FRANCAIS Then
 cmdLangage.Caption = S_DEMANDE_ANGLAIS
 Else
 cmdLangage.Caption = S_DEMANDE_FRANCAIS
 End If

 Exit Sub

Oups:

 wOups "frmChoixDemande", "lvwFournisseur_ItemClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwNouvellesPieces_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

 On Error GoTo Oups

 lvwNouvellesPieces.Sorted = True
 
 If lvwNouvellesPieces.SortOrder = lvwAscending Then
 lvwNouvellesPieces.SortOrder = lvwDescending
 Else
 lvwNouvellesPieces.SortOrder = lvwAscending
 End If
 
 lvwNouvellesPieces.SortKey = ColumnHeader.Index - 1

 Exit Sub

Oups:

 wOups "frmChoixDemande", "lvwNouvellesPieces_ColumnClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwNouvellesPieces_KeyDown(KeyCode As Integer, Shift As Integer)

 On Error GoTo Oups

 If KeyCode = vbKeyDelete Then
 Call lvwNouvellesPieces.ListItems.Remove(lvwNouvellesPieces.SelectedItem.Index)
 End If

 Exit Sub

Oups:

 wOups "frmChoixDemande", "lvwNouvellesPieces_KeyDown", Err, Err.number, Err.Description
End Sub

Private Sub lvwPiece_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

 On Error GoTo Oups

 lvwPiece.Sorted = True
 
 If lvwPiece.SortOrder = lvwAscending Then
 lvwPiece.SortOrder = lvwDescending
 Else
 lvwPiece.SortOrder = lvwAscending
 End If
 
 lvwPiece.SortKey = ColumnHeader.Index - 1

 Exit Sub

Oups:

 wOups "frmChoixDemande", "lvwPiece_ColumnClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwPiece_DblClick()

 On Error GoTo Oups

 Dim sQuantite As String
 
 If lvwPiece.ListItems.count > 0 Then
 sQuantite = InputBox("Quelle est la quantité?")
 
 If sQuantite <> vbNullString Then
 If IsNumeric(sQuantite) Then
 lvwPiece.SelectedItem.Text = sQuantite
 End If
 End If
 End If

 Exit Sub

Oups:

  wOups "frmChoixDemande", "lvwPiece_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub AfficherFournisseur()

 On Error GoTo Oups

 If lvwPiece.ListItems.count > 0 Then
 If VerifierSiChecked = True Then
 
 Call RemplirListViewFournisseur(True)
 
 If lvwfournisseur.ListItems.count > 0 Then
 Call AfficherControles(Fournisseur)
 Else
 Call MsgBox("Il n'y a aucun fournisseur pour cette pièce!", vbOKOnly, "Erreur")
 End If
 End If
 End If

  Exit Sub

Oups:

  wOups "frmChoixDemande", "AfficherFournisseur", Err, Err.number, Err.Description
End Sub

Private Sub AjouterPieceCollection()

 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim iCompteurAs Integer
 Dim bPieceExiste As Boolean
 Dim iQuantite As Integer
 Dim rstTempDP As ADODB.Recordset

 If m_eCatalogue = ELECTRIQUE Then
 Call g_connData.Execute("DELETE * FROM GrbTempDP WHERE TYPE = 'E'")
 Else
 Call g_connData.Execute("DELETE * FROM GrbTempDP WHERE TYPE = 'M'")
 End If

  Set rstTempDP = New ADODB.Recordset

  Call rstTempDP.Open("SELECT * FROM GrbTempDP", g_connData, adOpenDynamic, adLockOptimistic)
 
  For iCompteur = 1 To lvwPiece.ListItems.count
  If lvwPiece.ListItems(iCompteur).Checked = True Then
  bPieceExiste = False
 
  For iCompteur2 = 1 To m_collPiece.count
  If lvwPiece.ListItems(iCompteur).SubItems(I_COL_PIECE) = m_collPiece(iCompteur2) Then
  bPieceExiste = True
 
 Exit For
 End If
 Next iCompteur2
 
 If bPieceExiste = False Then
 Call m_collCategorie.Add(m_sCategorie)
 Call m_collPiece.Add(lvwPiece.ListItems(iCompteur).SubItems(I_COL_PIECE))
 Call m_collDescriptionFR.Add(lvwPiece.ListItems(iCompteur).SubItems(I_COL_DESC_FR))
 Call m_collDescriptionEN.Add(lvwPiece.ListItems(iCompteur).SubItems(I_COL_DESC_EN))
 Call m_collManufacturier.Add(lvwPiece.ListItems(iCompteur).SubItems(I_COL_FABRICANT))
 
 If lvwPiece.ListItems(iCompteur).Text <> vbNullString Then
 Call m_collQuantite.Add(lvwPiece.ListItems(iCompteur).Text)
 Else
 Call m_collQuantite.Add("1")
 End If

 Call rstTempDP.AddNew

 rstTempDP.Fields("PIECE") = lvwPiece.ListItems(iCompteur).SubItems(I_COL_PIECE)
 rstTempDP.Fields("ORDRE") = iCompteur

 If m_eCatalogue = ELECTRIQUE Then
 rstTempDP.Fields("TYPE") = "E"
1  Else
 rstTempDP.Fields("TYPE") = "M"
 End If

 Call rstTempDP.Update
 Else
 'Ajoute la quantité si c'est une demande de prix à partir d'un projet
 If m_bProjSoum = True Then
 iQuantite = Val(m_collQuantite(iCompteur2)) + Val(lvwPiece.ListItems(iCompteur).Text)
 
 Call m_collQuantite.Remove(iCompteur2)
 
 If m_collQuantite.count > 0 Then
 If m_collQuantite.count < iCompteur2 Then
 Call m_collQuantite.Add(iQuantite)
 Else
 If iCompteur2 > 1 Then
 Call m_collQuantite.Add(iQuantite, , , iCompteur2 - 1)
 Else
 Call m_collQuantite.Add(iQuantite, , , 1)
 End If
 End If
 Else
 Call m_collQuantite.Add(iQuantite)
 End If
 End If
End If
 End If
Next iCompteur

Call rstTempDP.Close
Set rstTempDP = Nothing

Exit Sub

Oups:

wOups "frmChoixDemande", "AjouterPieceCollection", Err, Err.number, Err.Description
End Sub

Private Function VerifierSiChecked() As Boolean

 On Error GoTo Oups

 Dim lvwSource As ListView
 Dim iCompteur As Integer
 
 If lvwPiece.Visible = True Then
 Set lvwSource = lvwPiece
 Else
 If lvwfournisseur.Visible = True Then
 Set lvwSource = lvwfournisseur
 Else
 Set lvwSource = lvwCategorie
 End If
  End If
 
  VerifierSiChecked = False
 
  For iCompteur = 1 To lvwSource.ListItems.count
  If lvwSource.ListItems(iCompteur).Checked = True Then
  VerifierSiChecked = True
 
  Exit For
  End If
  Next

10 Exit Function

Oups:

wOups "frmChoixDemande", "VerifierSiChecked", Err, Err.number, Err.Description
End Function

Private Sub EnregistrerDemandePrixNouvellesPieces()

 On Error GoTo Oups

 Dim rstImpDP As ADODB.Recordset
 Dim sNomTable As String
 Dim iCompteur As Integer
 
 If m_eCatalogue = ELECTRIQUE Then
 sNomTable = "GrbImpressionDemandePrixElec"
 Else
 sNomTable = "GrbImpressionDemandePrixMec"
 End If
 
 Set rstImpDP = New ADODB.Recordset
 
 rstImpDP.CursorLocation = adUseClient
 
  Call rstImpDP.Open("SELECT * FROM " & sNomTable, g_connData, adOpenDynamic, adLockOptimistic)
 
  For iCompteur = 1 To m_collPiece.count
  Call rstImpDP.AddNew
 
  rstImpDP.Fields("NoPiece") = m_collPiece(iCompteur)
  rstImpDP.Fields("Qte") = m_collQuantite(iCompteur)
  rstImpDP.Fields("Description") = m_collDescriptionFR(iCompteur)
  rstImpDP.Fields("Manufacturier") = m_collManufacturier(iCompteur)
 
  Call rstImpDP.Update
10 Next
 
Call rstImpDP.Requery
 
For iCompteur = 15 To rstImpDP.RecordCount Step -1
 Call rstImpDP.AddNew
 
 rstImpDP.Fields("NoPiece") = vbNullString
 rstImpDP.Fields("Qte") = vbNullString
 rstImpDP.Fields("Description") = vbNullString
 rstImpDP.Fields("Manufacturier") = vbNullString
 
 Call rstImpDP.Update
Next
 
Call rstImpDP.Close
Set rstImpDP = Nothing

1  Exit Sub

Oups:

wOups "frmChoixDemande", "EnregistrerDemandePrixNouvellesPieces", Err, Err.number, Err.Description
End Sub

Private Sub EnregistrerDemandePrix(ByVal iIDFRS As Integer)

 On Error GoTo Oups

 Dim rstImpDP As ADODB.Recordset
 Dim rstPieceFRS As ADODB.Recordset
 Dim rstPiece As ADODB.Recordset
 Dim sNomTable As String
 Dim sWhere As String
 Dim sCategorie As String
 Dim iCompteur As Integer
 
 If m_eCatalogue = ELECTRIQUE Then
 sNomTable = "GrbImpressionDemandePrixElec"
 Else
  sNomTable = "GrbImpressionDemandePrixMec"
  End If

  Set rstImpDP = New ADODB.Recordset
  Set rstPiece = New ADODB.Recordset
  Set rstPieceFRS = New ADODB.Recordset

  rstImpDP.CursorLocation = adUseClient

  Call rstImpDP.Open("SELECT * FROM " & sNomTable, g_connData, adOpenDynamic, adLockOptimistic)
 
  If m_eDemande <> MODE_PIECE Then
Select Case m_eDemande
 Case MODE_FOURNISSEUR:
 Call rstPieceFRS.Open("SELECT * FROM GrbPiecesFRS WHERE IDFRS = " & iIDFRS & " ORDER BY PIECE", g_connData, adOpenDynamic, adLockOptimistic)
 
 Case MODE_CATEGORIE:
 If m_eCatalogue = ELECTRIQUE Then
 sWhere = " AND GrbCatalogueElec.CATEGORIE In ("
 Else
 sWhere = " AND GrbCatalogueMec.CATEGORIE In ("
 End If

 For iCompteur = 1 To m_collCategorie.count
 sCategorie = Replace(m_collCategorie(iCompteur), "'", "''")
 
 If iCompteur <> m_collCategorie.count Then
 sWhere = sWhere & "'" & sCategorie & "',"
 Else
 sWhere = sWhere & "'" & sCategorie & "')"
 End If
 Next
 
 Call rstPieceFRS.Open("SELECT GrbPiecesFRS.* FROM GrbPiecesFRS INNER JOIN GrbCatalogueMec ON GrbPiecesFRS.PIECE = GrbCatalogueMec.PIECE WHERE GrbPiecesFRS.IDFRS = " & iIDFRS & sWhere & " ORDER BY GrbPiecesFRS.PIECE", g_connData, adOpenDynamic, adLockOptimistic)
 End Select
 
 Do While Not rstPieceFRS.EOF
1  If rstPieceFRS.Fields("Type") = "M" Then
 Call rstPiece.Open("SELECT FABRICANT, DESC_FR, DESC_EN FROM GrbCatalogueElec WHERE PIECE = '" & Replace(rstPieceFRS.Fields("PIECE"), "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstPiece.Open("SELECT FABRICANT, DESC_FR, DESC_EN FROM GrbCatalogueMec WHERE PIECE = '" & Replace(rstPieceFRS.Fields("PIECE"), "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If

 If Not rstPiece.EOF Then
 If m_collManufacturier.count > 0 Then
 For iCompteur = 1 To m_collManufacturier.count
 If m_collManufacturier(iCompteur) = rstPiece.Fields("FABRICANT") Then
 Call rstImpDP.AddNew
 
 rstImpDP.Fields("NoPiece") = rstPieceFRS.Fields("PIECE")
 rstImpDP.Fields("Qte") = "1"
255
 If m_sLangage = S_FRANCAIS Then
 rstImpDP.Fields("Description") = rstPiece.Fields("DESC_FR")
 Else
 rstImpDP.Fields("Description") = rstPiece.Fields("DESC_EN")
 End If
 
 rstImpDP.Fields("Manufacturier") = rstPiece.Fields("FABRICANT")
 
 Call rstImpDP.Update
 End If
 Next
 Else
 Call rstImpDP.AddNew

 rstImpDP.Fields("NoPiece") = rstPieceFRS.Fields("PIECE")
 rstImpDP.Fields("Qte") = "1"

 If m_sLangage = S_FRANCAIS Then
 rstImpDP.Fields("Description") = rstPiece.Fields("DESC_FR")
 Else
 rstImpDP.Fields("Description") = rstPiece.Fields("DESC_EN")
 End If

 rstImpDP.Fields("Manufacturier") = rstPiece.Fields("FABRICANT")

 Call rstImpDP.Update
 End If
 Else
 Call rstPieceFRS.Delete
 End If

 Call rstPiece.Close
 
 Call rstPieceFRS.MoveNext
 Loop

 Set rstPiece = Nothing
 
Call rstPieceFRS.Close
4 Set rstPieceFRS = Nothing
4 Else
4 If m_eCatalogue = ELECTRIQUE Then
4 Call rstPieceFRS.Open("SELECT GrbPiecesFRS.PIECE FROM GrbPiecesFRS INNER JOIN GrbTempDP ON GrbPiecesFRS.PIECE = GrbTempDP.PIECE WHERE GrbPiecesFRS.IDFRS = " & iIDFRS & " AND GrbTempDP.TYPE = 'E' GROUP BY GrbPiecesFRS.PIECE, ORDRE ORDER BY GrbTempDP.ORDRE", g_connData, adOpenDynamic, adLockOptimistic)
4 Else
4 Call rstPieceFRS.Open("SELECT GrbPiecesFRS.PIECE FROM GrbPiecesFRS INNER JOIN GrbTempDP ON GrbPiecesFRS.PIECE = GrbTempDP.PIECE WHERE GrbPiecesFRS.IDFRS = " & iIDFRS & " AND GrbTempDP.TYPE = 'M' GROUP BY GrbPiecesFRS.PIECE, ORDRE ORDER BY GrbTempDP.ORDRE", g_connData, adOpenDynamic, adLockOptimistic)
4 End If

4 Do While Not rstPieceFRS.EOF
4 For iCompteur = 1 To m_collPiece.count
4 If m_collPiece(iCompteur) = Trim(rstPieceFRS.Fields("PIECE")) Then
4 Call rstImpDP.AddNew

4  rstImpDP.Fields("NoPiece") = m_collPiece(iCompteur)

4  rstImpDP.Fields("Qte") = m_collQuantite(iCompteur)

4  If m_sLangage = S_FRANCAIS Then
4  rstImpDP.Fields("Description") = m_collDescriptionFR(iCompteur)
4  Else
4  rstImpDP.Fields("Description") = m_collDescriptionEN(iCompteur)
4  End If
 
4  rstImpDP.Fields("Manufacturier") = m_collManufacturier(iCompteur)

50 Call rstImpDP.Update

 Exit For
 End If
 Next

 Call rstPieceFRS.MoveNext
 Loop

 Call rstPieceFRS.Close
 Set rstPieceFRS = Nothing
 End If
 
 Call rstImpDP.Requery

 For iCompteur = 15 To (rstImpDP.RecordCount + 1) Step -1
 Call rstImpDP.AddNew

5  rstImpDP.Fields("NoPiece") = vbNullString
5  rstImpDP.Fields("Qte") = vbNullString
5  rstImpDP.Fields("Description") = vbNullString
5  rstImpDP.Fields("Manufacturier") = vbNullString
 
5  Call rstImpDP.Update
5  Next
 
5  Call rstImpDP.Close
5  Set rstImpDP = Nothing

60 Exit Sub

Oups:

60 wOups "frmChoixDemande", "EnregistrerDemandePrix", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerDemandePrix(ByVal iIDFRS As Integer)

 On Error GoTo Oups

 Dim rstImpDP As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim sNomTable As String
 
 If m_eCatalogue = ELECTRIQUE Then
 sNomTable = "GrbImpressionDemandePrixElec"
 Else
 sNomTable = "GrbImpressionDemandePrixMec"
 End If
 
 Set rstFRS = New ADODB.Recordset
 Set rstImpDP = New ADODB.Recordset
 
  Call rstFRS.Open("SELECT * FROM GrbFournisseur WHERE IDFRS = " & iIDFRS, g_connData, adOpenDynamic, adLockOptimistic)
 
  Call rstImpDP.Open("SELECT * FROM " & sNomTable, g_connData, adOpenDynamic, adLockOptimistic)
 
  Set DR_DemandePrix.DataSource = rstImpDP
 
 'Entête
 'Vérifie si c'est Anglais ou Francais
 'On modifie seulement si c'est Anglais
  If m_sLangage = S_ANGLAIS Then
  DR_DemandePrix.Sections("Section2").Controls("lblTitreDemande").Caption = "Price and Delivery Request"
  DR_DemandePrix.Sections("Section2").Controls("lblTitreFournisseur").Caption = "Supplier :"
  DR_DemandePrix.Sections("Section2").Controls("lblTitreContact").Caption = "Contact :"
  DR_DemandePrix.Sections("Section2").Controls("lblTitreTransport").Caption = "Transport :"
DR_DemandePrix.Sections("Section2").Controls("lblTitreDateReq").Caption = "Required Date :"
1 DR_DemandePrix.Sections("Section2").Controls("lblTitreNoSoum").Caption = "Your Ref # :"
 DR_DemandePrix.Sections("Section2").Controls("lblTitreTel").Caption = "Telephone :"
 DR_DemandePrix.Sections("Section2").Controls("lblTitreFax").Caption = "Fax :"
 DR_DemandePrix.Sections("Section2").Controls("lblTitreNoGRB").Caption = "OUR # :"
 DR_DemandePrix.Sections("Section2").Controls("lblTitreDate").Caption = "Date :"
 DR_DemandePrix.Sections("Section2").Controls("lblTitreComPar").Caption = "Purchaser :"
 DR_DemandePrix.Sections("Section2").Controls("lblTitrePage").Caption = "Page :"
 DR_DemandePrix.Sections("Section2").Controls("lblTitreQte").Caption = "Qty"
 DR_DemandePrix.Sections("Section2").Controls("lblTitrePiece").Caption = "Part Number"
 DR_DemandePrix.Sections("Section2").Controls("lblTitreDescription").Caption = "Description"
 DR_DemandePrix.Sections("Section2").Controls("lblTitreManufact").Caption = "Manufacturer"
DR_DemandePrix.Sections("Section2").Controls("lblTitrePrix").Caption = "Unit Price"
 DR_DemandePrix.Sections("Section2").Controls("lblTitreDelais").Caption = "Delay"
 DR_DemandePrix.Sections("Section3").Controls("lblTitreCommentaire").Caption = "Comments :"
 DR_DemandePrix.Sections("Section3").Controls("lblPrixValide").Caption = "Valid price for"
 DR_DemandePrix.Sections("Section3").Controls("lblJours").Caption = "Days"
 DR_DemandePrix.Sections("Section3").Controls("lblPiedPage").Caption = "THIS IS NOT AN ORDER"

 DR_DemandePrix.Sections("Section2").Controls("imgLogoFrancais").Visible = False
1  DR_DemandePrix.Sections("Section2").Controls("imgLogoAnglais").Visible = True
 End If
 
 DR_DemandePrix.Sections("Section2").Controls("lblFournisseur").Caption = rstFRS.Fields("NomFournisseur")
DR_DemandePrix.Sections("Section2").Controls("lblContact").Caption = m_sContact

If Not IsNull(rstFRS.Fields("CondTransport")) Then
 DR_DemandePrix.Sections("Section2").Controls("lblTransport").Caption = rstFRS.Fields("CondTransport")
Else
 DR_DemandePrix.Sections("Section2").Controls("lblTransport").Caption = ""
End If

DR_DemandePrix.Sections("Section2").Controls("lblDateRequise").Caption = mskDateRequise.Text
DR_DemandePrix.Sections("Section2").Controls("lblTel").Caption = rstFRS.Fields("Telephonne")
DR_DemandePrix.Sections("Section2").Controls("lblFax").Caption = rstFRS.Fields("Fax")
DR_DemandePrix.Sections("Section2").Controls("lblNoGRB").Caption = txtNoGRB.Text
2  DR_DemandePrix.Sections("Section2").Controls("lblDate").Caption = ConvertDate(Date)
DR_DemandePrix.Sections("Section2").Controls("lblCommandePar").Caption = g_sEmploye
2  DR_DemandePrix.Sections("Section3").Controls("lblCommentaire").Caption = txtcommentaire.Text

DR_DemandePrix.Orientation = rptOrientLandscape
 
2  Call DR_DemandePrix.Show(vbModal)
 
Call g_connData.Execute("DELETE * FROM " & sNomTable)
 
2  Call rstImpDP.Close
Set rstImpDP = Nothing
 
30 Call rstFRS.Close
Set rstFRS = Nothing

Exit Sub

Oups:

wOups "frmChoixDemande", "ImprimerDemandePrix", Err, Err.number, Err.Description
End Sub

Private Sub mskDateRequise_GotFocus()

 On Error GoTo Oups

 If Len(mskDateRequise.Text) = 10 Then
 mskDateRequise.Text = Right$(mskDateRequise.Text, 8)
 End If
 
 mskDateRequise.mask = "##-##-##"

 Exit Sub

Oups:

 wOups "frmChoixDemande", "mskDateRequise_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDateRequise_LostFocus()

 On Error GoTo Oups

 mskDateRequise.mask = vbNullString
 
 If mskDateRequise.Text = "__-__-__" Then
 mskDateRequise.Text = vbNullString
 Else
 If Len(mskDateRequise.Text) =   Then
 If IsDate(mskDateRequise.Text) Then
 mskDateRequise.Text = Year(DateSerial(Left$(mskDateRequise.Text, 2), Mid$(mskDateRequise.Text, 4, 2), Right$(mskDateRequise.Text, 2))) & Mid$(mskDateRequise.Text, 3, 8)
 End If
 End If
 End If

  Exit Sub

Oups:

  wOups "frmChoixDemande", "mskDateRequise_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtRechercher_Change()

 On Error GoTo Oups

 cmdRechercher.Caption = "Rechercher"

 Exit Sub

Oups:

 wOups "frmChoixDemande", "txtRechercher_Change", Err, Err.number, Err.Description
End Sub
