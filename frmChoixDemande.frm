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
   Picture         =   "frmChoixDemande.frx":0000
   ScaleHeight     =   6480
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
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
      ItemData        =   "frmChoixDemande.frx":2F0D
      Left            =   4440
      List            =   "frmChoixDemande.frx":2F1D
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
Private Const I_CMB_PIECE          As Integer = 0
Private Const I_CMB_DESCRIPTION_FR As Integer = 1
Private Const I_CMB_DESCRIPTION_EN As Integer = 2
Private Const I_CMB_FABRICANT      As Integer = 3

'Index des colonnes de lvwPiece
Private Const I_COL_QUANTITE       As Integer = 0
Private Const I_COL_PIECE          As Integer = 1
Private Const I_COL_DESC_FR        As Integer = 2
Private Const I_COL_DESC_EN        As Integer = 3
Private Const I_COL_FABRICANT      As Integer = 4
                                
'Index des colonnes de lvwFournisseur
Private Const I_COL_NOM_FRS        As Integer = 0
Private Const I_COL_LANGAGE        As Integer = 1
                                
'Index des colonnes de lvwNouvellesPieces
Private Const I_COL_QTE            As Integer = 0
Private Const I_COL_NO_PIECE       As Integer = 1
Private Const I_COL_DESCRIPTION    As Integer = 2
Private Const I_COL_MANUFACT       As Integer = 3
Private Const I_COL_CATEGORIE      As Integer = 4
                                
'Caption du bouton cmdLangage
Private Const S_DEMANDE_FRANCAIS   As String = "En français"
Private Const S_DEMANDE_ANGLAIS    As String = "En anglais"
                                
'Texte de la colonne Langage de la demande de lvwFournisseur
Private Const S_FRANCAIS           As String = "Français"
Private Const S_ANGLAIS            As String = "Anglais"

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
  MODE_PIECE = 0        'Pour une pièce
  MODE_FOURNISSEUR = 1  'Pour toutes les pièces d'un fournisseur
  MODE_CATEGORIE = 2    'Pour catégorie
  MODE_NOUVELLE = 3
End Enum

Private m_eMode             As enumMode

Private m_eDemande          As enumModeDemande

'Contient la valeur électrique ou mécanique
Private m_eCatalogue        As enumCatalogue

'Pour conserver en mémoire les pièces choisies
Private m_collPiece         As Collection
Private m_collQuantite      As Collection
Private m_collDescriptionFR As Collection
Private m_collDescriptionEN As Collection
Private m_collCategorie     As Collection
Private m_collManufacturier As Collection

'Pour savoir si les fournisseurs sont affichés après avoir choisit des pièces
Private m_bPiece            As Boolean

'Pour savoir si la catégorie a changé ou non,
'sert également pour mettre dans m_collCategorie
Private m_sCategorie        As String

Private m_sLangage          As String

'Pour savoir si la demande a été fait à partir des Projets / Soumissions
Private m_bProjSoum         As Boolean

'Pour savoir si la demande a été fait à partir des achats
Private m_bAchat            As Boolean

'Contient le numéro du projet si la demande à été fait à partir d'un Projet
Private m_sNoProjSoum       As String

'Contient le numéro de l'achat si la demande a été fait à partir des achats
Private m_sNoAchat          As String

'Contient l'index de l'achat si la demande a été fait à partir des achats
Private m_iIndexAchat       As Integer

'Pour savoir à quel index la rechercher est rendu
Private m_iIndexRecherche   As Integer

'Pour savoir par quoi trier le ListView
Private m_sTri              As String

'Pour savoir si c'est une soumission ou un projet
Private m_eType             As enumType

Public m_bAnnulerContact    As Boolean
Public m_sContact           As String

Public Sub Afficher(ByVal eCatalogue As enumCatalogue, ByVal eDemande As enumModeDemande)

5       On Error GoTo AfficherErreur

10      m_eCatalogue = eCatalogue
15      m_eDemande = eDemande
20      m_bProjSoum = False
        
25      Select Case eDemande
          Case MODE_FOURNISSEUR:
30          Call RemplirListViewFournisseur(False)
        
35          Call AfficherControles(Fournisseur)
      
40        Case MODE_PIECE:
45          Call RemplirComboCategorie
    
50          Call AfficherControles(PIECE)
      
55        Case MODE_CATEGORIE:
60          Call RemplirListViewCatalogue
      
65          Call AfficherControles(Categorie)
      
70        Case MODE_NOUVELLE:
75          Call RemplirComboCategorie
      
80          Call AfficherControles(NOUVELLE_PIECE)

85          If m_eDemande = MODE_NOUVELLE Then
90            Call RemplirComboManufacturiers
95          End If
100     End Select
  
105     Call Show(vbModal)

110     Exit Sub

AfficherErreur:

115     woups "frmChoixDemande", "Afficher", Err, Erl
End Sub

Public Sub AfficherProjetSoumission(ByVal sNoProjSoum As String, ByVal eCatalogue As enumCatalogue, ByVal eDemande As enumModeDemande, ByVal iType As Integer)

5       On Error GoTo AfficherErreur

10      m_eCatalogue = eCatalogue
15      m_eDemande = eDemande
20      m_sNoProjSoum = sNoProjSoum
25      m_eType = iType

30      txtNoGRB.Text = sNoProjSoum
  
35      Call RemplirListViewPieceProjetSoumission
  
40      Call AfficherControles(PIECE)
  
45      cmbTri.Visible = False
50      cmdTri.Visible = False
55      cmdRafraichir.Visible = False
  
60      cmbCategorie.Visible = False
65      lblCategorie.Visible = False
  
70      m_bProjSoum = True
  
75      Call Me.Show(vbModal)

80      Exit Sub

AfficherErreur:

85      woups "frmChoixDemande", "AfficherProjet", Err, Erl
End Sub

Public Sub AfficherAchat(ByVal sNoAchat As String, ByVal eCatalogue As enumCatalogue, ByVal eDemande As enumModeDemande)

5       On Error GoTo AfficherErreur

10      m_eCatalogue = eCatalogue
15      m_eDemande = eDemande

20      txtNoGRB.Text = sNoAchat

25      m_sNoAchat = Left$(sNoAchat, 9)
30      m_iIndexAchat = CInt(Right$(sNoAchat, 3))
  
35      Call RemplirListViewPieceAchat
  
40      Call AfficherControles(PIECE)
  
45      cmbTri.Visible = False
50      cmdTri.Visible = False
55      cmdRafraichir.Visible = False
  
60      cmbCategorie.Visible = False
65      lblCategorie.Visible = False
  
70      m_bAchat = True
  
75      Call Show(vbModal)

80      Exit Sub

AfficherErreur:

85      woups "frmChoixDemande", "AfficherProjet", Err, Erl
End Sub

Private Sub AfficherControles(ByVal eMode As enumMode)

5       On Error GoTo AfficherErreur

10      Dim bCategorie        As Boolean
15      Dim bLvwPiece         As Boolean
20      Dim bLvwFournisseur   As Boolean
25      Dim bLvwCategorie     As Boolean
30      Dim bLvwManufacturier As Boolean
35      Dim bLvwNouvelle      As Boolean
40      Dim bCmdOK            As Boolean
45      Dim bCmdImprimer      As Boolean
50      Dim bCmdLangage       As Boolean
55      Dim bNoGRB            As Boolean
60      Dim bDate             As Boolean
65      Dim bCommentaire      As Boolean
70      Dim bNoPiece          As Boolean
75      Dim bManufact         As Boolean
80      Dim bDescription      As Boolean
85      Dim bCmdAjouter       As Boolean
90      Dim iHeight           As Integer
95      Dim bRechercher       As Boolean
100     Dim bTri              As Boolean
105     Dim bSelectAll        As Boolean
110     Dim bDeselectAll      As Boolean
      
115     m_eMode = eMode
  
120     Select Case eMode
          Case Fournisseur:
125         bLvwFournisseur = True
130         bCmdImprimer = True
135         bCmdLangage = True
140         bNoGRB = True
145         bDate = True
150         bCommentaire = True
155         bRechercher = True
160         bSelectAll = True
165         bDeselectAll = True
      
170         iHeight = 6855
    
          Case PIECE:
175         bCategorie = True
180         bLvwPiece = True
185         bCmdOK = True
190         bTri = True
195         bSelectAll = True
200         bDeselectAll = True
            
205         iHeight = 6150

          Case Manufacturier:
210         bLvwManufacturier = True
215         bCmdOK = True
            
220         bSelectAll = True
225         bDeselectAll = True
            
230         iHeight = 6150
      
          Case Categorie:
235         bLvwCategorie = True
240         bCmdOK = True
245         bSelectAll = True
250         bDeselectAll = True
      
255         iHeight = 6150
      
          Case NOUVELLE_PIECE:
260         bLvwNouvelle = True
265         bNoPiece = True
270         bManufact = True
275         bDescription = True
280         bCategorie = True
285         bCmdOK = True
290         bCmdAjouter = True
      
295         iHeight = 6150
300     End Select
  
305     Me.Height = iHeight
  
310     lblCategorie.Visible = bCategorie
315     cmbCategorie.Visible = bCategorie
  
320     lvwPiece.Visible = bLvwPiece
325     lvwfournisseur.Visible = bLvwFournisseur
330     lvwCategorie.Visible = bLvwCategorie
335     lvwNouvellesPieces.Visible = bLvwNouvelle
340     lvwManufacturier.Visible = bLvwManufacturier
  
345     cmdSelectAll.Visible = bSelectAll
350     cmdDeselectAll.Visible = bDeselectAll
  
355     cmdOk.Visible = bCmdOK
  
360     lblNoPiece.Visible = bNoPiece
365     txtNoPiece.Visible = bNoPiece
  
370     lblManufacturier.Visible = bManufact
375     cmbManufacturier.Visible = bManufact
  
380     lblDescription.Visible = bDescription
385     txtDescription.Visible = bDescription
  
390     Cmdajouter.Visible = bCmdAjouter
    
395     cmdImprimer.Visible = bCmdImprimer
400     cmdLangage.Visible = bCmdLangage

405     lblNoGRB.Visible = bNoGRB
410     txtNoGRB.Visible = bNoGRB
  
415     lblDateRequise.Visible = bDate
420     mskDateRequise.Visible = bDate
425     lblFormatDate.Visible = bDate
  
430     lblCommentaire.Visible = bCommentaire
435     txtcommentaire.Visible = bCommentaire
  
440     txtRechercher.Visible = bRechercher
445     cmdRechercher.Visible = bRechercher
  
450     cmbTri.Visible = bTri
455     cmdRafraichir.Visible = bTri
460     cmdTri.Visible = bTri

465     Exit Sub

AfficherErreur:

470     woups "frmChoixDemande", "AfficherControles", Err, Erl
End Sub

Private Sub RemplirListViewManufacturier()
  
5       On Error GoTo AfficherErreur

10      Dim rstManufact As ADODB.Recordset
15      Dim sWhere      As String
20      Dim iCompteur   As Integer

25      Call lvwManufacturier.ListItems.Clear

30      lvwManufacturier.Sorted = True
35      lvwManufacturier.SortKey = 0

40      sWhere = "CATEGORIE In ("

45      For iCompteur = 1 To m_collCategorie.count
50        If iCompteur <> m_collCategorie.count Then
55          sWhere = sWhere & "'" & Replace(m_collCategorie(iCompteur), "'", "''") & "',"
60        Else
65          sWhere = sWhere & "'" & Replace(m_collCategorie(iCompteur), "'", "''") & "')"
70        End If
75      Next

80      Set rstManufact = New ADODB.Recordset

85      If m_eCatalogue = ELECTRIQUE Then
90        Call rstManufact.Open("SELECT DISTINCT FABRICANT FROM GRB_CatalogueElec WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
95      Else
100       Call rstManufact.Open("SELECT DISTINCT FABRICANT FROM GRB_CatalogueMec WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
105     End If

110     Do While Not rstManufact.EOF
115       Call lvwManufacturier.ListItems.Add(, , rstManufact.Fields("FABRICANT"))
        
120       Call rstManufact.MoveNext
125     Loop

130     Call rstManufact.Close
135     Set rstManufact = Nothing

140     Exit Sub

AfficherErreur:

145     woups "frmChoixDemande", "RemplirListViewManufacturier", Err, Erl
End Sub

Private Sub RemplirListViewFournisseur(ByVal bPiece As Boolean)

5       On Error GoTo AfficherErreur

10      Dim rstFRS    As ADODB.Recordset
15      Dim itmFRS    As ListItem
20      Dim sWhere    As String
25      Dim iCompteur As Integer
  
30      m_bPiece = bPiece
    
35      Call lvwfournisseur.ListItems.Clear
  
40      Set rstFRS = New ADODB.Recordset
  
45      If bPiece = False Then
50        Call rstFRS.Open("SELECT NomFournisseur, IDFRS FROM GRB_Fournisseur WHERE Supprimé = False ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
55      Else
60        sWhere = "PIECE In ("

65        For iCompteur = 1 To m_collPiece.count
70          If iCompteur <> m_collPiece.count Then
75            sWhere = sWhere & "'" & Replace(m_collPiece(iCompteur), "'", "''") & "',"
80          Else
85            sWhere = sWhere & "'" & Replace(m_collPiece(iCompteur), "'", "''") & "')"
90          End If
95        Next
  
100       Call rstFRS.Open("SELECT DISTINCT GRB_Fournisseur.NomFournisseur, GRB_Fournisseur.IDFRS FROM GRB_PiecesFRS INNER JOIN GRB_Fournisseur ON GRB_PiecesFRS.IDFRS = GRB_Fournisseur.IDFRS WHERE " & sWhere & " ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
105     End If
  
110     Do While Not rstFRS.EOF
115       Set itmFRS = lvwfournisseur.ListItems.Add
    
120       itmFRS.Text = rstFRS.Fields("NomFournisseur")
    
125       itmFRS.Tag = rstFRS.Fields("IDFRS")
    
130       itmFRS.SubItems(I_COL_LANGAGE) = S_FRANCAIS
    
135       Call rstFRS.MoveNext
140     Loop
  
145     Call rstFRS.Close
150     Set rstFRS = Nothing

155     Exit Sub

AfficherErreur:

160     woups "frmChoixDemande", "RemplirListViewFournisseur", Err, Erl
End Sub

Private Sub RemplirComboCategorie()

5       On Error GoTo AfficherErreur

10      Dim rstCategorie As ADODB.Recordset
  
15      Call cmbCategorie.Clear
  
20      Set rstCategorie = New ADODB.Recordset
  
25      If m_eCatalogue = ELECTRIQUE Then
30        Call rstCategorie.Open("SELECT DISTINCT CATEGORIE FROM GRB_CatalogueElec ORDER BY Categorie", g_connData, adOpenDynamic, adLockOptimistic)
35      Else
40        Call rstCategorie.Open("SELECT DISTINCT CATEGORIE FROM GRB_CatalogueMec ORDER BY Categorie", g_connData, adOpenDynamic, adLockOptimistic)
45      End If
   
50      Do While Not rstCategorie.EOF
55        If Not IsNull(rstCategorie.Fields("CATEGORIE")) Then
60          Call cmbCategorie.AddItem(rstCategorie.Fields("CATEGORIE"))
65        End If
    
70        Call rstCategorie.MoveNext
75      Loop
  
80      Call rstCategorie.Close
85      Set rstCategorie = Nothing
  
90      If cmbCategorie.ListCount > 0 Then
95        cmbCategorie.ListIndex = 0
100     End If

105     Exit Sub

AfficherErreur:

110     woups "frmChoixDemande", "RemplirComboCategorie", Err, Erl
End Sub

Private Sub RemplirListViewCatalogue()

5       On Error GoTo AfficherErreur

10      Dim rstCategorie As ADODB.Recordset
  
15      Call lvwCategorie.ListItems.Clear
  
20      Set rstCategorie = New ADODB.Recordset
  
25      If m_eCatalogue = ELECTRIQUE Then
30        Call rstCategorie.Open("SELECT DISTINCT CATEGORIE FROM GRB_CatalogueElec ORDER BY CATEGORIE", g_connData, adOpenDynamic, adLockOptimistic)
35      Else
40        Call rstCategorie.Open("SELECT DISTINCT CATEGORIE FROM GRB_CatalogueMec ORDER BY CATEGORIE", g_connData, adOpenDynamic, adLockOptimistic)
45      End If
  
50      Do While Not rstCategorie.EOF
55        Call lvwCategorie.ListItems.Add(, , rstCategorie.Fields("CATEGORIE"))
 
60        Call rstCategorie.MoveNext
65      Loop
  
70      Call rstCategorie.Close
75      Set rstCategorie = Nothing

80      Exit Sub

AfficherErreur:

85      woups "frmChoixDemande", "RemplirListViewCatalogue", Err, Erl
End Sub

Private Function TrouverIndexPiece(ByVal sPiece As String, ByVal sDescriptionFR As String, ByVal sDescriptionEN As String, ByVal sFabricant As String, ByVal iIndexActuel As Integer)

5       On Error GoTo AfficherErreur

10      Dim iIndex As Integer
15      Dim sTri   As String
20      Dim bDebut As Boolean
  
25      sTri = UCase(m_sTri)
  
30      sPiece = UCase(sPiece)
35      sDescriptionFR = UCase(sDescriptionFR)
40      sDescriptionEN = UCase(sDescriptionEN)
45      sFabricant = UCase(sFabricant)
    
50      If sTri <> vbNullString Then
55        bDebut = False
      
          'Selon le tri
60        Select Case cmbTri.ListIndex
            'Si c'est trier par PIECE
            Case I_CMB_PIECE:
              'Si la PIECE contient la recherche
65            If InStr(1, sPiece, sTri) > 0 Then
                'On met la variable à true pour l'ajouter au début
70              bDebut = True
75            End If
                      
            'Si c'est trier par DESCR_FR
            Case I_CMB_DESCRIPTION_FR:
              'Si la description contient la recherche
80            If InStr(1, sDescriptionFR, sTri) > 0 Then
                'On met la variable à true pour l'ajouter au début
85              bDebut = True
90            End If
            
            'Si c'est trier par DESCR_EN
            Case I_CMB_DESCRIPTION_EN:
              'Si la description contient la recherche
95            If InStr(1, sDescriptionEN, sTri) > 0 Then
                'On met la variable à true pour l'ajouter au début
100             bDebut = True
105           End If
            
            'Si c'est la colonne Manufacturier
            Case I_CMB_FABRICANT:
              'Si le manufacturier contient la recherche
110           If InStr(1, sFabricant, sTri) > 0 Then
               'On met la variable à true pour l'ajouter au début
115             bDebut = True
120           End If
125       End Select
      
130       If bDebut = True Then
135         iIndex = iIndexActuel + 1
140       Else
145         iIndex = 0
150       End If
155     Else
160       iIndex = 0
165     End If
  
170     TrouverIndexPiece = iIndex

175     Exit Function

AfficherErreur:

180     woups "frmChoixDemande", "TrouverIndexPiece", Err, Erl
End Function

Private Sub RemplirListViewPiece()

5       On Error GoTo AfficherErreur

10      Dim rstPiece   As ADODB.Recordset
15      Dim itmPiece   As ListItem
20      Dim sCategorie As String
25      Dim sOrderBy   As String
30      Dim iIndex     As Integer
35      Dim iCompteur  As Integer
      
40      sCategorie = Replace(cmbCategorie.Text, "'", "''")
  
45      Call lvwPiece.ListItems.Clear
  
        'Pour savoir par quoi trier le recordset
50      Select Case cmbTri.ListIndex
          Case I_CMB_PIECE:          sOrderBy = "PIECE"
55        Case I_CMB_DESCRIPTION_FR: sOrderBy = "DESC_FR"
60        Case I_CMB_DESCRIPTION_EN: sOrderBy = "DESC_EN"
65        Case I_CMB_FABRICANT:      sOrderBy = "FABRICANT"
70      End Select
    
75      Set rstPiece = New ADODB.Recordset
    
80      If m_eCatalogue = ELECTRIQUE Then
85        Call rstPiece.Open("SELECT * FROM GRB_CatalogueElec WHERE CATEGORIE = '" & sCategorie & "' ORDER BY " & sOrderBy, g_connData, adOpenDynamic, adLockOptimistic)
90      Else
95        Call rstPiece.Open("SELECT * FROM GRB_CatalogueMec WHERE CATEGORIE = '" & sCategorie & "' ORDER BY " & sOrderBy, g_connData, adOpenDynamic, adLockOptimistic)
100     End If
      
105     Do While Not rstPiece.EOF
110       If Not IsNull(rstPiece.Fields("DESC_FR")) And Not IsNull(rstPiece.Fields("DESC_EN")) Then
115         iIndex = TrouverIndexPiece(rstPiece.Fields("PIECE"), rstPiece.Fields("DESC_FR"), rstPiece.Fields("DESC_EN"), rstPiece.Fields("FABRICANT"), iIndex)
120       Else
125         If Not IsNull(rstPiece.Fields("DESC_FR")) Then
130           iIndex = TrouverIndexPiece(rstPiece.Fields("PIECE"), rstPiece.Fields("DESC_FR"), vbNullString, rstPiece.Fields("FABRICANT"), iIndex)
135         Else
140           If Not IsNull(rstPiece.Fields("DESC_EN")) Then
145             iIndex = TrouverIndexPiece(rstPiece.Fields("PIECE"), vbNullString, rstPiece.Fields("DESC_EN"), rstPiece.Fields("FABRICANT"), iIndex)
150           Else
155             iIndex = TrouverIndexPiece(rstPiece.Fields("PIECE"), vbNullString, vbNullString, rstPiece.Fields("FABRICANT"), iIndex)
160           End If
165         End If
170       End If

175       If iIndex = 0 Then
180         Set itmPiece = lvwPiece.ListItems.Add
185       Else
190         Set itmPiece = lvwPiece.ListItems.Add(iIndex)
195       End If
         
200       For iCompteur = 1 To m_collPiece.count
205         If m_collCategorie(iCompteur) = cmbCategorie.Text Then
210           If m_collPiece(iCompteur) = rstPiece.Fields("PIECE") Then
215             If m_collDescriptionFR(iCompteur) = rstPiece.Fields("DESC_FR") Then
220               If m_collDescriptionEN(iCompteur) = rstPiece.Fields("DESC_EN") Then
225                 If m_collManufacturier(iCompteur) = rstPiece.Fields("FABRICANT") Then
230                   itmPiece.Checked = True

235                   Exit For
240                 End If
245               End If
250             End If
255           End If
260         End If
265       Next
          
270       itmPiece.SubItems(I_COL_PIECE) = rstPiece.Fields("PIECE")

275       If Not IsNull(rstPiece.Fields("DESC_FR")) Then
280         itmPiece.SubItems(I_COL_DESC_FR) = rstPiece.Fields("DESC_FR")
285       Else
290         itmPiece.SubItems(I_COL_DESC_FR) = vbNullString
295       End If

300       If Not IsNull(rstPiece.Fields("DESC_EN")) Then
305         itmPiece.SubItems(I_COL_DESC_EN) = rstPiece.Fields("DESC_EN")
310       Else
315         itmPiece.SubItems(I_COL_DESC_EN) = vbNullString
320       End If

325       itmPiece.SubItems(I_COL_FABRICANT) = rstPiece.Fields("FABRICANT")
       
330       Call rstPiece.MoveNext
335     Loop
  
340     Call rstPiece.Close
345     Set rstPiece = Nothing

350     Exit Sub

AfficherErreur:

355     woups "frmChoixDemande", "RemplirListViewPiece", Err, Erl
End Sub

Private Sub RemplirListViewPieceProjetSoumission()

5       On Error GoTo AfficherErreur

10      Dim rstPiece As ADODB.Recordset
15      Dim itmPiece As ListItem
  
20      Call lvwPiece.ListItems.Clear
            
25      lvwPiece.Sorted = False
            
30      Set rstPiece = New ADODB.Recordset
            
        'Si c'est un projet
35      If m_eType = TYPE_PROJET Then
40        Call rstPiece.Open("SELECT Qté, NumItem, Desc_FR, Desc_EN, Manufact, IDFRS, PieceExtraChargeable, PieceExtraNonChargeable FROM GRB_Projet_Pieces WHERE (IDProjet = '" & m_sNoProjSoum & "') AND (IDFRS = 0 AND NumItem <> 'Texte') ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
45      Else
          'Si c'est une soumission
50        Call rstPiece.Open("SELECT Qté, NumItem, Desc_FR, Desc_En, Manufact, IDFRS, PieceExtraChargeable, PieceExtraNonChargeable FROM GRB_Soumission_Pieces WHERE (IDSoumission = '" & m_sNoProjSoum & "') AND (IDFRS = 0 AND NumItem <> 'Texte') ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
55      End If
  
60      Do While Not rstPiece.EOF
65        If rstPiece.Fields("PieceExtraChargeable") = False And rstPiece.Fields("PieceExtraNonChargeable") = False Then
70          Set itmPiece = lvwPiece.ListItems.Add
             
75          itmPiece.Text = rstPiece.Fields("Qté")
          
80          itmPiece.SubItems(I_COL_PIECE) = rstPiece.Fields("NumItem")
    
85          If Not IsNull(rstPiece.Fields("Desc_FR")) Then
90            itmPiece.SubItems(I_COL_DESC_FR) = rstPiece.Fields("Desc_FR")
95          Else
100           itmPiece.SubItems(I_COL_DESC_FR) = vbNullString
105         End If

110         If Not IsNull(rstPiece.Fields("Desc_En")) Then
115           itmPiece.SubItems(I_COL_DESC_EN) = rstPiece.Fields("Desc_En")
120         Else
125           itmPiece.SubItems(I_COL_DESC_EN) = vbNullString
130         End If
    
135         itmPiece.SubItems(I_COL_FABRICANT) = rstPiece.Fields("Manufact")
140       End If
    
145       Call rstPiece.MoveNext
150     Loop
  
155     Call rstPiece.Close
160     Set rstPiece = Nothing

165     Exit Sub

AfficherErreur:

170     woups "frmChoixDemande", "RemplirListViewPieceProjet", Err, Erl
End Sub

Private Sub RemplirListViewPieceAchat()

5       On Error GoTo AfficherErreur

10      Dim rstPiece As ADODB.Recordset
15      Dim itmPiece As ListItem
  
20      Call lvwPiece.ListItems.Clear
            
25      Set rstPiece = New ADODB.Recordset
            
30      Call rstPiece.Open("SELECT Qté, PIECE, Desc_FR, Desc_EN, Manufact, IDFRS FROM GRB_Achat_Pieces WHERE IDAchat = '" & m_sNoAchat & "' AND IndexAchat = " & m_iIndexAchat & " AND IDFRS = 0 ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
  
35      Do While Not rstPiece.EOF
40        Set itmPiece = lvwPiece.ListItems.Add
             
45        itmPiece.Text = rstPiece.Fields("Qté")
        
50        itmPiece.SubItems(I_COL_PIECE) = rstPiece.Fields("PIECE")
    
55        If Not IsNull(rstPiece.Fields("Desc_FR")) Then
60          itmPiece.SubItems(I_COL_DESC_FR) = rstPiece.Fields("Desc_FR")
65        Else
70          itmPiece.SubItems(I_COL_DESC_FR) = vbNullString
75        End If

80        If Not IsNull(rstPiece.Fields("Desc_En")) Then
85          itmPiece.SubItems(I_COL_DESC_EN) = rstPiece.Fields("Desc_En")
90        Else
95          itmPiece.SubItems(I_COL_DESC_EN) = vbNullString
100       End If
    
105       itmPiece.SubItems(I_COL_FABRICANT) = rstPiece.Fields("Manufact")
    
110       Call rstPiece.MoveNext
115     Loop
  
120     Call rstPiece.Close
125     Set rstPiece = Nothing

130     Exit Sub

AfficherErreur:

135     woups "frmChoixDemande", "RemplirListViewPieceProjet", Err, Erl
End Sub

Private Sub cmbCategorie_Click()

5       On Error GoTo AfficherErreur

10      Call AjouterPieceCollection

15      m_sCategorie = cmbCategorie.Text
  
20      Call RemplirListViewPiece
  
25      Call CocherCases

30      Exit Sub

AfficherErreur:

35      woups "frmChoixDemande", "cmbCategorie_Click", Err, Erl
End Sub

Private Sub RemplirComboManufacturiers()

5       On Error GoTo AfficherErreur

10      Dim rstPiece As ADODB.Recordset
  
15      Call cmbManufacturier.Clear
     
20      Set rstPiece = New ADODB.Recordset
     
25      If m_eCatalogue = ELECTRIQUE Then
30        Call rstPiece.Open("SELECT DISTINCT FABRICANT FROM GRB_CatalogueElec", g_connData, adOpenDynamic, adLockOptimistic)
35      Else
40        Call rstPiece.Open("SELECT DISTINCT FABRICANT FROM GRB_CatalogueMec", g_connData, adOpenDynamic, adLockOptimistic)
45      End If
    
50      Do While Not rstPiece.EOF
55        If Not IsNull(rstPiece.Fields("FABRICANT")) Then
60          Call cmbManufacturier.AddItem(rstPiece.Fields("FABRICANT"))
65        End If
    
70        Call rstPiece.MoveNext
75      Loop
  
80      Call rstPiece.Close
85      Set rstPiece = Nothing

90      Exit Sub

AfficherErreur:

95      woups "frmChoixDemande", "RemplirComboManufacturiers", Err, Erl
End Sub

Private Sub CocherCases()

5       On Error GoTo AfficherErreur

10      Dim iCompteur  As Integer
15      Dim iCompteur2 As Integer
  
20      For iCompteur = 1 To m_collCategorie.count
25        If m_collCategorie(iCompteur) = cmbCategorie.Text Then
30          For iCompteur2 = 1 To lvwPiece.ListItems.count
35            If lvwPiece.ListItems(iCompteur2).SubItems(I_COL_PIECE) = m_collPiece(iCompteur) Then
40              lvwPiece.ListItems(iCompteur2).Checked = True
45            End If
50          Next iCompteur2
55        End If
60      Next iCompteur

65      Exit Sub

AfficherErreur:

70      woups "frmChoixDemande", "CocherCases", Err, Erl
End Sub

Private Sub cmbManufacturier_KeyPress(KeyAscii As Integer)

5       On Error GoTo AfficherErreur

10      If KeyAscii <= 122 And KeyAscii >= 97 Then
15        KeyAscii = KeyAscii - 32
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmChoixDemande", "cmbManufacturier_KeyPress", Err, Erl
End Sub

Private Sub Cmdajouter_Click()

5       On Error GoTo AfficherErreur

10      Dim rstPiece  As ADODB.Recordset
15      Dim iCompteur As Integer
20      Dim bTrouver  As Boolean
25      Dim itmPiece  As ListItem
30      Dim sQuantite As String
  
35      If txtNoPiece.Text = vbNullString Or cmbManufacturier.Text = vbNullString Or txtDescription.Text = vbNullString Then
40        Call MsgBox("Vous devez absolument remplir tous les champs!", vbOKOnly, "Erreur")
    
45        Exit Sub
50      End If
  
55      If InStr(1, txtNoPiece.Text, "'") > 0 Then
60        Call MsgBox("Numéro invalide! Le numéro ne doit pas contenir d'appostrophes!", vbOKOnly, "Erreur")
    
65        Exit Sub
70      End If
  
75      Set rstPiece = New ADODB.Recordset
  
80      If m_eCatalogue = ELECTRIQUE Then
85        Call rstPiece.Open("SELECT * FROM GRB_CatalogueElec WHERE PIECE = '" & Replace(txtNoPiece.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
90      Else
95        Call rstPiece.Open("SELECT * FROM GRB_CatalogueMec WHERE PIECE = '" & Replace(txtNoPiece.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
100     End If
  
105     If rstPiece.EOF = False Then
110       bTrouver = True
115     End If
    
120     Call rstPiece.Close
125     Set rstPiece = Nothing
  
135     If bTrouver = True Then
140       Call MsgBox("Le numéro de pièce existe déjà!", vbOKOnly, "Erreur")
    
145       Exit Sub
150     End If
  
155     sQuantite = InputBox("Quelle est la quantité?")
  
160     If sQuantite <> vbNullString Then
165       If Not IsNumeric(sQuantite) Then
170         Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")
     
175         Exit Sub
180       End If
185     Else
190       sQuantite = "1"
195     End If
  
200     Set itmPiece = lvwNouvellesPieces.ListItems.Add
  
205     itmPiece.Text = sQuantite
210     itmPiece.SubItems(I_COL_NO_PIECE) = txtNoPiece.Text
215     itmPiece.SubItems(I_COL_DESCRIPTION) = txtDescription.Text
220     itmPiece.SubItems(I_COL_MANUFACT) = cmbManufacturier.Text
225     itmPiece.SubItems(I_COL_CATEGORIE) = cmbCategorie.Text

230     Exit Sub

AfficherErreur:

235     woups "frmChoixDemande", "cmdAjouter_Click", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmChoixDemande", "cmdFermer_Click", Err, Erl
End Sub

Private Sub cmdImprimer_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
15      Dim itmFRS    As ListItem
         
20      If lvwfournisseur.ListItems.count > 0 Then
25        If VerifierSiChecked = True Then
30          For iCompteur = 1 To lvwfournisseur.ListItems.count
35            If lvwfournisseur.ListItems(iCompteur).Checked = True Then
40              Set itmFRS = lvwfournisseur.ListItems(iCompteur)
            
45              m_sLangage = itmFRS.SubItems(I_COL_LANGAGE)

50              Call frmChoixContactFRS.Afficher(itmFRS.Tag)
          
55              If m_bAnnulerContact = False Then
60                If m_eDemande = MODE_NOUVELLE Then
65                  Call EnregistrerDemandePrixNouvellesPieces
70                Else
75                  Call EnregistrerDemandePrix(itmFRS.Tag)
80                End If
          
85                Call ImprimerDemandePrix(itmFRS.Tag)
90              End If
95            End If
100         Next
      
105         If m_eDemande = MODE_NOUVELLE Then
110           Call EnregistrerPieces
115         End If
120       End If
125     End If

130     Exit Sub

AfficherErreur:

135     woups "frmChoixDemande", "cmdImprimer_Click", Err, Erl
End Sub

Private Sub EnregistrerPieces()

5       On Error GoTo AfficherErreur

10      Dim rstPiece        As ADODB.Recordset
15      Dim rstPiecesFRS    As ADODB.Recordset
20      Dim iCompteurPieces As Integer
25      Dim iCompteurFRS    As Integer
  
30      Set rstPiece = New ADODB.Recordset
35      Set rstPiecesFRS = New ADODB.Recordset

40      For iCompteurPieces = 1 To m_collPiece.count
45        If m_eCatalogue = ELECTRIQUE Then
50          Call rstPiece.Open("SELECT * FROM GRB_CatalogueElec", g_connData, adOpenDynamic, adLockOptimistic)
55        Else
60          Call rstPiece.Open("SELECT * FROM GRB_CatalogueMec", g_connData, adOpenDynamic, adLockOptimistic)
65        End If
      
70        Call rstPiece.AddNew
      
75        rstPiece.Fields("PIECE") = m_collPiece(iCompteurPieces)
80        rstPiece.Fields("PIECE_GRB") = m_collPiece(iCompteurPieces) & "GRB"
85        rstPiece.Fields("DESC_FR") = m_collDescriptionFR(iCompteurPieces)
90        rstPiece.Fields("FABRICANT") = m_collManufacturier(iCompteurPieces)
95        rstPiece.Fields("CATEGORIE") = m_collCategorie(iCompteurPieces)
    
100       Call rstPiece.Update
      
105       Call rstPiecesFRS.Open("SELECT * FROM GRB_PiecesFRS", g_connData, adOpenDynamic, adLockOptimistic)
      
110       For iCompteurFRS = 1 To lvwfournisseur.ListItems.count
115         If lvwfournisseur.ListItems(iCompteurFRS).Checked = True Then
120           Call rstPiecesFRS.AddNew
        
125           rstPiecesFRS.Fields("IDFRS") = lvwfournisseur.ListItems(iCompteurFRS).Tag
130           rstPiecesFRS.Fields("PIECE") = m_collPiece(iCompteurPieces)
135           rstPiecesFRS.Fields("DATE") = ConvertDate(Date)
140           rstPiecesFRS.Fields("ENTRER_PAR") = g_sInitiale
145           rstPiecesFRS.Fields("PRIX_SP") = vbNullString
150           rstPiecesFRS.Fields("PERS_RESS") = vbNullString
155           rstPiecesFRS.Fields("PRIX_LIST") = "0"
160           rstPiecesFRS.Fields("ESCOMPTE") = "0"
165           rstPiecesFRS.Fields("PRIX_NET") = "0"
170           rstPiecesFRS.Fields("DeviseMonétaire") = "CAN"
175           rstPiecesFRS.Fields("PrixReel") = "0"
180           rstPiecesFRS.Fields("Type") = "M"
         
185           Call rstPiecesFRS.Update
190         End If
195       Next
     
200       Call rstPiecesFRS.Close
    
205       Call rstPiece.Close
210     Next

215     Set rstPiece = Nothing
220     Set rstPiecesFRS = Nothing
  
225     Exit Sub

AfficherErreur:

230     woups "frmChoixDemande", "EnregistrerPieces", Err, Erl
End Sub

Private Sub cmdLangage_Click()

5       On Error GoTo AfficherErreur

10      If cmdLangage.Caption = S_DEMANDE_FRANCAIS Then
15        lvwfournisseur.SelectedItem.SubItems(I_COL_LANGAGE) = S_FRANCAIS
    
20        cmdLangage.Caption = S_DEMANDE_ANGLAIS
25      Else
30        lvwfournisseur.SelectedItem.SubItems(I_COL_LANGAGE) = S_ANGLAIS
    
35        cmdLangage.Caption = S_DEMANDE_FRANCAIS
40      End If
  
45      Call lvwfournisseur.SetFocus

50      Exit Sub

AfficherErreur:

55      woups "frmChoixDemande", "cmdLangage_Click", Err, Erl
End Sub

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur

10      If m_eMode = PIECE Then
15        Call AjouterPieceCollection
  
20        Call AfficherFournisseur
25      Else
30        If m_eMode = NOUVELLE_PIECE Then
35          Call AjouterNouvellePieceCollection
      
40          Call RemplirListViewFournisseur(False)
                    
45          Call AfficherControles(Fournisseur)
50        Else
55          If m_eMode = Manufacturier Then
60            Call AjouterManufacturierCollection

65            Call RemplirListViewFournisseur(False)

70            Call AfficherControles(Fournisseur)
75          Else
80            Call AjouterCategorieCollection
    
85            If VerifierSiChecked = True Then
90              Call RemplirListViewManufacturier
    
95              Call AfficherControles(Manufacturier)
100           End If
105         End If
110       End If
115     End If

120     Exit Sub

AfficherErreur:

125     woups "frmChoixDemande", "cmdOK_Click", Err, Erl
End Sub

Private Sub AjouterNouvellePieceCollection()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
15      Dim itmPiece  As ListItem
  
20      For iCompteur = 1 To lvwNouvellesPieces.ListItems.count
25        Set itmPiece = lvwNouvellesPieces.ListItems(iCompteur)
      
30        Call m_collQuantite.Add(itmPiece.Text)
35        Call m_collPiece.Add(itmPiece.SubItems(I_COL_NO_PIECE))
40        Call m_collDescriptionFR.Add(itmPiece.SubItems(I_COL_DESCRIPTION))
45        Call m_collManufacturier.Add(itmPiece.SubItems(I_COL_MANUFACT))
50        Call m_collCategorie.Add(itmPiece.SubItems(I_COL_CATEGORIE))
55      Next

60      Exit Sub

AfficherErreur:

65      woups "frmChoixDemande", "AjouterNouvellePieceCollection", Err, Erl
End Sub

Private Sub AjouterManufacturierCollection()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      For iCompteur = 1 To lvwManufacturier.ListItems.count
20        If lvwManufacturier.ListItems(iCompteur).Checked = True Then
25          Call m_collManufacturier.Add(lvwManufacturier.ListItems(iCompteur).Text)
30        End If
35      Next

40      Exit Sub

AfficherErreur:

45      woups "frmChoixDemande", "AjouterManufacturierCollection", Err, Erl
End Sub

Private Sub AjouterCategorieCollection()

5       On Error GoTo AfficherErreur
      
10      Dim iCompteur As Integer
  
15      For iCompteur = 1 To lvwCategorie.ListItems.count
20        If lvwCategorie.ListItems(iCompteur).Checked = True Then
25          Call m_collCategorie.Add(lvwCategorie.ListItems(iCompteur).Text)
30        End If
35      Next

40      Exit Sub

AfficherErreur:

45      woups "frmChoixDemande", "AjouterCategorieCollection", Err, Erl
End Sub

Private Sub cmdRechercher_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
15      Dim bTrouver  As Boolean
  
        'Si le texte du bouton est rechercher
20      If cmdRechercher.Caption = "Rechercher" Then
          'Pour chaque élément du listview
25        For iCompteur = 1 To lvwfournisseur.ListItems.count
            'si le nom du fournisseur contient le texte à rechercher
30          If InStr(1, UCase(lvwfournisseur.ListItems(iCompteur).Text), UCase(txtRechercher.Text)) > 0 Then
35            bTrouver = True
      
40            lvwfournisseur.ListItems(iCompteur).Selected = True
        
45            Call lvwfournisseur.ListItems(iCompteur).EnsureVisible
        
50            Call lvwfournisseur.SetFocus
        
55            m_iIndexRecherche = iCompteur
        
60            Exit For
65          End If
70        Next
    
75        If bTrouver = True Then
80          cmdRechercher.Caption = "Suivant"
85        Else
90          Call MsgBox("Aucun fournisseur trouvé!", vbOKOnly)
95        End If
100     Else
          'Pour chaque élément restant du listview
105       For iCompteur = m_iIndexRecherche + 1 To lvwfournisseur.ListItems.count
            'Si le nom du fournisseur contient le texte à rechercher
110         If InStr(1, UCase(lvwfournisseur.ListItems(iCompteur).Text), UCase(txtRechercher.Text)) > 0 Then
115           bTrouver = True
      
120           lvwfournisseur.ListItems(iCompteur).Selected = True
        
125           Call lvwfournisseur.ListItems(iCompteur).EnsureVisible
        
130           Call lvwfournisseur.SetFocus
        
135           m_iIndexRecherche = iCompteur
        
140           Exit For
145         End If
150       Next
    
155       If bTrouver = False Then
160         Call MsgBox("Aucun fournisseur trouvé!", vbOKOnly)
      
165         cmdRechercher.Caption = "Rechercher"
170       End If
175     End If

180     Exit Sub

AfficherErreur:

185     woups "frmChoixDemande", "cmdRechercher_Click", Err, Erl
End Sub

Private Sub cmdSelectAll_Click()

5       On Error GoTo AfficherErreur

10      Dim lvwSource As ListView
15      Dim iCompteur As Integer

        Select Case m_eMode
          Case PIECE:         Set lvwSource = lvwPiece
          Case Categorie:     Set lvwSource = lvwCategorie
          Case Fournisseur:   Set lvwSource = lvwfournisseur
          Case Manufacturier: Set lvwSource = lvwManufacturier
        End Select

20      For iCompteur = 1 To lvwSource.ListItems.count
25        lvwSource.ListItems(iCompteur).Checked = True
30      Next

35      Exit Sub

AfficherErreur:

40      woups "frmChoixDemande", "cmdSelectAll_Click", Err, Erl
End Sub

Private Sub cmdDeSelectAll_Click()

5       On Error GoTo AfficherErreur

10      Dim lvwSource As ListView
15      Dim iCompteur As Integer

        Select Case m_eMode
          Case PIECE:         Set lvwSource = lvwPiece
          Case Categorie:     Set lvwSource = lvwCategorie
          Case Fournisseur:   Set lvwSource = lvwfournisseur
          Case Manufacturier: Set lvwSource = lvwManufacturier
        End Select

20      For iCompteur = 1 To lvwSource.ListItems.count
25        lvwSource.ListItems(iCompteur).Checked = False
30      Next

35      Exit Sub

AfficherErreur:

40      woups "frmChoixDemande", "cmdSelectAll_Click", Err, Erl
End Sub


Private Sub cmdTri_Click()

5       On Error GoTo AfficherErreur

10      m_sTri = InputBox("Quel est la pièce à trier?")
    
15      If m_sTri <> vbNullString Then
20        lvwCategorie.Sorted = False
25        lvwPiece.Sorted = False
30        lvwfournisseur.Sorted = False
35        lvwNouvellesPieces.Sorted = False

40        Call AjouterPieceCollection

45        Call RemplirListViewPiece
50      End If

55      Exit Sub

AfficherErreur:

60     woups "frmChoixDemande", "cmdTri_Click", Err, Erl
End Sub

Private Sub cmdRafraichir_Click()

5       On Error GoTo AfficherErreur

10      If m_sTri <> vbNullString Then
15        m_sTri = vbNullString
    
20        Call RemplirListViewPiece
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmChoixDemande", "cmdRafraichir_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Set m_collPiece = New Collection
15      Set m_collQuantite = New Collection
20      Set m_collDescriptionFR = New Collection
25      Set m_collDescriptionEN = New Collection
30      Set m_collCategorie = New Collection
35      Set m_collManufacturier = New Collection
    
40      cmbTri.ListIndex = I_CMB_PIECE

45      Exit Sub

AfficherErreur:

50      woups "frmChoixDemande", "Form_Load", Err, Erl
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
5       On Error GoTo AfficherErreur
  
10      m_sTri = vbNullString

15      Exit Sub

AfficherErreur:

20      woups "frmChoixDemande", "Form_Unload", Err, Erl
End Sub

Private Sub lvwCategorie_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

5       On Error GoTo AfficherErreur

10      lvwCategorie.Sorted = True
  
15      If lvwCategorie.SortOrder = lvwAscending Then
20        lvwCategorie.SortOrder = lvwDescending
25      Else
30        lvwCategorie.SortOrder = lvwAscending
35      End If
  
40      lvwCategorie.SortKey = ColumnHeader.Index - 1

45      Exit Sub

AfficherErreur:

50      woups "frmChoixDemande", "lvwCategorie_ColumnClick", Err, Erl
End Sub

Private Sub lvwFournisseur_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

5       On Error GoTo AfficherErreur

10      lvwfournisseur.Sorted = True
  
15      If lvwfournisseur.SortOrder = lvwAscending Then
20        lvwfournisseur.SortOrder = lvwDescending
25      Else
30        lvwfournisseur.SortOrder = lvwAscending
35      End If
  
40      lvwfournisseur.SortKey = ColumnHeader.Index - 1

45      Exit Sub

AfficherErreur:

50      woups "frmChoixDemande", "lvwFournisseur_ColumnClick", Err, Erl
End Sub

Private Sub lvwFournisseur_ItemClick(ByVal Item As MSComctlLib.ListItem)

5       On Error GoTo AfficherErreur

10      If Item.SubItems(I_COL_LANGAGE) = S_FRANCAIS Then
15        cmdLangage.Caption = S_DEMANDE_ANGLAIS
20      Else
25        cmdLangage.Caption = S_DEMANDE_FRANCAIS
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmChoixDemande", "lvwFournisseur_ItemClick", Err, Erl
End Sub

Private Sub lvwNouvellesPieces_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

5       On Error GoTo AfficherErreur

10      lvwNouvellesPieces.Sorted = True
  
15      If lvwNouvellesPieces.SortOrder = lvwAscending Then
20        lvwNouvellesPieces.SortOrder = lvwDescending
25      Else
30        lvwNouvellesPieces.SortOrder = lvwAscending
35      End If
  
40      lvwNouvellesPieces.SortKey = ColumnHeader.Index - 1

45      Exit Sub

AfficherErreur:

50      woups "frmChoixDemande", "lvwNouvellesPieces_ColumnClick", Err, Erl
End Sub

Private Sub lvwNouvellesPieces_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      If KeyCode = vbKeyDelete Then
15        Call lvwNouvellesPieces.ListItems.Remove(lvwNouvellesPieces.SelectedItem.Index)
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmChoixDemande", "lvwNouvellesPieces_KeyDown", Err, Erl
End Sub

Private Sub lvwPiece_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

5       On Error GoTo AfficherErreur

10      lvwPiece.Sorted = True
  
15      If lvwPiece.SortOrder = lvwAscending Then
20        lvwPiece.SortOrder = lvwDescending
25      Else
30        lvwPiece.SortOrder = lvwAscending
35      End If
  
40      lvwPiece.SortKey = ColumnHeader.Index - 1

45      Exit Sub

AfficherErreur:

50      woups "frmChoixDemande", "lvwPiece_ColumnClick", Err, Erl
End Sub

Private Sub lvwPiece_DblClick()

5       On Error GoTo AfficherErreur

10      Dim sQuantite As String
  
15      If lvwPiece.ListItems.count > 0 Then
20        sQuantite = InputBox("Quelle est la quantité?")
  
25        If sQuantite <> vbNullString Then
30          If IsNumeric(sQuantite) Then
35            lvwPiece.SelectedItem.Text = sQuantite
40          End If
45        End If
50      End If

55      Exit Sub

AfficherErreur:

60      woups "frmChoixDemande", "lvwPiece_DblClick", Err, Erl
End Sub

Private Sub AfficherFournisseur()

5       On Error GoTo AfficherErreur

10      If lvwPiece.ListItems.count > 0 Then
15        If VerifierSiChecked = True Then
      
20          Call RemplirListViewFournisseur(True)
  
25          If lvwfournisseur.ListItems.count > 0 Then
30            Call AfficherControles(Fournisseur)
35          Else
40            Call MsgBox("Il n'y a aucun fournisseur pour cette pièce!", vbOKOnly, "Erreur")
45          End If
50        End If
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmChoixDemande", "AfficherFournisseur", Err, Erl
End Sub

Private Sub AjouterPieceCollection()

5       On Error GoTo AfficherErreur

10      Dim iCompteur    As Integer
15      Dim iCompteur2   As Integer
20      Dim bPieceExiste As Boolean
25      Dim iQuantite    As Integer
30      Dim rstTempDP    As ADODB.Recordset

35      If m_eCatalogue = ELECTRIQUE Then
40        Call g_connData.Execute("DELETE * FROM GRB_TempDP WHERE TYPE = 'E'")
45      Else
50        Call g_connData.Execute("DELETE * FROM GRB_TempDP WHERE TYPE = 'M'")
55      End If

60      Set rstTempDP = New ADODB.Recordset

65      Call rstTempDP.Open("SELECT * FROM GRB_TempDP", g_connData, adOpenDynamic, adLockOptimistic)
  
70      For iCompteur = 1 To lvwPiece.ListItems.count
75        If lvwPiece.ListItems(iCompteur).Checked = True Then
80          bPieceExiste = False
         
85          For iCompteur2 = 1 To m_collPiece.count
90            If lvwPiece.ListItems(iCompteur).SubItems(I_COL_PIECE) = m_collPiece(iCompteur2) Then
95              bPieceExiste = True
       
100             Exit For
105           End If
110         Next iCompteur2
        
115         If bPieceExiste = False Then
120           Call m_collCategorie.Add(m_sCategorie)
125           Call m_collPiece.Add(lvwPiece.ListItems(iCompteur).SubItems(I_COL_PIECE))
130           Call m_collDescriptionFR.Add(lvwPiece.ListItems(iCompteur).SubItems(I_COL_DESC_FR))
135           Call m_collDescriptionEN.Add(lvwPiece.ListItems(iCompteur).SubItems(I_COL_DESC_EN))
140           Call m_collManufacturier.Add(lvwPiece.ListItems(iCompteur).SubItems(I_COL_FABRICANT))
      
145           If lvwPiece.ListItems(iCompteur).Text <> vbNullString Then
150             Call m_collQuantite.Add(lvwPiece.ListItems(iCompteur).Text)
155           Else
160             Call m_collQuantite.Add("1")
165           End If

170           Call rstTempDP.AddNew

175           rstTempDP.Fields("PIECE") = lvwPiece.ListItems(iCompteur).SubItems(I_COL_PIECE)
180           rstTempDP.Fields("ORDRE") = iCompteur

185           If m_eCatalogue = ELECTRIQUE Then
190             rstTempDP.Fields("TYPE") = "E"
195           Else
200             rstTempDP.Fields("TYPE") = "M"
205           End If

210           Call rstTempDP.Update
215         Else
              'Ajoute la quantité si c'est une demande de prix à partir d'un projet
220           If m_bProjSoum = True Then
225             iQuantite = Val(m_collQuantite(iCompteur2)) + Val(lvwPiece.ListItems(iCompteur).Text)
          
230             Call m_collQuantite.Remove(iCompteur2)
          
235             If m_collQuantite.count > 0 Then
240               If m_collQuantite.count < iCompteur2 Then
245                 Call m_collQuantite.Add(iQuantite)
250               Else
255                 If iCompteur2 > 1 Then
260                   Call m_collQuantite.Add(iQuantite, , , iCompteur2 - 1)
265                 Else
270                   Call m_collQuantite.Add(iQuantite, , , 1)
275                 End If
280               End If
285             Else
290               Call m_collQuantite.Add(iQuantite)
295             End If
300           End If
305         End If
310       End If
315     Next iCompteur

320     Call rstTempDP.Close
325     Set rstTempDP = Nothing

330     Exit Sub

AfficherErreur:

335     woups "frmChoixDemande", "AjouterPieceCollection", Err, Erl
End Sub

Private Function VerifierSiChecked() As Boolean

5       On Error GoTo AfficherErreur

10      Dim lvwSource As ListView
15      Dim iCompteur As Integer
  
20      If lvwPiece.Visible = True Then
25        Set lvwSource = lvwPiece
30      Else
35        If lvwfournisseur.Visible = True Then
40          Set lvwSource = lvwfournisseur
45        Else
50          Set lvwSource = lvwCategorie
55        End If
60      End If
  
65      VerifierSiChecked = False
  
70      For iCompteur = 1 To lvwSource.ListItems.count
75        If lvwSource.ListItems(iCompteur).Checked = True Then
80          VerifierSiChecked = True
      
85          Exit For
90        End If
95      Next

100     Exit Function

AfficherErreur:

105     woups "frmChoixDemande", "VerifierSiChecked", Err, Erl
End Function

Private Sub EnregistrerDemandePrixNouvellesPieces()

5       On Error GoTo AfficherErreur

10      Dim rstImpDP  As ADODB.Recordset
15      Dim sNomTable As String
20      Dim iCompteur As Integer
    
25      If m_eCatalogue = ELECTRIQUE Then
30        sNomTable = "GRB_ImpressionDemandePrixElec"
35      Else
40        sNomTable = "GRB_ImpressionDemandePrixMec"
45      End If
  
50      Set rstImpDP = New ADODB.Recordset
  
55      rstImpDP.CursorLocation = adUseClient
  
60      Call rstImpDP.Open("SELECT * FROM " & sNomTable, g_connData, adOpenDynamic, adLockOptimistic)
  
65      For iCompteur = 1 To m_collPiece.count
70        Call rstImpDP.AddNew
    
75        rstImpDP.Fields("NoPiece") = m_collPiece(iCompteur)
80        rstImpDP.Fields("Qte") = m_collQuantite(iCompteur)
85        rstImpDP.Fields("Description") = m_collDescriptionFR(iCompteur)
90        rstImpDP.Fields("Manufacturier") = m_collManufacturier(iCompteur)
             
95        Call rstImpDP.Update
100     Next
  
105     Call rstImpDP.Requery
  
110     For iCompteur = 15 To rstImpDP.RecordCount Step -1
115       Call rstImpDP.AddNew
    
120       rstImpDP.Fields("NoPiece") = vbNullString
125       rstImpDP.Fields("Qte") = vbNullString
130       rstImpDP.Fields("Description") = vbNullString
135       rstImpDP.Fields("Manufacturier") = vbNullString
    
140       Call rstImpDP.Update
145     Next
      
150     Call rstImpDP.Close
155     Set rstImpDP = Nothing

160     Exit Sub

AfficherErreur:

165     woups "frmChoixDemande", "EnregistrerDemandePrixNouvellesPieces", Err, Erl
End Sub

Private Sub EnregistrerDemandePrix(ByVal iIDFRS As Integer)

5       On Error GoTo AfficherErreur

10      Dim rstImpDP    As ADODB.Recordset
15      Dim rstPieceFRS As ADODB.Recordset
20      Dim rstPiece    As ADODB.Recordset
25      Dim sNomTable   As String
30      Dim sWhere      As String
35      Dim sCategorie  As String
40      Dim iCompteur   As Integer
        
45      If m_eCatalogue = ELECTRIQUE Then
50        sNomTable = "GRB_ImpressionDemandePrixElec"
55      Else
60        sNomTable = "GRB_ImpressionDemandePrixMec"
65      End If

70      Set rstImpDP = New ADODB.Recordset
75      Set rstPiece = New ADODB.Recordset
80      Set rstPieceFRS = New ADODB.Recordset

85      rstImpDP.CursorLocation = adUseClient

90      Call rstImpDP.Open("SELECT * FROM " & sNomTable, g_connData, adOpenDynamic, adLockOptimistic)
  
95      If m_eDemande <> MODE_PIECE Then
100       Select Case m_eDemande
            Case MODE_FOURNISSEUR:
105           Call rstPieceFRS.Open("SELECT * FROM GRB_PiecesFRS WHERE IDFRS = " & iIDFRS & " ORDER BY PIECE", g_connData, adOpenDynamic, adLockOptimistic)
  
110         Case MODE_CATEGORIE:
115           If m_eCatalogue = ELECTRIQUE Then
120             sWhere = " AND GRB_CatalogueElec.CATEGORIE In ("
125           Else
130             sWhere = " AND GRB_CatalogueMec.CATEGORIE In ("
135           End If

140           For iCompteur = 1 To m_collCategorie.count
145             sCategorie = Replace(m_collCategorie(iCompteur), "'", "''")
      
150             If iCompteur <> m_collCategorie.count Then
155               sWhere = sWhere & "'" & sCategorie & "',"
160             Else
165               sWhere = sWhere & "'" & sCategorie & "')"
170             End If
175           Next
  
180           Call rstPieceFRS.Open("SELECT GRB_PiecesFRS.* FROM GRB_PiecesFRS INNER JOIN GRB_CatalogueMec ON GRB_PiecesFRS.PIECE = GRB_CatalogueMec.PIECE WHERE GRB_PiecesFRS.IDFRS = " & iIDFRS & sWhere & " ORDER BY GRB_PiecesFRS.PIECE", g_connData, adOpenDynamic, adLockOptimistic)
185       End Select
  
190       Do While Not rstPieceFRS.EOF
195         If rstPieceFRS.Fields("Type") = "M" Then
200           Call rstPiece.Open("SELECT FABRICANT, DESC_FR, DESC_EN FROM GRB_CatalogueElec WHERE PIECE = '" & Replace(rstPieceFRS.Fields("PIECE"), "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
205         Else
210           Call rstPiece.Open("SELECT FABRICANT, DESC_FR, DESC_EN FROM GRB_CatalogueMec WHERE PIECE = '" & Replace(rstPieceFRS.Fields("PIECE"), "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
215         End If

220         If Not rstPiece.EOF Then
225           If m_collManufacturier.count > 0 Then
230             For iCompteur = 1 To m_collManufacturier.count
235               If m_collManufacturier(iCompteur) = rstPiece.Fields("FABRICANT") Then
240                 Call rstImpDP.AddNew
        
245                 rstImpDP.Fields("NoPiece") = rstPieceFRS.Fields("PIECE")
250                 rstImpDP.Fields("Qte") = "1"
255
260                 If m_sLangage = S_FRANCAIS Then
265                   rstImpDP.Fields("Description") = rstPiece.Fields("DESC_FR")
270                 Else
275                   rstImpDP.Fields("Description") = rstPiece.Fields("DESC_EN")
280                 End If
   
285                 rstImpDP.Fields("Manufacturier") = rstPiece.Fields("FABRICANT")
    
290                 Call rstImpDP.Update
295               End If
300             Next
305           Else
310             Call rstImpDP.AddNew

315             rstImpDP.Fields("NoPiece") = rstPieceFRS.Fields("PIECE")
320             rstImpDP.Fields("Qte") = "1"

325             If m_sLangage = S_FRANCAIS Then
330               rstImpDP.Fields("Description") = rstPiece.Fields("DESC_FR")
335             Else
340               rstImpDP.Fields("Description") = rstPiece.Fields("DESC_EN")
345             End If

350             rstImpDP.Fields("Manufacturier") = rstPiece.Fields("FABRICANT")

355             Call rstImpDP.Update
360           End If
365         Else
370           Call rstPieceFRS.Delete
375         End If

380         Call rstPiece.Close
        
385         Call rstPieceFRS.MoveNext
390       Loop

395       Set rstPiece = Nothing
    
400       Call rstPieceFRS.Close
405       Set rstPieceFRS = Nothing
410     Else
415       If m_eCatalogue = ELECTRIQUE Then
420         Call rstPieceFRS.Open("SELECT GRB_PiecesFRS.PIECE FROM GRB_PiecesFRS INNER JOIN GRB_TempDP ON GRB_PiecesFRS.PIECE = GRB_TempDP.PIECE WHERE GRB_PiecesFRS.IDFRS = " & iIDFRS & " AND GRB_TempDP.TYPE = 'E' GROUP BY GRB_PiecesFRS.PIECE, ORDRE ORDER BY GRB_TempDP.ORDRE", g_connData, adOpenDynamic, adLockOptimistic)
425       Else
430         Call rstPieceFRS.Open("SELECT GRB_PiecesFRS.PIECE FROM GRB_PiecesFRS INNER JOIN GRB_TempDP ON GRB_PiecesFRS.PIECE = GRB_TempDP.PIECE WHERE GRB_PiecesFRS.IDFRS = " & iIDFRS & " AND GRB_TempDP.TYPE = 'M' GROUP BY GRB_PiecesFRS.PIECE, ORDRE ORDER BY GRB_TempDP.ORDRE", g_connData, adOpenDynamic, adLockOptimistic)
435       End If

440       Do While Not rstPieceFRS.EOF
445         For iCompteur = 1 To m_collPiece.count
450           If m_collPiece(iCompteur) = Trim(rstPieceFRS.Fields("PIECE")) Then
455             Call rstImpDP.AddNew

460             rstImpDP.Fields("NoPiece") = m_collPiece(iCompteur)

465             rstImpDP.Fields("Qte") = m_collQuantite(iCompteur)

470             If m_sLangage = S_FRANCAIS Then
475               rstImpDP.Fields("Description") = m_collDescriptionFR(iCompteur)
480             Else
485               rstImpDP.Fields("Description") = m_collDescriptionEN(iCompteur)
490             End If
          
495             rstImpDP.Fields("Manufacturier") = m_collManufacturier(iCompteur)

500             Call rstImpDP.Update

505             Exit For
510           End If
515         Next

520         Call rstPieceFRS.MoveNext
525       Loop

530       Call rstPieceFRS.Close
535       Set rstPieceFRS = Nothing
540     End If
      
545     Call rstImpDP.Requery

550     For iCompteur = 15 To (rstImpDP.RecordCount + 1) Step -1
555       Call rstImpDP.AddNew

560       rstImpDP.Fields("NoPiece") = vbNullString
565       rstImpDP.Fields("Qte") = vbNullString
570       rstImpDP.Fields("Description") = vbNullString
575       rstImpDP.Fields("Manufacturier") = vbNullString
  
580       Call rstImpDP.Update
585     Next
     
590     Call rstImpDP.Close
595     Set rstImpDP = Nothing

600     Exit Sub

AfficherErreur:

605     woups "frmChoixDemande", "EnregistrerDemandePrix", Err, Erl
End Sub

Private Sub ImprimerDemandePrix(ByVal iIDFRS As Integer)

5       On Error GoTo AfficherErreur

10      Dim rstImpDP  As ADODB.Recordset
15      Dim rstFRS    As ADODB.Recordset
20      Dim sNomTable As String
  
25      If m_eCatalogue = ELECTRIQUE Then
30        sNomTable = "GRB_ImpressionDemandePrixElec"
35      Else
40        sNomTable = "GRB_ImpressionDemandePrixMec"
45      End If
  
50      Set rstFRS = New ADODB.Recordset
55      Set rstImpDP = New ADODB.Recordset
  
60      Call rstFRS.Open("SELECT * FROM GRB_Fournisseur WHERE IDFRS = " & iIDFRS, g_connData, adOpenDynamic, adLockOptimistic)
  
65      Call rstImpDP.Open("SELECT * FROM " & sNomTable, g_connData, adOpenDynamic, adLockOptimistic)
  
70      Set DR_DemandePrix.DataSource = rstImpDP
  
        'Entête
        'Vérifie si c'est Anglais ou Francais
        'On modifie seulement si c'est Anglais
75      If m_sLangage = S_ANGLAIS Then
80        DR_DemandePrix.Sections("Section2").Controls("lblTitreDemande").Caption = "Price and Delivery Request"
85        DR_DemandePrix.Sections("Section2").Controls("lblTitreFournisseur").Caption = "Supplier :"
90        DR_DemandePrix.Sections("Section2").Controls("lblTitreContact").Caption = "Contact :"
95        DR_DemandePrix.Sections("Section2").Controls("lblTitreTransport").Caption = "Transport :"
100       DR_DemandePrix.Sections("Section2").Controls("lblTitreDateReq").Caption = "Required Date :"
105       DR_DemandePrix.Sections("Section2").Controls("lblTitreNoSoum").Caption = "Your Ref # :"
110       DR_DemandePrix.Sections("Section2").Controls("lblTitreTel").Caption = "Telephone :"
115       DR_DemandePrix.Sections("Section2").Controls("lblTitreFax").Caption = "Fax :"
120       DR_DemandePrix.Sections("Section2").Controls("lblTitreNoGRB").Caption = "OUR # :"
125       DR_DemandePrix.Sections("Section2").Controls("lblTitreDate").Caption = "Date :"
130       DR_DemandePrix.Sections("Section2").Controls("lblTitreComPar").Caption = "Purchaser :"
135       DR_DemandePrix.Sections("Section2").Controls("lblTitrePage").Caption = "Page :"
140       DR_DemandePrix.Sections("Section2").Controls("lblTitreQte").Caption = "Qty"
145       DR_DemandePrix.Sections("Section2").Controls("lblTitrePiece").Caption = "Part Number"
150       DR_DemandePrix.Sections("Section2").Controls("lblTitreDescription").Caption = "Description"
155       DR_DemandePrix.Sections("Section2").Controls("lblTitreManufact").Caption = "Manufacturer"
160       DR_DemandePrix.Sections("Section2").Controls("lblTitrePrix").Caption = "Unit Price"
165       DR_DemandePrix.Sections("Section2").Controls("lblTitreDelais").Caption = "Delay"
170       DR_DemandePrix.Sections("Section3").Controls("lblTitreCommentaire").Caption = "Comments :"
175       DR_DemandePrix.Sections("Section3").Controls("lblPrixValide").Caption = "Valid price for"
180       DR_DemandePrix.Sections("Section3").Controls("lblJours").Caption = "Days"
185       DR_DemandePrix.Sections("Section3").Controls("lblPiedPage").Caption = "THIS IS NOT AN ORDER"

190       DR_DemandePrix.Sections("Section2").Controls("imgLogoFrancais").Visible = False
195       DR_DemandePrix.Sections("Section2").Controls("imgLogoAnglais").Visible = True
200     End If
  
205     DR_DemandePrix.Sections("Section2").Controls("lblFournisseur").Caption = rstFRS.Fields("NomFournisseur")
210     DR_DemandePrix.Sections("Section2").Controls("lblContact").Caption = m_sContact

215     If Not IsNull(rstFRS.Fields("CondTransport")) Then
220       DR_DemandePrix.Sections("Section2").Controls("lblTransport").Caption = rstFRS.Fields("CondTransport")
225     Else
230       DR_DemandePrix.Sections("Section2").Controls("lblTransport").Caption = ""
235     End If

240     DR_DemandePrix.Sections("Section2").Controls("lblDateRequise").Caption = mskDateRequise.Text
245     DR_DemandePrix.Sections("Section2").Controls("lblTel").Caption = rstFRS.Fields("Telephonne")
250     DR_DemandePrix.Sections("Section2").Controls("lblFax").Caption = rstFRS.Fields("Fax")
255     DR_DemandePrix.Sections("Section2").Controls("lblNoGRB").Caption = txtNoGRB.Text
260     DR_DemandePrix.Sections("Section2").Controls("lblDate").Caption = ConvertDate(Date)
265     DR_DemandePrix.Sections("Section2").Controls("lblCommandePar").Caption = g_sEmploye
270     DR_DemandePrix.Sections("Section3").Controls("lblCommentaire").Caption = txtcommentaire.Text

275     DR_DemandePrix.Orientation = rptOrientLandscape
  
280     Call DR_DemandePrix.Show(vbModal)
   
285     Call g_connData.Execute("DELETE * FROM " & sNomTable)
  
290     Call rstImpDP.Close
295     Set rstImpDP = Nothing
  
300     Call rstFRS.Close
305     Set rstFRS = Nothing

310     Exit Sub

AfficherErreur:

315     woups "frmChoixDemande", "ImprimerDemandePrix", Err, Erl
End Sub

Private Sub mskDateRequise_GotFocus()

5       On Error GoTo AfficherErreur

10      If Len(mskDateRequise.Text) = 10 Then
15        mskDateRequise.Text = Right$(mskDateRequise.Text, 8)
20      End If
  
25      mskDateRequise.mask = "##-##-##"

30      Exit Sub

AfficherErreur:

35      woups "frmChoixDemande", "mskDateRequise_GotFocus", Err, Erl
End Sub

Private Sub mskDateRequise_LostFocus()

5       On Error GoTo AfficherErreur

10      mskDateRequise.mask = vbNullString
  
15      If mskDateRequise.Text = "__-__-__" Then
20        mskDateRequise.Text = vbNullString
25      Else
30        If Len(mskDateRequise.Text) = 8 Then
35          If IsDate(mskDateRequise.Text) Then
40            mskDateRequise.Text = Year(DateSerial(Left$(mskDateRequise.Text, 2), Mid$(mskDateRequise.Text, 4, 2), Right$(mskDateRequise.Text, 2))) & Mid$(mskDateRequise.Text, 3, 8)
45          End If
50        End If
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmChoixDemande", "mskDateRequise_LostFocus", Err, Erl
End Sub

Private Sub txtRechercher_Change()

5       On Error GoTo AfficherErreur

10      cmdRechercher.Caption = "Rechercher"

15      Exit Sub

AfficherErreur:

20      woups "frmChoixDemande", "txtRechercher_Change", Err, Erl
End Sub
