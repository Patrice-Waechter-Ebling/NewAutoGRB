VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRechercheInventaire 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recherche d'inventaire"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmRechercheInventaire.frx":0000
   ScaleHeight     =   6555
   ScaleWidth      =   9255
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   7800
      TabIndex        =   9
      Top             =   6120
      Width           =   1335
   End
   Begin VB.ComboBox cmbCategorie 
      Height          =   315
      ItemData        =   "frmRechercheInventaire.frx":2F0D
      Left            =   1440
      List            =   "frmRechercheInventaire.frx":2F0F
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   3855
   End
   Begin VB.CommandButton cmdAfficher 
      Caption         =   "Afficher"
      Default         =   -1  'True
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   1020
      Width           =   1335
   End
   Begin VB.ComboBox cmbRecherche 
      Height          =   315
      ItemData        =   "frmRechercheInventaire.frx":2F11
      Left            =   5520
      List            =   "frmRechercheInventaire.frx":2F21
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtRecherche 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   3855
   End
   Begin MSComctlLib.ListView lvwInventaire 
      Height          =   4215
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7435
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
         Text            =   "Qté"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No.Item"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fabricant"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Catégorie"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Localisation"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Prix listé"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Escompte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Prix net"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8160
      TabIndex        =   8
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL  :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   7
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Rechercher dans :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblTitreRecherche 
      BackStyle       =   0  'Transparent
      Caption         =   "Texte à rechercher :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "frmRechercheInventaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const I_CMB_NO_ITEM As Integer = 0
Private Const I_CMB_FABRICANT As Integer = 1
Private Const I_CMB_DESCRIPTION As Integer = 2
Private Const I_CMB_CATEGORIE As Integer = 3

Private Const I_COL_QTE As Integer = 0
Private Const I_COL_NO_ITEM As Integer = 1
Private Const I_COL_FABRICANT As Integer = 2
Private Const I_COL_DESCRIPTION As Integer = 3
Private Const I_COL_CATEGORIE As Integer = 4
Private Const I_COL_LOCALISATION As Integer = 5
Private Const I_COL_PRIX_LIST As Integer = 6
Private Const I_COL_ESCOMPTE As Integer = 7
Private Const I_COL_PRIX_NET As Integer = 8
Private Const I_COL_TOTAL As Integer = 9

Private m_eCatalogue As enumCatalogue

Private Sub RemplirComboCategories()

 On Error GoTo Oups

 'Remplir le combo des catégories
 Dim rstCategorie As ADODB.Recordset
 Dim sNomCategorie As String
 
 Set rstCategorie = New ADODB.Recordset
 
 If m_eCatalogue = ELECTRIQUE Then
 Call rstCategorie.Open("SELECT DISTINCT Categorie FROM GrbCatalogueElec", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstCategorie.Open("SELECT DISTINCT Categorie FROM GrbCatalogueMec", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 
 Do While Not rstCategorie.EOF
 sNomCategorie = rstCategorie.Fields("Categorie")
 
  Call cmbCategorie.AddItem(sNomCategorie)
 
  Call rstCategorie.MoveNext
  Loop
 
  Call rstCategorie.Close
  Set rstCategorie = Nothing

  Exit Sub

Oups:

  wOups "frmRechercheInventaire", "RemplirComboCategories", Err, Err.number, Err.Description
End Sub

Public Sub Afficher(ByVal eCatalogue As enumCatalogue)

 On Error GoTo Oups

 m_eCatalogue = eCatalogue
 
 Call RemplirComboCategories

 cmbRecherche.ListIndex = I_CMB_NO_ITEM

 Call Me.Show

 Exit Sub

Oups:

 wOups "frmRechercheInventaire", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub cmbRecherche_Click()

 On Error GoTo Oups

 If cmbRecherche.ListIndex = I_CMB_CATEGORIE Then
 txtRecherche.Visible = False
 cmbCategorie.Visible = True
 lblTitreRecherche.Caption = "Catégorie à rechercher"
 Else
 cmbCategorie.Visible = False
 txtRecherche.Visible = True
 lblTitreRecherche.Caption = "Texte à rechercher"
 End If

 Exit Sub

Oups:

  wOups "frmRechercheInventaire", "cmbRecherche_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAfficher_Click()

 On Error GoTo Oups

 Call RemplirListView

 Exit Sub

Oups:

 wOups "frmRechercheInventaire", "cmdAfficher_Click", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListView()

 On Error GoTo Oups

 Dim sWhere As String
 Dim rstInv As ADODB.Recordset
 Dim rstCatalogue As ADODB.Recordset
 Dim iCompteur As Integer

 If (cmbRecherche.ListIndex <> I_CMB_CATEGORIE And txtRecherche.Text <> "") Or (cmbRecherche.ListIndex = I_CMB_CATEGORIE And cmbCategorie.ListIndex > -1) Then
 Select Case cmbRecherche.ListIndex
 Case I_CMB_NO_ITEM:
 sWhere = "INSTR(1, PIECE, '" & txtRecherche.Text & "') > 0"
 
 Case I_CMB_FABRICANT:
 sWhere = "INSTR(1, FABRICANT, '" & txtRecherche.Text & "') > 0"
 
 Case I_CMB_DESCRIPTION:
  sWhere = "INSTR(1, DESC_FR, '" & txtRecherche.Text & "') > 0"
 
  Case I_CMB_CATEGORIE:
  sWhere = "INSTR(1, CATEGORIE, '" & Replace(cmbCategorie.Text, "'", "''") & "') > 0"
  End Select
 
  Set rstCatalogue = New ADODB.Recordset
  Set rstInv = New ADODB.Recordset
 
  Call lvwInventaire.ListItems.Clear

  lblTotal.Caption = ""
 
Screen.MousePointer = vbHourglass
 
1 If m_eCatalogue = ELECTRIQUE Then
 Call rstCatalogue.Open("SELECT * FROM GrbCatalogueElec WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstCatalogue.Open("SELECT * FROM GrbCatalogueMec WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
 End If

 Do While Not rstCatalogue.EOF
 If m_eCatalogue = ELECTRIQUE Then
 Call rstInv.Open("SELECT * FROM GrbInventaireElec WHERE NoItem = '" & Replace(rstCatalogue.Fields("PIECE"), "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstInv.Open("SELECT * FROM GrbInventaireMec WHERE NoItem = '" & Replace(rstCatalogue.Fields("PIECE"), "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If

 If Not rstInv.EOF Then
 Call AjouterItemListView(rstInv, rstCatalogue.Fields("CATEGORIE"))
 End If
 
 Call rstInv.Close
 
 Call rstCatalogue.MoveNext

 DoEvents
 Loop
 
1  Call CalculerTotal
 
 Call rstCatalogue.Close
 
 Set rstCatalogue = Nothing
 Set rstInv = Nothing

 Screen.MousePointer = vbDefault
Else
 If cmbRecherche.ListIndex = I_CMB_CATEGORIE Then
 Call MsgBox("La catégorie à rechercher est obligatoire!", vbOKOnly, "Erreur")
 Else
 Call MsgBox("Le texte à rechercher ne doit pas être vide!", vbOKOnly, "Erreur")
 End If
End If
 
Exit Sub

Oups:

2  wOups "frmRechercheInventaire", "RemplirListView", Err, Err.number, Err.Description
End Sub

Private Sub AjouterItemListView(ByVal rstInv As ADODB.Recordset, ByVal sCategorie As String)

 On Error GoTo Oups

 Dim itmInv As ListItem

 Set itmInv = lvwInventaire.ListItems.Add

 itmInv.Text = rstInv.Fields("QuantitéStock")
 itmInv.SubItems(I_COL_NO_ITEM) = rstInv.Fields("NoItem")
 itmInv.SubItems(I_COL_FABRICANT) = rstInv.Fields("Manufacturier")
 itmInv.SubItems(I_COL_DESCRIPTION) = rstInv.Fields("Description")
 itmInv.SubItems(I_COL_CATEGORIE) = sCategorie
 itmInv.SubItems(I_COL_LOCALISATION) = rstInv.Fields("Localisation")
 itmInv.SubItems(I_COL_PRIX_LIST) = Conversion(rstInv.Fields("Prix Liste"), MODE_ARGENT, 4)
 itmInv.SubItems(I_COL_ESCOMPTE) = Conversion(rstInv.Fields("Escompte"), MODE_POURCENT)
  itmInv.SubItems(I_COL_PRIX_NET) = Conversion(rstInv.Fields("Prix net"), MODE_ARGENT, 4)
  itmInv.SubItems(I_COL_TOTAL) = Conversion(Replace(rstInv.Fields("Prix net"), ".", ",") * Replace(rstInv.Fields("QuantitéStock"), ".", ","), MODE_ARGENT)

  Exit Sub

Oups:

  wOups "frmRechercheInventaire", "AjouterItemListView", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTotal()

 On Error GoTo Oups

 Dim dblTotal As Double
 Dim iCompteur As Integer

 For iCompteur = 1 To lvwInventaire.ListItems.count
 dblTotal = dblTotal + lvwInventaire.ListItems(iCompteur).SubItems(I_COL_TOTAL)
 Next

 lblTotal.Caption = Conversion(dblTotal, MODE_ARGENT)

 Exit Sub

Oups:

 wOups "frmRechercheInventaire", "CalculerTotal", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmRechercheInventaire", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub
