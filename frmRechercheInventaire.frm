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
   Picture         =   "frmRechercheInventaire.frx":0000
   ScaleHeight     =   6555
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
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

Private Const I_CMB_NO_ITEM      As Integer = 0
Private Const I_CMB_FABRICANT    As Integer = 1
Private Const I_CMB_DESCRIPTION  As Integer = 2
Private Const I_CMB_CATEGORIE    As Integer = 3

Private Const I_COL_QTE          As Integer = 0
Private Const I_COL_NO_ITEM      As Integer = 1
Private Const I_COL_FABRICANT    As Integer = 2
Private Const I_COL_DESCRIPTION  As Integer = 3
Private Const I_COL_CATEGORIE    As Integer = 4
Private Const I_COL_LOCALISATION As Integer = 5
Private Const I_COL_PRIX_LIST    As Integer = 6
Private Const I_COL_ESCOMPTE     As Integer = 7
Private Const I_COL_PRIX_NET     As Integer = 8
Private Const I_COL_TOTAL        As Integer = 9

Private m_eCatalogue As enumCatalogue

Private Sub RemplirComboCategories()

5       On Error GoTo AfficherErreur

        'Remplir le combo des catégories
10      Dim rstCategorie  As ADODB.Recordset
15      Dim sNomCategorie As String
        
20      Set rstCategorie = New ADODB.Recordset
        
25      If m_eCatalogue = ELECTRIQUE Then
30        Call rstCategorie.Open("SELECT DISTINCT Categorie FROM GRB_CatalogueElec", g_connData, adOpenDynamic, adLockOptimistic)
35      Else
40        Call rstCategorie.Open("SELECT DISTINCT Categorie FROM GRB_CatalogueMec", g_connData, adOpenDynamic, adLockOptimistic)
45      End If
    
50      Do While Not rstCategorie.EOF
55        sNomCategorie = rstCategorie.Fields("Categorie")
      
60        Call cmbCategorie.AddItem(sNomCategorie)
    
65        Call rstCategorie.MoveNext
70      Loop
    
75      Call rstCategorie.Close
80      Set rstCategorie = Nothing

85      Exit Sub

AfficherErreur:

90      woups "frmRechercheInventaire", "RemplirComboCategories", Err, Erl
End Sub

Public Sub Afficher(ByVal eCatalogue As enumCatalogue)

5       On Error GoTo AfficherErreur

10      m_eCatalogue = eCatalogue
  
15      Call RemplirComboCategories

20      cmbRecherche.ListIndex = I_CMB_NO_ITEM

25      Call Me.Show

30      Exit Sub

AfficherErreur:

35      woups "frmRechercheInventaire", "Afficher", Err, Erl
End Sub

Private Sub cmbRecherche_Click()

5       On Error GoTo AfficherErreur

10      If cmbRecherche.ListIndex = I_CMB_CATEGORIE Then
15        txtRecherche.Visible = False
20        cmbCategorie.Visible = True
25        lblTitreRecherche.Caption = "Catégorie à rechercher"
30      Else
35        cmbCategorie.Visible = False
40        txtRecherche.Visible = True
45        lblTitreRecherche.Caption = "Texte à rechercher"
50      End If

55      Exit Sub

AfficherErreur:

60      woups "frmRechercheInventaire", "cmbRecherche_Click", Err, Erl
End Sub

Private Sub cmdAfficher_Click()

5       On Error GoTo AfficherErreur

10      Call RemplirListView

15      Exit Sub

AfficherErreur:

20      woups "frmRechercheInventaire", "cmdAfficher_Click", Err, Erl
End Sub

Private Sub RemplirListView()

5       On Error GoTo AfficherErreur

10      Dim sWhere       As String
15      Dim rstInv       As ADODB.Recordset
20      Dim rstCatalogue As ADODB.Recordset
25      Dim iCompteur    As Integer

30      If (cmbRecherche.ListIndex <> I_CMB_CATEGORIE And txtRecherche.Text <> "") Or (cmbRecherche.ListIndex = I_CMB_CATEGORIE And cmbCategorie.ListIndex > -1) Then
35        Select Case cmbRecherche.ListIndex
            Case I_CMB_NO_ITEM:
40            sWhere = "INSTR(1, PIECE, '" & txtRecherche.Text & "') > 0"
        
45          Case I_CMB_FABRICANT:
50            sWhere = "INSTR(1, FABRICANT, '" & txtRecherche.Text & "') > 0"
         
55          Case I_CMB_DESCRIPTION:
60            sWhere = "INSTR(1, DESC_FR, '" & txtRecherche.Text & "') > 0"
          
65          Case I_CMB_CATEGORIE:
70            sWhere = "INSTR(1, CATEGORIE, '" & Replace(cmbCategorie.Text, "'", "''") & "') > 0"
75        End Select
      
80        Set rstCatalogue = New ADODB.Recordset
85        Set rstInv = New ADODB.Recordset
   
90        Call lvwInventaire.ListItems.Clear

95        lblTotal.Caption = ""
  
100       Screen.MousePointer = vbHourglass
  
105       If m_eCatalogue = ELECTRIQUE Then
110         Call rstCatalogue.Open("SELECT * FROM GRB_CatalogueElec WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
115       Else
120         Call rstCatalogue.Open("SELECT * FROM GRB_CatalogueMec WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
125       End If

130       Do While Not rstCatalogue.EOF
135         If m_eCatalogue = ELECTRIQUE Then
140           Call rstInv.Open("SELECT * FROM GRB_InventaireElec WHERE NoItem = '" & Replace(rstCatalogue.Fields("PIECE"), "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
145         Else
150           Call rstInv.Open("SELECT * FROM GRB_InventaireMec WHERE NoItem = '" & Replace(rstCatalogue.Fields("PIECE"), "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
155         End If

160         If Not rstInv.EOF Then
165           Call AjouterItemListView(rstInv, rstCatalogue.Fields("CATEGORIE"))
170         End If
         
175         Call rstInv.Close
         
180         Call rstCatalogue.MoveNext

185         DoEvents
190       Loop
      
195       Call CalculerTotal
      
200       Call rstCatalogue.Close
      
205       Set rstCatalogue = Nothing
210       Set rstInv = Nothing

215       Screen.MousePointer = vbDefault
220     Else
225       If cmbRecherche.ListIndex = I_CMB_CATEGORIE Then
230         Call MsgBox("La catégorie à rechercher est obligatoire!", vbOKOnly, "Erreur")
235       Else
240         Call MsgBox("Le texte à rechercher ne doit pas être vide!", vbOKOnly, "Erreur")
245       End If
250     End If
 
255     Exit Sub

AfficherErreur:

260     woups "frmRechercheInventaire", "RemplirListView", Err, Erl
End Sub

Private Sub AjouterItemListView(ByVal rstInv As ADODB.Recordset, ByVal sCategorie As String)

5       On Error GoTo AfficherErreur

10      Dim itmInv As ListItem

15      Set itmInv = lvwInventaire.ListItems.Add

20      itmInv.Text = rstInv.Fields("QuantitéStock")
25      itmInv.SubItems(I_COL_NO_ITEM) = rstInv.Fields("NoItem")
30      itmInv.SubItems(I_COL_FABRICANT) = rstInv.Fields("Manufacturier")
35      itmInv.SubItems(I_COL_DESCRIPTION) = rstInv.Fields("Description")
40      itmInv.SubItems(I_COL_CATEGORIE) = sCategorie
45      itmInv.SubItems(I_COL_LOCALISATION) = rstInv.Fields("Localisation")
50      itmInv.SubItems(I_COL_PRIX_LIST) = Conversion(rstInv.Fields("Prix Liste"), MODE_ARGENT, 4)
55      itmInv.SubItems(I_COL_ESCOMPTE) = Conversion(rstInv.Fields("Escompte"), MODE_POURCENT)
60      itmInv.SubItems(I_COL_PRIX_NET) = Conversion(rstInv.Fields("Prix net"), MODE_ARGENT, 4)
65      itmInv.SubItems(I_COL_TOTAL) = Conversion(Replace(rstInv.Fields("Prix net"), ".", ",") * Replace(rstInv.Fields("QuantitéStock"), ".", ","), MODE_ARGENT)

70      Exit Sub

AfficherErreur:

75      woups "frmRechercheInventaire", "AjouterItemListView", Err, Erl
End Sub

Private Sub CalculerTotal()

5       On Error GoTo AfficherErreur

10      Dim dblTotal  As Double
15      Dim iCompteur As Integer

20      For iCompteur = 1 To lvwInventaire.ListItems.count
25        dblTotal = dblTotal + lvwInventaire.ListItems(iCompteur).SubItems(I_COL_TOTAL)
30      Next

35      lblTotal.Caption = Conversion(dblTotal, MODE_ARGENT)

40      Exit Sub

AfficherErreur:

45      woups "frmRechercheInventaire", "CalculerTotal", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmRechercheInventaire", "cmdFermer_Click", Err, Erl
End Sub
