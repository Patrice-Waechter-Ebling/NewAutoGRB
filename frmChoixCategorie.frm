VERSION 5.00
Begin VB.Form frmChoixCategorie 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChoixCategorie.frx":0000
   ScaleHeight     =   2400
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbCategorie 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Text            =   "cmbCategorie"
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dans quelle catégorie ?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "frmChoixCategorie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_eCatalogue As enumCatalogue

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      If m_eCatalogue = ELECTRIQUE Then
15        FrmCatalogueElec.m_bAnnulerCopie = True
20      Else
25        FrmCatalogueMec.m_bAnnulerCopie = True
30      End If

35      Call Unload(Me)

40      Exit Sub

AfficherErreur:

45      woups "frmChoixCategorie", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur

10      If m_eCatalogue = ELECTRIQUE Then
15        FrmCatalogueElec.m_bAnnulerCopie = False
20        FrmCatalogueElec.m_sCategorieCopie = cmbCategorie.Text
25      Else
30        FrmCatalogueMec.m_bAnnulerCopie = False
35        FrmCatalogueMec.m_sCategorieCopie = cmbCategorie.Text
40      End If
  
45      Call Unload(Me)

50      Exit Sub

AfficherErreur:

55      woups "frmChoixCategorie", "cmdOK_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Call RemplirComboCategorie

15      Exit Sub

AfficherErreur:

20      woups "frmChoixCategorie", "Form_Load", Err, Erl
End Sub

Private Sub RemplirComboCategorie()

5       On Error GoTo AfficherErreur
        
        'Remplir le combo des catégories
10      Dim rstCategorie As ADODB.Recordset
  
        'Il faut vider le combo avant de le remplir
15      Call cmbCategorie.Clear
     
20      Set rstCategorie = New ADODB.Recordset
     
25      If m_eCatalogue = ELECTRIQUE Then
30        Call rstCategorie.Open("SELECT DISTINCT CATEGORIE FROM GRB_CatalogueElec ORDER BY CATEGORIE", g_connData, adOpenDynamic, adLockOptimistic)
35      Else
40        Call rstCategorie.Open("SELECT DISTINCT CATEGORIE FROM GRB_CatalogueMec ORDER BY CATEGORIE", g_connData, adOpenDynamic, adLockOptimistic)
45      End If
  
        'Tant que ce n'est pas la fin des enregistrements
50      Do While Not rstCategorie.EOF
55        Call cmbCategorie.AddItem(rstCategorie.Fields("CATEGORIE"))
          
60        Call rstCategorie.MoveNext
65      Loop
  
70      Call rstCategorie.Close
75      Set rstCategorie = Nothing
  
        'Si le combo n'est pas vide, on sélectionne le premier
80      If cmbCategorie.ListCount > 0 Then
85        cmbCategorie.ListIndex = 0
90      End If

95      Exit Sub

AfficherErreur:

100     woups "frmChoixCategorie", "RemplirComboCategorie", Err, Erl
End Sub

Public Sub Afficher(ByVal eCatalogue As enumCatalogue)
        
5       On Error GoTo AfficherErreur

10      m_eCatalogue = eCatalogue

15      Call Me.Show(vbModal)

20      Exit Sub

AfficherErreur:

25      woups "frmChoixCategorie", "Afficher", Err, Erl
End Sub
