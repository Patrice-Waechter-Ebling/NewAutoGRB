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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmChoixCategorie.frx":0000
   ScaleHeight     =   2400
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
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

 On Error GoTo Oups

 If m_eCatalogue = ELECTRIQUE Then
 FrmCatalogueElec.m_bAnnulerCopie = True
 Else
 FrmCatalogueMec.m_bAnnulerCopie = True
 End If

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixCategorie", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdOK_Click()

 On Error GoTo Oups

 If m_eCatalogue = ELECTRIQUE Then
 FrmCatalogueElec.m_bAnnulerCopie = False
 FrmCatalogueElec.m_sCategorieCopie = cmbCategorie.Text
 Else
 FrmCatalogueMec.m_bAnnulerCopie = False
 FrmCatalogueMec.m_sCategorieCopie = cmbCategorie.Text
 End If
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixCategorie", "cmdOK_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Call RemplirComboCategorie

 Exit Sub

Oups:

 wOups "frmChoixCategorie", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboCategorie()

 On Error GoTo Oups
 
 'Remplir le combo des catégories
 Dim rstCategorie As ADODB.Recordset
 
 'Il faut vider le combo avant de le remplir
 Call cmbCategorie.Clear
 
 Set rstCategorie = New ADODB.Recordset
 
 If m_eCatalogue = ELECTRIQUE Then
 Call rstCategorie.Open("SELECT DISTINCT CATEGORIE FROM GrbCatalogueElec ORDER BY CATEGORIE", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstCategorie.Open("SELECT DISTINCT CATEGORIE FROM GrbCatalogueMec ORDER BY CATEGORIE", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstCategorie.EOF
 Call cmbCategorie.AddItem(rstCategorie.Fields("CATEGORIE"))
 
  Call rstCategorie.MoveNext
  Loop
 
  Call rstCategorie.Close
  Set rstCategorie = Nothing
 
 'Si le combo n'est pas vide, on sélectionne le premier
  If cmbCategorie.ListCount > 0 Then
  cmbCategorie.ListIndex = 0
  End If

  Exit Sub

Oups:

10 wOups "frmChoixCategorie", "RemplirComboCategorie", Err, Err.number, Err.Description
End Sub

Public Sub Afficher(ByVal eCatalogue As enumCatalogue)
 
 On Error GoTo Oups

 m_eCatalogue = eCatalogue

 Call Me.Show(vbModal)

 Exit Sub

Oups:

 wOups "frmChoixCategorie", "Afficher", Err, Err.number, Err.Description
End Sub
