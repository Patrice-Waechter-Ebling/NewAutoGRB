VERSION 5.00
Begin VB.Form frmChoixVille 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sélection de la ville"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   3885
   Begin VB.ComboBox cmbVille 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Text            =   "cmbVille"
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
      Caption         =   "Choisissez la ville"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "frmChoixVille"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 FrmClient.m_bAnnulerVille = True
 FrmClient.m_sVille = ""

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixVille", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdOK_Click()

 On Error GoTo Oups

 FrmClient.m_bAnnulerVille = False
 FrmClient.m_sVille = cmbVille.Text
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixVille", "cmdOK_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Call RemplirComboVille

 Exit Sub

Oups:

 wOups "frmChoixVille", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboVille()

 On Error GoTo Oups
 
 'Remplir le combo des catégories
 Dim rstVille As ADODB.Recordset
 
 'Il faut vider le combo avant de le remplir
 Call cmbVille.Clear
 
 Set rstVille = New ADODB.Recordset
 
 Call rstVille.Open("SELECT DISTINCT VilleLiv FROM GrbClient ORDER BY VilleLiv", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstVille.EOF
 If Not IsNull(rstVille.Fields("VilleLiv")) Then
  If Trim(rstVille.Fields("VilleLiv")) <> "" Then
  Call cmbVille.AddItem(rstVille.Fields("VilleLiv"))
  End If
  End If
 
  Call rstVille.MoveNext
  Loop
 
  Call rstVille.Close
  Set rstVille = Nothing
 
 'Si le combo n'est pas vide, on sélectionne le premier
10 If cmbVille.ListCount > 0 Then
1 cmbVille.ListIndex = 0
End If

Exit Sub

Oups:

wOups "frmChoixVille", "RemplirComboVille", Err, Err.number, Err.Description
End Sub
