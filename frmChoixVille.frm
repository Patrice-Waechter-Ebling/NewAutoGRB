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
   MinButton       =   0   'False
   Picture         =   "frmChoixVille.frx":0000
   ScaleHeight     =   2400
   ScaleWidth      =   3885
   StartUpPosition =   2  'CenterScreen
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

5       On Error GoTo AfficherErreur

10      FrmClient.m_bAnnulerVille = True
15      FrmClient.m_sVille = ""

20      Call Unload(Me)

25      Exit Sub

AfficherErreur:

30      woups "frmChoixVille", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur

10      FrmClient.m_bAnnulerVille = False
15      FrmClient.m_sVille = cmbVille.Text
  
20      Call Unload(Me)

25      Exit Sub

AfficherErreur:

30      woups "frmChoixVille", "cmdOK_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Call RemplirComboVille

15      Exit Sub

AfficherErreur:

20      woups "frmChoixVille", "Form_Load", Err, Erl
End Sub

Private Sub RemplirComboVille()

5       On Error GoTo AfficherErreur
        
        'Remplir le combo des catégories
10      Dim rstVille As ADODB.Recordset
  
        'Il faut vider le combo avant de le remplir
15      Call cmbVille.Clear
     
20      Set rstVille = New ADODB.Recordset
     
25      Call rstVille.Open("SELECT DISTINCT VilleLiv FROM GRB_Client ORDER BY VilleLiv", g_connData, adOpenDynamic, adLockOptimistic)
  
        'Tant que ce n'est pas la fin des enregistrements
50      Do While Not rstVille.EOF
55        If Not IsNull(rstVille.Fields("VilleLiv")) Then
60          If Trim(rstVille.Fields("VilleLiv")) <> "" Then
65            Call cmbVille.AddItem(rstVille.Fields("VilleLiv"))
70          End If
75        End If
          
80        Call rstVille.MoveNext
85      Loop
  
90      Call rstVille.Close
95      Set rstVille = Nothing
  
        'Si le combo n'est pas vide, on sélectionne le premier
100     If cmbVille.ListCount > 0 Then
105       cmbVille.ListIndex = 0
110     End If

115     Exit Sub

AfficherErreur:

120     woups "frmChoixVille", "RemplirComboVille", Err, Erl
End Sub
