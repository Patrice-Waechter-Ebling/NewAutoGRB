VERSION 5.00
Begin VB.Form frmChoixClient 
   BackColor       =   &H00000000&
   Caption         =   "Choix du client"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmChoixClient.frx":0000
   ScaleHeight     =   5070
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescription 
      Height          =   645
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3720
      Width           =   3375
   End
   Begin VB.Frame fraRecherche 
      BackColor       =   &H00000000&
      Caption         =   "Recherche"
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
      Height          =   1215
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
      Begin VB.CommandButton cmdRafraichir 
         Caption         =   "Rafraichir"
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdRecherche 
         Caption         =   "OK"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtRecherche 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.ComboBox cmbClient 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2760
      Width           =   3375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pour quel client ?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmChoixClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur

10      If cmbclient.ListIndex <> -1 Then
15        If Trim$(txtDescription.Text) <> vbNullString Then
20          frmFacturation.m_iIDClient = cmbclient.ItemData(cmbclient.ListIndex)
25          frmFacturation.m_sDescription = txtDescription.Text
  
30          Call Unload(Me)
35        Else
40          Call MsgBox("La description est obligatoire!", vbOKOnly, "Erreur")
45        End If
50      Else
55        Call MsgBox("Le client est obligatoire!", vbOKOnly, "Erreur")
60      End If

65      Exit Sub

AfficherErreur:

70      woups "frmChoixClient", "cmdOK_Click", Err, Erl
End Sub

Private Sub cmdRafraichir_Click()

5       On Error GoTo AfficherErreur

10      Call RemplirComboClient(vbNullString)

15      Exit Sub

AfficherErreur:

20      woups "frmChoixClient", "cmdRafraichir_Click", Err, Erl
End Sub

Private Sub cmdRecherche_Click()

5       On Error GoTo AfficherErreur

10      Call RemplirComboClient(txtRecherche.Text)

15      Exit Sub

AfficherErreur:

20      woups "frmChoixClient", "cmdRecherche_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
15      Dim sIDClient As String

20      Call RemplirComboClient(vbNullString)

25      If frmFacturation.m_bModifClient = True Then
30        sIDClient = frmFacturation.txtClient.Tag

35        For iCompteur = 0 To cmbclient.ListCount - 1
40          If cmbclient.ItemData(iCompteur) = sIDClient Then
45            cmbclient.ListIndex = iCompteur

50            Exit For
55          End If
60        Next

65        txtDescription.Text = frmFacturation.txtDescription.Text
70      End If

75      Exit Sub

AfficherErreur:

80      woups "frmChoixClient", "Form_Load", Err, Erl
End Sub

Private Sub RemplirComboClient(ByVal sRecherche As String)

5       On Error GoTo AfficherErreur

        'Remplir le combo des clients
10      Dim rstClient As ADODB.Recordset
    
        'Il faut vider le combo avant de le remplir
15      Call cmbclient.Clear
      
20      Set rstClient = New ADODB.Recordset
      
25      If Trim$(sRecherche) = vbNullString Then
30        Call rstClient.Open("SELECT NomClient, IDClient FROM GRB_Client WHERE Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)
35      Else
40        If InStr(1, sRecherche, "'") > 0 Then
45          Call rstClient.Open("SELECT NomClient, IDClient FROM GRB_Client WHERE NomClient LIKE '%" & Replace(sRecherche, "'", "''") & "%' AND Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)
50        Else
55          Call rstClient.Open("SELECT NomClient, IDClient FROM GRB_Client WHERE INSTR(1, NomClient, '" & sRecherche & "') > 0 AND Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)
60        End If
65      End If
    
        'Tant que ce n'est pas la fin des enregistrements
70      Do While Not rstClient.EOF
75        Call cmbclient.AddItem(rstClient.Fields("NomClient"))
        
80        cmbclient.ItemData(cmbclient.newIndex) = rstClient.Fields("IDClient")
        
85        Call rstClient.MoveNext
90      Loop
  
95      Call rstClient.Close
100     Set rstClient = Nothing
  
        'Si le combo n'est pas vide, on sélectionne le premier
105     If cmbclient.ListCount > 0 Then
110       cmbclient.ListIndex = 0
115     End If

120     Exit Sub

AfficherErreur:

125     woups "frmChoixClient", "RemplirComboClient", Err, Erl
End Sub
