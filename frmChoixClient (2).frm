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
   MDIChild        =   -1  'True
   ScaleHeight     =   5070
   ScaleWidth      =   3600
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

 On Error GoTo Oups

 If cmbclient.ListIndex <> -1 Then
 If Trim$(txtDescription.Text) <> vbNullString Then
 frmFacturation.m_iIDClient = cmbclient.ItemData(cmbclient.ListIndex)
 frmFacturation.m_sDescription = txtDescription.Text
 
 Call Unload(Me)
 Else
 Call MsgBox("La description est obligatoire!", vbOKOnly, "Erreur")
 End If
 Else
 Call MsgBox("Le client est obligatoire!", vbOKOnly, "Erreur")
  End If

  Exit Sub

Oups:

  wOups "frmChoixClient", "cmdOK_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRafraichir_Click()

 On Error GoTo Oups

 Call RemplirComboClient(vbNullString)

 Exit Sub

Oups:

 wOups "frmChoixClient", "cmdRafraichir_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRecherche_Click()

 On Error GoTo Oups

 Call RemplirComboClient(txtRecherche.Text)

 Exit Sub

Oups:

 wOups "frmChoixClient", "cmdRecherche_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim sIDClient As String

 Call RemplirComboClient(vbNullString)

 If frmFacturation.m_bModifClient = True Then
 sIDClient = frmFacturation.txtClient.Tag

 For iCompteur = 0 To cmbclient.ListCount - 1
 If cmbclient.ItemData(iCompteur) = sIDClient Then
 cmbclient.ListIndex = iCompteur

 Exit For
 End If
  Next

  txtDescription.Text = frmFacturation.txtDescription.Text
  End If

  Exit Sub

Oups:

  wOups "frmChoixClient", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboClient(ByVal sRecherche As String)

 On Error GoTo Oups

 'Remplir le combo des clients
 Dim rstClient As ADODB.Recordset
 
 'Il faut vider le combo avant de le remplir
 Call cmbclient.Clear
 
 Set rstClient = New ADODB.Recordset
 
 If Trim$(sRecherche) = vbNullString Then
 Call rstClient.Open("SELECT NomClient, IDClient FROM GrbClient WHERE Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 If InStr(1, sRecherche, "'") > 0 Then
 Call rstClient.Open("SELECT NomClient, IDClient FROM GrbClient WHERE NomClient LIKE '%" & Replace(sRecherche, "'", "''") & "%' AND Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstClient.Open("SELECT NomClient, IDClient FROM GrbClient WHERE INSTR(1, NomClient, '" & sRecherche & "') > 0 AND Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)
  End If
  End If
 
 'Tant que ce n'est pas la fin des enregistrements
  Do While Not rstClient.EOF
  Call cmbclient.AddItem(rstClient.Fields("NomClient"))
 
  cmbclient.ItemData(cmbclient.newIndex) = rstClient.Fields("IDClient")
 
  Call rstClient.MoveNext
  Loop
 
  Call rstClient.Close
10 Set rstClient = Nothing
 
 'Si le combo n'est pas vide, on sélectionne le premier
If cmbclient.ListCount > 0 Then
 cmbclient.ListIndex = 0
End If

Exit Sub

Oups:

wOups "frmChoixClient", "RemplirComboClient", Err, Err.number, Err.Description
End Sub
