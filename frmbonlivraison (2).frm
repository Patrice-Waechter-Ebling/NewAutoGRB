VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmbonlivraison 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bon livraison"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   7965
   Begin VB.CommandButton CmdQuit 
      Caption         =   "&Fermer"
      Height          =   495
      Left            =   3480
      TabIndex        =   16
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton CmdSupp 
      Caption         =   "&Supprimer tout"
      Height          =   495
      Left            =   1680
      TabIndex        =   14
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Ajouter"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Frame fraqte 
      Caption         =   "QTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   7455
      Begin VB.TextBox txtManufacturier 
         Height          =   285
         Left            =   4320
         TabIndex        =   8
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtQteBo 
         Height          =   285
         Left            =   3000
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtQteCom 
         Height          =   285
         Left            =   480
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sauvegarde"
         Height          =   495
         Left            =   2640
         TabIndex        =   11
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton cmdFermerQte 
         Caption         =   "Fermer"
         Height          =   495
         Left            =   4440
         TabIndex        =   12
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtQteLivr 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtDescription 
         Height          =   525
         Left            =   480
         TabIndex        =   10
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Manufacturier"
         Height          =   255
         Left            =   4320
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Qte bo"
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label label2 
         Caption         =   "Qte com"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Qte livr"
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Description"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
   End
   Begin MSComctlLib.ListView lvwBonLivraison 
      Height          =   2175
      Left            =   0
      TabIndex        =   13
      Top             =   1200
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   3836
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
         Text            =   "qte com"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "qte livr"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "qte bo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "description"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "manufacturier"
         Object.Width           =   2646
      EndProperty
   End
End
Attribute VB_Name = "frmbonlivraison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const I_COL_COMMANDE As Integer = 0
Private Const I_COL_LIVRAISON As Integer = 1
Private Const I_COL_BACK_ORDER As Integer = 2
Private Const I_COL_DESCRIPTION As Integer = 3
Private Const I_COL_MANUFACTURIER As Integer = 4

Private m_bModeAjouter As Boolean

Private Sub RemplirListView()

 On Error GoTo Oups
 
 'rempli le ListView
 Dim rstImpression As ADODB.Recordset
 Dim itmImpression As ListItem

 CmdAdd.Visible = True
 CmdSupp.Visible = True
 
 'vide lister
 Call lvwBonLivraison.ListItems.Clear
 
 lvwBonLivraison.Sorted = False
 
 'ouvre la table pour client
 Set rstImpression = New ADODB.Recordset
 
 Call rstImpression.Open("SELECT * FROM Grbimpression_bonlivraison WHERE user = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

 'tant que pas a la fin de la table
 Do While Not rstImpression.EOF
 If Not IsNull(rstImpression.Fields("qte_com")) Or _
 Not IsNull(rstImpression.Fields("qte_livr")) Or _
 Not IsNull(rstImpression.Fields("qte_bo")) Or _
 Not IsNull(rstImpression.Fields("Description")) Or _
 Not IsNull(rstImpression.Fields("Manufacturier")) Then
 'ajoute au lister
  Set itmImpression = lvwBonLivraison.ListItems.Add
 
 'no du client
  itmImpression.Tag = rstImpression.Fields("no")
 
 'qte_com
  If Not IsNull(rstImpression.Fields("qte_com")) Then
  itmImpression.Text = rstImpression.Fields("qte_com")
  Else
  itmImpression.Text = vbNullString
  End If
 
 'qte_livr
  If Not IsNull(rstImpression.Fields("qte_livr")) Then
 itmImpression.SubItems(I_COL_LIVRAISON) = rstImpression.Fields("qte_livr")
Else
 itmImpression.SubItems(I_COL_LIVRAISON) = vbNullString
 End If
 
 'qte_bo
 If Not IsNull(rstImpression.Fields("qte_bo")) Then
 itmImpression.SubItems(I_COL_BACK_ORDER) = rstImpression.Fields("qte_bo")
 Else
 itmImpression.SubItems(I_COL_BACK_ORDER) = vbNullString
 End If
 
 'description
 If Not IsNull(rstImpression.Fields("Description")) Then
 itmImpression.SubItems(I_COL_DESCRIPTION) = rstImpression.Fields("Description")
 Else
 itmImpression.SubItems(I_COL_DESCRIPTION) = vbNullString
 End If
 
 'manufacturier
 If Not IsNull(rstImpression.Fields("manufacturier")) Then
 itmImpression.SubItems(I_COL_MANUFACTURIER) = rstImpression.Fields("Manufacturier")
 Else
 itmImpression.SubItems(I_COL_MANUFACTURIER) = vbNullString
 End If
1  Else
 Call rstImpression.Delete
 End If
 
 'prochaine enreg
 Call rstImpression.MoveNext
Loop
 
 'fermeture table
Call rstImpression.Close
Set rstImpression = Nothing

Exit Sub

Oups:

wOups "frmbonlivraison", "RemplirListView", Err, Err.number, Err.Description
End Sub

Private Sub CmdAdd_Click()

 On Error GoTo Oups
 'ajoute une qte
 'met visible fenetre pour ajouter
 fraqte.Visible = True

 'mode ajoouter ou editer
 m_bModeAjouter = True

 'valeur par defaut sur l'ouverture
 txtQteCom.Text = vbNullString
 txtQteLivr.Text = vbNullString
 txtQteBo.Text = vbNullString
 txtDescription.Text = vbNullString
 txtManufacturier.Text = vbNullString
 
 Call txtQteCom.SetFocus

 Exit Sub
 
Oups:

 wOups "frmbonlivraison", "CmdAdd_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdfermerqte_Click()

 On Error GoTo Oups
 'quitte liste qte
 'cache fenetre
 fraqte.Visible = False
 
 Call RemplirListView

 Exit Sub

Oups:

 wOups "frmbonlivraison", "cmdfermerqte_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdquit_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmbonlivraison", "cmdquit_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdsave_Click()

 On Error GoTo Oups
 
 'pour sauver l'enregistrement
 Dim rstImpression As ADODB.Recordset
 Dim iNoIndex As Integer
 
 Set rstImpression = New ADODB.Recordset
 
 'si le mode est ajouter
 If m_bModeAjouter = True Then
 'table impression ouvert
 Call rstImpression.Open("SELECT * FROM Grbimpression_bonlivraison WHERE user = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstImpression.EOF = False Then
 If rstImpression.RecordCount >= 10 Then
 With rstImpression
 Do While Not .EOF
 If IsNull(.Fields("qte_com")) And IsNull(.Fields("qte_livr")) And IsNull(.Fields("qte_bo")) And IsNull(.Fields("description")) And IsNull(.Fields("manufacturier")) Then
  iNoIndex = .Fields("No")
 
  Exit Do
  End If
 
  Call .MoveNext
  Loop
  End With
 
  If iNoIndex = 0 Then
  iNoIndex = rstImpression.RecordCount + 1
 
 Call rstImpression.AddNew
 End If
 Else
 rstImpression.MoveLast
 
 iNoIndex = rstImpression.Fields("No") + 1
 
 Call rstImpression.AddNew
 End If
 Else
 iNoIndex = 1
 
 Call rstImpression.AddNew
 End If
 
 rstImpression.Fields("no") = iNoIndex
1  Else
 Call rstImpression.Open("SELECT * FROM GrbImpression_BonLivraison WHERE user = '" & g_sUserID & "' AND [No] = " & lvwBonLivraison.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
 End If
 
If txtQteCom = vbNullString Then
 rstImpression.Fields("qte_com") = Null
Else
 rstImpression.Fields("qte_com") = txtQteCom.Text
1  End If
 
 If txtQteLivr = vbNullString Then
 rstImpression.Fields("qte_livr") = Null
Else
 rstImpression.Fields("qte_livr") = txtQteLivr.Text
End If
 
If txtQteBo = vbNullString Then
 rstImpression.Fields("qte_bo") = Null
Else
 rstImpression.Fields("qte_bo") = txtQteBo.Text
End If
 
If txtDescription = vbNullString Then
 rstImpression.Fields("Description") = Null
2  Else
 rstImpression.Fields("Description") = txtDescription.Text
2  End If
 
If txtManufacturier = vbNullString Then
rstImpression.Fields("manufacturier") = Null
Else
rstImpression.Fields("manufacturier") = txtManufacturier.Text
End If
 
30 rstImpression.Fields("user") = g_sUserID
 
Call rstImpression.Update
 
 'ferme la table
Call rstImpression.Close
Set rstImpression = Nothing
 
 'cache la petite fenetre
fraqte.Visible = False
 
 'rempli le lister
Call RemplirListView

Exit Sub

Oups:

wOups "frmbonlivraison", "cmdsave_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdSupp_Click()

 On Error GoTo Oups
 '################################################
 'supprime l'enregistrement sélectionné
 Call g_connData.Execute("DELETE * FROM Grbimpression_bonlivraison")
 
 'initialise le lister
 Call RemplirListView

 Exit Sub

Oups:

 wOups "frmbonlivraison", "CmdSupp_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups
 
 Screen.MousePointer = vbDefault
 
 'Remplir lister
 Call RemplirListView

 Exit Sub

Oups:

 wOups "frmbonlivraison", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)

 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim iNbreItem As Integer
 Dim rstImpression As ADODB.Recordset

 'ouvre les tables
 Set rstImpression = New ADODB.Recordset
 
 Call rstImpression.Open("SELECT * FROM Grbimpression_bonlivraison WHERE user = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

 iNbreItem = lvwBonLivraison.ListItems.count
 
 If iNbreItem > 0 Then
 For iCompteur = iNbreItem + 1 To 10
 Call rstImpression.AddNew
 
 rstImpression.Fields("no") = iCompteur
  rstImpression.Fields("user") = g_sUserID
 
  Call rstImpression.Update
  Next
  End If
 
  Call rstImpression.Close
  Set rstImpression = Nothing

  Exit Sub

Oups:

  wOups "frmbonlivraison", "Form_Unload", Err, Err.number, Err.Description
End Sub

Private Sub lvwBonLivraison_DblClick()

 On Error GoTo Oups
 'sur dbclick, affiche fenetre pour modifié l'enreg selectionné dans lister

 'si lister pas vide
 If lvwBonLivraison.ListItems.count <> 0 Then
 'met fenetre visible
 fraqte.Visible = True
 
 txtQteCom.Text = lvwBonLivraison.SelectedItem.Text
 txtQteLivr.Text = lvwBonLivraison.SelectedItem.SubItems(I_COL_LIVRAISON)
 txtQteBo.Text = lvwBonLivraison.SelectedItem.SubItems(I_COL_BACK_ORDER)
 txtDescription.Text = lvwBonLivraison.SelectedItem.SubItems(I_COL_DESCRIPTION)
 txtManufacturier.Text = lvwBonLivraison.SelectedItem.SubItems(I_COL_MANUFACTURIER)
 
 'met en mode edition
 m_bModeAjouter = False
 End If

 Exit Sub

Oups:

  wOups "frmbonlivraison", "lvwBonLivraison_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwBonLivraison_KeyDown(KeyCode As Integer, Shift As Integer)

 On Error GoTo Oups

 If lvwBonLivraison.ListItems.count > 0 Then
 If KeyCode = vbKeyDelete Then
 Call g_connData.Execute("DELETE * FROM GrbImpression_BonLivraison WHERE [no] = " & lvwBonLivraison.SelectedItem.Tag & " AND User = '" & g_sUserID & "'")
 
 Call CorrigerNumeros
 
 Call RemplirListView
 End If
 End If

 Exit Sub

Oups:

 wOups "frmbonlivraison", "lvwBonLivraison_KeyDown", Err, Err.number, Err.Description
End Sub

Private Sub CorrigerNumeros()

 On Error GoTo Oups

 Dim rstNo As ADODB.Recordset
 Dim iNo As Integer
 
 Set rstNo = New ADODB.Recordset
 
 Call rstNo.Open("SELECT * FROM GrbImpression_BonLivraison WHERE user = '" & g_sUserID & "' ORDER BY [no]", g_connData, adOpenDynamic, adLockOptimistic)
 
 iNo = 1
 
 Do While Not rstNo.EOF
 rstNo.Fields("No") = iNo
 
 iNo = iNo + 1
 
 Call rstNo.Update
 
 Call rstNo.MoveNext
  Loop
 
  Call rstNo.Close
  Set rstNo = Nothing

  Exit Sub

Oups:

  wOups "frmbonlivraison", "CorrigerNumeros", Err, Err.number, Err.Description
End Sub
