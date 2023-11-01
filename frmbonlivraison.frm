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

Private Const I_COL_COMMANDE      As Integer = 0
Private Const I_COL_LIVRAISON     As Integer = 1
Private Const I_COL_BACK_ORDER    As Integer = 2
Private Const I_COL_DESCRIPTION   As Integer = 3
Private Const I_COL_MANUFACTURIER As Integer = 4

Private m_bModeAjouter As Boolean

Private Sub RemplirListView()

5       On Error GoTo AfficherErreur
        
        'rempli le ListView
10      Dim rstImpression As ADODB.Recordset
15      Dim itmImpression As ListItem

20      CmdAdd.Visible = True
25      CmdSupp.Visible = True
  
        'vide lister
30      Call lvwBonLivraison.ListItems.Clear
  
35      lvwBonLivraison.Sorted = False
  
        'ouvre la table pour client
40      Set rstImpression = New ADODB.Recordset
        
45      Call rstImpression.Open("SELECT * FROM GRB_impression_bonlivraison WHERE user = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

        'tant que pas a la fin de la table
50      Do While Not rstImpression.EOF
55        If Not IsNull(rstImpression.Fields("qte_com")) Or _
             Not IsNull(rstImpression.Fields("qte_livr")) Or _
             Not IsNull(rstImpression.Fields("qte_bo")) Or _
             Not IsNull(rstImpression.Fields("Description")) Or _
             Not IsNull(rstImpression.Fields("Manufacturier")) Then
            'ajoute au lister
60          Set itmImpression = lvwBonLivraison.ListItems.Add
        
            'no du client
65          itmImpression.Tag = rstImpression.Fields("no")
        
            'qte_com
70          If Not IsNull(rstImpression.Fields("qte_com")) Then
75            itmImpression.Text = rstImpression.Fields("qte_com")
80          Else
85            itmImpression.Text = vbNullString
90          End If
        
            'qte_livr
95          If Not IsNull(rstImpression.Fields("qte_livr")) Then
100           itmImpression.SubItems(I_COL_LIVRAISON) = rstImpression.Fields("qte_livr")
105         Else
110           itmImpression.SubItems(I_COL_LIVRAISON) = vbNullString
115         End If
        
            'qte_bo
120         If Not IsNull(rstImpression.Fields("qte_bo")) Then
125           itmImpression.SubItems(I_COL_BACK_ORDER) = rstImpression.Fields("qte_bo")
130         Else
135           itmImpression.SubItems(I_COL_BACK_ORDER) = vbNullString
140         End If
        
            'description
145         If Not IsNull(rstImpression.Fields("Description")) Then
150           itmImpression.SubItems(I_COL_DESCRIPTION) = rstImpression.Fields("Description")
155         Else
160           itmImpression.SubItems(I_COL_DESCRIPTION) = vbNullString
165         End If
        
            'manufacturier
170         If Not IsNull(rstImpression.Fields("manufacturier")) Then
175           itmImpression.SubItems(I_COL_MANUFACTURIER) = rstImpression.Fields("Manufacturier")
180         Else
185           itmImpression.SubItems(I_COL_MANUFACTURIER) = vbNullString
190         End If
195       Else
200         Call rstImpression.Delete
205       End If
      
          'prochaine enreg
210       Call rstImpression.MoveNext
215     Loop
      
        'fermeture table
220     Call rstImpression.Close
225     Set rstImpression = Nothing

230     Exit Sub

AfficherErreur:

235     woups "frmbonlivraison", "RemplirListView", Err, Erl
End Sub

Private Sub CmdAdd_Click()

5       On Error GoTo AfficherErreur
              'ajoute une qte
              'met visible fenetre pour ajouter
10      fraqte.Visible = True

              'mode ajoouter ou editer
15      m_bModeAjouter = True

              'valeur par defaut sur l'ouverture
20      txtQteCom.Text = vbNullString
25      txtQteLivr.Text = vbNullString
30      txtQteBo.Text = vbNullString
35      txtDescription.Text = vbNullString
40      txtManufacturier.Text = vbNullString
  
45      Call txtQteCom.SetFocus

50      Exit Sub
 
AfficherErreur:

55      woups "frmbonlivraison", "CmdAdd_Click", Err, Erl
End Sub

Private Sub cmdfermerqte_Click()

5       On Error GoTo AfficherErreur
              'quitte liste qte
              'cache fenetre
10      fraqte.Visible = False
  
15      Call RemplirListView

20      Exit Sub

AfficherErreur:

25      woups "frmbonlivraison", "cmdfermerqte_Click", Err, Erl
End Sub

Private Sub cmdquit_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmbonlivraison", "cmdquit_Click", Err, Erl
End Sub

Private Sub cmdsave_Click()

5       On Error GoTo AfficherErreur
        
        'pour sauver l'enregistrement
10      Dim rstImpression As ADODB.Recordset
15      Dim iNoIndex      As Integer
  
20      Set rstImpression = New ADODB.Recordset
  
        'si le mode est ajouter
25      If m_bModeAjouter = True Then
          'table impression ouvert
30        Call rstImpression.Open("SELECT * FROM GRB_impression_bonlivraison WHERE user = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
35        If rstImpression.EOF = False Then
40          If rstImpression.RecordCount >= 10 Then
45            With rstImpression
50              Do While Not .EOF
55                If IsNull(.Fields("qte_com")) And IsNull(.Fields("qte_livr")) And IsNull(.Fields("qte_bo")) And IsNull(.Fields("description")) And IsNull(.Fields("manufacturier")) Then
60                  iNoIndex = .Fields("No")
            
65                  Exit Do
70                End If
            
75                Call .MoveNext
80              Loop
85            End With
                
90            If iNoIndex = 0 Then
95              iNoIndex = rstImpression.RecordCount + 1
          
100             Call rstImpression.AddNew
105           End If
110         Else
115           rstImpression.MoveLast
        
120           iNoIndex = rstImpression.Fields("No") + 1
        
125           Call rstImpression.AddNew
130         End If
135       Else
140         iNoIndex = 1
      
145         Call rstImpression.AddNew
150       End If
    
155       rstImpression.Fields("no") = iNoIndex
160     Else
165       Call rstImpression.Open("SELECT * FROM GRB_Impression_BonLivraison WHERE user = '" & g_sUserID & "' AND [No] = " & lvwBonLivraison.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
170     End If
      
175     If txtQteCom = vbNullString Then
180       rstImpression.Fields("qte_com") = Null
185     Else
190       rstImpression.Fields("qte_com") = txtQteCom.Text
195     End If
    
200     If txtQteLivr = vbNullString Then
205       rstImpression.Fields("qte_livr") = Null
210     Else
215       rstImpression.Fields("qte_livr") = txtQteLivr.Text
220     End If
      
225     If txtQteBo = vbNullString Then
230       rstImpression.Fields("qte_bo") = Null
235     Else
240       rstImpression.Fields("qte_bo") = txtQteBo.Text
245     End If
      
250     If txtDescription = vbNullString Then
255       rstImpression.Fields("Description") = Null
260     Else
265       rstImpression.Fields("Description") = txtDescription.Text
270     End If
      
275     If txtManufacturier = vbNullString Then
280       rstImpression.Fields("manufacturier") = Null
285     Else
290       rstImpression.Fields("manufacturier") = txtManufacturier.Text
295     End If
    
300     rstImpression.Fields("user") = g_sUserID
    
305     Call rstImpression.Update
      
        'ferme la table
310     Call rstImpression.Close
315     Set rstImpression = Nothing
  
        'cache la petite fenetre
320     fraqte.Visible = False
  
        'rempli le lister
325     Call RemplirListView

330     Exit Sub

AfficherErreur:

335     woups "frmbonlivraison", "cmdsave_Click", Err, Erl
End Sub

Private Sub CmdSupp_Click()

5       On Error GoTo AfficherErreur
              '################################################
              'supprime l'enregistrement sélectionné
10      Call g_connData.Execute("DELETE * FROM GRB_impression_bonlivraison")
  
              'initialise le lister
15      Call RemplirListView

20      Exit Sub

AfficherErreur:

25      woups "frmbonlivraison", "CmdSupp_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur
        
10      Screen.MousePointer = vbDefault
        
        'Remplir lister
15      Call RemplirListView

20      Exit Sub

AfficherErreur:

25      woups "frmbonlivraison", "Form_Load", Err, Erl
End Sub

Private Sub Form_Unload(Cancel As Integer)

5       On Error GoTo AfficherErreur

10      Dim iCompteur     As Integer
15      Dim iNbreItem     As Integer
20      Dim rstImpression As ADODB.Recordset

        'ouvre les tables
25      Set rstImpression = New ADODB.Recordset
        
30      Call rstImpression.Open("SELECT * FROM GRB_impression_bonlivraison WHERE user = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

35      iNbreItem = lvwBonLivraison.ListItems.count
  
40      If iNbreItem > 0 Then
45        For iCompteur = iNbreItem + 1 To 10
50          Call rstImpression.AddNew
      
55          rstImpression.Fields("no") = iCompteur
60          rstImpression.Fields("user") = g_sUserID
    
65          Call rstImpression.Update
70        Next
75      End If
  
80      Call rstImpression.Close
85      Set rstImpression = Nothing

90      Exit Sub

AfficherErreur:

95      woups "frmbonlivraison", "Form_Unload", Err, Erl
End Sub

Private Sub lvwBonLivraison_DblClick()

5       On Error GoTo AfficherErreur
              'sur dbclick, affiche fenetre pour modifié l'enreg selectionné dans lister

              'si lister pas vide
10      If lvwBonLivraison.ListItems.count <> 0 Then
                'met fenetre visible
15        fraqte.Visible = True
        
20        txtQteCom.Text = lvwBonLivraison.SelectedItem.Text
25        txtQteLivr.Text = lvwBonLivraison.SelectedItem.SubItems(I_COL_LIVRAISON)
30        txtQteBo.Text = lvwBonLivraison.SelectedItem.SubItems(I_COL_BACK_ORDER)
35        txtDescription.Text = lvwBonLivraison.SelectedItem.SubItems(I_COL_DESCRIPTION)
40        txtManufacturier.Text = lvwBonLivraison.SelectedItem.SubItems(I_COL_MANUFACTURIER)
    
                'met en mode edition
45        m_bModeAjouter = False
50      End If

55      Exit Sub

AfficherErreur:

60      woups "frmbonlivraison", "lvwBonLivraison_DblClick", Err, Erl
End Sub

Private Sub lvwBonLivraison_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      If lvwBonLivraison.ListItems.count > 0 Then
15        If KeyCode = vbKeyDelete Then
20          Call g_connData.Execute("DELETE * FROM GRB_Impression_BonLivraison WHERE [no] = " & lvwBonLivraison.SelectedItem.Tag & " AND User = '" & g_sUserID & "'")
      
25          Call CorrigerNumeros
      
30          Call RemplirListView
35        End If
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmbonlivraison", "lvwBonLivraison_KeyDown", Err, Erl
End Sub

Private Sub CorrigerNumeros()

5       On Error GoTo AfficherErreur

10      Dim rstNo As ADODB.Recordset
15      Dim iNo   As Integer
  
20      Set rstNo = New ADODB.Recordset
  
25      Call rstNo.Open("SELECT * FROM GRB_Impression_BonLivraison WHERE user = '" & g_sUserID & "' ORDER BY [no]", g_connData, adOpenDynamic, adLockOptimistic)
  
30      iNo = 1
  
35      Do While Not rstNo.EOF
40        rstNo.Fields("No") = iNo
    
45        iNo = iNo + 1
    
50        Call rstNo.Update
    
55        Call rstNo.MoveNext
60      Loop
  
65      Call rstNo.Close
70      Set rstNo = Nothing

75      Exit Sub

AfficherErreur:

80      woups "frmbonlivraison", "CorrigerNumeros", Err, Erl
End Sub
