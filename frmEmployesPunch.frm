VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmployesPunch 
   BackColor       =   &H00000000&
   Caption         =   "Catégories"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmEmployesPunch.frx":0000
   ScaleHeight     =   4845
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwEmployes 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employé"
         Object.Width           =   11112
      EndProperty
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   4200
      Width           =   1335
   End
End
Attribute VB_Name = "frmEmployesPunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFermer_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      Set frmPunch.m_collEmployes = New Collection

20      For iCompteur = 1 To lvwEmployes.ListItems.Count
25        If lvwEmployes.ListItems(iCompteur).Checked = True Then
30          Call frmPunch.m_collEmployes.Add(lvwEmployes.ListItems(iCompteur).Tag)
35        End If
40      Next

45      Call Unload(Me)

50      Exit Sub

AfficherErreur:

55      Call AfficherErreur(Me, "cmdFermer_Click", Err, Erl)
End Sub

Public Sub AfficherEmployes(ByVal iNoEmploye As Integer)

5       On Error GoTo AfficherErreur

10      Dim rstEmployePunch As ADODB.Recordset
15      Dim itmEmploye      As ListItem
20      Dim iCompteurLVW    As Integer
25      Dim iCompteurCOLL   As Integer

30      Set rstEmployePunch = New ADODB.Recordset

35      Call rstEmployePunch.Open("SELECT GRB_Employés.NoEmploye, GRB_Employés.Employe FROM GRB_Employés INNER JOIN GRB_AutorisationPunch ON GRB_Employés.NoEmploye = GRB_AutorisationPunch.NoEmploye WHERE GRB_AutorisationPunch.AutoriserPar = " & iNoEmploye & " ORDER BY GRB_Employés.Employe", g_connData, adOpenDynamic, adLockOptimistic)

40      Do While Not rstEmployePunch.EOF
45        Set itmEmploye = lvwEmployes.ListItems.Add

50        itmEmploye.Text = rstEmployePunch.Fields("Employe")

55        itmEmploye.Tag = rstEmployePunch.Fields("NoEmploye")

60        Call rstEmployePunch.MoveNext
65      Loop
           
70      Call rstEmployePunch.Close
75      Set rstEmployePunch = Nothing
           
80      If Not frmPunch.m_collEmployes Is Nothing Then
85        For iCompteurCOLL = 1 To frmPunch.m_collEmployes.Count
90          For iCompteurLVW = 1 To lvwEmployes.ListItems.Count
95            If lvwEmployes.ListItems(iCompteurLVW).Tag = frmPunch.m_collEmployes(iCompteurCOLL) Then
100             lvwEmployes.ListItems(iCompteurLVW).Checked = True

105             Exit For
110           End If
115         Next
120       Next
125     End If
           
130     Call Me.Show(vbModal)

135     Exit Sub

AfficherErreur:

140     Call AfficherErreur(Me, "AfficherEmployes", Err, Erl)
End Sub
