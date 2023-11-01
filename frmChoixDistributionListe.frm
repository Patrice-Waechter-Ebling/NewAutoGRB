VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChoixMailList 
   BackColor       =   &H00000000&
   Caption         =   "Choix de la liste de distribution"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmChoixDistributionListe.frx":0000
   ScaleHeight     =   4065
   ScaleWidth      =   6690
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvwDistList 
      Height          =   1695
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Liste"
         Object.Width           =   4392
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre de contacts"
         Object.Width           =   2858
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Dossier"
         Object.Width           =   3175
      EndProperty
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dans quelle liste de distribution voulez-vous l'ajouter ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   6495
   End
End
Attribute VB_Name = "frmChoixMailList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const I_LVW_LISTE   As Integer = 0
Private Const I_LVW_NOMBRE  As Integer = 1
Private Const I_LVW_DOSSIER As Integer = 2

Private m_frmSource As Form
Private m_otlApp    As Outlook.Application

Public Sub Afficher(ByVal frmSource As Form, ByVal otlApp As Outlook.Application)

5       On Error GoTo AfficherErreur

10      Set m_frmSource = frmSource
  
15      Set m_otlApp = otlApp
  
20      Call Me.Show(vbModal)

25      Exit Sub

AfficherErreur:

30      woups "frmChoixMailList", "Afficher", Err, Erl
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      m_frmSource.m_bAnnulerDistList = True

15      Call Unload(Me)

20      Exit Sub

AfficherErreur:

25      woups "frmChoixMailList", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur

10      Dim objItem As Object
15      Dim folGRB  As Outlook.MAPIFolder
20      Dim myItems As Outlook.Items

25      If lvwDistList.ListItems.count > 0 Then
30        m_frmSource.m_bAnnulerDistList = False

35        Set folGRB = GetFolder(m_otlApp, lvwDistList.SelectedItem.SubItems(I_LVW_DOSSIER))

40        Set myItems = folGRB.Items.Restrict("[MessageClass] = 'IPM.DistList'")

45        For Each objItem In folGRB.Items
50          If objItem.Class = olDistributionList Then
55            If objItem = lvwDistList.SelectedItem.Text Then
60              Set m_frmSource.m_otlDistList = objItem

65              Exit For
70            End If
75          End If
80        Next

85        Call Unload(Me)
90      Else
95        Call MsgBox("Il n'y a aucune liste de distribution!", vbOKOnly, "Erreur")
100     End If

105     Exit Sub

AfficherErreur:

110     woups "frmChoixMailList", "cmdOK_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Dim folGRB  As Outlook.MAPIFolder
15      Dim myItems As Outlook.Items
20      Dim objItem As Object
25      Dim itmDL   As ListItem

30      Set folGRB = GetFolder(m_otlApp, "Contacts GRB")

35      Set myItems = folGRB.Items.Restrict("[MessageClass] = 'IPM.DistList'")

40      For Each objItem In myItems
45        Set itmDL = lvwDistList.ListItems.Add

50        itmDL.Text = objItem
55        itmDL.SubItems(I_LVW_NOMBRE) = objItem.MemberCount
60        itmDL.SubItems(I_LVW_DOSSIER) = "Contacts GRB"
65      Next

70      Set folGRB = GetFolder(m_otlApp, "Clients GRB")

75      Set myItems = folGRB.Items.Restrict("[MessageClass] = 'IPM.DistList'")

80      For Each objItem In myItems
85        Set itmDL = lvwDistList.ListItems.Add

90        itmDL.Text = objItem
95        itmDL.SubItems(I_LVW_NOMBRE) = objItem.MemberCount
100       itmDL.SubItems(I_LVW_DOSSIER) = "Clients GRB"
105      Next

110     Set folGRB = GetFolder(m_otlApp, "Fournisseurs GRB")

115     Set myItems = folGRB.Items.Restrict("[MessageClass] = 'IPM.DistList'")

120     For Each objItem In myItems
125       Set itmDL = lvwDistList.ListItems.Add

130       itmDL.Text = objItem
135       itmDL.SubItems(I_LVW_NOMBRE) = objItem.MemberCount
140       itmDL.SubItems(I_LVW_DOSSIER) = "Fournisseurs GRB"
145     Next

150     m_frmSource.fraEtatOutlook.Visible = False

155     Exit Sub

AfficherErreur:

160     woups "frmChoixMailList", "Form_Load", Err, Erl
End Sub
