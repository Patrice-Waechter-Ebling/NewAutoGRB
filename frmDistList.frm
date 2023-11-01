VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDistList 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listes de distribution"
   ClientHeight    =   11370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmDistList.frx":0000
   ScaleHeight     =   11370
   ScaleWidth      =   13125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCreerListe 
      Caption         =   "Recréer les listes"
      Height          =   375
      Left            =   10920
      TabIndex        =   4
      Top             =   10800
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Contacts"
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
      Height          =   6015
      Left            =   240
      TabIndex        =   1
      Top             =   4560
      Width           =   12735
      Begin MSComctlLib.ListView lvwContacts 
         Height          =   5535
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   9763
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
            Text            =   "Nom"
            Object.Width           =   10769
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Courriel"
            Object.Width           =   6932
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Liste"
            Object.Width           =   3545
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Listes"
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
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   11775
      Begin VB.CommandButton cmdAfficher 
         Caption         =   "Afficher"
         Height          =   375
         Left            =   9480
         TabIndex        =   6
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton cmdRafraichir 
         Caption         =   "Rafraîchir"
         Height          =   375
         Left            =   9480
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin MSComctlLib.ListView lvwDistList 
         Height          =   2655
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
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
            Text            =   "Nom de la liste"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Dossier Outlook"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre de membres"
            Object.Width           =   4410
         EndProperty
      End
   End
End
Attribute VB_Name = "frmDistList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const I_COL_DISTLIST_NOM     As Integer = 0
Private Const I_COL_DISTLIST_DOSSIER As Integer = 1
Private Const I_COL_DISTLIST_NBRE    As Integer = 2

Private Const I_COL_CONTACT_NOM      As Integer = 0
Private Const I_COL_CONTACT_COURRIEL As Integer = 1
Private Const I_COL_CONTACT_DISTLIST As Integer = 2

Private Const I_ITM_TOUTES_LISTES    As Integer = 1

Private m_otlApp      As Outlook.Application
Private m_bDejaOuvert As Boolean

Private m_folClients  As MAPIFolder
Private m_folContacts As MAPIFolder
Private m_folFRS      As MAPIFolder

Private Sub RemplirListViewContacts()

5       On Error GoTo AfficherErreur

10      Dim otlDistList As Outlook.DistListItem
15      Dim myRecipient As Outlook.Recipient
20      Dim itmContact  As ListItem
25      Dim sDistList   As String
30      Dim iCompteur   As Integer
35      Dim iListe      As Integer
40      Dim bAfficher   As Boolean

45      Call lvwContacts.ListItems.Clear

50      If lvwDistList.ListItems.count > 0 Then
55        Screen.MousePointer = vbHourglass

60        lvwDistList.Enabled = False
65        cmdCreerListe.Enabled = False
70        cmdRafraichir.Enabled = False
75        cmdAfficher.Enabled = False
          
80        For iListe = 2 To lvwDistList.ListItems.count
85          bAfficher = False

90          If lvwDistList.ListItems(I_ITM_TOUTES_LISTES).Selected = True Then
95            bAfficher = True
100         Else
105           If lvwDistList.ListItems(iListe).Selected = True Then
110             bAfficher = True
115           End If
120         End If

125         If bAfficher = True Then
130           sDistList = lvwDistList.ListItems(iListe).Text

135           Set otlDistList = m_otlApp.GetNamespace("MAPI").GetItemFromID(lvwDistList.ListItems(iListe).Tag)

140           For iCompteur = 1 To otlDistList.MemberCount
145             Set myRecipient = otlDistList.GetMember(iCompteur)

150             Set itmContact = lvwContacts.ListItems.Add

155             itmContact.Text = myRecipient.Name
160             itmContact.SubItems(I_COL_CONTACT_COURRIEL) = myRecipient.Address
165             itmContact.SubItems(I_COL_CONTACT_DISTLIST) = sDistList
170           Next

175           DoEvents
180         End If
185       Next

190       lvwDistList.Enabled = True
195       cmdCreerListe.Enabled = True
200       cmdRafraichir.Enabled = True
205       cmdAfficher.Enabled = True

210       Screen.MousePointer = vbDefault
215     End If

220     Exit Sub

AfficherErreur:

225     woups "frmDistList", "RemplirListViewContacts", Err, Erl
End Sub

Private Sub RemplirListViewDistList()

5       On Error GoTo AfficherErreur

10      Dim folGRB   As Outlook.MAPIFolder
15      Dim myItems  As Outlook.Items
20      Dim objItem  As Object
25      Dim itmDL    As ListItem
30      Dim dblTotal As Double

35      Call lvwDistList.ListItems.Clear

40      Set itmDL = lvwDistList.ListItems.Add

45      itmDL.Text = "(Toutes les listes)"

50      Set folGRB = m_folContacts

55      Set myItems = folGRB.Items.Restrict("[MessageClass] = 'IPM.DistList'")

60      For Each objItem In myItems
65        Set itmDL = lvwDistList.ListItems.Add

70        itmDL.Tag = objItem.EntryID
75        itmDL.Text = objItem.DLName
80        itmDL.SubItems(I_COL_DISTLIST_DOSSIER) = "Contacts GRB"
85        itmDL.SubItems(I_COL_DISTLIST_NBRE) = objItem.MemberCount

90        dblTotal = dblTotal + objItem.MemberCount
95      Next

100     Set folGRB = m_folClients

105     Set myItems = folGRB.Items.Restrict("[MessageClass] = 'IPM.DistList'")

110     For Each objItem In myItems
115       Set itmDL = lvwDistList.ListItems.Add

120       itmDL.Tag = objItem.EntryID
125       itmDL.Text = objItem.DLName
130       itmDL.SubItems(I_COL_DISTLIST_DOSSIER) = "Clients GRB"
135       itmDL.SubItems(I_COL_DISTLIST_NBRE) = objItem.MemberCount

140       dblTotal = dblTotal + objItem.MemberCount
145      Next

150     Set folGRB = m_folFRS

155     Set myItems = folGRB.Items.Restrict("[MessageClass] = 'IPM.DistList'")

160     For Each objItem In myItems
165       Set itmDL = lvwDistList.ListItems.Add

170       itmDL.Tag = objItem.EntryID
175       itmDL.Text = objItem.DLName
180       itmDL.SubItems(I_COL_DISTLIST_DOSSIER) = "Fournisseurs GRB"
185       itmDL.SubItems(I_COL_DISTLIST_NBRE) = objItem.MemberCount

190       dblTotal = dblTotal + objItem.MemberCount
195     Next

200     lvwDistList.ListItems(1).SubItems(I_COL_DISTLIST_NBRE) = dblTotal

205     Exit Sub

AfficherErreur:

210     woups "frmDistList", "RemplirListViewDistList", Err, Erl
End Sub

Private Sub cmdAfficher_Click()

5       On Error GoTo AfficherErreur

10      Call RemplirListViewContacts

15      Exit Sub

AfficherErreur:

20      woups "frmDistList", "cmdAfficher_Click", Err, Erl
End Sub

Private Sub cmdCreerListe_Click()

5       On Error GoTo AfficherErreur

10      Call OuvrirForm(frmAjoutDL, False)

15      Exit Sub

AfficherErreur:

20      woups "frmDistList", "cmdCreerListe_Click", Err, Erl
End Sub

Private Sub cmdRafraichir_Click()

5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbHourglass

15      Call RemplirListViewDistList

20      Screen.MousePointer = vbDefault

25      Exit Sub

AfficherErreur:

30      woups "frmDistList", "cmdRafraichir_Click", Err, Erl
End Sub

Private Sub Form_Load()
        
5       On Error GoTo AfficherErreur

10      Set m_otlApp = OuvrirOutlook(m_bDejaOuvert)

15      Set m_folClients = GetFolder(m_otlApp, "Clients GRB")
20      Set m_folContacts = GetFolder(m_otlApp, "Contacts GRB")
25      Set m_folFRS = GetFolder(m_otlApp, "Fournisseurs GRB")

30      Call RemplirListViewDistList

35      Screen.MousePointer = vbDefault

40      Exit Sub

AfficherErreur:

45      woups "frmDistList", "Form_Load", Err, Erl
End Sub

Private Sub lvwContacts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

5       On Error GoTo AfficherErreur

10      lvwContacts.Sorted = False

15      lvwContacts.SortKey = ColumnHeader.Index - 1

20      If lvwContacts.SortOrder = lvwAscending Then
25        lvwContacts.SortOrder = lvwDescending
30      Else
35        lvwContacts.SortOrder = lvwAscending
40      End If

45      lvwContacts.Sorted = True

50      Exit Sub

AfficherErreur:

55      woups "frmDistList", "lvwContacts_ColumnClick", Err, Erl
End Sub

Private Sub lvwDistList_DblClick()

5       On Error GoTo AfficherErreur

10      Call RemplirListViewContacts

15      Exit Sub

AfficherErreur:

20      woups "frmDistList", "lvwDistList_DblClick", Err, Erl
End Sub

Private Sub lvwContacts_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      Dim otlDistList As Outlook.DistListItem
15      Dim iIndex      As Integer
20      Dim iCompteur   As Integer
25      Dim sAdresse    As String

30      If lvwContacts.ListItems.count > 0 Then
35        If KeyCode = vbKeyDelete Then
40          If MsgBox("Voulez-vous vraiment effacer l'enregistrement '" & lvwContacts.SelectedItem.Text & "' ?", vbYesNo) = vbYes Then
45            If lvwDistList.SelectedItem.Index = 1 Then
50              For iCompteur = 2 To lvwDistList.ListItems.count
55                If lvwDistList.ListItems(iCompteur).Text = lvwContacts.SelectedItem.SubItems(I_COL_CONTACT_DISTLIST) Then
60                  Set otlDistList = m_otlApp.GetNamespace("MAPI").GetItemFromID(lvwDistList.ListItems(iCompteur).Tag)

65                  Exit For
70                End If
75              Next
80            Else
85              Set otlDistList = m_otlApp.GetNamespace("MAPI").GetItemFromID(lvwDistList.SelectedItem.Tag)
90            End If

95            iIndex = lvwContacts.SelectedItem.Index

100           sAdresse = lvwContacts.SelectedItem.SubItems(I_COL_CONTACT_COURRIEL)

105           Call otlDistList.RemoveMember(otlDistList.GetMember(iIndex))

110           Call otlDistList.Save

115           lvwDistList.SelectedItem.SubItems(I_COL_DISTLIST_NBRE) = otlDistList.MemberCount

120           If MsgBox("Voulez-vous ajouter ce contact à la liste des exceptions ? " & vbNewLine & _
                        vbNewLine & _
                        "Ceci évitera de l'ajouter lors de la création de nouvelles listes.", vbYesNo) = vbYes Then
125             Call AjouterException(sAdresse)
130           End If

135           Call RemplirListViewContacts

140           If iIndex > lvwContacts.ListItems.count Then
145             If lvwContacts.ListItems.count > 0 Then
150               lvwContacts.ListItems(lvwContacts.ListItems.count).Selected = True
155             End If
160           Else
165             lvwContacts.ListItems(iIndex).Selected = True
170           End If
175         End If
180       End If
185     End If

190     Exit Sub

AfficherErreur:

195     woups "frmDistList", "lvwDistList_KeyDown", Err, Erl
End Sub

Private Sub AjouterException(ByVal sAdresse As String)

5       On Error GoTo AfficherErreur

10      Dim rstExceptions As ADODB.Recordset

15      Set rstExceptions = New ADODB.Recordset

20      Call rstExceptions.Open("SELECT * FROM GRB_ExceptionsDL WHERE [Exception] = '" & Replace(sAdresse, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

25      If rstExceptions.EOF Then
30        Call rstExceptions.AddNew

35        rstExceptions.Fields("Exception") = sAdresse

40        Call rstExceptions.Update
45      End If

50      Call rstExceptions.Close
55      Set rstExceptions = Nothing

60      Exit Sub

AfficherErreur:

65      woups "frmDistList", "AjouterException", Err, Erl
End Sub
