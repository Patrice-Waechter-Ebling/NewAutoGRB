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
   MDIChild        =   -1  'True
   ScaleHeight     =   11370
   ScaleWidth      =   13125
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
Private Const I_COL_DISTLIST_NOM As Integer = 0
Private Const I_COL_DISTLIST_DOSSIER As Integer = 1
Private Const I_COL_DISTLIST_NBRE As Integer = 2

Private Const I_COL_CONTACT_NOM As Integer = 0
Private Const I_COL_CONTACT_COURRIEL As Integer = 1
Private Const I_COL_CONTACT_DISTLIST As Integer = 2

Private Const I_ITM_TOUTES_LISTES As Integer = 1

Private m_bDejaOuvert As Boolean

Private Sub RemplirListViewContacts()

 On Error GoTo Oups

 Dim otlDistList As Outlook.DistListItem
 Dim myRecipient As Outlook.Recipient
 Dim itmContact As ListItem
 Dim sDistList As String
 Dim iCompteur As Integer
 Dim iListe As Integer
 Dim bAfficher As Boolean

 Call lvwContacts.ListItems.Clear

 If lvwDistList.ListItems.count > 0 Then
 Screen.MousePointer = vbHourglass

  lvwDistList.Enabled = False
  cmdCreerListe.Enabled = False
  cmdRafraichir.Enabled = False
  cmdAfficher.Enabled = False
 
  For iListe = 2 To lvwDistList.ListItems.count
  bAfficher = False

  If lvwDistList.ListItems(I_ITM_TOUTES_LISTES).Selected = True Then
  bAfficher = True
 Else
 If lvwDistList.ListItems(iListe).Selected = True Then
 bAfficher = True
 End If
 End If

 If bAfficher = True Then
 sDistList = lvwDistList.ListItems(iListe).Text

 Set otlDistList = m_otlApp.GetNamespace("MAPI").GetItemFromID(lvwDistList.ListItems(iListe).Tag)

 For iCompteur = 1 To otlDistList.MemberCount
 Set myRecipient = otlDistList.GetMember(iCompteur)

 Set itmContact = lvwContacts.ListItems.Add

 itmContact.Text = myRecipient.Name
 itmContact.SubItems(I_COL_CONTACT_COURRIEL) = myRecipient.Address
 itmContact.SubItems(I_COL_CONTACT_DISTLIST) = sDistList
 Next

 DoEvents
 End If
 Next

 lvwDistList.Enabled = True
1  cmdCreerListe.Enabled = True
 cmdRafraichir.Enabled = True
 cmdAfficher.Enabled = True

 Screen.MousePointer = vbDefault
End If

Exit Sub

Oups:

wOups "frmDistList", "RemplirListViewContacts", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewDistList()

 On Error GoTo Oups

 Dim folGRB As Outlook.MAPIFolder
 Dim myItems As Outlook.Items
 Dim objItem As Object
 Dim itmDL As ListItem
 Dim dblTotal As Double

 Call lvwDistList.ListItems.Clear

 Set itmDL = lvwDistList.ListItems.Add

 itmDL.Text = "(Toutes les listes)"

 Set folGRB = m_folContacts

 Set myItems = folGRB.Items.Restrict("[MessageClass] = 'IPM.DistList'")

  For Each objItem In myItems
  Set itmDL = lvwDistList.ListItems.Add

  itmDL.Tag = objItem.EntryID
  itmDL.Text = objItem.DLName
  itmDL.SubItems(I_COL_DISTLIST_DOSSIER) = "Contacts GRB"
  itmDL.SubItems(I_COL_DISTLIST_NBRE) = objItem.MemberCount

  dblTotal = dblTotal + objItem.MemberCount
  Next

10 Set folGRB = m_folClients

Set myItems = folGRB.Items.Restrict("[MessageClass] = 'IPM.DistList'")

For Each objItem In myItems
 Set itmDL = lvwDistList.ListItems.Add

 itmDL.Tag = objItem.EntryID
 itmDL.Text = objItem.DLName
 itmDL.SubItems(I_COL_DISTLIST_DOSSIER) = "Clients GRB"
 itmDL.SubItems(I_COL_DISTLIST_NBRE) = objItem.MemberCount

 dblTotal = dblTotal + objItem.MemberCount
 Next

Set folGRB = m_folFRS

Set myItems = folGRB.Items.Restrict("[MessageClass] = 'IPM.DistList'")

1  For Each objItem In myItems
 Set itmDL = lvwDistList.ListItems.Add

 itmDL.Tag = objItem.EntryID
 itmDL.Text = objItem.DLName
 itmDL.SubItems(I_COL_DISTLIST_DOSSIER) = "Fournisseurs GRB"
 itmDL.SubItems(I_COL_DISTLIST_NBRE) = objItem.MemberCount

 dblTotal = dblTotal + objItem.MemberCount
1  Next

 lvwDistList.ListItems(1).SubItems(I_COL_DISTLIST_NBRE) = dblTotal

 Exit Sub

Oups:

wOups "frmDistList", "RemplirListViewDistList", Err, Err.number, Err.Description
End Sub

Private Sub cmdAfficher_Click()

 On Error GoTo Oups

 Call RemplirListViewContacts

 Exit Sub

Oups:

 wOups "frmDistList", "cmdAfficher_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdCreerListe_Click()

 On Error GoTo Oups

 Call OuvrirForm(frmAjoutDL, False)

 Exit Sub

Oups:

 wOups "frmDistList", "cmdCreerListe_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRafraichir_Click()

 On Error GoTo Oups

 Screen.MousePointer = vbHourglass

 Call RemplirListViewDistList

 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmDistList", "cmdRafraichir_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()
 
 On Error GoTo Oups

 Set m_otlApp = OuvrirOutlook(m_bDejaOuvert)

 Set m_folClients = GetFolder(m_otlApp, "Clients GRB")
 Set m_folContacts = GetFolder(m_otlApp, "Contacts GRB")
 Set m_folFRS = GetFolder(m_otlApp, "Fournisseurs GRB")

 Call RemplirListViewDistList

 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmDistList", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub lvwContacts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

 On Error GoTo Oups

 lvwContacts.Sorted = False

 lvwContacts.SortKey = ColumnHeader.Index - 1

 If lvwContacts.SortOrder = lvwAscending Then
 lvwContacts.SortOrder = lvwDescending
 Else
 lvwContacts.SortOrder = lvwAscending
 End If

 lvwContacts.Sorted = True

 Exit Sub

Oups:

 wOups "frmDistList", "lvwContacts_ColumnClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwDistList_DblClick()

 On Error GoTo Oups

 Call RemplirListViewContacts

 Exit Sub

Oups:

 wOups "frmDistList", "lvwDistList_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwContacts_KeyDown(KeyCode As Integer, Shift As Integer)

 On Error GoTo Oups

 Dim otlDistList As Outlook.DistListItem
 Dim iIndex As Integer
 Dim iCompteur As Integer
 Dim sAdresse As String

 If lvwContacts.ListItems.count > 0 Then
 If KeyCode = vbKeyDelete Then
 If MsgBox("Voulez-vous vraiment effacer l'enregistrement '" & lvwContacts.SelectedItem.Text & "' ?", vbYesNo) = vbYes Then
 If lvwDistList.SelectedItem.Index = 1 Then
 For iCompteur = 2 To lvwDistList.ListItems.count
 If lvwDistList.ListItems(iCompteur).Text = lvwContacts.SelectedItem.SubItems(I_COL_CONTACT_DISTLIST) Then
  Set otlDistList = m_otlApp.GetNamespace("MAPI").GetItemFromID(lvwDistList.ListItems(iCompteur).Tag)

  Exit For
  End If
  Next
  Else
  Set otlDistList = m_otlApp.GetNamespace("MAPI").GetItemFromID(lvwDistList.SelectedItem.Tag)
  End If

  iIndex = lvwContacts.SelectedItem.Index

 sAdresse = lvwContacts.SelectedItem.SubItems(I_COL_CONTACT_COURRIEL)

 Call otlDistList.RemoveMember(otlDistList.GetMember(iIndex))

 Call otlDistList.Save

 lvwDistList.SelectedItem.SubItems(I_COL_DISTLIST_NBRE) = otlDistList.MemberCount

 If MsgBox("Voulez-vous ajouter ce contact à la liste des exceptions ? " & vbNewLine & _
 vbNewLine & _
 "Ceci évitera de l'ajouter lors de la création de nouvelles listes.", vbYesNo) = vbYes Then
 Call AjouterException(sAdresse)
 End If

 Call RemplirListViewContacts

 If iIndex > lvwContacts.ListItems.count Then
 If lvwContacts.ListItems.count > 0 Then
 lvwContacts.ListItems(lvwContacts.ListItems.count).Selected = True
 End If
 Else
 lvwContacts.ListItems(iIndex).Selected = True
 End If
 End If
 End If
End If

 Exit Sub

Oups:

1  wOups "frmDistList", "lvwDistList_KeyDown", Err, Err.number, Err.Description
End Sub

Private Sub AjouterException(ByVal sAdresse As String)

 On Error GoTo Oups

 Dim rstExceptions As ADODB.Recordset

 Set rstExceptions = New ADODB.Recordset

 Call rstExceptions.Open("SELECT * FROM GrbExceptionsDL WHERE [Exception] = '" & Replace(sAdresse, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstExceptions.EOF Then
 Call rstExceptions.AddNew

 rstExceptions.Fields("Exception") = sAdresse

 Call rstExceptions.Update
 End If

 Call rstExceptions.Close
 Set rstExceptions = Nothing

  Exit Sub

Oups:

  wOups "frmDistList", "AjouterException", Err, Err.number, Err.Description
End Sub
