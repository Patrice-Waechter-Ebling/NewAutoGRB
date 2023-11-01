VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSoumissionSectionMec 
   BackColor       =   &H00000000&
   Caption         =   "Sections mécaniques"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   6090
   Begin VB.Frame fraAjout 
      BackColor       =   &H00000000&
      Caption         =   "Ajout de nouvelles sections"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   4575
      Begin VB.TextBox txtFrancais 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtAnglais 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   3480
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdAnnuler 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Français"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Anglais"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.CommandButton CmdSupp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Supprimer"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton CmdFerme 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Fermer"
      Height          =   495
      Left            =   4440
      TabIndex        =   14
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Ç"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "È"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton CmdAdd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Ajouter"
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdModifier 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Modifier"
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   2280
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvwSection 
      Height          =   2895
      Left            =   720
      TabIndex        =   13
      Top             =   1080
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5106
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Français"
         Object.Width           =   2924
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Anglais"
         Object.Width           =   2924
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Section"
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
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmSoumissionSectionMec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const I_COL_FRANCAIS As Integer = 0
Private Const I_COL_ANGLAIS As Integer = 1

Private m_bAjout As Boolean

Private Sub Sauvegarde()

 On Error GoTo Oups

 'Sauvegarde des données dans l'ordre du lister
 Dim rstSection As ADODB.Recordset
 Dim iCompteur As Integer
 
 Set rstSection = New ADODB.Recordset
 
 'Pour toutes les données dans lister
 For iCompteur = 1 To lvwSection.ListItems.count
 Call rstSection.Open("SELECT NomSectionFR, NomSectionEN, Ordre FROM GrbSoumProjSectionMec WHERE Id = " & lvwSection.ListItems(iCompteur).Tag, g_connData, adOpenDynamic, adLockOptimistic)
 
 rstSection.Fields("Ordre") = iCompteur
 
 Call rstSection.Update
 
 'Ferme la table
 Call rstSection.Close
 Next

 Set rstSection = Nothing

  Exit Sub

Oups:

  wOups "frmSoumissionSectionMec", "Sauvegarde", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewSection()

 On Error GoTo Oups

 'Rempli le ListView des Sections
 Dim rstSection As ADODB.Recordset
 Dim itmSection As ListItem
 
 Set rstSection = New ADODB.Recordset
 
 Call rstSection.Open("SELECT * FROM GrbSoumProjSectionMec ORDER BY Ordre", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Il faut vider le ListView avant de le remplir
 Call lvwSection.ListItems.Clear
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstSection.EOF
 Set itmSection = lvwSection.ListItems.Add
 
 itmSection.Tag = rstSection.Fields("Id")
 
 'Nom en francais
 itmSection.Text = rstSection.Fields("NomSectionFR")
 
 'Nom en anglais
 If Not IsNull(rstSection.Fields("NomSectionEN")) Then
  itmSection.SubItems(I_COL_ANGLAIS) = rstSection.Fields("NomSectionEN")
  Else
  itmSection.SubItems(I_COL_ANGLAIS) = vbNullString
  End If
 
  Call rstSection.MoveNext
  Loop
 
  Call rstSection.Close
  Set rstSection = Nothing

10 Exit Sub

Oups:

wOups "frmSoumissionSectionMec", "RemplirListViewSection", Err, Err.number, Err.Description
End Sub

Private Sub CmdAdd_Click()

 On Error GoTo Oups

 m_bAjout = True

 txtAnglais.Text = vbNullString
 txtFrancais.Text = vbNullString

 fraAjout.Visible = True

 Call txtFrancais.SetFocus

 Exit Sub

Oups:

 wOups "frmSoumissionSectionMec", "CmdAdd_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups
 
 fraAjout.Visible = False

 Exit Sub

Oups:

 wOups "frmSoumissionSectionMec", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdOK_Click()

 On Error GoTo Oups

 Dim rstSection As ADODB.Recordset
 Dim rstMaxOrdre As ADODB.Recordset
 
 If Trim$(txtFrancais.Text) = vbNullString Or Trim$(txtAnglais.Text) = vbNullString Then
 Call MsgBox("Le nom en français et en anglais est obligatoire!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
 
 Screen.MousePointer = vbHourglass
 
 Set rstSection = New ADODB.Recordset
 
 If m_bAjout = True Then
 'ouvre la table
 Call rstSection.Open("SELECT * FROM GrbSoumProjSectionMec WHERE NomSectionFR = '" & Replace(txtFrancais.Text, "'", "''") & "' OR NomSectionEN = '" & Replace(txtAnglais.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 'si n'existe pas
  If rstSection.EOF Then
  Set rstMaxOrdre = New ADODB.Recordset

  Call rstMaxOrdre.Open("SELECT Max(Ordre) as MaxOrdre FROM GrbSoumProjSectionMec", g_connData, adOpenDynamic, adLockOptimistic)
 
 'ajoute la section
  Call rstSection.AddNew
 
  rstSection.Fields("NomSectionFR") = txtFrancais.Text
  rstSection.Fields("NomSectionEN") = txtAnglais.Text
  rstSection.Fields("Ordre") = rstMaxOrdre.Fields("MaxOrdre") + 1
 
  Call rstMaxOrdre.Close
 Set rstMaxOrdre = Nothing
 
Call rstSection.Update
 
 m_bAjout = False
 Else
 Call MsgBox("Cette section existe déjà!")
 End If
 
 Call rstSection.Close
 Set rstSection = Nothing
Else
 Call rstSection.Open("SELECT * FROM GrbSoumProjSectionMec WHERE Id = " & lvwSection.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
 
 rstSection.Fields("NomSectionFR") = txtFrancais.Text
 rstSection.Fields("NomSectionEN") = txtAnglais.Text
 
Call rstSection.Update
 
 'ferme la table
 Call rstSection.Close
 Set rstSection = Nothing
End If
 
 Call RemplirListViewSection
 
fraAjout.Visible = False
 
 Screen.MousePointer = vbDefault

1  Exit Sub

Oups:

 wOups "frmSoumissionSectionMec", "CmdOK_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDown_Click()

 On Error GoTo Oups

 '''''''''''''''''''''''''''''''''''''''''
 'descend la selection d'une ligne
 '''''''''''''''''''''''''''''''''''''''''
 Dim sTagAvant As String
 Dim sTagApres As String
 Dim sFrancaisAvant As String
 Dim sFrancaisApres As String
 Dim sAnglaisAvant As String
 Dim sAnglaisApres As String
 Dim iIndex As Integer
 
 iIndex = lvwSection.SelectedItem.Index
 
 If iIndex < lvwSection.ListItems.count Then
 'garde en memoire les données qui vont se repositionné dans la list
 sTagAvant = lvwSection.ListItems(iIndex).Tag
  sFrancaisAvant = lvwSection.ListItems(iIndex).Text
  sAnglaisAvant = lvwSection.ListItems(iIndex).SubItems(I_COL_ANGLAIS)
 
  sTagApres = lvwSection.ListItems(iIndex + 1).Tag
  sFrancaisApres = lvwSection.ListItems(iIndex + 1).Text
  sAnglaisApres = lvwSection.ListItems(iIndex + 1).SubItems(I_COL_ANGLAIS)
 
 'reposition dans la liste
  lvwSection.ListItems(iIndex).Tag = sTagApres
  lvwSection.ListItems(iIndex).Text = sFrancaisApres
  lvwSection.ListItems(iIndex).SubItems(I_COL_ANGLAIS) = sAnglaisApres
 
lvwSection.ListItems(iIndex + 1).Tag = sTagAvant
1 lvwSection.ListItems(iIndex + 1).Text = sFrancaisAvant
 lvwSection.ListItems(iIndex + 1).SubItems(I_COL_ANGLAIS) = sAnglaisAvant

 'descend la selection
 lvwSection.ListItems(iIndex + 1).Selected = True
 
 Call lvwSection.SelectedItem.EnsureVisible
End If

Exit Sub

Oups:

wOups "frmSoumissionSectionMec", "cmdDown_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdUp_Click()

 On Error GoTo Oups

 '''''''''''''''''''''''''''''''''''''''''
 'monte la selection d'une ligne
 '''''''''''''''''''''''''''''''''''''''''
 Dim sTagAvant As String
 Dim sTagApres As String
 Dim sFrancaisAvant As String
 Dim sFrancaisApres As String
 Dim sAnglaisAvant As String
 Dim sAnglaisApres As String
 Dim iIndex As Integer
 
 iIndex = lvwSection.SelectedItem.Index
 
 If iIndex > 1 Then
 'garde en memoire les données qui vont se repositionné dans la list
 sTagAvant = lvwSection.ListItems(iIndex).Tag
  sFrancaisAvant = lvwSection.ListItems(iIndex).Text
  sAnglaisAvant = lvwSection.ListItems(iIndex).SubItems(I_COL_ANGLAIS)
 
  sTagApres = lvwSection.ListItems(iIndex - 1).Tag
  sFrancaisApres = lvwSection.ListItems(iIndex - 1).Text
  sAnglaisApres = lvwSection.ListItems(iIndex - 1).SubItems(I_COL_ANGLAIS)
 
 'reposition dans la liste
  lvwSection.ListItems(iIndex).Tag = sTagApres
  lvwSection.ListItems(iIndex).Text = sFrancaisApres
  lvwSection.ListItems(iIndex).SubItems(I_COL_ANGLAIS) = sAnglaisApres
 
lvwSection.ListItems(iIndex - 1).Tag = sTagAvant
1 lvwSection.ListItems(iIndex - 1).Text = sFrancaisAvant
 lvwSection.ListItems(iIndex - 1).SubItems(I_COL_ANGLAIS) = sAnglaisAvant

 'monte la selection
 lvwSection.ListItems(iIndex - 1).Selected = True
 
 Call lvwSection.SelectedItem.EnsureVisible
End If

Exit Sub

Oups:

wOups "frmSoumissionSectionMec", "cmdUp_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdFerme_Click()

 On Error GoTo Oups

 Call Sauvegarde

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmSoumissionSectionMec", "CmdFerme_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdModifier_Click()

 On Error GoTo Oups

 txtFrancais.Text = lvwSection.SelectedItem.Text
 txtAnglais.Text = lvwSection.SelectedItem.SubItems(I_COL_ANGLAIS)
 
 fraAjout.Visible = True

 Exit Sub

Oups:

 wOups "frmSoumissionSectionMec", "cmdModifier_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdSupp_Click()

 On Error GoTo Oups

 Dim rstSoumission As ADODB.Recordset

 'fonction qui supprime lenregistrement courant
 If lvwSection.ListItems.count > 0 Then
 If MsgBox("Etes-vous sur de supprimer cette enregistrement?", vbYesNo, "Supprimer") = vbYes Then
 Screen.MousePointer = vbHourglass
 
 Set rstSoumission = New ADODB.Recordset
 
 Call rstSoumission.Open("SELECT Id FROM GrbSoumission_pieces WHERE Id = " & lvwSection.SelectedItem.Tag & " AND Type = 'M'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstSoumission.EOF Then
 'Efface l'enregistrement
 Call g_connData.Execute("DELETE * FROM GrbSoumProjSectionMec WHERE Id = " & lvwSection.SelectedItem.Tag)
 
 'Si le combo n'est pas vide, on sélectionne le premier
 If lvwSection.ListItems.count > 0 Then
 lvwSection.ListItems(1).Selected = True
  End If
  Else
  Call MsgBox("Impossible de supprimer une section déjà utilisé dans une soumission!", vbOKOnly, "Erreur")
  End If
 
 'Ferme la table
  Call rstSoumission.Close
  Set rstSoumission = Nothing
 
  Call RemplirListViewSection
  End If
10 Else
1 Call MsgBox("Aucun enregistrement sélectionné!")
End If
 
Screen.MousePointer = vbDefault

Exit Sub

Oups:

wOups "frmSoumissionSectionMec", "CmdSupp_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 'Ouverture de la fenêtre
 Call RemplirListViewSection

 Exit Sub

Oups:

 wOups "frmSoumissionSectionMec", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub lvwSection_Click()

 On Error GoTo Oups

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 'met les fleche enabled depandant si au debut ou a la fin
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If lvwSection.ListItems.count > 0 Then
 If lvwSection.SelectedItem.Index = lvwSection.ListItems.count Then
 cmdDown.Enabled = False
 Else
 cmdDown.Enabled = True
 End If
 
 If lvwSection.SelectedItem.Index = 1 Then
 cmdUp.Enabled = False
 Else
 cmdUp.Enabled = True
  End If
  End If

  Exit Sub

Oups:

  wOups "frmSoumissionSectionMec", "lvwSection_Click", Err, Err.number, Err.Description
End Sub

Private Sub lvwSection_DblClick()

 On Error GoTo Oups

 txtFrancais.Text = lvwSection.SelectedItem.Text
 txtAnglais.Text = lvwSection.SelectedItem.SubItems(I_COL_ANGLAIS)
 
 fraAjout.Visible = True

 Exit Sub

Oups:

 wOups "frmSoumissionSectionMec", "lvwSection_DblClick", Err, Err.number, Err.Description
End Sub
