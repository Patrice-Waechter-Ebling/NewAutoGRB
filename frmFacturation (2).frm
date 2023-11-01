VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacturation 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturation"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9360
   Begin VB.CommandButton cmd_export 
      Appearance      =   0  'Flat
      Caption         =   "Exporter vers Excel"
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   840
      Width           =   3615
   End
   Begin VB.Frame fraType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   33
      Top             =   1140
      Width           =   1695
      Begin VB.OptionButton optType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tous"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Électrique"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Mécanique"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdVerrouiller 
      Appearance      =   0  'Flat
      Caption         =   "Verrouiller Soum"
      Height          =   375
      Left            =   2880
      TabIndex        =   32
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdSommaire 
      Appearance      =   0  'Flat
      Caption         =   "Sommaire"
      Height          =   375
      Left            =   4800
      TabIndex        =   31
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtDateOuverture 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdCommentaire 
      Appearance      =   0  'Flat
      Caption         =   "Commentaires"
      Height          =   375
      Left            =   2160
      TabIndex        =   29
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdModifier 
      Appearance      =   0  'Flat
      Caption         =   "Modifier"
      Height          =   375
      Left            =   7080
      TabIndex        =   28
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdNCRectifier 
      Appearance      =   0  'Flat
      Caption         =   "NC"
      Height          =   375
      Left            =   8280
      TabIndex        =   21
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmdReouverture 
      Appearance      =   0  'Flat
      Caption         =   "Annuler Fermeture"
      Height          =   375
      Left            =   6600
      TabIndex        =   24
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmdImprimer 
      Appearance      =   0  'Flat
      Caption         =   "Imprimer"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtRaisonFermeture 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtDateFermeture 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtClient 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton cmdSupprimer 
      Appearance      =   0  'Flat
      Caption         =   "Supprimer"
      Height          =   375
      Left            =   6120
      TabIndex        =   19
      Top             =   5760
      Width           =   975
   End
   Begin VB.Frame fraMontrer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Soumissions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   1695
      Begin VB.OptionButton optMontrer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fermées"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optMontrer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ouvertes"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optMontrer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Toutes"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOuvrirProjSoum 
      Appearance      =   0  'Flat
      Caption         =   "Ouvrir Soum"
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdFermerProjSoum 
      Appearance      =   0  'Flat
      Caption         =   "Fermer Soum"
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      Top             =   5760
      Width           =   1215
   End
   Begin VB.ComboBox cmbProjSoum 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmFacturation.frx":0000
      Left            =   4800
      List            =   "frmFacturation.frx":000A
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdFacturerRectifier 
      Appearance      =   0  'Flat
      Caption         =   "Facturer"
      Height          =   375
      Left            =   7200
      TabIndex        =   20
      Top             =   5760
      Width           =   975
   End
   Begin VB.ComboBox cmbNoProjSoum 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6480
      TabIndex        =   14
      Text            =   "cmbNoProjSoum"
      Top             =   1800
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvwProjets 
      Height          =   3375
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employé"
         Object.Width           =   1376
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   1746
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Début"
         Object.Width           =   1085
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fin"
         Object.Width           =   1085
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Description"
         Object.Width           =   6826
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Total"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "No. Facture"
         Object.Width           =   1799
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Type"
         Object.Width           =   3000
      EndProperty
   End
   Begin VB.CommandButton cmdFermer 
      Appearance      =   0  'Flat
      Caption         =   "Fermer"
      Height          =   375
      Left            =   8280
      TabIndex        =   25
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox txtNoProjSoum 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6480
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Description :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4440
      TabIndex        =   38
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblRaisonFermeture 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Raison de la fermeture :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblDateFermeture 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date de fermeture : "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblDateOuverture 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date d'ouverture : "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblClient 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Client : "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblHeuresNonFacturees 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   27
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lblHeuresFacturees 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   23
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Heures non facturées :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Heures facturées :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label lblTitreProjSoum 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Numéro de projet :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   10
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "frmFacturation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

'Index des colonnes du listview
Private Const I_LVW_EMPLOYE As Integer = 0
Private Const I_LVW_DATE As Integer = 1
Private Const I_LVW_DEBUT As Integer = 2
Private Const I_LVW_FIN As Integer = 3
Private Const I_LVW_DESCRIPTION As Integer = 4
Private Const I_LVW_TOTAL As Integer = 5
Private Const I_LVW_NO_FACTURE As Integer = 6
Private Const I_LVW_TYPE As Integer = 7


'Index de optType
Private Const I_OPT_TYPE_ELECTRIQUE As Integer = 0
Private Const I_OPT_TYPE_MECANIQUE As Integer = 1
Private Const I_OPT_TYPE_TOUS As Integer = 2

'Index du combo
Private Const I_CMB_PROJET As Integer = 0
Private Const I_CMB_SOUMISSION As Integer = 1

'Caption du bouton cmdFacturerRectifier
Private Const S_FACTURER As String = "Facturer"
Private Const S_RECTIFIER As String = "Rectifier"
Private Const S_NC As String = "NC"

'Caption des Option Buttons
'Si c'est un projet
Private Const S_PROJ_OUVERT As String = "Ouverts"
Private Const S_PROJ_FERME As String = "Fermés"
Private Const S_PROJ_TOUS As String = "Tous"

'Si c'est une soumission
Private Const S_SOUM_OUVERT As String = "Ouvertes"
Private Const S_SOUM_FERME As String = "Fermées"
Private Const S_SOUM_TOUS As String = "Toutes"

'Caption de fraMontrer
Private Const S_FRA_PROJ As String = "Projets"
Private Const S_FRA_SOUM As String = "Soumissions"

'Index des Option Buttons
Private Const I_OPT_TOUS As Integer = 0
Private Const I_OPT_OUVERT As Integer = 1
Private Const I_OPT_FERME As Integer = 2

Private Enum enumType
 TYPE_PROJET = 0
 TYPE_SOUMISSION = 1
End Enum

Private m_eType As enumType

Public m_iIDClient As Integer
Public m_sDescription As String

Public m_bModifClient As Boolean

Private m_bLoading As Boolean

Private Sub cmbNoProjSoum_Click()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset

 Set rstProjSoum = New ADODB.Recordset

 txtNoProjSoum.Text = cmbNoProjSoum.Text
 
 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstProjSoum.Fields("Ouvert") = True Then
 cmdFermerProjSoum.Enabled = True
 Else
 cmdFermerProjSoum.Enabled = False
 End If

 If rstProjSoum.Fields("Verrouillé") = True Then
  If m_eType = TYPE_SOUMISSION Then
  cmdVerrouiller.Caption = "Déverrouiller Soum"
  Else
  cmdVerrouiller.Caption = "Déverrouiller Proj"
  End If
  Else
  If m_eType = TYPE_SOUMISSION Then
  cmdVerrouiller.Caption = "Verrouiller Soum"
Else
cmdVerrouiller.Caption = "Verrouiller Proj"
 End If
End If

Call rstProjSoum.Close
Set rstProjSoum = Nothing
 
Call AfficherProjSoum

Exit Sub

Oups:

wOups "frmFacturation", "cmbNoProjSoum_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbNoProjSoum_KeyUp(KeyCode As Integer, Shift As Integer)
 
 On Error GoTo Oups

 Dim iCompteur As Integer
 
 For iCompteur = 0 To cmbNoProjSoum.ListCount - 1
 If UCase(cmbNoProjSoum.LIST(iCompteur)) = UCase(cmbNoProjSoum.Text) Then
 cmbNoProjSoum.ListIndex = iCompteur
 
 Exit For
 End If
 Next

 Exit Sub

Oups:

 wOups "frmFacturation", "cmbProjSoum_KeyUp", Err, Err.number, Err.Description
End Sub


Private Sub AfficherProjSoum()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim rstClient As ADODB.Recordset
 
 Set rstProjSoum = New ADODB.Recordset
 Set rstClient = New ADODB.Recordset
 
 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call rstClient.Open("SELECT NomClient FROM GrbClient WHERE IDClient = " & rstProjSoum.Fields("NoClient"), g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstClient.EOF Then
 txtClient.Text = rstClient.Fields("NomClient")
 txtClient.Tag = rstProjSoum.Fields("NoClient")
 End If

  Call rstClient.Close
  Set rstClient = Nothing
 
  If Not IsNull(rstProjSoum.Fields("DateOuverture")) Then
  txtDateOuverture.Text = rstProjSoum.Fields("DateOuverture")
  Else
  txtDateOuverture.Text = vbNullString
  End If
 
  If optMontrer(I_OPT_TOUS).Value = True Or optMontrer(I_OPT_FERME).Value = True Then
If Not IsNull(rstProjSoum.Fields("DateFermeture")) Then
txtDateFermeture.Text = rstProjSoum.Fields("DateFermeture")
 Else
 txtDateFermeture.Text = vbNullString
 End If
 
 If Not IsNull(rstProjSoum.Fields("RaisonFermeture")) Then
 txtRaisonFermeture.Text = rstProjSoum.Fields("RaisonFermeture")
 Else
 txtRaisonFermeture.Text = vbNullString
 End If
End If

If Not IsNull(rstProjSoum.Fields("Description")) Then
txtDescription.Text = rstProjSoum.Fields("Description")
Else
 txtDescription.Text = vbNullString
End If

 Call rstProjSoum.Close
Set rstProjSoum = Nothing
 
 Call RemplirListView

1  Exit Sub

Oups:

 wOups "frmFacturation", "AfficherProjSoum", Err, Err.number, Err.Description
End Sub

Private Sub cmd_export_Click()
Call vb_to_excel

End Sub

Private Sub cmdCommentaire_Click()

 On Error GoTo Oups

 If cmbProjSoum.ListIndex = I_CMB_PROJET Then
 Call frmCommentairesProjSoum.Afficher(cmbNoProjSoum.Text, True)
 Else
 Call frmCommentairesProjSoum.Afficher(cmbNoProjSoum.Text, False)
 End If

 Exit Sub

Oups:

 wOups "frmFacturation", "cmdCommentaire_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdFacturerRectifier_Click()

 On Error GoTo Oups

 Dim sNoFacture As String
 Dim rstFacture As ADODB.Recordset
 Dim sWhere As String
 Dim iCompteur As Integer
 
 If cmdFacturerRectifier.Caption = S_FACTURER Then
 'Change la valeur du champs "Facturé" dans la table GrbPunch pour True et
 'ajoute le numéro de la facture dans le champs "NoFacture"
 
 sNoFacture = InputBox("Entrez le numéro de la facture")
 
 'Le numéro de facture peut être vide, mais si il ne l'est pas, il doit
 'être numérique
 If sNoFacture <> vbNullString Then
 If Not IsNumeric(sNoFacture) Then
 Call MsgBox("Le numéro de facture est invalide", vbOKOnly, "Erreur")
 
 Exit Sub
  End If
 
  sWhere = "IDPunch In ("
 
  For iCompteur = 1 To lvwProjets.ListItems.count
 'Si l'élément est sélectionné
  If lvwProjets.ListItems(iCompteur).Selected = True Then
 'Si la condition where est vide, c'est parce que c'est le premier élément
 'sélectionné
  If sWhere = "IDPunch In (" Then
  sWhere = sWhere & lvwProjets.ListItems(iCompteur).Tag
  Else
  sWhere = sWhere & "," & lvwProjets.ListItems(iCompteur).Tag
 End If
 End If
 Next

 sWhere = sWhere & ")"
 
 Set rstFacture = New ADODB.Recordset
 
 'Ouverture des enregistrements sélectionnés dans le ListView
 Call rstFacture.Open("SELECT * FROM GrbPunch WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
 
 Do While Not rstFacture.EOF
 'Mettre la facturation à true et remplir le numéro de facture
 rstFacture.Fields("Facturé") = True
 rstFacture.Fields("NoFacture") = sNoFacture
 
 Call rstFacture.Update
 
 Call rstFacture.MoveNext
 Loop
 
 Call rstFacture.Close
 Set rstFacture = Nothing
 
 Call RemplirListView(lvwProjets.SelectedItem.Index)
 End If
 Else
 'Change la valeur du champs "Facturé" dans la table GrbPunch pour False et
 'ajoute le numéro de la facture dans le champs "NoFacture"
 
 sWhere = "IDPunch In ("
 
 For iCompteur = 1 To lvwProjets.ListItems.count
 'Si l'élément est sélectionné
1  If lvwProjets.ListItems(iCompteur).Selected = True Then
 'Si la condition where est vide, c'est parce que c'est le premier élément
 'sélectionné
 If sWhere = "IDPunch In (" Then
 sWhere = sWhere & lvwProjets.ListItems(iCompteur).Tag
 Else
 sWhere = sWhere & "," & lvwProjets.ListItems(iCompteur).Tag
 End If
 End If
 Next
 
 sWhere = sWhere & ")"
 
 Set rstFacture = New ADODB.Recordset
 
 'Ouverture des enregistrements sélectionnés dans le ListView
 Call rstFacture.Open("SELECT Facturé, NoFacture FROM GrbPunch WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin du recordset
 Do While Not rstFacture.EOF
 'Mettre la facturation à false et vider le numéro de facture
 rstFacture.Fields("Facturé") = False
 rstFacture.Fields("NoFacture") = vbNullString
 
 Call rstFacture.Update
 
 Call rstFacture.MoveNext
 Loop
 
Call rstFacture.Close
 Set rstFacture = Nothing
 
Call RemplirListView(lvwProjets.SelectedItem.Index)
End If

30 Exit Sub

Oups:

wOups "frmFacturation", "cmdFacturerRectifier_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdModifier_Click()

 On Error GoTo Oups

 Dim rstPunch As ADODB.Recordset
 Dim rstProjSoum As ADODB.Recordset

 If PeutModifier = True Then
 m_bModifClient = True

 Call frmChoixClient.Show(vbModal)

 m_bModifClient = False

 If m_iIDClient <> txtClient.Tag Then
 Set rstProjSoum = New ADODB.Recordset

 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

 rstProjSoum.Fields("NoClient") = m_iIDClient
  rstProjSoum.Fields("Description") = m_sDescription

  Call rstProjSoum.Update

  Call rstProjSoum.Close
  Set rstProjSoum = Nothing

  Set rstPunch = New ADODB.Recordset

  Call rstPunch.Open("SELECT * FROM GrbPunch WHERE NoProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

  If Not rstPunch.EOF Then
  If MsgBox("Voulez-vous modifier tous les punch ?", vbYesNo) = vbYes Then
 Do While Not rstPunch.EOF
 rstPunch.Fields("NoClient") = m_iIDClient

 Call rstPunch.Update

 Call rstPunch.MoveNext
 Loop
 End If
 End If

 Call rstPunch.Close
 Set rstPunch = Nothing

 Call AfficherProjSoum
 End If
End If

1  Exit Sub

Oups:

wOups "frmFacturation", "cmdModifier_Click", Err, Err.number, Err.Description
End Sub

Private Function PeutModifier() As Boolean

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim rstProjet As ADODB.Recordset
 Dim rstSoumission As ADODB.Recordset
 Dim bPeutModifier As Boolean

 Set rstProjSoum = New ADODB.Recordset

 Call rstProjSoum.Open("SELECT Ouvert, Type FROM GrbProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

 If rstProjSoum.Fields("Ouvert") = True Then
 If rstProjSoum.Fields("Type") = "P" Then
 Set rstProjet = New ADODB.Recordset

 Call rstProjet.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

  If rstProjet.EOF Then
  Call rstProjet.Close

  Call rstProjet.Open("SELECT * FROM GrbProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

  If rstProjet.EOF Then
  bPeutModifier = True
  Else
  Call MsgBox("Le client doit être modifié dans l'écran des projets mécaniques!", vbOKOnly, "Erreur")

  bPeutModifier = False
 End If

 Call rstProjet.Close
 Else
 Call MsgBox("Le client doit être modifié dans l'écran des projets électriques!", vbOKOnly, "Erreur")

 Call rstProjet.Close

 bPeutModifier = False
 End If

 Set rstProjet = Nothing
 Else
 Set rstSoumission = New ADODB.Recordset

 Call rstSoumission.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

 If rstSoumission.EOF Then
 Call rstSoumission.Close

 Call rstSoumission.Open("SELECT * FROM GrbSoumissionMec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

 If rstSoumission.EOF Then
 bPeutModifier = True
 Else
 Call MsgBox("Le client doit être modifié dans l'écran des soumissions mécaniques!", vbOKOnly, "Erreur")

 bPeutModifier = False
1  End If

 Call rstSoumission.Close
 Else
 Call MsgBox("Le client doit être modifié dans l'écran des soumissions électriques!", vbOKOnly, "Erreur")

 Call rstSoumission.Close

 bPeutModifier = False
 End If

 Set rstSoumission = Nothing
 End If
Else
 If rstProjSoum.Fields("Type") = "P" Then
 Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
 Else
 Call MsgBox("Cette soumission est fermée!", vbOKOnly, "Erreur")
 End If

bPeutModifier = False
End If

2  Call rstProjSoum.Close
Set rstProjSoum = Nothing

2  PeutModifier = bPeutModifier

Exit Function

Oups:

30 wOups "frmFacturation", "PeutModifier", Err, Err.number, Err.Description
End Function

Private Sub cmdNCRectifier_Click()

 On Error GoTo Oups

 Dim rstFacture As ADODB.Recordset
 Dim sWhere As String
 Dim iCompteur As Integer
 
 If cmdNCRectifier.Caption = S_NC Then
 'Change la valeur du champs "Facturé" dans la table GrbPunch pour True et
 'ajoute NC dans le champs "NoFacture"
 sWhere = "IDPunch In ("
 
 For iCompteur = 1 To lvwProjets.ListItems.count
 'Si l'élément est sélectionné
 If lvwProjets.ListItems(iCompteur).Selected = True Then
 'Si la condition where est vide, c'est parce que c'est le premier élément
 'sélectionné
 If sWhere = "IDPunch In (" Then
 sWhere = sWhere & lvwProjets.ListItems(iCompteur).Tag
 Else
  sWhere = sWhere & "," & lvwProjets.ListItems(iCompteur).Tag
  End If
  End If
  Next
 
  sWhere = sWhere & ")"

  Set rstFacture = New ADODB.Recordset
 
 'Ouverture des enregistrements sélectionnés dans le ListView
  Call rstFacture.Open("SELECT * FROM GrbPunch WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
 
  Do While Not rstFacture.EOF
 'Mettre la facturation à true et remplir le numéro de facture
 rstFacture.Fields("Facturé") = True
rstFacture.Fields("NoFacture") = "NC"
 
 Call rstFacture.Update
 
 Call rstFacture.MoveNext
 Loop
 
 Call rstFacture.Close
 Set rstFacture = Nothing
 
 Call RemplirListView(lvwProjets.SelectedItem.Index)
Else
 'Change la valeur du champs "Facturé" dans la table GrbPunch pour False et
 'enlève NC
 sWhere = "IDPunch In ("
 
 For iCompteur = 1 To lvwProjets.ListItems.count
 'Si l'élément est sélectionné
 If lvwProjets.ListItems(iCompteur).Selected = True Then
 'Si la condition where est vide, c'est parce que c'est le premier élément
 'sélectionné
 If sWhere = "IDPunch In (" Then
 sWhere = sWhere & lvwProjets.ListItems(iCompteur).Tag
 Else
 sWhere = sWhere & "," & lvwProjets.ListItems(iCompteur).Tag
 End If
 End If
 Next
 
1  sWhere = sWhere & ")"

 Set rstFacture = New ADODB.Recordset
 
 'Ouverture des enregistrements sélectionnés dans le ListView
 Call rstFacture.Open("SELECT * FROM GrbPunch WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin du recordset
 Do While Not rstFacture.EOF
 'Mettre la facturation à false et vider le numéro de facture
 rstFacture.Fields("Facturé") = False
 rstFacture.Fields("NoFacture") = vbNullString
 
 Call rstFacture.Update
 
 Call rstFacture.MoveNext
 Loop
 
 Call rstFacture.Close
 Set rstFacture = Nothing
 
 Call RemplirListView(lvwProjets.SelectedItem.Index)
End If

2  Exit Sub

Oups:

wOups "frmFacturation", "cmdFacturerRectifier_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 'Fermer de la fênêtre
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmFacturation", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdImprimer_Click()
Dim intdummie As Integer



 On Error GoTo Oups

 If m_eType = TYPE_PROJET Then
 Call frmChoixDateImpressionFacturation.Afficher(txtNoProjSoum.Text, True, txtClient.Text, txtDescription.Text)
 Else
 Call frmChoixDateImpressionFacturation.Afficher(txtNoProjSoum.Text, False, txtClient.Text, txtDescription.Text)
 End If

 Exit Sub

Oups:

 wOups "frmFacturation", "cmdImprimer_Click", Err, Err.number, Err.Description
End Sub

Private Function vb_to_excel()



  Dim iCount As Integer
 Dim oXLApp As Excel.Application 'Declare the object variables
 Dim oXLBook As Excel.Workbook
 Dim oXLSheet As Excel.Worksheet
 Dim data_array(1 To 1500, 1 To 7) As Variant 'modifier pour intégré une nouvelle onglet
 Dim r As Integer
 Set oXLApp = New Excel.Application 'Create a new instance of Excel
 Set oXLBook = oXLApp.Workbooks.Add 'Add a new workbook
 Set oXLSheet = oXLBook.Worksheets(1) 'Work with the first worksheet
 oXLApp.Visible = False

'on inscrit les valeurs du listbox dans un tableau
r = 1
Do While r <= lvwProjets.ListItems.count <> Empty
 data_array(r, 1) = lvwProjets.ListItems(r)
 data_array(r, 2) = lvwProjets.ListItems(r).SubItems(1)
 data_array(r, 3) = lvwProjets.ListItems(r).SubItems(2)
 data_array(r, 4) = lvwProjets.ListItems(r).SubItems(3)
 data_array(r, 5) = lvwProjets.ListItems(r).SubItems(4) 'Ajouter la description a la table excel 2017-06-2  GLL
 data_array(r, 6) = CDbl(lvwProjets.ListItems(r).SubItems(5))
 data_array(r, 7) = lvwProjets.ListItems(r).SubItems(7) 'Ajouter pour avoir le tableau complet en Excel
 
 r = r + 1
 
Loop




'creation en-tête de colonne
oXLSheet.range("A1: G1").Font.Bold = True
oXLSheet.range("A:G").HorizontalAlignment = xlCenter
oXLSheet.range("A1: G1").Value = Array("Employé", "Date", "Debut", "Fin", "Description", "Total", "Type") 'GLL


'inscription des valeur du tableau dans excel
oXLSheet.range("A2").Resize(r, 7).Value = data_array
'ajustement largeur des colonne
oXLSheet.range("A:G").Columns.AutoFit
oXLApp.Visible = True

 







End Function


Private Sub cmdOuvrirProjSoum_Click()

 On Error GoTo Oups

 Dim sNumero As String
 Dim rstProjSoum As Recordset
 Dim sQuestion As String
 Dim sType As String
 Dim bNoValide As Boolean

 Select Case m_eType
 Case TYPE_PROJET:
 sQuestion = "Quel est le numéro du projet?"
 sType = "P"
 Case TYPE_SOUMISSION:
 sQuestion = "Quel est le numéro de la soumission?"
  sType = "S"
  End Select
 
  sNumero = InputBox(sQuestion)
 
  If sNumero <> vbNullString Then
  bNoValide = True

  If ValiderFormatNumeroProjSoum(sNumero) = False Then
  bNoValide = False
  End If

If bNoValide = True Then
If m_eType = TYPE_PROJET Then
 If ValiderFormatProjet(sNumero) = False Then
 bNoValide = False
 End If
 Else
 If ValiderFormatSoumission(sNumero) = False Then
 bNoValide = False
 End If
 End If
 End If

 If bNoValide = False Then
 Exit Sub
 End If

 Call frmChoixClient.Show(vbModal)
 
 Set rstProjSoum = New ADODB.Recordset
 
 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstProjSoum.EOF Then
 Call rstProjSoum.AddNew
 
1  rstProjSoum.Fields("IDProjSoum") = sNumero
 rstProjSoum.Fields("NoClient") = m_iIDClient
 rstProjSoum.Fields("Description") = m_sDescription
 rstProjSoum.Fields("DateOuverture") = ConvertDate(Date)
 rstProjSoum.Fields("Ouvert") = True
 rstProjSoum.Fields("Type") = sType
 
 Call rstProjSoum.Update
 
 Call RemplirComboProjSoum
 Else
 Call MsgBox("Ce numéro existe déjà!", vbOKOnly, "Erreur")
 End If
 
 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
2  End If

Exit Sub

Oups:

2  wOups "frmFacturation", "cmdOuvrirProjSoum_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdFermerProjSoum_Click()

 On Error GoTo Oups

 Dim rstProjSoum As Recordset
 Dim sQuestion As String
 Dim sRaison As String
 
 Select Case m_eType
 Case TYPE_PROJET: sQuestion = "Voulez-vous vraiment fermer ce projet?"
 Case TYPE_SOUMISSION: sQuestion = "Voulez-vous vraiment fermer cette soumission?"
 End Select
 
 If MsgBox(sQuestion, vbYesNo) = vbYes Then
 Set rstProjSoum = New ADODB.Recordset

 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 rstProjSoum.Fields("Ouvert") = False
  rstProjSoum.Fields("DateFermeture") = ConvertDate(Date)
 
  If m_eType = TYPE_SOUMISSION Then
  sRaison = InputBox("Quelle est la raison de la fermeture?")
 
  rstProjSoum.Fields("RaisonFermeture") = sRaison
  End If
 
  Call rstProjSoum.Update
 
  Call rstProjSoum.Close
  Set rstProjSoum = Nothing
 
Call RemplirComboProjSoum
End If

Exit Sub

Oups:

wOups "frmFacturation", "cmdFermerProjSoum_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdReouverture_Click()

 On Error GoTo Oups

 Dim rstProjSoum As Recordset
 Dim sQuestion As String
 
 If cmbNoProjSoum.ListIndex <> -1 Then
 Select Case m_eType
 Case TYPE_PROJET: sQuestion = "Voulez-vous vraiment annuler la fermeture de ce projet?"
 Case TYPE_SOUMISSION: sQuestion = "Voulez-vous vraiment annuler la fermeture de cette soumission?"
 End Select
 
 If MsgBox(sQuestion, vbYesNo) = vbYes Then
 Set rstProjSoum = New ADODB.Recordset

 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 rstProjSoum.Fields("Ouvert") = True
  rstProjSoum.Fields("RaisonFermeture") = Null
 
  Call rstProjSoum.Update
 
  Call rstProjSoum.Close
  Set rstProjSoum = Nothing

  Call ViderValeurs
 
  Call RemplirComboProjSoum
  End If
  End If

10 Exit Sub

Oups:

wOups "frmFacturation", "cmdReouverture_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdSommaire_Click()

 On Error GoTo Oups

 Call frmChoixDateSommairePunch.Show(vbModal)

 Exit Sub

Oups:

 wOups "frmFacturation", "cmdImprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdsupprimer_Click()

 On Error GoTo Oups

 Dim sMessage As String
 Dim sErreur As String
 
 Call RemplirListView
 
 If m_eType = TYPE_PROJET Then
 sMessage = "Voulez-vous vraiment effacer le projet " & txtNoProjSoum.Text & "?"
 sErreur = "Impossible de supprimer ce projet car il y a déjà des punchs!"
 Else
 sMessage = "Voulez-vous vraiment effacer la soumission " & txtNoProjSoum.Text & "?"
 sErreur = "Impossible de supprimer cette soumission car il y a déjà des punchs!"
 End If

  If lvwProjets.ListItems.count = 0 Then
  If MsgBox(sMessage, vbYesNo) = vbYes Then
  Call g_connData.Execute("DELETE * FROM GrbProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'")
 
  Call ViderValeurs
 
  Call RemplirComboProjSoum

  If cmbNoProjSoum.ListCount = 0 Then
  Call ViderValeurs
  End If
End If
Else
 Call MsgBox(sErreur, vbOKOnly, "Erreur")
End If

Exit Sub

Oups:

wOups "frmFacturation", "cmdSupprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub ViderValeurs()

 On Error GoTo Oups

 txtClient.Text = vbNullString
 txtDescription.Text = vbNullString
 txtDateOuverture.Text = vbNullString
 txtDateFermeture.Text = vbNullString
 txtRaisonFermeture.Text = vbNullString

 Exit Sub

Oups:

 wOups "frmFacturation", "ViderValeurs", Err, Err.number, Err.Description
End Sub

Private Sub cmdVerrouiller_Click()
 
 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 
 Set rstProjSoum = New ADODB.Recordset
 
 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

 Select Case cmdVerrouiller.Caption
 Case "Verrouiller Soum": rstProjSoum.Fields("Verrouillé") = True
 Case "Verrouiller Proj": rstProjSoum.Fields("Verrouillé") = True
 Case "Déverrouiller Soum": rstProjSoum.Fields("Verrouillé") = False
 Case "Déverrouiller Proj": rstProjSoum.Fields("Verrouillé") = False
 End Select
 
 Call rstProjSoum.Update
 
 Call rstProjSoum.Close
  Set rstProjSoum = Nothing
 
  Call cmbNoProjSoum_Click

  Exit Sub

Oups:

  wOups "frmFacturation", "cmdVerrouiller_Click", Err, Err.number, Err.Description
End Sub

Private Sub Command1_Click()

Call vb_to_excel

End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 m_bLoading = True

 cmbProjSoum.ListIndex = I_CMB_PROJET
 cmdFacturerRectifier.Enabled = False
 cmdNCRectifier.Enabled = False
 optMontrer(I_OPT_OUVERT).Value = True
 optType(I_OPT_TYPE_TOUS).Value = True

 m_bLoading = False

 Call RemplirComboProjSoum

 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

  wOups "frmFacturation", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub cmbProjSoum_Click()

 On Error GoTo Oups

 'Rempli le cmbNoProjet avec les numéros de projet
 Select Case cmbProjSoum.ListIndex
 Case I_CMB_PROJET:
 m_eType = TYPE_PROJET
 lblTitreProjSoum.Caption = "Numéro de projet"
 cmdOuvrirProjSoum.Caption = "Ouvrir Projet"
 cmdFermerProjSoum.Caption = "Fermer Projet"
 fraMontrer.Caption = S_FRA_PROJ
 optMontrer(I_OPT_TOUS).Caption = S_PROJ_TOUS
 optMontrer(I_OPT_OUVERT).Caption = S_PROJ_OUVERT
 optMontrer(I_OPT_FERME).Caption = S_PROJ_FERME
 
 Case I_CMB_SOUMISSION:
  m_eType = TYPE_SOUMISSION
  lblTitreProjSoum.Caption = "Numéro de soumission"
  cmdOuvrirProjSoum.Caption = "Ouvrir Soum"
  cmdFermerProjSoum.Caption = "Fermer Soum"
  fraMontrer.Caption = S_FRA_SOUM
  optMontrer(I_OPT_TOUS).Caption = S_SOUM_TOUS
  optMontrer(I_OPT_OUVERT).Caption = S_SOUM_OUVERT
  optMontrer(I_OPT_FERME).Caption = S_SOUM_FERME
10 End Select
 
Call RemplirComboProjSoum

Exit Sub

Oups:

wOups "frmFacturation", "cmbProjSoum_Click", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboProjSoum()

 On Error GoTo Oups

 'Rempli le cmbNoProjet avec les numéros de projet et soumissions
 'Si bOuvert est à True, on affiche seulement ceux qui sont ouverts actuellement
 Dim rstProjet As ADODB.Recordset
 Dim sType As String
 Dim sWhere As String
 
 If m_bLoading = False Then
 Select Case m_eType
 Case TYPE_PROJET: sType = "P"
 Case TYPE_SOUMISSION: sType = "S"
 End Select
 
 If optMontrer(I_OPT_TOUS).Value = True Then
 sWhere = "Type = '" & sType & "'"
 Else
  If optMontrer(I_OPT_OUVERT).Value = True Then
  sWhere = "Ouvert = True AND Type = '" & sType & "'"
  Else
  sWhere = "Ouvert = False AND Type = '" & sType & "'"
  End If
  End If
 
  If optType(I_OPT_TYPE_ELECTRIQUE).Value = True Then
  sWhere = sWhere & " AND Left(IDProjSoum, 1) = 'E'"
Else
If optType(I_OPT_TYPE_MECANIQUE).Value = True Then
 sWhere = sWhere & " AND Left(IDProjSoum, 1) = 'M'"
 End If
 End If
 
 'Il faut vider le Combo avant de le remplir
 Call cmbNoProjSoum.Clear
 
 Set rstProjet = New ADODB.Recordset
 
 'Ouverture d'un recordset contenant les NoProjet
 Call rstProjet.Open("SELECT IDProjSoum, Ouvert FROM GrbProjSoum WHERE " & sWhere & " ORDER BY IDProjSoum", g_connData, adOpenDynamic, adLockOptimistic)
 
 Do While Not rstProjet.EOF
 'Ajout du numéro de projet dans le Combo
 Call cmbNoProjSoum.AddItem(rstProjet.Fields("IDProjSoum"))
 
 Call rstProjet.MoveNext
 Loop
 
Call rstProjet.Close
 Set rstProjet = Nothing
 
 'Si il y a des éléments dans le combo, on sélectionne le premier
 If cmbNoProjSoum.ListCount > 0 Then
 cmbNoProjSoum.ListIndex = 0
 Else
 Call lvwProjets.ListItems.Clear
 End If
1  End If
 
 Exit Sub

Oups:

 wOups "frmFacturation", "RemplirComboProjSoum", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListView(Optional ByVal o_iIndex As Integer = 1)

 On Error GoTo Oups

 'Remplissage du listView dépendamment du no de projet choisi
 Dim rstProjet As ADODB.Recordset
 Dim itmProjet As ListItem
 Dim lColor As Long
 Dim sDateDebut As String
 Dim sDateFin As String
 Dim sTotal As String
 
 'Il faut vider le listView avant de le remplir
 Call lvwProjets.ListItems.Clear
 
 sDateDebut = "TIMESERIAL(Left(GrbPunch.HeureDébut,2),RIGHT(GrbPunch.HeureDébut,2),0)"
 
 sDateFin = "TIMESERIAL(Left(GrbPunch.HeureFin,2),RIGHT(GrbPunch.HeureFin,2),0)"
 
 sTotal = "((" & sDateFin & " - " & sDateDebut & ")* 24) As Total"
 
 'Ouverture des enregistrements avec comme filtre, le numéro du projet
  Set rstProjet = New ADODB.Recordset
 
  rstProjet.CursorLocation = adUseServer
 
  Call rstProjet.Open("SELECT GrbPunch.*, " & sTotal & ", Grbemployés.initiale FROM Grbemployés INNER JOIN GrbPunch ON Grbemployés.noemploye = GrbPunch.NoEmploye WHERE NoProjet = '" & txtNoProjSoum.Text & "' ORDER BY [Date], HeureDébut, HeureFin", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
  Do While Not rstProjet.EOF
 'Vérification du champs "Facturé", si il est à vrai, il faut l'inscrire en
 'rouge dans le ListView, sinon, il faut l'inscrire en noir
  If rstProjet.Fields("Facturé") = "Vrai" Then
  lColor = COLOR_ROUGE
  Else
  lColor = COLOR_NOIR
End If
 
1 Set itmProjet = lvwProjets.ListItems.Add
 
 itmProjet.Tag = rstProjet.Fields("IDPunch")
 
 'Initiales de l'employé
 itmProjet.Text = rstProjet.Fields("Initiale")
 itmProjet.ForeColor = lColor
 
 'Date
 itmProjet.SubItems(I_LVW_DATE) = rstProjet.Fields("Date")
 itmProjet.ListSubItems(I_LVW_DATE).ForeColor = lColor
 
 'Début
 If Not IsNull(rstProjet.Fields("HeureDébut")) Then
 itmProjet.SubItems(I_LVW_DEBUT) = rstProjet.Fields("HeureDébut")
 Else
 itmProjet.SubItems(I_LVW_DEBUT) = vbNullString
 End If
 
itmProjet.ListSubItems(I_LVW_DEBUT).ForeColor = lColor
 
 'Fin
 If Not IsNull(rstProjet.Fields("HeureFin")) Then
 itmProjet.SubItems(I_LVW_FIN) = rstProjet.Fields("HeureFin")
 Else
 itmProjet.SubItems(I_LVW_FIN) = vbNullString
 End If
 
 itmProjet.ListSubItems(I_LVW_FIN).ForeColor = lColor
 
 'Description
1  If Not IsNull(rstProjet.Fields("Commentaire")) Then
 itmProjet.SubItems(I_LVW_DESCRIPTION) = rstProjet.Fields("Commentaire")
 Else
 itmProjet.SubItems(I_LVW_DESCRIPTION) = vbNullString
 End If
 
 itmProjet.ListSubItems(I_LVW_DESCRIPTION).ForeColor = lColor
 
 'Total
 If Not IsNull(rstProjet.Fields("Total")) Then
 itmProjet.SubItems(I_LVW_TOTAL) = Round(rstProjet.Fields("Total"), 2)
 Else
 itmProjet.SubItems(I_LVW_TOTAL) = vbNullString
 End If
 
 itmProjet.ListSubItems(I_LVW_TOTAL).ForeColor = lColor
 
 'Numéro de facture
 If Not IsNull(rstProjet.Fields("NoFacture")) Then
 itmProjet.SubItems(I_LVW_NO_FACTURE) = rstProjet.Fields("NoFacture")
 Else
 itmProjet.SubItems(I_LVW_NO_FACTURE) = vbNullString
 End If
 
itmProjet.ListSubItems(I_LVW_NO_FACTURE).ForeColor = lColor
 
 'Type
 If Not IsNull(rstProjet.Fields("Type")) Then
 itmProjet.SubItems(I_LVW_TYPE) = rstProjet.Fields("Type")
 Else
 itmProjet.SubItems(I_LVW_TYPE) = vbNullString
 End If
 
 itmProjet.ListSubItems(I_LVW_TYPE).ForeColor = lColor

 Call rstProjet.MoveNext
2  Loop

If lvwProjets.ListItems.count > 0 Then
Call lvwProjets.ListItems(o_iIndex).EnsureVisible
End If
 
Call rstProjet.Close
Set rstProjet = Nothing
 
Call CalculerTotaux

Exit Sub

Oups:

wOups "frmFacturation", "RemplirListView", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTotaux()

 On Error GoTo Oups

 Dim dblTotalFacture As Double
 Dim dblTotalNonFacture As Double
 Dim iCompteur As Integer
 
 For iCompteur = 1 To lvwProjets.ListItems.count
 If lvwProjets.ListItems(iCompteur).SubItems(I_LVW_NO_FACTURE) <> vbNullString Then
 dblTotalFacture = dblTotalFacture + CDbl(lvwProjets.ListItems(iCompteur).SubItems(I_LVW_TOTAL))
 Else
 dblTotalNonFacture = dblTotalNonFacture + CDbl(lvwProjets.ListItems(iCompteur).SubItems(I_LVW_TOTAL))
 End If
 Next
 
  lblHeuresFacturees.Caption = Round(dblTotalFacture, 2)
  lblHeuresNonFacturees.Caption = Round(dblTotalNonFacture, 2)

  Exit Sub

Oups:

  wOups "frmFacturation", "CalculerTotaux", Err, Err.number, Err.Description
End Sub

Private Sub VerifierSelection()

 On Error GoTo Oups

 'D'après les éléments sélectionner dans le ListView, cette méthode active
 'le bon bouton
 Dim iCompteur As Integer
 Dim iSelected As Integer
 Dim iFacture As Integer
 Dim iNC As Integer
 Dim iNon As Integer
 
 'Boucle servant à compter le nombre d'éléments sélectionnés dans le ListView,
 'le nombre d'éléments facturés et nombre d'éléments non facturés
 For iCompteur = 1 To lvwProjets.ListItems.count
 If lvwProjets.ListItems(iCompteur).Selected Then
 'Compte les éléments sélectionnés
 iSelected = iSelected + 1
 
 If lvwProjets.ListItems(iCompteur).SubItems(I_LVW_NO_FACTURE) = "NC" Then
 'Compte les nc
 iNC = iNC + 1
  Else
  If lvwProjets.ListItems(iCompteur).SubItems(I_LVW_NO_FACTURE) <> vbNullString Then
 'Compte les factures
  iFacture = iFacture + 1
  Else
 'Compte les non
  iNon = iNon + 1
  End If
  End If
  End If
10 Next
 
 'Si tous les éléments sélectionnés ont été facturés
If iSelected = iFacture Then
 cmdFacturerRectifier.Enabled = True
 cmdNCRectifier.Enabled = False

 cmdFacturerRectifier.Caption = S_RECTIFIER
Else
 'Si tous les éléments sélectionnés sont NC
 If iSelected = iNC Then
 cmdFacturerRectifier.Enabled = False
 cmdNCRectifier.Enabled = True

 cmdNCRectifier.Caption = S_RECTIFIER
 Else
 'Si tous les éléments sélectionnés n'ont pas été facturés
 If iSelected = iNon Then
 cmdFacturerRectifier.Enabled = True
 cmdNCRectifier.Enabled = True

 cmdFacturerRectifier.Caption = S_FACTURER
 cmdNCRectifier.Caption = S_NC
 Else
 'Si les éléments sélectionnés sont facturés ou non
 cmdFacturerRectifier.Enabled = False
 cmdNCRectifier.Enabled = False
1  End If
 End If
 End If

Exit Sub

Oups:

wOups "frmFacturation", "VerifierSelection", Err, Err.number, Err.Description
End Sub

Private Sub lvwProjets_ItemClick(ByVal Item As MSComctlLib.ListItem)

 On Error GoTo Oups

 'Vérification de la sélection lorsque un Item dans le ListView est cliqué
 Call VerifierSelection

 Exit Sub

Oups:

 wOups "frmFacturation", "lvwProjets_ItemClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwProjets_Click()

 On Error GoTo Oups

 'Vérification de la sélection lorsque un Item dans le ListView est cliqué
 'Cette méthode est importante puisque si l'utilisateur déclique un Item en tenant
 'la touche "Ctrl" enfoncé, ça ne passe pas dans l'événement ItemClick
 Call VerifierSelection

 Exit Sub

Oups:

 wOups "frmFacturation", "lvwProjets_Click", Err, Err.number, Err.Description
End Sub

Private Sub optMontrer_Click(Index As Integer)

 On Error GoTo Oups

 Dim bFermeture As Boolean

 Call ViderValeurs

 Select Case Index
 Case I_OPT_TOUS:
 Call RemplirComboProjSoum

 bFermeture = True
 
 Case I_OPT_OUVERT:
 Call RemplirComboProjSoum

 bFermeture = False
 
 Case I_OPT_FERME:
 Call RemplirComboProjSoum

 bFermeture = True
 End Select

  lblDateFermeture.Visible = bFermeture
  txtDateFermeture.Visible = bFermeture
  lblRaisonFermeture.Visible = bFermeture
  txtRaisonFermeture.Visible = bFermeture

  cmdReouverture.Visible = bFermeture

  Exit Sub

Oups:

  wOups "frmFacturation", "optMontrer_Click", Err, Err.number, Err.Description
End Sub

Private Sub optType_Click(Index As Integer)
 
 On Error GoTo Oups

 Call RemplirComboProjSoum

 Exit Sub

Oups:

 wOups "frmFacturation", "optType_Click", Err, Err.number, Err.Description
End Sub

Private Function ValiderFormatSoumission(ByVal sNoSoumission As String) As Boolean
 
 On Error GoTo Oups

 If Mid$(sNoSoumission, 3, 1) = "1" Then
 ValiderFormatSoumission = True
 Else
 Call MsgBox("Une soumission doit absolument avoir '1' comme 3e caractère !", vbOKOnly, "Erreur")

 ValiderFormatSoumission = False
 End If

 Exit Function

Oups:

 wOups "FrmFacturation", "ValiderFormatSoumission", Err, Err.number, Err.Description
End Function

Private Function ValiderFormatProjet(ByVal sNoProjet As String) As Boolean
 
 On Error GoTo Oups

 Dim iType As Integer

 iType = Mid$(sNoProjet, 3, 1)

 If iType = 4 Or iType = 5 Or iType =   Or iType =   Then
 ValiderFormatProjet = True
 Else
 Call MsgBox("Un projet ouvert dans cet écran doit absolument avoir '4', '5', '7' ou '9' comme 3e caractère !", vbOKOnly, "Erreur")

 ValiderFormatProjet = False
 End If

 Exit Function

Oups:

 wOups "FrmFacturationElec", "ValiderFormatProjet", Err, Err.number, Err.Description
End Function
