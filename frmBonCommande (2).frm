VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBonCommande 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bon de commande"
   ClientHeight    =   7050
   ClientLeft      =   -15
   ClientTop       =   -15
   ClientWidth     =   9750
   Icon            =   "frmBonCommande.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   9750
   Begin MSComCtl2.MonthView mvwDate 
      Height          =   2370
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      StartOfWeek     =   152633345
      CurrentDate     =   38310
   End
   Begin VB.TextBox txtTotal 
      Height          =   285
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txtCommentaire 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   31
      Text            =   "frmBonCommande.frx":000C
      Top             =   6120
      Width           =   5415
   End
   Begin VB.CommandButton cmdInstructions 
      Caption         =   "Modifier"
      Height          =   375
      Left            =   5280
      TabIndex        =   26
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Height          =   375
      Left            =   7080
      TabIndex        =   32
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   8400
      TabIndex        =   33
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdDate 
      Caption         =   "..."
      Height          =   285
      Left            =   3120
      TabIndex        =   21
      Top             =   1680
      Width           =   375
   End
   Begin VB.CheckBox chkAfficherInstructions 
      BackColor       =   &H00000000&
      Caption         =   "Afficher les instructions de livraison"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Frame fraImpression 
      BackColor       =   &H00000000&
      Caption         =   "Impression"
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
      Height          =   735
      Left            =   7080
      TabIndex        =   23
      Top             =   1560
      Width           =   2535
      Begin VB.OptionButton optImpression 
         BackColor       =   &H00000000&
         Caption         =   "Anglais"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton optImpression 
         BackColor       =   &H00000000&
         Caption         =   "Français"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.TextBox txtComPar 
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtNoBC 
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtFax 
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtTelephone 
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtVotreNoSoum 
      Height          =   285
      Left            =   4440
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtDateRequise 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Top             =   1200
      Width           =   1815
   End
   Begin VB.ComboBox cmbContact 
      Height          =   315
      Left            =   1200
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.ComboBox cmbFournisseur 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin MSComctlLib.ListView lvwBonCommande 
      Height          =   3375
      Left            =   120
      TabIndex        =   27
      Top             =   2400
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   5953
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Qté"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No Item"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Manufact"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Prix"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Escompte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "TOTAL :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   30
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Commentaires :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Commandé par :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   17
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblNoBC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "# BC :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fax :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Téléphone :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblVotreNoSoum 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Votre # Soum :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date Requise :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Transport :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contacts :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fournisseurs :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmBonCommande"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Index des colonnes du listview
Private Const I_COL_QUANTITE As Integer = 0
Private Const I_COL_NO_ITEM As Integer = 1
Private Const I_COL_DESCR As Integer = 2
Private Const I_COL_MANUFACT As Integer = 3
Private Const I_COL_PRIX As Integer = 4
Private Const I_COL_ESCOMPTE As Integer = 5
Private Const I_COL_TOTAL As Integer = 6

'Index de optImpression
Private Const I_IMP_FRANCAIS As Integer = 0
Private Const I_IMP_ANGLAIS As Integer = 1

Public Enum enumFormSource
 I_PROJET_MEC = 0
 I_PROJET_ELEC = 1
 I_ACHAT_MEC = 2
 I_ACHAT_ELEC = 3
 I_INVENTAIRE_MEC = 4
 I_INVENTAIRE_ELEC = 5
 I_RETOUR_MARCHANDISE = 6
End Enum

Private Enum enumLangage
 FRANCAIS = 0
 ANGLAIS = 1
End Enum
 
'Pour savoir le no de projet
Private m_sNoProjet As String

Private m_sNoAchat As String
Private m_iIndexAchat As Integer

'Pour savoir le type (Électrique ou mécanique)
Public m_eForm As enumFormSource

'Pour savoir si le form vient d'être ouvert
Private m_bOuverture As Boolean

'No du fournisseur sélectionné dans le combo
Private m_iNoFRS As Integer

'Pièces à partir d'un projet
Private m_collPieces As Collection
Private m_collNoLigne As Collection

Private m_sEmploye As String

Private m_eImpRetour As enumImpressionRetour

Private m_eLangage As enumLangage

Private Sub cmbFournisseur_Click()

 On Error GoTo Oups

 Screen.MousePointer = vbHourglass
 
 'Il ne faut pas enregistrer les modifications sur l'ouverture du form
 'parce qu'il n'y a pas eu de modifications encore
 
 'Si le form ne vient pas d'etre ouvert
 If m_bOuverture = False Then
 'On enregistre les modifications
 Call EnregistrerModifFournisseur
 Else
 m_bOuverture = False
 End If

 m_iNoFRS = cmbFournisseur.ItemData(cmbFournisseur.ListIndex)

 'Rempli le combo des contacts
 Call RemplirComboContacts
 
 'On affiche le contenu du fournisseur sélectionné
 Call AfficherContenuFournisseur
 
 Screen.MousePointer = vbDefault

  Exit Sub

Oups:

  wOups "frmBonCommande", "cmbFournisseur_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDate_Click()

 On Error GoTo Oups

 mvwDate.Value = Date
 
 mvwDate.Visible = True
 
 Call mvwDate.SetFocus

 Exit Sub

Oups:

 wOups "frmBonCommande", "cmdDate_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 Dim sNoBC As String

 If m_eForm = I_RETOUR_MARCHANDISE Then
 sNoBC = txtVotreNoSoum.Text
 Else
 sNoBC = txtNoBC.Text
 End If

 Call g_connData.Execute("DELETE * FROM GrbBonsCommandes_Pieces WHERE NoBonCommande = '" & sNoBC & "'")
 Call g_connData.Execute("DELETE * FROM GrbBonsCommandes WHERE NoBonCommande = '" & sNoBC & "'")

 Call Unload(Me)

 Exit Sub

Oups:

  wOups "frmBonCommande", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Public Sub AfficherFormProjetAchat(ByVal sNoProjet As String, ByVal sNoBonCommande As String, ByVal collPiece As Collection, ByVal collNoLigne As Collection, ByVal eForm As enumFormSource, ByVal iLangage As Integer)

 On Error GoTo Oups

 If eForm = I_ACHAT_ELEC Or eForm = I_ACHAT_MEC Then
 m_sNoAchat = Left$(sNoProjet, 9)
 m_iIndexAchat = CInt(Right$(sNoProjet, 3))
 Else
 m_sNoProjet = sNoProjet
 End If
 
 m_eForm = eForm

 m_eLangage = iLangage
 
 Set m_collPieces = collPiece

 Set m_collNoLigne = collNoLigne
 
  m_bOuverture = True
 
  txtNoBC.Text = sNoBonCommande

  Call g_connData.Execute("DELETE * FROM GrbBonsCommandes WHERE NoBonCommande = '" & txtNoBC.Text & "'")
  Call g_connData.Execute("DELETE * FROM GrbBonsCommandes_Pieces WHERE NoBonCommande = '" & txtNoBC.Text & "'")
 
 'Enregistrement du bon de commande
  If eForm = I_ACHAT_ELEC Or eForm = I_ACHAT_MEC Then
  Call EnregistrerBonCommandeAchat
  Else
  Call EnregistrerBonCommandeProjet
10 End If
 
 'On rempli les fournisseurs
Call RemplirComboFournisseurs
 
 'Affichage du form modalement
Call Me.Show(vbModal)

Exit Sub

Oups:

wOups "frmBonCommande", "AfficherFormProjet", Err, Err.number, Err.Description
End Sub

Public Sub AfficherFormRetourMarchandiseProjet(ByVal sNoProjet As String, ByVal sNoBonCommande As String, ByVal collPiece As Collection, ByVal collNoLigne As Collection, ByVal sUserID As String, ByVal eImpRetour As enumImpressionRetour)

 On Error GoTo Oups

 Dim rstEmploye As ADODB.Recordset

 Me.Caption = "Retour de marchandise"

 lblVotreNoSoum.Caption = "Notre # : "

 lblNoBC.Caption = "# RMA : "

 m_eImpRetour = eImpRetour

 m_sNoProjet = Right$(sNoProjet, Len(sNoProjet) - 1)
 
 m_eForm = I_RETOUR_MARCHANDISE
 
 Set m_collPieces = collPiece

 Set m_collNoLigne = collNoLigne

 m_bOuverture = True

  Set rstEmploye = New ADODB.Recordset

  Call rstEmploye.Open("SELECT Employe FROM GrbEmployés WHERE loginname = '" & sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
  m_sEmploye = rstEmploye.Fields("Employe")

  Call rstEmploye.Close
  Set rstEmploye = Nothing
 
  txtVotreNoSoum.Text = sNoBonCommande

  txtVotreNoSoum.Locked = True

  txtNoBC.Locked = False
 
10 Call g_connData.Execute("DELETE * FROM GrbBonsCommandes WHERE NoBonCommande = '" & txtVotreNoSoum.Text & "'")
Call g_connData.Execute("DELETE * FROM GrbBonsCommandes_Pieces WHERE NoBonCommande = '" & txtVotreNoSoum.Text & "'")
 
 'Enregistrement du bon de commande
Call EnregistrerBonCommandeRetourMarchandiseProjet
 
 'On rempli les fournisseurs
Call RemplirComboFournisseurs
 
 'Affichage du form modalement
Call Me.Show(vbModal)

Exit Sub

Oups:

wOups "frmBonCommande", "AfficherFormRetourMarchandiseProjet", Err, Err.number, Err.Description
End Sub

Public Sub AfficherFormRetourMarchandiseAchat(ByVal sNoAchat As String, ByVal iIndexAchat As Integer, ByVal sNoBonCommande As String, ByVal collPiece As Collection, ByVal collNoLigne As Collection, ByVal sUserID As String, ByVal eImpRetour As enumImpressionRetour)

 On Error GoTo Oups

 Dim rstEmploye As ADODB.Recordset

 Me.Caption = "Retour de marchandise"

 lblVotreNoSoum.Caption = "Notre # : "

 lblNoBC.Caption = "# RMA : "

 m_eImpRetour = eImpRetour

 m_sNoAchat = sNoAchat
 m_iIndexAchat = iIndexAchat
 
 m_eForm = I_RETOUR_MARCHANDISE
 
 Set m_collPieces = collPiece

 Set m_collNoLigne = collNoLigne

  m_bOuverture = True

  Set rstEmploye = New ADODB.Recordset

  Call rstEmploye.Open("SELECT Employe FROM GrbEmployés WHERE loginname = '" & sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
  m_sEmploye = rstEmploye.Fields("Employe")

  Call rstEmploye.Close
  Set rstEmploye = Nothing
 
  txtVotreNoSoum.Text = sNoBonCommande

  txtVotreNoSoum.Locked = True

10 txtNoBC.Locked = False

Call g_connData.Execute("DELETE * FROM GrbBonsCommandes WHERE NoBonCommande = '" & txtVotreNoSoum.Text & "'")
Call g_connData.Execute("DELETE * FROM GrbBonsCommandes_Pieces WHERE NoBonCommande = '" & txtVotreNoSoum.Text & "'")
 
 'Enregistrement du bon de commande
Call EnregistrerBonCommandeRetourMarchandiseAchat
 
 'On rempli les fournisseurs
Call RemplirComboFournisseurs
 
 'Affichage du form modalement
Call Me.Show(vbModal)

Exit Sub

Oups:

wOups "frmBonCommande", "AfficherFormRetourMarchandiseAchat", Err, Err.number, Err.Description
End Sub


Private Sub AfficherContenuFournisseur()

 On Error GoTo Oups

 Dim rstBC As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim sNoBC As String

 If m_eForm = I_RETOUR_MARCHANDISE Then
 sNoBC = txtVotreNoSoum.Text
 Else
 sNoBC = txtNoBC.Text
 End If

 Set rstBC = New ADODB.Recordset
 
 Call rstBC.Open("SELECT * FROM GrbBonsCommandes WHERE NoFournisseur = " & m_iNoFRS & " AND NoBonCommande = '" & sNoBC & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 'À l'attention de
  If rstBC.Fields("Attention") <> vbNullString Then
  If ComboContient(cmbContact, rstBC.Fields("Attention")) = True Then
  cmbContact.Text = rstBC.Fields("Attention")
  Else
  cmbContact.ListIndex = -1
  End If
  Else
  cmbContact.ListIndex = -1
10 End If
 
 'Transport
If Not IsNull(rstBC.Fields("Transport")) Or Trim(rstBC.Fields("Transport")) <> vbNullString Then
 txtTransport.Text = rstBC.Fields("Transport")
Else
 txtTransport.Text = vbNullString
End If
 
 'Date requise
If Not IsNull(rstBC.Fields("DateRequise")) Or Trim(rstBC.Fields("DateRequise")) <> vbNullString Then
 txtDateRequise.Text = rstBC.Fields("DateRequise")
Else
 txtDateRequise.Text = vbNullString
End If
 
 'Votre # Soum
If m_eForm = I_RETOUR_MARCHANDISE Then
If Not IsNull(rstBC.Fields("VotreNoSoum")) Then
 txtNoBC.Text = rstBC.Fields("VotreNoSoum")
 Else
 txtNoBC.Text = vbNullString
 End If
Else
 If Not IsNull(rstBC.Fields("VotreNoSoum")) Then
1  txtVotreNoSoum.Text = rstBC.Fields("VotreNoSoum")
 Else
 txtVotreNoSoum.Text = vbNullString
 End If
End If
 
 'Numéro de tel et fax du fournisseur
Set rstFRS = New ADODB.Recordset

Call rstFRS.Open("SELECT Telephonne, Fax FROM GrbFournisseur WHERE IDFRS = " & cmbFournisseur.ItemData(cmbFournisseur.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
 
txtTelephone.Text = rstFRS.Fields("Telephonne")
txtFax.Text = rstFRS.Fields("Fax")

Call rstFRS.Close
Set rstFRS = Nothing
 
 'Date
txtDate.Text = rstBC.Fields("DateCommande")
 
 'Commandé par
txtComPar.Text = rstBC.Fields("CommandePar")
 
 'Commentaire
2  If Not IsNull(rstBC.Fields("Commentaire")) Then
 txtcommentaire.Text = rstBC.Fields("Commentaire")
2  Else
 txtcommentaire.Text = vbNullString
2  End If
 
 'Total
txtTotal.Text = Conversion(rstBC.Fields("Total"), MODE_ARGENT)
 
 'Afficher les instructions de livraison
2  chkAfficherInstructions.Value = Abs(CInt(rstBC.Fields("AffichageInstructions")))
 
 'Langue d'impression
If rstBC.Fields("LangueImpression") = "Français" Then
optImpression(I_IMP_FRANCAIS).Value = True
Else
 optImpression(I_IMP_ANGLAIS).Value = True
End If
 
Call rstBC.Close
Set rstBC = Nothing
 
Call RemplirListView

Exit Sub

Oups:

wOups "frmBonCommande", "AfficherContenuFournisseur", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListView()

 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 Dim itmPiece As ListItem
 Dim iCompteur As Integer
 Dim dblEscompte As Double
 Dim dblPrix As Double
 Dim sNoBC As String

 If m_eForm = I_RETOUR_MARCHANDISE Then
 sNoBC = txtVotreNoSoum.Text
 Else
 sNoBC = txtNoBC.Text
  End If
 
  Call lvwBonCommande.ListItems.Clear
 
  Set rstPiece = New ADODB.Recordset
 
  Call rstPiece.Open("SELECT * FROM GrbBonsCommandes_Pieces WHERE NoBonCommande = '" & sNoBC & "' AND NoFournisseur = " & m_iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)
 
  Do While Not rstPiece.EOF
  If Not IsNull(rstPiece.Fields("Qté")) Then
  Set itmPiece = lvwBonCommande.ListItems.Add
 
 'Quantité
  itmPiece.Text = rstPiece.Fields("Qté")

 If Not IsNull(rstPiece.Fields("NuméroLigne")) Then
 itmPiece.Tag = rstPiece.Fields("NuméroLigne")
 End If

 'No. Item
 itmPiece.SubItems(I_COL_NO_ITEM) = rstPiece.Fields("NoItem")
 
 'Description
 If Not IsNull(rstPiece.Fields("Description")) Then
 itmPiece.SubItems(I_COL_DESCR) = rstPiece.Fields("Description")
 Else
 itmPiece.SubItems(I_COL_DESCR) = ""
 End If
 
 'Manufacturier
 itmPiece.SubItems(I_COL_MANUFACT) = rstPiece.Fields("Manufact")
 
 'Prix/unité
 If Not IsNull(rstPiece.Fields("Prix")) Then
 itmPiece.SubItems(I_COL_PRIX) = Conversion(rstPiece.Fields("Prix"), MODE_ARGENT, 4)
 Else
 itmPiece.SubItems(I_COL_PRIX) = Conversion(0, MODE_ARGENT, 4)
 End If
 
 'Escompte
 If Trim(rstPiece.Fields("Escompte")) <> vbNullString Then
 itmPiece.SubItems(I_COL_ESCOMPTE) = Conversion(rstPiece.Fields("Escompte"), MODE_POURCENT)
 Else
 itmPiece.SubItems(I_COL_ESCOMPTE) = " "
1  End If
 
 'Total
 If Not IsNull(rstPiece.Fields("Total")) Then
 itmPiece.SubItems(I_COL_TOTAL) = Conversion(rstPiece.Fields("Total"), MODE_ARGENT)
 Else
 itmPiece.SubItems(I_COL_TOTAL) = Conversion(0, MODE_ARGENT)
 End If
 End If

 Call rstPiece.MoveNext
Loop
 
Call rstPiece.Close
Set rstPiece = Nothing

Exit Sub

Oups:

wOups "frmBonCommande", "RemplirListView", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboContacts()

 On Error GoTo Oups

 Dim rstContact As ADODB.Recordset
 Dim rstContactFRS As ADODB.Recordset
 
 Call cmbContact.Clear
 
 Set rstContactFRS = New ADODB.Recordset
 Set rstContact = New ADODB.Recordset
 
 Call rstContactFRS.Open("SELECT * FROM GrbContactFRS WHERE NoFRS = " & m_iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)

 Do While Not rstContactFRS.EOF
 Call rstContact.Open("SELECT NomContact FROM GrbContact WHERE IDContact = " & rstContactFRS.Fields("NoContact"), g_connData, adOpenDynamic, adLockOptimistic)

 If Not rstContact.EOF Then
 Call cmbContact.AddItem(rstContact.Fields("NomContact"))
  End If

  Call rstContact.Close

  Call rstContactFRS.MoveNext
  Loop

  Call rstContactFRS.Close
 
  If cmbContact.ListCount = 0 Then
  Call rstContact.Open("SELECT NomContact FROM GrbContact WHERE Supprimé = False ORDER BY NomContact", g_connData, adOpenDynamic, adLockOptimistic)
 
  Do While Not rstContact.EOF
 Call cmbContact.AddItem(rstContact.Fields("NomContact"))
 
Call rstContact.MoveNext
 Loop
 
 Call rstContact.Close
End If

Set rstContact = Nothing

Exit Sub

Oups:

wOups "frmBonCommande", "RemplirComboContact", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboFournisseurs()

 On Error GoTo Oups

 Dim rstBC As ADODB.Recordset
 Dim sNoBC As String
 
 If m_eForm = I_RETOUR_MARCHANDISE Then
 sNoBC = txtVotreNoSoum.Text
 Else
 sNoBC = txtNoBC.Text
 End If

 Set rstBC = New ADODB.Recordset
 
 Call rstBC.Open("SELECT NoFournisseur, NomFournisseur FROM GrbBonsCommandes INNER JOIN GrbFournisseur ON GrbBonsCommandes.NoFournisseur = GrbFournisseur.IDFRS WHERE NoBonCommande = '" & sNoBC & "' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Pour chaque enregistrements du recordset
 Do While Not rstBC.EOF
 'On ajoute le nom dans le combo
  Call cmbFournisseur.AddItem(rstBC.Fields("NomFournisseur"))
 
 'On ajoute le no dans l'itemdata du combo
  cmbFournisseur.ItemData(cmbFournisseur.newIndex) = rstBC.Fields("NoFournisseur")
 
  Call rstBC.MoveNext
  Loop
 
  Call rstBC.Close
  Set rstBC = Nothing

 'Si le combo n'est pas vide
  If cmbFournisseur.ListCount > 0 Then
 'On sélectionne le premier enregistrement
  cmbFournisseur.ListIndex = 0
10 End If

Exit Sub

Oups:

wOups "frmBonCommande", "RemplirComboFournisseurs", Err, Err.number, Err.Description
End Sub

Private Sub cmdImprimer_Click()

 On Error GoTo Oups

 Dim rstConfig As ADODB.Recordset
 Dim rstBC As ADODB.Recordset
 Dim rstBCPiece As ADODB.Recordset
 Dim rstFournisseur As ADODB.Recordset
 Dim bGRB As Boolean
 Dim iCompteur As Integer
 Dim sNoBC As String
 
 Screen.MousePointer = vbHourglass
 
 If m_eForm = I_RETOUR_MARCHANDISE Then
 sNoBC = txtVotreNoSoum.Text
  Else
  sNoBC = txtNoBC.Text
  End If
 
 'Sur l'impression, on enregistre une dernière fois le bon de commande
  Call EnregistrerModifFournisseur
 
  Set rstBC = New ADODB.Recordset
 
  If m_eForm = I_RETOUR_MARCHANDISE Then
  If m_eImpRetour = MODE_RETOUR Then
  Call rstBC.Open("SELECT * FROM GrbBonsCommandes WHERE NoBonCommande = '" & sNoBC & "' AND NoFournisseur = " & cmbFournisseur.ItemData(cmbFournisseur.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
Else
Call rstBC.Open("SELECT * FROM GrbBonsCommandes WHERE NoBonCommande = '" & sNoBC & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If
Else
 Call rstBC.Open("SELECT * FROM GrbBonsCommandes WHERE NoBonCommande = '" & sNoBC & "'", g_connData, adOpenDynamic, adLockOptimistic)
End If
 
If m_eForm = I_ACHAT_ELEC Or m_eForm = I_PROJET_ELEC Then
 Do While Not rstBC.EOF
 If rstBC.Fields("DateRequise") = "" Then
 Set rstFournisseur = New ADODB.Recordset

 Call rstFournisseur.Open("SELECT NomFournisseur FROM GrbFournisseur WHERE IDFRS = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)

 Call MsgBox("Le date requise est nécessaire pour le fournisseur " & rstFournisseur.Fields("NomFournisseur") & "!", vbOKOnly, "Erreur")
 
 Call rstFournisseur.Close
 Set rstFournisseur = Nothing

 Call rstBC.Close
 Set rstBC = Nothing

 Screen.MousePointer = vbDefault

 Exit Sub
 End If

1  Call rstBC.MoveNext
 Loop
 
 Call rstBC.MoveFirst
End If
 
Set rstBCPiece = New ADODB.Recordset
Set rstFournisseur = New ADODB.Recordset
Set rstConfig = New ADODB.Recordset
 
rstBCPiece.CursorLocation = adUseClient
 
Do While Not rstBC.EOF
 Call rstBCPiece.Open("SELECT * FROM GrbBonsCommandes_Pieces WHERE NoBonCommande = '" & sNoBC & "' AND NoFournisseur = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)
 ''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' Met au minimum 15 lignes pour un bon de commande '
 ''''''''''''''''''''''''''''''''''''''''''''''''''''
 If rstBCPiece.RecordCount < 15 Then
 iCompteur = 15 - rstBCPiece.RecordCount
 
 Do While Not iCompteur = 0
 'Ajoute une ligne vide
 Call rstBCPiece.AddNew
 
 rstBCPiece.Fields("NoBonCommande") = rstBC.Fields("NoBonCommande")
 rstBCPiece.Fields("NoFournisseur") = rstBC.Fields("NoFournisseur")
 rstBCPiece.Fields("Type") = rstBC.Fields("Type")
 
 Call rstBCPiece.Update
 
 iCompteur = iCompteur - 1
 Loop
 End If
 
 'Ouvre la table fournisseur
Call rstFournisseur.Open("SELECT * FROM GrbFournisseur WHERE IDFRS = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'Ouvre la table config
3 Call rstConfig.Open("SELECT parcel_label_line1, parcel_label_line2, parcel_label_line3, ParcelAssist, ParcelEtat FROM GrbConfig", g_connData, adOpenDynamic, adLockOptimistic)
 
 ''''''''''''''''''''''''''''''''''''''
 ' U.S. PARCEL SERVICE SHIPMENTS ONLY '
 ''''''''''''''''''''''''''''''''''''''
 If rstBC.Fields("AffichageInstructions") = True Then
 'Orientation de la page
 'Printer.Orientation = vbPRORPortrait
 DR_Commande_parcel.Orientation = rptOrientPortrait
 
 'Affiche les données
 DR_Commande_parcel.Sections("section4").Controls("lblcompagnie").Caption = rstConfig.Fields("parcel_label_line1")
 DR_Commande_parcel.Sections("section4").Controls("lbladresse").Caption = rstConfig.Fields("parcel_label_line2")
 DR_Commande_parcel.Sections("section4").Controls("lblpays").Caption = rstConfig.Fields("parcel_label_line3")
 DR_Commande_parcel.Sections("section4").Controls("lblassist").Caption = "Should you have any questions, do not hesitate to call " & rstConfig.Fields("ParcelAssist") & ", it will be our pleasure to assist you."
 DR_Commande_parcel.Sections("section4").Controls("lblreminder").Caption = "Please note that you are shipping to a " & rstConfig.Fields("ParcelEtat") & " address and therefore your shipment is considered as domestic."
 
 'Ouvre le rapport
 Set DR_Commande_parcel.DataSource = rstConfig
 
 Call DR_Commande_parcel.Show(vbModal)
 End If
 
 ''''''''''''
 ' Commande '
 ''''''''''''
If m_eForm = I_RETOUR_MARCHANDISE Then
 If m_eImpRetour = MODE_DEMANDE_RETOUR Then
 DR_Commande.Caption = "Demande de retour de marchandise"
 Else
 DR_Commande.Caption = "Retour de marchandise"
 End If
 Else
 DR_Commande.Caption = "Commande"
End If
 
4 If rstBC.Fields("LangueImpression") = "Anglais" Then
4 If m_eForm = I_RETOUR_MARCHANDISE Then
4 DR_Commande.Sections("Section2").Controls("lblTitrebc").Caption = "RMA #"

4 If m_eImpRetour = MODE_DEMANDE_RETOUR Then
4 DR_Commande.Sections("Section2").Controls("lblTitreCommande").Caption = "RMA REQUEST"
4 Else
4 DR_Commande.Sections("Section2").Controls("lblTitreCommande").Caption = "RETURN ORDER"
4 End If

4 DR_Commande.Sections("section2").Controls("lbltitreNoSoum").Caption = "Our #"
4 Else
4 DR_Commande.Sections("Section2").Controls("lbltitrebc").Caption = "PO #"
4  DR_Commande.Sections("Section2").Controls("lbltitrecommande").Caption = "PURCHASE ORDER"
4  DR_Commande.Sections("section2").Controls("lbltitreNoSoum").Caption = "Your ref #"
4  End If

4  DR_Commande.Sections("Section3").Controls("lbltitreCommentaire").Caption = "Comments:"
4  DR_Commande.Sections("section2").Controls("lbltitrecompar").Caption = "Purchaser:"
4  DR_Commande.Sections("section2").Controls("lbltitrecontact").Caption = "ATT:"
4  DR_Commande.Sections("section2").Controls("lbltitredate").Caption = "Date:"
4  DR_Commande.Sections("section2").Controls("lbltitredatereq").Caption = "Date required"
50 DR_Commande.Sections("section2").Controls("lbltitredescription").Caption = "Description"
DR_Commande.Sections("section2").Controls("lbltitreescompte").Caption = "Discount"
 DR_Commande.Sections("section2").Controls("lbltitrefax").Caption = "Fax"
 DR_Commande.Sections("section2").Controls("lbltitrefournisseur").Caption = "SUPPLIER:"
 DR_Commande.Sections("section2").Controls("lbltitremanufact").Caption = "Manufacturer"
 DR_Commande.Sections("section2").Controls("lbltitrePiece").Caption = "Part Number"
 DR_Commande.Sections("section2").Controls("lbltitrepage").Caption = "Page:"
 DR_Commande.Sections("section2").Controls("lblPage").Caption = "%p of %P"
 DR_Commande.Sections("section2").Controls("lbltitreprix").Caption = "Unit Price"
 DR_Commande.Sections("section2").Controls("lbltitreqte").Caption = "Qty"
 DR_Commande.Sections("section2").Controls("lbltitretel").Caption = "Telephone"
 DR_Commande.Sections("section2").Controls("lbltitretotal").Caption = "Total"
5  DR_Commande.Sections("Section3").Controls("lbltitretotalfin").Caption = "TOTAL"
5  DR_Commande.Sections("section2").Controls("lbltitretransport").Caption = "TRANSPORT"
5  DR_Commande.Sections("Section3").Controls("lbltypeprix").Caption = rstFournisseur.Fields("pays") + " Funds"
5  DR_Commande.Sections("Section3").Controls("lblPiedPage").Caption = "CONFIRM THE ORDER AND SHIPPING DATE"

5  DR_Commande.Sections("Section2").Controls("imgLogoFrancais").Visible = False
5  DR_Commande.Sections("Section2").Controls("imgLogoAnglais").Visible = True
5  Else
5  If m_eForm = I_RETOUR_MARCHANDISE Then
60 DR_Commande.Sections("Section2").Controls("lblTitreBC").Caption = "# RMA"

  If m_eImpRetour = MODE_DEMANDE_RETOUR Then
  DR_Commande.Sections("Section2").Controls("lblTitreCommande").Caption = "DEMANDE DE RETOUR DE MARCHANDISE"
  Else
  DR_Commande.Sections("Section2").Controls("lblTitreCommande").Caption = "RETOUR DE MARCHANDISE"
  End If

  DR_Commande.Sections("Section2").Controls("lblTitreNoSoum").Caption = "Notre #"
  Else
  DR_Commande.Sections("Section2").Controls("lblTitreBC").Caption = "BC #"
  DR_Commande.Sections("Section2").Controls("lblTitreCommande").Caption = "COMMANDE"
  DR_Commande.Sections("Section2").Controls("lblTitreNoSoum").Caption = "Votre # soum"
  End If
6  End If

6  If m_eForm = I_RETOUR_MARCHANDISE Then
6  If m_eImpRetour = MODE_RETOUR Then
6  DR_Commande.Sections("Section3").Controls("lblCopieCredit").Visible = True
6  End If
6  End If

6  If m_eForm = I_RETOUR_MARCHANDISE Then
6  If Not IsNull(rstBC.Fields("VotreNoSoum")) Then
70 DR_Commande.Sections("Section2").Controls("lblNoBC").Caption = rstBC.Fields("VotreNoSoum")
  Else
  DR_Commande.Sections("Section2").Controls("lblNoBC").Caption = vbNullString
  End If
  Else
  DR_Commande.Sections("section2").Controls("lblNoBC").Caption = rstBC.Fields("NoBonCommande")
  End If

  If m_eForm = I_RETOUR_MARCHANDISE Then
  DR_Commande.Sections("Section3").Controls("lblPiedPage").Visible = False
  Else
  DR_Commande.Sections("Section3").Controls("lblPiedPage").Visible = True
   End If

7  If Not IsNull(rstBC.Fields("Commentaire")) Then
7  DR_Commande.Sections("Section3").Controls("lblCommentaire").Caption = rstBC.Fields("Commentaire")
7  Else
7  DR_Commande.Sections("Section3").Controls("lblCommentaire").Caption = vbNullString
7  End If

7  DR_Commande.Sections("section2").Controls("lblCommandePar").Caption = rstBC.Fields("CommandePar")

80 If Not IsNull(rstBC.Fields("Attention")) Then
  DR_Commande.Sections("section2").Controls("lblContact").Caption = rstBC.Fields("Attention")
  Else
  DR_Commande.Sections("Section2").Controls("lblContact").Caption = vbNullString
  End If

  DR_Commande.Sections("Section2").Controls("lblDate").Caption = rstBC.Fields("DateCommande")
 
  If Not IsNull(rstBC.Fields("DateRequise")) Then
  DR_Commande.Sections("Section2").Controls("lblDateRequise").Caption = rstBC.Fields("DateRequise")
  Else
  DR_Commande.Sections("Section2").Controls("lblDateRequise").Caption = vbNullString
  End If
 
  DR_Commande.Sections("section2").Controls("lblFax").Caption = rstFournisseur.Fields("Fax")
   DR_Commande.Sections("section2").Controls("lblFournisseur").Caption = rstFournisseur.Fields("NomFournisseur")
   DR_Commande.Sections("section2").Controls("lblTel").Caption = rstFournisseur.Fields("telephonne")

   If m_eForm = I_RETOUR_MARCHANDISE Then
   DR_Commande.Sections("Section2").Controls("lblNoSoum").Caption = rstBC.Fields("NoBonCommande")
8  Else
8  If Not IsNull(rstBC.Fields("VotreNoSoum")) Then
8  DR_Commande.Sections("Section2").Controls("lblNoSoum").Caption = rstBC.Fields("VotreNoSoum")
8  Else
90 DR_Commande.Sections("Section2").Controls("lblNoSoum").Caption = vbNullString
  End If
  End If

  DR_Commande.Sections("Section3").Controls("lblTotalFin").Caption = Conversion(rstBC.Fields("total"), MODE_ARGENT)
 
  If Not IsNull(rstBC.Fields("Transport")) Then
  DR_Commande.Sections("section2").Controls("lblTransport").Caption = rstBC.Fields("Transport")
  Else
  DR_Commande.Sections("section2").Controls("lblTransport").Caption = " "
  End If

  If m_eForm = I_ACHAT_ELEC Or m_eForm = I_INVENTAIRE_ELEC Or m_eForm = I_PROJET_ELEC Then
  DR_Commande.Sections("Section3").Controls("lblCSA").Visible = True
  End If
 
 'Si on affiche adresse livraison dans commentaire
 If rstBC.Fields("AffichageInstructions") = True Then
   Call rstConfig.Requery
 
 DR_Commande.Sections("Section3").Controls("lblCommentaire").Caption = "Shipping Address:" & vbNewLine & DR_Commande.Sections("Section3").Controls("lblCommentaire").Caption
   End If
 rstBCPiece.MoveFirst
 Do While rstBCPiece.EOF = False
 
 If rstBCPiece.Fields("NoItem") <> vbNull Then
 If Len(rstBCPiece.Fields("NoItem")) > 2 Then
 DR_Commande.Sections("section1").Controls("text2").Font.SIZE = 8
 End If
 End If
 rstBCPiece.MoveNext
 Loop
 Set DR_Commande.DataSource = rstBCPiece
 
   DR_Commande.Orientation = rptOrientLandscape
 
 Call DR_Commande.Show(vbModal)

9  If m_eForm <> I_RETOUR_MARCHANDISE Then
 If UCase(rstFournisseur.Fields("NomFournisseur")) = "SOLUTION GRB INC." Then
 DR_Commande_recu.Orientation = rptOrientLandscape

 If m_eForm = I_PROJET_ELEC Or m_eForm = I_PROJET_MEC Then
 Call rstBCPiece.Close

 If m_eForm = I_PROJET_ELEC Then
 Call rstBCPiece.Open("SELECT * FROM GrbBonsCommandes_Pieces LEFT JOIN GrbInventaireElec ON GrbBonsCommandes_Pieces.NoItem = GrbInventaireElec.NoItem WHERE NoBonCommande = '" & sNoBC & "' AND NoFournisseur = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstBCPiece.Open("SELECT * FROM GrbBonsCommandes_Pieces LEFT JOIN GrbInventaireMec ON GrbBonsCommandes_Pieces.NoItem = GrbInventaireMec.NoItem WHERE NoBonCommande = '" & sNoBC & "' AND NoFournisseur = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)
 End If

 DR_Commande_recu.Sections("Section1").Controls("txtNoItem").DataField = "GrbBonsCommandes_Pieces.NoItem"
 DR_Commande_recu.Sections("Section1").Controls("txtDescription").DataField = "GrbBonsCommandes_Pieces.Description"
 Else
10  If m_eForm = I_ACHAT_ELEC Or m_eForm = I_ACHAT_MEC Then
10  Call rstBCPiece.Close

10  If m_eForm = I_ACHAT_ELEC Then
10  Call rstBCPiece.Open("SELECT * FROM GrbBonsCommandes_Pieces LEFT JOIN GrbInventaireElec ON GrbBonsCommandes_Pieces.NoItem = GrbInventaireElec.NoItem WHERE NoBonCommande = '" & sNoBC & "' AND NoFournisseur = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)
10  Else
10  Call rstBCPiece.Open("SELECT * FROM GrbBonsCommandes_Pieces LEFT JOIN GrbInventaireMec ON GrbBonsCommandes_Pieces.NoItem = GrbInventaireMec.NoItem WHERE NoBonCommande = '" & sNoBC & "' AND NoFournisseur = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)
10  End If
 Dim testgll As String
 testgll = "GrbBonsCommandes_Pieces.NoItem"
10  DR_Commande_recu.Sections("Section1").Controls("txtNoItem").DataField = "GrbBonsCommandes_Pieces.NoItem"
1 DR_Commande_recu.Sections("Section1").Controls("txtDescription").DataField = "GrbBonsCommandes_Pieces.Description"

1 End If
1 End If

1 Set DR_Commande_recu.DataSource = rstBCPiece

1 DR_Commande_recu.Sections("Section2").Controls("lblfournisseur").Caption = rstFournisseur.Fields("NomFournisseur")
1 DR_Commande_recu.Sections("Section2").Controls("lblprojet").Caption = rstBC.Fields("NoProjet")
1 DR_Commande_recu.Sections("Section5").Controls("lbldatereq").Caption = rstBC.Fields("DateRequise")

1 Call DR_Commande_recu.Show(vbModal)
1 End If
1 Else
1 If m_eForm = I_RETOUR_MARCHANDISE Then
1 If m_eImpRetour = MODE_RETOUR Then
1 Call ImprimerRetour(rstBC.Fields("NoBonCommande"), rstBC.Fields("NoFournisseur"), rstBC.Fields("VotreNoSoum"))
1 Call ImprimerRetourDossier(rstBC.Fields("NoBonCommande"), rstBC.Fields("NoFournisseur"))
 End If
1 End If
 End If

 'Prochain enregistrement
1 Call rstBC.MoveNext
 
 Call rstBCPiece.Close
 
11  Call rstConfig.Close

 Call rstFournisseur.Close
1 Loop

12 Set rstFournisseur = Nothing
12 Set rstConfig = Nothing
12 Set rstBCPiece = Nothing

12 Call g_connData.Execute("DELETE * FROM GrbBonsCommandes_Pieces WHERE NoBonCommande = '" & sNoBC & "'")
12 Call g_connData.Execute("DELETE * FROM GrbBonsCommandes WHERE NoBonCommande = '" & sNoBC & "'")

12 Call Unload(Me)
 
12 Screen.MousePointer = vbDefault

12 Exit Sub

Oups:

12 wOups "frmBonCommande", "cmdImprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerRetour(ByVal sNoRetour As String, ByVal iNoFRS As Integer, ByVal sNoRMA As String)
 
 On Error GoTo Oups

 Dim rstBCPiece As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 
 Set rstBCPiece = New ADODB.Recordset
 
 Call rstBCPiece.Open("SELECT * FROM GrbBonsCommandes_Pieces WHERE NoBonCommande = '" & sNoRetour & "' AND NoFournisseur = " & iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)
 
 Set DR_Retour.DataSource = rstBCPiece
 
 DR_Retour.Orientation = rptOrientLandscape
 
 Set rstFRS = New ADODB.Recordset
 
 Call rstFRS.Open("SELECT NomFournisseur FROM GrbFournisseur WHERE IDFRS = " & iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)
 
 DR_Retour.Sections("Section2").Controls("lblFournisseur").Caption = rstFRS.Fields("NomFournisseur")
 
 Call rstFRS.Close
  Set rstFRS = Nothing
 
  DR_Retour.Sections("Section2").Controls("lblNoProjet").Caption = sNoRetour
  DR_Retour.Sections("Section2").Controls("lblNoRMA").Caption = sNoRMA
  DR_Retour.Sections("Section2").Controls("lblDate").Caption = ConvertDate(Date)
 
  Call DR_Retour.Show(vbModal)
 
  Call rstBCPiece.Close
  Set rstBCPiece = Nothing

  Exit Sub

Oups:

10 wOups "frmBonCommande", "ImprimerRetour", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerRetourDossier(ByVal sNoRetour As String, ByVal iNoFRS As Integer)

 On Error GoTo Oups

 Dim rstBC As ADODB.Recordset
 Dim rstBCPiece As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset

 Set rstBC = New ADODB.Recordset
 Set rstBCPiece = New ADODB.Recordset

 Call rstBC.Open("SELECT * FROM GrbBonsCommandes WHERE NoBonCommande = '" & sNoRetour & "' AND NoFournisseur = " & iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)

 Call rstBCPiece.Open("SELECT * FROM GrbBonsCommandes_Pieces WHERE NoBonCommande = '" & sNoRetour & "' AND NoFournisseur = " & iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)
 
 Set DR_Retour.DataSource = rstBCPiece
 
 DR_Retour.Orientation = rptOrientLandscape
 
 Set rstFRS = New ADODB.Recordset
 
  Call rstFRS.Open("SELECT * FROM GrbFournisseur WHERE IDFRS = " & iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)
 
  DR_Commande.Orientation = rptOrientLandscape

  DR_Commande.Caption = "Retour de marchandise"
 
  DR_Commande.Sections("Section2").Controls("lblTitreBC").Caption = "# RMA"

  DR_Commande.Sections("Section2").Controls("lblTitreCommande").Caption = "RETOUR DE MARCHANDISE"

  DR_Commande.Sections("Section2").Controls("lblTitreNoSoum").Caption = "Notre #"

  If Not IsNull(rstBC.Fields("VotreNoSoum")) Then
  DR_Commande.Sections("Section2").Controls("lblNoBC").Caption = rstBC.Fields("VotreNoSoum")
10 Else
1 DR_Commande.Sections("Section2").Controls("lblNoBC").Caption = vbNullString
End If
 
If Not IsNull(rstBC.Fields("Commentaire")) Then
 DR_Commande.Sections("Section3").Controls("lblCommentaire").Caption = rstBC.Fields("Commentaire")
Else
 DR_Commande.Sections("Section3").Controls("lblCommentaire").Caption = vbNullString
End If
 
DR_Commande.Sections("section2").Controls("lblCommandePar").Caption = rstBC.Fields("CommandePar")

If Not IsNull(rstBC.Fields("Attention")) Then
 DR_Commande.Sections("section2").Controls("lblContact").Caption = rstBC.Fields("Attention")
Else
DR_Commande.Sections("Section2").Controls("lblContact").Caption = vbNullString
End If

 DR_Commande.Sections("section2").Controls("lblDate").Caption = rstBC.Fields("DateCommande")
 
If Not IsNull(rstBC.Fields("DateRequise")) Then
 DR_Commande.Sections("section2").Controls("lblDateRequise").Caption = rstBC.Fields("DateRequise")
Else
 DR_Commande.Sections("section2").Controls("lblDateRequise").Caption = vbNullString
1  End If
 
 DR_Commande.Sections("section2").Controls("lblFax").Caption = rstFRS.Fields("Fax")
 DR_Commande.Sections("section2").Controls("lblFournisseur").Caption = rstFRS.Fields("NomFournisseur")
DR_Commande.Sections("section2").Controls("lblTel").Caption = rstFRS.Fields("telephonne")
 
DR_Commande.Sections("Section2").Controls("lblNoSoum").Caption = rstBC.Fields("NoBonCommande")
 
DR_Commande.Sections("Section3").Controls("lblTotalFin").Caption = Conversion(rstBC.Fields("total"), MODE_ARGENT)
 
If Not IsNull(rstBC.Fields("Transport")) Then
 DR_Commande.Sections("section2").Controls("lblTransport").Caption = rstBC.Fields("Transport")
Else
 DR_Commande.Sections("section2").Controls("lblTransport").Caption = " "
End If
 
DR_Commande.Sections("Section3").Controls("lblPiedPage").Visible = False
 
2  Set DR_Commande.DataSource = rstBCPiece
 
Call DR_Commande.Show(vbModal)
 
2  Call rstFRS.Close
Set rstFRS = Nothing

2  Call rstBC.Close
Set rstBC = Nothing

2  Call rstBCPiece.Close
Set rstBCPiece = Nothing
 
30 Screen.MousePointer = vbDefault

Exit Sub

Oups:

wOups "frmBonCommande", "ImprimerRetourDossier", Err, Err.number, Err.Description
End Sub

Private Sub cmdInstructions_Click()

 On Error GoTo Oups

 Call OuvrirForm(FrmBonCommande_Instruction, True)

 Exit Sub

Oups:

 wOups "frmBonCommande", "cmdInstructions_Click", Err, Err.number, Err.Description
End Sub



Private Sub mvwDate_DateClick(ByVal DateClicked As Date)

 On Error GoTo Oups

 txtDateRequise.Text = ConvertDate(DateClicked)
 
 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmBonCommande", "mvwDate_DateClick", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_LostFocus()

 On Error GoTo Oups

 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmBonCommande", "mvwDate_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub EnregistrerBonCommandeRetourMarchandiseProjet()

 On Error GoTo Oups

 Dim rstBC As ADODB.Recordset
 Dim rstBCPiece As ADODB.Recordset
 Dim rstPiece As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim iCompteur As Integer
 Dim dblTotal As Double
 Dim sWhere As String
 Dim sWherePiece As String
 Dim sWhereNoLigne As String
 Dim sEscompte As String
 
 'Recordset source
  sWhere = "(IDProjet = '" & m_sNoProjet & "')"
 
  sWherePiece = "GrbProjet_Pieces.NumItem In ("
  sWhereNoLigne = "GrbProjet_Pieces.NuméroLigne In ("
 
  For iCompteur = 1 To m_collPieces.count
  If iCompteur <> m_collPieces.count Then
  sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "', "
  sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ", "
  Else
 sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "')"
sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ")"
 End If
Next

Set rstFRS = New ADODB.Recordset
Set rstBC = New ADODB.Recordset
Set rstPiece = New ADODB.Recordset
Set rstBCPiece = New ADODB.Recordset
 
sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne
 
Call rstFRS.Open("SELECT DISTINCT GrbProjet_Pieces.IDFRS, GrbFournisseur.CondTransport FROM GrbProjet_Pieces INNER JOIN GrbFournisseur ON GrbProjet_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
 
 'Recordsets destinations
Call rstBC.Open("SELECT * FROM GrbBonsCommandes", g_connData, adOpenDynamic, adLockOptimistic)
 
Do While Not rstFRS.EOF
Call rstBC.AddNew
 
 'Enregistrement du bon
 rstBC.Fields("NoBonCommande") = txtVotreNoSoum.Text
 rstBC.Fields("NoFournisseur") = rstFRS.Fields("IDFRS")
 rstBC.Fields("NoProjet") = m_sNoProjet
 rstBC.Fields("Attention") = ""
 
 If Not IsNull(rstFRS.Fields("CondTransport")) And rstFRS.Fields("CondTransport") <> vbNullString Then
 rstBC.Fields("Transport") = rstFRS.Fields("CondTransport")
1  Else
 rstBC.Fields("Transport") = "Votre camion"
 End If
 
 rstBC.Fields("DateRequise") = ConvertDate(Date)
 rstBC.Fields("DateCommande") = ConvertDate(Date)

 If m_eForm = I_RETOUR_MARCHANDISE Then
 rstBC.Fields("CommandePar") = m_sEmploye
 Else
 rstBC.Fields("CommandePar") = g_sEmploye
 End If

 rstBC.Fields("LangueImpression") = "Français"
 
 sWhere = "(IDProjet = '" & m_sNoProjet & "' AND IDFRS = " & rstFRS.Fields("IDFRS") & ")"

 sWherePiece = "NumItem In ("
sWhereNoLigne = "NuméroLigne In ("
 
 For iCompteur = 1 To m_collPieces.count
 If iCompteur <> m_collPieces.count Then
 sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "', "
 sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ", "
 Else
 sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "')"
 sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ")"
 End If
3 Next
 
 sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne
 
 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
 
 dblTotal = 0
 
 'Enregistrement des pièces
 Do While Not rstPiece.EOF
 Call rstBCPiece.Open("SELECT * FROM GrbBonsCommandes_Pieces WHERE NoItem = '" & Replace(rstPiece.Fields("NumItem"), "'", "''") & "' AND NoFournisseur = " & rstPiece.Fields("IDFRS") & " AND NoBonCommande = '" & txtVotreNoSoum.Text & "' AND Prix = '" & rstPiece.Fields("PrixOrigine") & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If Trim(rstPiece.Fields("ESCOMPTE")) <> "" Then
 sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
 
 Do While CDbl(sEscompte) > 1
 sEscompte = CDbl(sEscompte) / 100
 Loop

 dblTotal = dblTotal + CDbl(Conversion(CStr((Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * (1 - CDbl(sEscompte))) * CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString))), MODE_DECIMAL))
 Else
 dblTotal = dblTotal + CDbl(Conversion(CStr(Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * Replace(rstPiece.Fields("Qté"), "-", vbNullString)), MODE_DECIMAL))
 End If
 
 If rstBCPiece.EOF Then
 Call rstBCPiece.AddNew

 rstBCPiece.Fields("NoBonCommande") = txtVotreNoSoum.Text
 rstBCPiece.Fields("NoFournisseur") = rstPiece.Fields("IDFRS")
 rstBCPiece.Fields("Qté") = Replace(rstPiece.Fields("Qté"), "-", vbNullString)
 
4 rstBCPiece.Fields("NoItem") = rstPiece.Fields("NumItem")

4 rstBCPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne")
 
4 rstBCPiece.Fields("Description") = rstPiece.Fields("Desc_fr")
4 rstBCPiece.Fields("Manufact") = rstPiece.Fields("Manufact")
 
4 rstBCPiece.Fields("Prix") = rstPiece.Fields("PrixOrigine")
 
4 If rstPiece.Fields("Escompte") <> vbNullString Then
4 rstBCPiece.Fields("Escompte") = rstPiece.Fields("Escompte")
4 Else
4 rstBCPiece.Fields("Escompte") = "0"
4 End If
 
4 If Trim(rstPiece.Fields("Escompte")) <> "" Then
4  sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
 
4  Do While CDbl(sEscompte) > 1
4  sEscompte = CDbl(sEscompte) / 100
4  Loop

4  rstBCPiece.Fields("Total") = Conversion(CStr((Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * (1 - CDbl(sEscompte))) * CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString))), MODE_DECIMAL)
4  Else
4  rstBCPiece.Fields("Total") = Conversion(CStr(Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * Replace(rstPiece.Fields("Qté"), "-", vbNullString)), MODE_DECIMAL)
4  End If
50 Else
 rstBCPiece.Fields("Qté") = CDbl(rstBCPiece.Fields("Qté")) + CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString))

 rstBCPiece.Fields("NuméroLigne") = rstBCPiece.Fields("NuméroLigne") & ", " & rstPiece.Fields("NuméroLigne")

 If Trim(rstPiece.Fields("Escompte")) <> "" Then
 sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
 
 Do While CDbl(sEscompte) > 1
 sEscompte = CDbl(sEscompte) / 100
 Loop

 rstBCPiece.Fields("Total") = Replace(rstBCPiece.Fields("Total"), ".", ",") + (CDbl(Replace(rstPiece.Fields("PrixOrigine"), ".", ",")) * (1 - CDbl(sEscompte)) * CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString)))
 Else
 rstBCPiece.Fields("Total") = Replace(rstBCPiece.Fields("Total"), ".", ",") + (CDbl(Replace(rstPiece.Fields("PrixOrigine"), ".", ",")) * CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString)))
 End If
5  End If
 
5  Call rstBCPiece.Update
 
5  Call rstBCPiece.Close
 
5  Call rstPiece.MoveNext
5  Loop

5  Call rstPiece.Close
 
5  rstBC.Fields("Total") = Conversion(CStr(dblTotal), MODE_DECIMAL)
 
5  Call rstBC.Update
 
60 Call rstFRS.MoveNext
60 Loop
 
  Call rstFRS.Close
  Set rstFRS = Nothing
 
  Call rstBC.Close
  Set rstBC = Nothing

  Set rstPiece = Nothing
  Set rstBCPiece = Nothing

  Exit Sub

Oups:

  wOups "frmBonCommande", "EnregistrerBonCommandeRetourMarchandiseProjet", Err, Err.number, Err.Description
End Sub

Private Sub EnregistrerBonCommandeRetourMarchandiseAchat()

 On Error GoTo Oups

 Dim rstBC As ADODB.Recordset
 Dim rstBCPiece As ADODB.Recordset
 Dim rstPiece As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim iCompteur As Integer
 Dim dblTotal As Double
 Dim sWhere As String
 Dim sWherePiece As String
 Dim sWhereNoLigne As String
 Dim sEscompte As String
 
 'Recordset source
  sWhere = "(IDAchat = '" & m_sNoAchat & "' AND IndexAchat = " & m_iIndexAchat & ")"

  sWherePiece = "GrbAchat_Pieces.PIECE In ("
  sWhereNoLigne = "GrbAchat_Pieces.NuméroLigne In ("
 
  For iCompteur = 1 To m_collPieces.count
  If iCompteur <> m_collPieces.count Then
  sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "', "
  sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ", "
  Else
 sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "')"
sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ")"
 End If
Next
 
Set rstFRS = New ADODB.Recordset
Set rstBC = New ADODB.Recordset
Set rstPiece = New ADODB.Recordset
Set rstBCPiece = New ADODB.Recordset
 
sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne
 
Call rstFRS.Open("SELECT DISTINCT GrbAchat_Pieces.IDFRS, GrbFournisseur.CondTransport FROM GrbAchat_Pieces INNER JOIN GrbFournisseur ON GrbAchat_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
 
 'Recordsets destinations
Call rstBC.Open("SELECT * FROM GrbBonsCommandes", g_connData, adOpenDynamic, adLockOptimistic)
 
Do While Not rstFRS.EOF
Call rstBC.AddNew
 
 'Enregistrement du bon
 rstBC.Fields("NoBonCommande") = txtVotreNoSoum.Text
 rstBC.Fields("NoFournisseur") = rstFRS.Fields("IDFRS")
 rstBC.Fields("NoProjet") = m_sNoAchat & " - " & m_iIndexAchat
 rstBC.Fields("Attention") = ""
 
 If Not IsNull(rstFRS.Fields("CondTransport")) And rstFRS.Fields("CondTransport") <> vbNullString Then
 rstBC.Fields("Transport") = rstFRS.Fields("CondTransport")
1  Else
 rstBC.Fields("Transport") = "Votre camion"
 End If
 
 rstBC.Fields("DateRequise") = ConvertDate(Date)
 rstBC.Fields("DateCommande") = ConvertDate(Date)

 If m_eForm = I_RETOUR_MARCHANDISE Then
 rstBC.Fields("CommandePar") = m_sEmploye
 Else
 rstBC.Fields("CommandePar") = g_sEmploye
 End If

 rstBC.Fields("LangueImpression") = "Français"
 
 sWhere = "(IDAchat = '" & m_sNoAchat & "' AND IndexAchat = " & m_iIndexAchat & " AND IDFRS = " & rstFRS.Fields("IDFRS") & ")"
 
 sWherePiece = "PIECE In ("
sWhereNoLigne = "NuméroLigne In ("
 
 For iCompteur = 1 To m_collPieces.count
 If iCompteur <> m_collPieces.count Then
 sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "', "
 sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ", "
 Else
 sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "')"
 sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ")"
 End If
3 Next
 
 sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne
 
 Call rstPiece.Open("SELECT * FROM GrbAchat_Pieces WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
 
 dblTotal = 0
 
 'Enregistrement des pièces
 Do While Not rstPiece.EOF
 Call rstBCPiece.Open("SELECT * FROM GrbBonsCommandes_Pieces WHERE NoItem = '" & Replace(rstPiece.Fields("PIECE"), "'", "''") & "' AND NoFournisseur = " & rstPiece.Fields("IDFRS") & " AND NoBonCommande = '" & txtVotreNoSoum.Text & "' AND Prix = '" & rstPiece.Fields("PrixOrigine") & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If Trim(rstPiece.Fields("ESCOMPTE")) <> "" Then
 sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
 
 Do While CDbl(sEscompte) > 1
 sEscompte = CDbl(sEscompte) / 100
 Loop

 dblTotal = dblTotal + CDbl(Conversion(CStr((Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * (1 - CDbl(sEscompte))) * CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString))), MODE_DECIMAL))
 Else
 dblTotal = dblTotal + CDbl(Conversion(CStr(Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * Replace(rstPiece.Fields("Qté"), "-", vbNullString)), MODE_DECIMAL))
 End If
 
 If rstBCPiece.EOF Then
 Call rstBCPiece.AddNew

 rstBCPiece.Fields("NoBonCommande") = txtVotreNoSoum.Text
 rstBCPiece.Fields("NoFournisseur") = rstPiece.Fields("IDFRS")
 rstBCPiece.Fields("Qté") = Replace(rstPiece.Fields("Qté"), "-", vbNullString)
 
4 rstBCPiece.Fields("NoItem") = rstPiece.Fields("PIECE")

4 rstBCPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne")
 
4 rstBCPiece.Fields("Description") = rstPiece.Fields("Desc_fr")
4 rstBCPiece.Fields("Manufact") = rstPiece.Fields("Manufact")
 
4 rstBCPiece.Fields("Prix") = rstPiece.Fields("PrixOrigine")
 
4 If rstPiece.Fields("Escompte") <> vbNullString Then
4 rstBCPiece.Fields("Escompte") = rstPiece.Fields("Escompte")
4 Else
4 rstBCPiece.Fields("Escompte") = "0"
4 End If
 
4 If Trim(rstPiece.Fields("Escompte")) <> "" Then
4  sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
 
4  Do While CDbl(sEscompte) > 1
4  sEscompte = CDbl(sEscompte) / 100
4  Loop

4  rstBCPiece.Fields("Total") = Conversion(CStr((Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * (1 - CDbl(sEscompte))) * CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString))), MODE_DECIMAL)
4  Else
4  rstBCPiece.Fields("Total") = Conversion(CStr(Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * Replace(rstPiece.Fields("Qté"), "-", vbNullString)), MODE_DECIMAL)
4  End If
50 Else
 rstBCPiece.Fields("Qté") = CDbl(rstBCPiece.Fields("Qté")) + CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString))

 rstBCPiece.Fields("NuméroLigne") = rstBCPiece.Fields("NuméroLigne") & ", " & rstPiece.Fields("NuméroLigne")

 If Trim(rstPiece.Fields("Escompte")) <> "" Then
 sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
 
 Do While CDbl(sEscompte) > 1
 sEscompte = CDbl(sEscompte) / 100
 Loop

 rstBCPiece.Fields("Total") = Replace(rstBCPiece.Fields("Total"), ".", ",") + (CDbl(Replace(rstPiece.Fields("PrixOrigine"), ".", ",")) * (1 - CDbl(sEscompte)) * CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString)))
 Else
 rstBCPiece.Fields("Total") = Replace(rstBCPiece.Fields("Total"), ".", ",") + (CDbl(Replace(rstPiece.Fields("PrixOrigine"), ".", ",")) * CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString)))
 End If
5  End If
 
5  Call rstBCPiece.Update
 
5  Call rstBCPiece.Close
 
5  Call rstPiece.MoveNext
5  Loop

5  Call rstPiece.Close
 
5  rstBC.Fields("Total") = Conversion(CStr(dblTotal), MODE_DECIMAL)
 
5  Call rstBC.Update
 
60 Call rstFRS.MoveNext
60 Loop

  Call rstFRS.Close
  Set rstFRS = Nothing
 
  Call rstBC.Close
  Set rstBC = Nothing

  Set rstPiece = Nothing
  Set rstBCPiece = Nothing

  Exit Sub

Oups:

  wOups "frmBonCommande", "EnregistrerBonCommandeRetourMarchandiseAchat", Err, Err.number, Err.Description
End Sub

Private Sub EnregistrerBonCommandeProjet()

 On Error GoTo Oups

 Dim rstBC As ADODB.Recordset
 Dim rstBCPiece As ADODB.Recordset
 Dim rstPiece As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim iCompteur As Integer
 Dim dblTotal As Double
 Dim sType As String
 Dim sWhere As String
 Dim sWherePiece As String
 Dim sWhereNoLigne As String
  Dim sEscompte As String
 
  If m_eForm = I_PROJET_ELEC Then
  sType = "E"
  Else
  sType = "M"
  End If
 
 'Recordset source
  sWhere = "(IDProjet = '" & m_sNoProjet & "' AND Type = '" & sType & "')"

  sWherePiece = "GrbProjet_Pieces.NumItem In ("
10 sWhereNoLigne = "GrbProjet_Pieces.NuméroLigne In ("
 
For iCompteur = 1 To m_collPieces.count
 If iCompteur <> m_collPieces.count Then
 sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "', "
 sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ", "
 Else
 sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "')"
 sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ")"
 End If
Next

Set rstBC = New ADODB.Recordset
Set rstFRS = New ADODB.Recordset
1  Set rstPiece = New ADODB.Recordset
Set rstBCPiece = New ADODB.Recordset

 sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne
 
Call rstFRS.Open("SELECT DISTINCT GrbProjet_Pieces.IDFRS, GrbFournisseur.CondTransport FROM GrbProjet_Pieces INNER JOIN GrbFournisseur ON GrbProjet_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
 
 'Recordsets destinations
 Call rstBC.Open("SELECT * FROM GrbBonsCommandes", g_connData, adOpenDynamic, adLockOptimistic)
 
Do While Not rstFRS.EOF
 Call rstBC.AddNew
 
 'Enregistrement du bon
1  rstBC.Fields("NoBonCommande") = txtNoBC.Text
 rstBC.Fields("NoFournisseur") = rstFRS.Fields("IDFRS")
 rstBC.Fields("NoProjet") = m_sNoProjet
 rstBC.Fields("Attention") = ""
 
 If Not IsNull(rstFRS.Fields("CondTransport")) And rstFRS.Fields("CondTransport") <> vbNullString Then
 rstBC.Fields("Transport") = rstFRS.Fields("CondTransport")
 Else
 rstBC.Fields("Transport") = "Votre camion"
 End If
 
 If m_eForm = I_PROJET_ELEC Then
 rstBC.Fields("DateRequise") = ""
 Else
 rstBC.Fields("DateRequise") = ConvertDate(Date)
End If

 rstBC.Fields("DateCommande") = ConvertDate(Date)
rstBC.Fields("CommandePar") = g_sEmploye

 If m_eLangage = FRANCAIS Then
 rstBC.Fields("LangueImpression") = "Français"
 Else
 rstBC.Fields("LangueImpression") = "Anglais"
 End If
 
rstBC.Fields("Type") = sType

3 sWhere = "(IDProjet = '" & m_sNoProjet & "' AND IDFRS = " & rstFRS.Fields("IDFRS") & ")"
 
 sWherePiece = "NumItem In ("
 sWhereNoLigne = "NuméroLigne In ("
 
 For iCompteur = 1 To m_collPieces.count
 If iCompteur <> m_collPieces.count Then
 sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "', "
 sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ", "
 Else
 sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "')"
 sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ")"
 End If
Next

 sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne
 
Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
 
 dblTotal = 0
 
 'Enregistrement des pièces
Do While Not rstPiece.EOF
 Call rstBCPiece.Open("SELECT * FROM GrbBonsCommandes_Pieces WHERE NoItem = '" & Replace(rstPiece.Fields("NumItem"), "'", "''") & "' AND NoFournisseur = " & rstPiece.Fields("IDFRS") & " AND NoBonCommande = '" & txtNoBC.Text & "' AND Prix = '" & rstPiece.Fields("PrixOrigine") & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If Trim(rstPiece.Fields("ESCOMPTE")) <> "" Then
 sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
 
 Do While CDbl(sEscompte) > 1
4 sEscompte = CDbl(sEscompte) / 100
4 Loop

4 dblTotal = dblTotal + CDbl(Conversion(CStr((Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * (1 - CDbl(sEscompte))) * CDbl(rstPiece.Fields("Qté"))), MODE_DECIMAL))
4 Else
4 dblTotal = dblTotal + CDbl(Conversion(CStr(Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * rstPiece.Fields("Qté")), MODE_DECIMAL))
4 End If

 'Si la pièce n'existe pas, on l'ajoute
 'sinon, on change la quantité et le total
4 If rstBCPiece.EOF Then
4 Call rstBCPiece.AddNew
 
4 rstBCPiece.Fields("Type") = sType
 
4 rstBCPiece.Fields("NoBonCommande") = txtNoBC.Text
4 rstBCPiece.Fields("NoFournisseur") = rstPiece.Fields("IDFRS")
4  rstBCPiece.Fields("Qté") = rstPiece.Fields("Qté")
 
4  rstBCPiece.Fields("NoItem") = rstPiece.Fields("NumItem")

4  rstBCPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne")
 
4  If rstBC.Fields("LangueImpression") = "Français" Then
4  rstBCPiece.Fields("Description") = rstPiece.Fields("DESC_FR")
4  Else
4  rstBCPiece.Fields("Description") = rstPiece.Fields("DESC_EN")
4  End If

50 rstBCPiece.Fields("Manufact") = rstPiece.Fields("Manufact")
 
 rstBCPiece.Fields("Prix") = rstPiece.Fields("PrixOrigine")
 
 If rstPiece.Fields("Escompte") <> vbNullString Then
 rstBCPiece.Fields("Escompte") = rstPiece.Fields("Escompte")
 Else
 rstBCPiece.Fields("Escompte") = "0"
 End If
 
 If Trim(rstPiece.Fields("Escompte")) <> "" Then
 sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
 
 Do While CDbl(sEscompte) > 1
 sEscompte = CDbl(sEscompte) / 100
 Loop

5  rstBCPiece.Fields("Total") = Conversion(CStr((Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * (1 - CDbl(sEscompte))) * CDbl(rstPiece.Fields("Qté"))), MODE_DECIMAL)
5  Else
5  rstBCPiece.Fields("Total") = Conversion(CStr(Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * rstPiece.Fields("Qté")), MODE_DECIMAL)
5  End If
5  Else
5  rstBCPiece.Fields("Qté") = CDbl(rstBCPiece.Fields("Qté")) + CDbl(rstPiece.Fields("Qté"))

5  rstBCPiece.Fields("NuméroLigne") = rstBCPiece.Fields("NuméroLigne") & ", " & rstPiece.Fields("NuméroLigne")

5  If Trim(rstPiece.Fields("Escompte")) <> "" Then
60 sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
 
  Do While CDbl(sEscompte) > 1
  sEscompte = CDbl(sEscompte) / 100
  Loop

  rstBCPiece.Fields("Total") = Replace(rstBCPiece.Fields("Total"), ".", ",") + (CDbl(Replace(rstPiece.Fields("PrixOrigine"), ".", ",")) * (1 - CDbl(sEscompte)) * CDbl(rstPiece.Fields("Qté")))
  Else
  rstBCPiece.Fields("Total") = Replace(rstBCPiece.Fields("Total"), ".", ",") + (CDbl(Replace(rstPiece.Fields("PrixOrigine"), ".", ",")) * CDbl(rstPiece.Fields("Qté")))
  End If
  End If
 
  Call rstBCPiece.Update
 
  Call rstBCPiece.Close
 
  Call rstPiece.MoveNext
6  Loop

6  Call rstPiece.Close
 
6  rstBC.Fields("Total") = Conversion(CStr(dblTotal), MODE_DECIMAL)
 
6  Call rstBC.Update
 
6  Call rstFRS.MoveNext
6  Loop
 
6  Call rstFRS.Close
6  Set rstFRS = Nothing
 
70 Call rstBC.Close
70 Set rstBC = Nothing

  Set rstBCPiece = Nothing
  Set rstPiece = Nothing

  Exit Sub

Oups:

  wOups "frmBonCommande", "EnregistrerBonCommandeProjet", Err, Err.number, Err.Description
End Sub

Private Sub EnregistrerBonCommandeAchat()

 On Error GoTo Oups

 Dim rstBC As ADODB.Recordset
 Dim rstBCPiece As ADODB.Recordset
 Dim rstPiece As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim iCompteur As Integer
 Dim dblTotal As Double
 Dim sType As String
 Dim sWhere As String
 Dim sWherePiece As String
 Dim sWhereNoLigne As String
  Dim sEscompte As String
 
  If m_eForm = I_ACHAT_ELEC Then
  sType = "E"
  Else
  sType = "M"
  End If
 
 'Recordset source
  sWhere = "(IDAchat = '" & m_sNoAchat & "' AND IndexAchat = " & m_iIndexAchat & ")"
 
  sWherePiece = "GrbAchat_Pieces.PIECE In ("
10 sWhereNoLigne = "GrbAchat_Pieces.NuméroLigne In ("
 
For iCompteur = 1 To m_collPieces.count
 If iCompteur <> m_collPieces.count Then
 sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "', "
 sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ", "
 Else
 sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "')"
 sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ")"
 End If
Next

Set rstFRS = New ADODB.Recordset
Set rstBC = New ADODB.Recordset
1  Set rstPiece = New ADODB.Recordset
Set rstBCPiece = New ADODB.Recordset

 sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne
 
 'Recordset source
Call rstFRS.Open("SELECT DISTINCT GrbAchat_Pieces.IDFRS, GrbFournisseur.CondTransport FROM GrbAchat_Pieces INNER JOIN GrbFournisseur ON GrbAchat_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)

 'Recordsets destinations
 Call rstBC.Open("SELECT * FROM GrbBonsCommandes", g_connData, adOpenDynamic, adLockOptimistic)
 
Do While Not rstFRS.EOF
 Call rstBC.AddNew
 
 'Enregistrement du bon
1  rstBC.Fields("NoBonCommande") = txtNoBC.Text
 rstBC.Fields("NoFournisseur") = rstFRS.Fields("IDFRS")
 rstBC.Fields("NoProjet") = m_sNoAchat
 rstBC.Fields("Attention") = ""
 
 If Not IsNull(rstFRS.Fields("CondTransport")) And rstFRS.Fields("CondTransport") <> vbNullString Then
 rstBC.Fields("Transport") = rstFRS.Fields("CondTransport")
 Else
 rstBC.Fields("Transport") = "Votre camion"
 End If

 If m_eForm = I_ACHAT_ELEC Then
 rstBC.Fields("DateRequise") = ""
 Else
 rstBC.Fields("DateRequise") = ConvertDate(Date)
End If

 rstBC.Fields("DateCommande") = ConvertDate(Date)
rstBC.Fields("CommandePar") = g_sEmploye
 rstBC.Fields("LangueImpression") = "Français"
 
rstBC.Fields("Type") = sType
 
 Call rstPiece.Open("SELECT * FROM GrbAchat_Pieces WHERE " & sWhere & " AND IDFRS = " & rstFRS.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
 
dblTotal = 0
 
 'Enregistrement des pièces
 Do While Not rstPiece.EOF
 Call rstBCPiece.Open("SELECT * FROM GrbBonsCommandes_Pieces WHERE NoItem = '" & Replace(rstPiece.Fields("PIECE"), "'", "''") & "' AND NoFournisseur = " & rstPiece.Fields("IDFRS") & " AND NoBonCommande = '" & txtNoBC.Text & "' AND Prix = '" & rstPiece.Fields("PrixOrigine") & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
If Trim(rstPiece.Fields("ESCOMPTE")) <> "" Then
 sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
 
 Do While CDbl(sEscompte) > 1
 sEscompte = CDbl(sEscompte) / 100
 Loop

 If Not IsNull(rstPiece.Fields("PrixOrigine")) Then
 dblTotal = dblTotal + CDbl(Conversion(CStr((Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * (1 - CDbl(sEscompte))) * CDbl(rstPiece.Fields("Qté"))), MODE_DECIMAL))
 End If
 Else
 If Not IsNull(rstPiece.Fields("PrixOrigine")) Then
 dblTotal = dblTotal + CDbl(Conversion(CStr(Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * rstPiece.Fields("Qté")), MODE_DECIMAL))
 End If
 End If
 
 'Si la pièce n'existe pas, on l'ajoute
 'sinon, on change la quantité et le total
 If rstBCPiece.EOF Then
 Call rstBCPiece.AddNew
 
 rstBCPiece.Fields("Type") = sType
 
 rstBCPiece.Fields("NoBonCommande") = txtNoBC.Text
 rstBCPiece.Fields("NoFournisseur") = rstPiece.Fields("IDFRS")
 rstBCPiece.Fields("Qté") = rstPiece.Fields("Qté")

 rstBCPiece.Fields("NoItem") = rstPiece.Fields("PIECE")

4 rstBCPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne")

4 rstBCPiece.Fields("Description") = rstPiece.Fields("Desc_fr")
4 rstBCPiece.Fields("Manufact") = rstPiece.Fields("Manufact")

4 rstBCPiece.Fields("Prix") = rstPiece.Fields("PrixOrigine")

4 If Not IsNull(rstPiece.Fields("Escompte")) And rstPiece.Fields("Escompte") <> vbNullString Then
4 rstBCPiece.Fields("Escompte") = rstPiece.Fields("Escompte")
4 Else
4 rstBCPiece.Fields("Escompte") = "0"
4 End If

4 If Trim(rstPiece.Fields("Escompte")) <> "" Then
4 sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
 
4  Do While CDbl(sEscompte) > 1
4  sEscompte = CDbl(sEscompte) / 100
4  Loop

4  If Not IsNull(rstPiece.Fields("PrixOrigine")) Then
4  rstBCPiece.Fields("Total") = Conversion(CStr((Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * (1 - CDbl(sEscompte))) * CDbl(rstPiece.Fields("Qté"))), MODE_DECIMAL)
4  End If
4  Else
4  If Not IsNull(rstPiece.Fields("PrixOrigine")) Then
50 rstBCPiece.Fields("Total") = Conversion(CStr(Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * rstPiece.Fields("Qté")), MODE_DECIMAL)
 End If
 End If
 Else
 rstBCPiece.Fields("Qté") = CDbl(rstBCPiece.Fields("Qté")) + CDbl(rstPiece.Fields("Qté"))

 rstBCPiece.Fields("NuméroLigne") = rstBCPiece.Fields("NuméroLigne") & ", " & rstPiece.Fields("NuméroLigne")

 If Trim(rstPiece.Fields("Escompte")) <> "" Then
 sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
 
 Do While CDbl(sEscompte) > 1
 sEscompte = CDbl(sEscompte) / 100
 Loop

 rstBCPiece.Fields("Total") = Replace(rstBCPiece.Fields("Total"), ".", ",") + (CDbl(Replace(rstPiece.Fields("PrixOrigine"), ".", ",")) * (1 - CDbl(sEscompte)) * CDbl(rstPiece.Fields("Qté")))
5  Else
5  rstBCPiece.Fields("Total") = Replace(rstBCPiece.Fields("Total"), ".", ",") + (CDbl(Replace(rstPiece.Fields("PrixOrigine"), ".", ",")) * CDbl(rstPiece.Fields("Qté")))
5  End If
5  End If

5  Call rstBCPiece.Update

5  Call rstBCPiece.Close

5  Call rstPiece.MoveNext
5  Loop

60 rstBC.Fields("Total") = Conversion(CStr(dblTotal), MODE_DECIMAL)

  Call rstBC.Update

  Call rstFRS.MoveNext

  Call rstPiece.Close
  Loop
 
  Call rstFRS.Close
  Set rstFRS = Nothing

  Call rstBC.Close
  Set rstBC = Nothing

  Set rstPiece = Nothing
  Set rstBCPiece = Nothing

  Exit Sub

Oups:

6  wOups "frmBonCommande", "EnregistrerBonCommande", Err, Err.number, Err.Description
End Sub

Private Sub EnregistrerModifFournisseur()

 On Error GoTo Oups

 Dim rstBC As ADODB.Recordset
 Dim rstBCPiece As ADODB.Recordset
 Dim iCompteur As Integer
 Dim itmBC As ListItem
 Dim sNoBC As String

 If m_eForm = I_RETOUR_MARCHANDISE Then
 sNoBC = txtVotreNoSoum.Text
 Else
 sNoBC = txtNoBC.Text
 End If

  Set rstBC = New ADODB.Recordset

  Call rstBC.Open("SELECT * FROM GrbBonsCommandes WHERE NoBonCommande = '" & sNoBC & "' AND NoFournisseur = " & m_iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)
 
 'Enregistre le bon de commande
  rstBC.Fields("Attention") = cmbContact.Text
  rstBC.Fields("Transport") = txtTransport.Text
  rstBC.Fields("DateRequise") = txtDateRequise.Text

  If m_eForm = I_RETOUR_MARCHANDISE Then
  rstBC.Fields("VotreNoSoum") = txtNoBC.Text
  Else
rstBC.Fields("VotreNoSoum") = txtVotreNoSoum.Text
End If

rstBC.Fields("Commentaire") = txtcommentaire.Text
rstBC.Fields("Total") = Conversion(txtTotal.Text, MODE_PAS_FORMAT)
rstBC.Fields("AffichageInstructions") = chkAfficherInstructions.Value
 
If optImpression(I_IMP_FRANCAIS).Value = True Then
 rstBC.Fields("LangueImpression") = "Français"
Else
 rstBC.Fields("LangueImpression") = "Anglais"
End If
 
Call rstBC.Update
 
Call rstBC.Close
1  Set rstBC = Nothing
 
Set rstBCPiece = New ADODB.Recordset
 
 If m_eForm <> I_PROJET_ELEC And m_eForm <> I_PROJET_MEC Then
 For iCompteur = 1 To lvwBonCommande.ListItems.count
 Set itmBC = lvwBonCommande.ListItems(iCompteur)
 
 'Enregistre les pièces
 Call rstBCPiece.Open("SELECT * FROM GrbBonsCommandes_Pieces WHERE NoBonCommande = '" & sNoBC & "' AND NoFournisseur = " & m_iNoFRS & " AND NoItem = '" & Replace(itmBC.SubItems(I_COL_NO_ITEM), "'", "''") & "' AND NuméroLigne = '" & itmBC.Tag & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstBCPiece.EOF Then
1  rstBCPiece.Fields("Qté") = itmBC.Text
 rstBCPiece.Fields("Total") = itmBC.SubItems(I_COL_TOTAL)
 
 Call rstBCPiece.Update
 Else
 Call MsgBox("Impossible d'enregistrer le bon de commande!", vbOKOnly, "Erreur")
 End If
 
 Call rstBCPiece.Close
 Next

 Set rstBCPiece = Nothing
End If

Exit Sub

Oups:

wOups "frmBonCommande", "EnregistrerModifFournisseur", Err, Err.number, Err.Description
End Sub

Private Sub txtDateRequise_KeyPress(KeyAscii As Integer)

 On Error GoTo Oups

 If KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
 KeyAscii = 0
 End If

 Exit Sub

Oups:

 wOups "frmBonCommande", "txtDateRequise_KeyPress", Err, Err.number, Err.Description
End Sub
