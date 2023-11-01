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
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      StartOfWeek     =   90243073
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
Private Const I_COL_NO_ITEM  As Integer = 1
Private Const I_COL_DESCR    As Integer = 2
Private Const I_COL_MANUFACT As Integer = 3
Private Const I_COL_PRIX     As Integer = 4
Private Const I_COL_ESCOMPTE As Integer = 5
Private Const I_COL_TOTAL    As Integer = 6

'Index de optImpression
Private Const I_IMP_FRANCAIS As Integer = 0
Private Const I_IMP_ANGLAIS  As Integer = 1

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
Private m_sNoProjet   As String

Private m_sNoAchat    As String
Private m_iIndexAchat As Integer

'Pour savoir le type (Électrique ou mécanique)
Public m_eForm        As enumFormSource

'Pour savoir si le form vient d'être ouvert
Private m_bOuverture  As Boolean

'No du fournisseur sélectionné dans le combo
Private m_iNoFRS      As Integer

'Pièces à partir d'un projet
Private m_collPieces  As Collection
Private m_collNoLigne As Collection

Private m_sEmploye    As String

Private m_eImpRetour  As enumImpressionRetour

Private m_eLangage    As enumLangage

Private Sub cmbFournisseur_Click()

5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbHourglass
  
        'Il ne faut pas enregistrer les modifications sur l'ouverture du form
        'parce qu'il n'y a pas eu de modifications encore
  
        'Si le form ne vient pas d'etre ouvert
15      If m_bOuverture = False Then
          'On enregistre les modifications
20        Call EnregistrerModifFournisseur
25      Else
30        m_bOuverture = False
35      End If

40      m_iNoFRS = cmbFournisseur.ItemData(cmbFournisseur.ListIndex)

        'Rempli le combo des contacts
45      Call RemplirComboContacts
    
        'On affiche le contenu du fournisseur sélectionné
50      Call AfficherContenuFournisseur
  
55      Screen.MousePointer = vbDefault

60      Exit Sub

AfficherErreur:

65      woups "frmBonCommande", "cmbFournisseur_Click", Err, Erl
End Sub

Private Sub cmdDate_Click()

5       On Error GoTo AfficherErreur

10      mvwDate.Value = Date
  
15      mvwDate.Visible = True
  
20      Call mvwDate.SetFocus

25      Exit Sub

AfficherErreur:

30      woups "frmBonCommande", "cmdDate_Click", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

10      Dim sNoBC As String

15      If m_eForm = I_RETOUR_MARCHANDISE Then
20        sNoBC = txtVotreNoSoum.Text
25      Else
30        sNoBC = txtNoBC.Text
35      End If

40      Call g_connData.Execute("DELETE * FROM GRB_BonsCommandes_Pieces WHERE NoBonCommande = '" & sNoBC & "'")
45      Call g_connData.Execute("DELETE * FROM GRB_BonsCommandes WHERE NoBonCommande = '" & sNoBC & "'")

50      Call Unload(Me)

55      Exit Sub

AfficherErreur:

60      woups "frmBonCommande", "cmdFermer_Click", Err, Erl
End Sub

Public Sub AfficherFormProjetAchat(ByVal sNoProjet As String, ByVal sNoBonCommande As String, ByVal collPiece As Collection, ByVal collNoLigne As Collection, ByVal eForm As enumFormSource, ByVal iLangage As Integer)

5       On Error GoTo AfficherErreur

10      If eForm = I_ACHAT_ELEC Or eForm = I_ACHAT_MEC Then
15        m_sNoAchat = Left$(sNoProjet, 9)
20        m_iIndexAchat = CInt(Right$(sNoProjet, 3))
25      Else
30        m_sNoProjet = sNoProjet
35      End If
  
40      m_eForm = eForm

45      m_eLangage = iLangage
  
50      Set m_collPieces = collPiece

55      Set m_collNoLigne = collNoLigne
    
60      m_bOuverture = True
    
65      txtNoBC.Text = sNoBonCommande

70      Call g_connData.Execute("DELETE * FROM GRB_BonsCommandes WHERE NoBonCommande = '" & txtNoBC.Text & "'")
75      Call g_connData.Execute("DELETE * FROM GRB_BonsCommandes_Pieces WHERE NoBonCommande = '" & txtNoBC.Text & "'")
  
        'Enregistrement du bon de commande
80      If eForm = I_ACHAT_ELEC Or eForm = I_ACHAT_MEC Then
85        Call EnregistrerBonCommandeAchat
90      Else
95        Call EnregistrerBonCommandeProjet
100     End If
  
        'On rempli les fournisseurs
105     Call RemplirComboFournisseurs
  
        'Affichage du form modalement
110     Call Me.Show(vbModal)

115     Exit Sub

AfficherErreur:

120     woups "frmBonCommande", "AfficherFormProjet", Err, Erl
End Sub

Public Sub AfficherFormRetourMarchandiseProjet(ByVal sNoProjet As String, ByVal sNoBonCommande As String, ByVal collPiece As Collection, ByVal collNoLigne As Collection, ByVal sUserID As String, ByVal eImpRetour As enumImpressionRetour)

5       On Error GoTo AfficherErreur

10      Dim rstEmploye As ADODB.Recordset

15      Me.Caption = "Retour de marchandise"

20      lblVotreNoSoum.Caption = "Notre # : "

25      lblNoBC.Caption = "# RMA : "

30      m_eImpRetour = eImpRetour

35      m_sNoProjet = Right$(sNoProjet, Len(sNoProjet) - 1)
  
40      m_eForm = I_RETOUR_MARCHANDISE
  
45      Set m_collPieces = collPiece

50      Set m_collNoLigne = collNoLigne

55      m_bOuverture = True

60      Set rstEmploye = New ADODB.Recordset

65      Call rstEmploye.Open("SELECT Employe FROM GRB_Employés WHERE loginname = '" & sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
70      m_sEmploye = rstEmploye.Fields("Employe")

75      Call rstEmploye.Close
80      Set rstEmploye = Nothing
    
85      txtVotreNoSoum.Text = sNoBonCommande

90      txtVotreNoSoum.Locked = True

95      txtNoBC.Locked = False
  
100     Call g_connData.Execute("DELETE * FROM GRB_BonsCommandes WHERE NoBonCommande = '" & txtVotreNoSoum.Text & "'")
105     Call g_connData.Execute("DELETE * FROM GRB_BonsCommandes_Pieces WHERE NoBonCommande = '" & txtVotreNoSoum.Text & "'")
  
        'Enregistrement du bon de commande
110     Call EnregistrerBonCommandeRetourMarchandiseProjet
  
        'On rempli les fournisseurs
115     Call RemplirComboFournisseurs
  
        'Affichage du form modalement
120     Call Me.Show(vbModal)

125     Exit Sub

AfficherErreur:

130     woups "frmBonCommande", "AfficherFormRetourMarchandiseProjet", Err, Erl
End Sub

Public Sub AfficherFormRetourMarchandiseAchat(ByVal sNoAchat As String, ByVal iIndexAchat As Integer, ByVal sNoBonCommande As String, ByVal collPiece As Collection, ByVal collNoLigne As Collection, ByVal sUserID As String, ByVal eImpRetour As enumImpressionRetour)

5       On Error GoTo AfficherErreur

10      Dim rstEmploye As ADODB.Recordset

15      Me.Caption = "Retour de marchandise"

20      lblVotreNoSoum.Caption = "Notre # : "

25      lblNoBC.Caption = "# RMA : "

30      m_eImpRetour = eImpRetour

35      m_sNoAchat = sNoAchat
40      m_iIndexAchat = iIndexAchat
  
45      m_eForm = I_RETOUR_MARCHANDISE
  
50      Set m_collPieces = collPiece

55      Set m_collNoLigne = collNoLigne

60      m_bOuverture = True

65      Set rstEmploye = New ADODB.Recordset

70      Call rstEmploye.Open("SELECT Employe FROM GRB_Employés WHERE loginname = '" & sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
75      m_sEmploye = rstEmploye.Fields("Employe")

80      Call rstEmploye.Close
85      Set rstEmploye = Nothing
    
90      txtVotreNoSoum.Text = sNoBonCommande

95      txtVotreNoSoum.Locked = True

100     txtNoBC.Locked = False

105     Call g_connData.Execute("DELETE * FROM GRB_BonsCommandes WHERE NoBonCommande = '" & txtVotreNoSoum.Text & "'")
110     Call g_connData.Execute("DELETE * FROM GRB_BonsCommandes_Pieces WHERE NoBonCommande = '" & txtVotreNoSoum.Text & "'")
  
        'Enregistrement du bon de commande
115     Call EnregistrerBonCommandeRetourMarchandiseAchat
  
        'On rempli les fournisseurs
120     Call RemplirComboFournisseurs
  
        'Affichage du form modalement
125     Call Me.Show(vbModal)

130     Exit Sub

AfficherErreur:

135     woups "frmBonCommande", "AfficherFormRetourMarchandiseAchat", Err, Erl
End Sub


Private Sub AfficherContenuFournisseur()

5       On Error GoTo AfficherErreur

10      Dim rstBC  As ADODB.Recordset
15      Dim rstFRS As ADODB.Recordset
20      Dim sNoBC  As String

25      If m_eForm = I_RETOUR_MARCHANDISE Then
30        sNoBC = txtVotreNoSoum.Text
35      Else
40        sNoBC = txtNoBC.Text
45      End If

50      Set rstBC = New ADODB.Recordset
  
55      Call rstBC.Open("SELECT * FROM GRB_BonsCommandes WHERE NoFournisseur = " & m_iNoFRS & " AND NoBonCommande = '" & sNoBC & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
        'À l'attention de
60      If rstBC.Fields("Attention") <> vbNullString Then
65        If ComboContient(cmbContact, rstBC.Fields("Attention")) = True Then
70          cmbContact.Text = rstBC.Fields("Attention")
75        Else
80          cmbContact.ListIndex = -1
85        End If
90      Else
95        cmbContact.ListIndex = -1
100     End If
  
        'Transport
105     If Not IsNull(rstBC.Fields("Transport")) Or Trim(rstBC.Fields("Transport")) <> vbNullString Then
110       txtTransport.Text = rstBC.Fields("Transport")
115     Else
120       txtTransport.Text = vbNullString
125     End If
  
        'Date requise
130     If Not IsNull(rstBC.Fields("DateRequise")) Or Trim(rstBC.Fields("DateRequise")) <> vbNullString Then
135       txtDateRequise.Text = rstBC.Fields("DateRequise")
140     Else
145       txtDateRequise.Text = vbNullString
150     End If
  
        'Votre # Soum
155     If m_eForm = I_RETOUR_MARCHANDISE Then
160       If Not IsNull(rstBC.Fields("VotreNoSoum")) Then
165         txtNoBC.Text = rstBC.Fields("VotreNoSoum")
170       Else
175         txtNoBC.Text = vbNullString
180       End If
185     Else
190       If Not IsNull(rstBC.Fields("VotreNoSoum")) Then
195         txtVotreNoSoum.Text = rstBC.Fields("VotreNoSoum")
200       Else
205         txtVotreNoSoum.Text = vbNullString
210       End If
215     End If
  
        'Numéro de tel et fax du fournisseur
220     Set rstFRS = New ADODB.Recordset

225     Call rstFRS.Open("SELECT Telephonne, Fax FROM GRB_Fournisseur WHERE IDFRS = " & cmbFournisseur.ItemData(cmbFournisseur.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
  
230     txtTelephone.Text = rstFRS.Fields("Telephonne")
235     txtFax.Text = rstFRS.Fields("Fax")

240     Call rstFRS.Close
245     Set rstFRS = Nothing
      
        'Date
250     txtDate.Text = rstBC.Fields("DateCommande")
  
        'Commandé par
255     txtComPar.Text = rstBC.Fields("CommandePar")
  
        'Commentaire
260     If Not IsNull(rstBC.Fields("Commentaire")) Then
265       txtcommentaire.Text = rstBC.Fields("Commentaire")
270     Else
275       txtcommentaire.Text = vbNullString
280     End If
  
        'Total
285     txtTotal.Text = Conversion(rstBC.Fields("Total"), MODE_ARGENT)
  
        'Afficher les instructions de livraison
290     chkAfficherInstructions.Value = Abs(CInt(rstBC.Fields("AffichageInstructions")))
  
        'Langue d'impression
295     If rstBC.Fields("LangueImpression") = "Français" Then
300       optImpression(I_IMP_FRANCAIS).Value = True
305     Else
310       optImpression(I_IMP_ANGLAIS).Value = True
315     End If
  
320     Call rstBC.Close
325     Set rstBC = Nothing
 
330     Call RemplirListView

335     Exit Sub

AfficherErreur:

340     woups "frmBonCommande", "AfficherContenuFournisseur", Err, Erl
End Sub

Private Sub RemplirListView()

5       On Error GoTo AfficherErreur

10      Dim rstPiece    As ADODB.Recordset
15      Dim itmPiece    As ListItem
20      Dim iCompteur   As Integer
25      Dim dblEscompte As Double
30      Dim dblPrix     As Double
35      Dim sNoBC       As String

40      If m_eForm = I_RETOUR_MARCHANDISE Then
45        sNoBC = txtVotreNoSoum.Text
50      Else
55        sNoBC = txtNoBC.Text
60      End If
  
65      Call lvwBonCommande.ListItems.Clear
  
70      Set rstPiece = New ADODB.Recordset
  
75      Call rstPiece.Open("SELECT * FROM GRB_BonsCommandes_Pieces WHERE NoBonCommande = '" & sNoBC & "' AND NoFournisseur = " & m_iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)
  
80      Do While Not rstPiece.EOF
85        If Not IsNull(rstPiece.Fields("Qté")) Then
90          Set itmPiece = lvwBonCommande.ListItems.Add
    
            'Quantité
95          itmPiece.Text = rstPiece.Fields("Qté")

100         If Not IsNull(rstPiece.Fields("NuméroLigne")) Then
105           itmPiece.Tag = rstPiece.Fields("NuméroLigne")
110         End If

            'No. Item
115         itmPiece.SubItems(I_COL_NO_ITEM) = rstPiece.Fields("NoItem")
          
            'Description
120         If Not IsNull(rstPiece.Fields("Description")) Then
125           itmPiece.SubItems(I_COL_DESCR) = rstPiece.Fields("Description")
130         Else
135           itmPiece.SubItems(I_COL_DESCR) = ""
140         End If
    
            'Manufacturier
145         itmPiece.SubItems(I_COL_MANUFACT) = rstPiece.Fields("Manufact")
    
            'Prix/unité
150         If Not IsNull(rstPiece.Fields("Prix")) Then
155           itmPiece.SubItems(I_COL_PRIX) = Conversion(rstPiece.Fields("Prix"), MODE_ARGENT, 4)
160         Else
165           itmPiece.SubItems(I_COL_PRIX) = Conversion(0, MODE_ARGENT, 4)
170         End If
    
            'Escompte
175         If Trim(rstPiece.Fields("Escompte")) <> vbNullString Then
180           itmPiece.SubItems(I_COL_ESCOMPTE) = Conversion(rstPiece.Fields("Escompte"), MODE_POURCENT)
185         Else
190           itmPiece.SubItems(I_COL_ESCOMPTE) = " "
195         End If
    
            'Total
200         If Not IsNull(rstPiece.Fields("Total")) Then
205           itmPiece.SubItems(I_COL_TOTAL) = Conversion(rstPiece.Fields("Total"), MODE_ARGENT)
210         Else
215           itmPiece.SubItems(I_COL_TOTAL) = Conversion(0, MODE_ARGENT)
220         End If
225       End If

230       Call rstPiece.MoveNext
235     Loop
  
240     Call rstPiece.Close
245     Set rstPiece = Nothing

250     Exit Sub

AfficherErreur:

255     woups "frmBonCommande", "RemplirListView", Err, Erl
End Sub

Private Sub RemplirComboContacts()

5       On Error GoTo AfficherErreur

10      Dim rstContact    As ADODB.Recordset
15      Dim rstContactFRS As ADODB.Recordset
    
20      Call cmbContact.Clear
    
25      Set rstContactFRS = New ADODB.Recordset
30      Set rstContact = New ADODB.Recordset
    
35      Call rstContactFRS.Open("SELECT * FROM GRB_ContactFRS WHERE NoFRS = " & m_iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)

40      Do While Not rstContactFRS.EOF
45        Call rstContact.Open("SELECT NomContact FROM GRB_Contact WHERE IDContact = " & rstContactFRS.Fields("NoContact"), g_connData, adOpenDynamic, adLockOptimistic)

50        If Not rstContact.EOF Then
55          Call cmbContact.AddItem(rstContact.Fields("NomContact"))
60        End If

65        Call rstContact.Close

70        Call rstContactFRS.MoveNext
75      Loop

80      Call rstContactFRS.Close
  
85      If cmbContact.ListCount = 0 Then
90        Call rstContact.Open("SELECT NomContact FROM GRB_Contact WHERE Supprimé = False ORDER BY NomContact", g_connData, adOpenDynamic, adLockOptimistic)
    
95        Do While Not rstContact.EOF
100         Call cmbContact.AddItem(rstContact.Fields("NomContact"))
    
105         Call rstContact.MoveNext
110       Loop
    
115       Call rstContact.Close
120     End If

125     Set rstContact = Nothing

130     Exit Sub

AfficherErreur:

135     woups "frmBonCommande", "RemplirComboContact", Err, Erl
End Sub

Private Sub RemplirComboFournisseurs()

5       On Error GoTo AfficherErreur

10      Dim rstBC As ADODB.Recordset
15      Dim sNoBC As String
    
20      If m_eForm = I_RETOUR_MARCHANDISE Then
25        sNoBC = txtVotreNoSoum.Text
30      Else
35        sNoBC = txtNoBC.Text
40      End If

45      Set rstBC = New ADODB.Recordset
    
50      Call rstBC.Open("SELECT NoFournisseur, NomFournisseur FROM GRB_BonsCommandes INNER JOIN GRB_Fournisseur ON GRB_BonsCommandes.NoFournisseur = GRB_Fournisseur.IDFRS WHERE NoBonCommande = '" & sNoBC & "' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
  
        'Pour chaque enregistrements du recordset
55      Do While Not rstBC.EOF
          'On ajoute le nom dans le combo
60        Call cmbFournisseur.AddItem(rstBC.Fields("NomFournisseur"))
    
          'On ajoute le no dans l'itemdata du combo
65        cmbFournisseur.ItemData(cmbFournisseur.newIndex) = rstBC.Fields("NoFournisseur")
   
70        Call rstBC.MoveNext
75      Loop
  
80      Call rstBC.Close
85      Set rstBC = Nothing

        'Si le combo n'est pas vide
90      If cmbFournisseur.ListCount > 0 Then
          'On sélectionne le premier enregistrement
95        cmbFournisseur.ListIndex = 0
100     End If

105     Exit Sub

AfficherErreur:

110     woups "frmBonCommande", "RemplirComboFournisseurs", Err, Erl
End Sub

Private Sub cmdImprimer_Click()

5       On Error GoTo AfficherErreur

10      Dim rstConfig      As ADODB.Recordset
15      Dim rstBC          As ADODB.Recordset
20      Dim rstBCPiece     As ADODB.Recordset
25      Dim rstFournisseur As ADODB.Recordset
30      Dim bGRB           As Boolean
35      Dim iCompteur      As Integer
40      Dim sNoBC          As String
     
45      Screen.MousePointer = vbHourglass
  
50      If m_eForm = I_RETOUR_MARCHANDISE Then
55        sNoBC = txtVotreNoSoum.Text
60      Else
65        sNoBC = txtNoBC.Text
70      End If
  
        'Sur l'impression, on enregistre une dernière fois le bon de commande
75      Call EnregistrerModifFournisseur
  
80      Set rstBC = New ADODB.Recordset
  
85      If m_eForm = I_RETOUR_MARCHANDISE Then
90        If m_eImpRetour = MODE_RETOUR Then
95          Call rstBC.Open("SELECT * FROM GRB_BonsCommandes WHERE NoBonCommande = '" & sNoBC & "' AND NoFournisseur = " & cmbFournisseur.ItemData(cmbFournisseur.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
100       Else
105         Call rstBC.Open("SELECT * FROM GRB_BonsCommandes WHERE NoBonCommande = '" & sNoBC & "'", g_connData, adOpenDynamic, adLockOptimistic)
110       End If
115     Else
120       Call rstBC.Open("SELECT * FROM GRB_BonsCommandes WHERE NoBonCommande = '" & sNoBC & "'", g_connData, adOpenDynamic, adLockOptimistic)
125     End If
    
130     If m_eForm = I_ACHAT_ELEC Or m_eForm = I_PROJET_ELEC Then
135       Do While Not rstBC.EOF
140         If rstBC.Fields("DateRequise") = "" Then
145           Set rstFournisseur = New ADODB.Recordset

150           Call rstFournisseur.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)

155           Call MsgBox("Le date requise est nécessaire pour le fournisseur " & rstFournisseur.Fields("NomFournisseur") & "!", vbOKOnly, "Erreur")
    
160           Call rstFournisseur.Close
165           Set rstFournisseur = Nothing

170           Call rstBC.Close
175           Set rstBC = Nothing

180           Screen.MousePointer = vbDefault

185           Exit Sub
190         End If

195         Call rstBC.MoveNext
200       Loop
    
205       Call rstBC.MoveFirst
210     End If
    
215     Set rstBCPiece = New ADODB.Recordset
220     Set rstFournisseur = New ADODB.Recordset
225     Set rstConfig = New ADODB.Recordset
    
230     rstBCPiece.CursorLocation = adUseClient
    
235     Do While Not rstBC.EOF
240       Call rstBCPiece.Open("SELECT * FROM GRB_BonsCommandes_Pieces WHERE NoBonCommande = '" & sNoBC & "' AND NoFournisseur = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)
          ''''''''''''''''''''''''''''''''''''''''''''''''''''
          ' Met au minimum 15 lignes pour un bon de commande '
          ''''''''''''''''''''''''''''''''''''''''''''''''''''
245       If rstBCPiece.RecordCount < 15 Then
250         iCompteur = 15 - rstBCPiece.RecordCount
      
255         Do While Not iCompteur = 0
              'Ajoute une ligne vide
260           Call rstBCPiece.AddNew
        
265           rstBCPiece.Fields("NoBonCommande") = rstBC.Fields("NoBonCommande")
270           rstBCPiece.Fields("NoFournisseur") = rstBC.Fields("NoFournisseur")
275           rstBCPiece.Fields("Type") = rstBC.Fields("Type")
        
280           Call rstBCPiece.Update
        
285           iCompteur = iCompteur - 1
290         Loop
295       End If
          
          'Ouvre la table fournisseur
300       Call rstFournisseur.Open("SELECT * FROM GRB_Fournisseur WHERE IDFRS = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)
    
          'Ouvre la table config
305       Call rstConfig.Open("SELECT parcel_label_line1, parcel_label_line2, parcel_label_line3, ParcelAssist, ParcelEtat FROM GRB_Config", g_connData, adOpenDynamic, adLockOptimistic)
  
          ''''''''''''''''''''''''''''''''''''''
          ' U.S. PARCEL SERVICE SHIPMENTS ONLY '
          ''''''''''''''''''''''''''''''''''''''
310       If rstBC.Fields("AffichageInstructions") = True Then
            'Orientation de la page
            'Printer.Orientation = vbPRORPortrait
315         DR_Commande_parcel.Orientation = rptOrientPortrait
  
            'Affiche les données
320         DR_Commande_parcel.Sections("section4").Controls("lblcompagnie").Caption = rstConfig.Fields("parcel_label_line1")
325         DR_Commande_parcel.Sections("section4").Controls("lbladresse").Caption = rstConfig.Fields("parcel_label_line2")
330         DR_Commande_parcel.Sections("section4").Controls("lblpays").Caption = rstConfig.Fields("parcel_label_line3")
335         DR_Commande_parcel.Sections("section4").Controls("lblassist").Caption = "Should you have any questions, do not hesitate to call " & rstConfig.Fields("ParcelAssist") & ", it will be our pleasure to assist you."
340         DR_Commande_parcel.Sections("section4").Controls("lblreminder").Caption = "Please note that you are shipping to a " & rstConfig.Fields("ParcelEtat") & " address and therefore your shipment is considered as domestic."
 
            'Ouvre le rapport
345         Set DR_Commande_parcel.DataSource = rstConfig
  
350         Call DR_Commande_parcel.Show(vbModal)
355       End If
    
          ''''''''''''
          ' Commande '
          ''''''''''''
360       If m_eForm = I_RETOUR_MARCHANDISE Then
365         If m_eImpRetour = MODE_DEMANDE_RETOUR Then
370           DR_Commande.Caption = "Demande de retour de marchandise"
375         Else
380           DR_Commande.Caption = "Retour de marchandise"
385         End If
390       Else
395         DR_Commande.Caption = "Commande"
400       End If
  
405       If rstBC.Fields("LangueImpression") = "Anglais" Then
410         If m_eForm = I_RETOUR_MARCHANDISE Then
415           DR_Commande.Sections("Section2").Controls("lblTitrebc").Caption = "RMA #"

420           If m_eImpRetour = MODE_DEMANDE_RETOUR Then
425             DR_Commande.Sections("Section2").Controls("lblTitreCommande").Caption = "RMA REQUEST"
430           Else
435             DR_Commande.Sections("Section2").Controls("lblTitreCommande").Caption = "RETURN ORDER"
440           End If

445           DR_Commande.Sections("section2").Controls("lbltitreNoSoum").Caption = "Our #"
450         Else
455           DR_Commande.Sections("Section2").Controls("lbltitrebc").Caption = "PO #"
460           DR_Commande.Sections("Section2").Controls("lbltitrecommande").Caption = "PURCHASE ORDER"
465           DR_Commande.Sections("section2").Controls("lbltitreNoSoum").Caption = "Your ref #"
470         End If

475         DR_Commande.Sections("Section3").Controls("lbltitreCommentaire").Caption = "Comments:"
480         DR_Commande.Sections("section2").Controls("lbltitrecompar").Caption = "Purchaser:"
485         DR_Commande.Sections("section2").Controls("lbltitrecontact").Caption = "ATT:"
490         DR_Commande.Sections("section2").Controls("lbltitredate").Caption = "Date:"
495         DR_Commande.Sections("section2").Controls("lbltitredatereq").Caption = "Date required"
500         DR_Commande.Sections("section2").Controls("lbltitredescription").Caption = "Description"
505         DR_Commande.Sections("section2").Controls("lbltitreescompte").Caption = "Discount"
510         DR_Commande.Sections("section2").Controls("lbltitrefax").Caption = "Fax"
515         DR_Commande.Sections("section2").Controls("lbltitrefournisseur").Caption = "SUPPLIER:"
520         DR_Commande.Sections("section2").Controls("lbltitremanufact").Caption = "Manufacturer"
525         DR_Commande.Sections("section2").Controls("lbltitrePiece").Caption = "Part Number"
530         DR_Commande.Sections("section2").Controls("lbltitrepage").Caption = "Page:"
535         DR_Commande.Sections("section2").Controls("lblPage").Caption = "%p of %P"
540         DR_Commande.Sections("section2").Controls("lbltitreprix").Caption = "Unit Price"
545         DR_Commande.Sections("section2").Controls("lbltitreqte").Caption = "Qty"
550         DR_Commande.Sections("section2").Controls("lbltitretel").Caption = "Telephone"
555         DR_Commande.Sections("section2").Controls("lbltitretotal").Caption = "Total"
560         DR_Commande.Sections("Section3").Controls("lbltitretotalfin").Caption = "TOTAL"
565         DR_Commande.Sections("section2").Controls("lbltitretransport").Caption = "TRANSPORT"
570         DR_Commande.Sections("Section3").Controls("lbltypeprix").Caption = rstFournisseur.Fields("pays") + " Funds"
575         DR_Commande.Sections("Section3").Controls("lblPiedPage").Caption = "CONFIRM THE ORDER AND SHIPPING DATE"

580         DR_Commande.Sections("Section2").Controls("imgLogoFrancais").Visible = False
585         DR_Commande.Sections("Section2").Controls("imgLogoAnglais").Visible = True
590       Else
595         If m_eForm = I_RETOUR_MARCHANDISE Then
600           DR_Commande.Sections("Section2").Controls("lblTitreBC").Caption = "# RMA"

605           If m_eImpRetour = MODE_DEMANDE_RETOUR Then
610             DR_Commande.Sections("Section2").Controls("lblTitreCommande").Caption = "DEMANDE DE RETOUR DE MARCHANDISE"
615           Else
620             DR_Commande.Sections("Section2").Controls("lblTitreCommande").Caption = "RETOUR DE MARCHANDISE"
625           End If

630           DR_Commande.Sections("Section2").Controls("lblTitreNoSoum").Caption = "Notre #"
635         Else
640           DR_Commande.Sections("Section2").Controls("lblTitreBC").Caption = "BC #"
645           DR_Commande.Sections("Section2").Controls("lblTitreCommande").Caption = "COMMANDE"
650           DR_Commande.Sections("Section2").Controls("lblTitreNoSoum").Caption = "Votre # soum"
655         End If
660       End If

665       If m_eForm = I_RETOUR_MARCHANDISE Then
670         If m_eImpRetour = MODE_RETOUR Then
675           DR_Commande.Sections("Section3").Controls("lblCopieCredit").Visible = True
680         End If
685       End If

690       If m_eForm = I_RETOUR_MARCHANDISE Then
695         If Not IsNull(rstBC.Fields("VotreNoSoum")) Then
700           DR_Commande.Sections("Section2").Controls("lblNoBC").Caption = rstBC.Fields("VotreNoSoum")
705         Else
710           DR_Commande.Sections("Section2").Controls("lblNoBC").Caption = vbNullString
715         End If
720       Else
725         DR_Commande.Sections("section2").Controls("lblNoBC").Caption = rstBC.Fields("NoBonCommande")
730       End If

735       If m_eForm = I_RETOUR_MARCHANDISE Then
740         DR_Commande.Sections("Section3").Controls("lblPiedPage").Visible = False
750       Else
755         DR_Commande.Sections("Section3").Controls("lblPiedPage").Visible = True
765       End If

770       If Not IsNull(rstBC.Fields("Commentaire")) Then
775         DR_Commande.Sections("Section3").Controls("lblCommentaire").Caption = rstBC.Fields("Commentaire")
780       Else
785         DR_Commande.Sections("Section3").Controls("lblCommentaire").Caption = vbNullString
790       End If

795       DR_Commande.Sections("section2").Controls("lblCommandePar").Caption = rstBC.Fields("CommandePar")

800       If Not IsNull(rstBC.Fields("Attention")) Then
805         DR_Commande.Sections("section2").Controls("lblContact").Caption = rstBC.Fields("Attention")
810       Else
815         DR_Commande.Sections("Section2").Controls("lblContact").Caption = vbNullString
820       End If

825       DR_Commande.Sections("Section2").Controls("lblDate").Caption = rstBC.Fields("DateCommande")
   
830       If Not IsNull(rstBC.Fields("DateRequise")) Then
835         DR_Commande.Sections("Section2").Controls("lblDateRequise").Caption = rstBC.Fields("DateRequise")
840       Else
845         DR_Commande.Sections("Section2").Controls("lblDateRequise").Caption = vbNullString
850       End If
   
855       DR_Commande.Sections("section2").Controls("lblFax").Caption = rstFournisseur.Fields("Fax")
860       DR_Commande.Sections("section2").Controls("lblFournisseur").Caption = rstFournisseur.Fields("NomFournisseur")
865       DR_Commande.Sections("section2").Controls("lblTel").Caption = rstFournisseur.Fields("telephonne")

870       If m_eForm = I_RETOUR_MARCHANDISE Then
875         DR_Commande.Sections("Section2").Controls("lblNoSoum").Caption = rstBC.Fields("NoBonCommande")
880       Else
885         If Not IsNull(rstBC.Fields("VotreNoSoum")) Then
890           DR_Commande.Sections("Section2").Controls("lblNoSoum").Caption = rstBC.Fields("VotreNoSoum")
895         Else
900           DR_Commande.Sections("Section2").Controls("lblNoSoum").Caption = vbNullString
905         End If
910       End If

915       DR_Commande.Sections("Section3").Controls("lblTotalFin").Caption = Conversion(rstBC.Fields("total"), MODE_ARGENT)
    
920       If Not IsNull(rstBC.Fields("Transport")) Then
925         DR_Commande.Sections("section2").Controls("lblTransport").Caption = rstBC.Fields("Transport")
930       Else
935         DR_Commande.Sections("section2").Controls("lblTransport").Caption = " "
940       End If

945       If m_eForm = I_ACHAT_ELEC Or m_eForm = I_INVENTAIRE_ELEC Or m_eForm = I_PROJET_ELEC Then
950         DR_Commande.Sections("Section3").Controls("lblCSA").Visible = True
955       End If
  
          'Si on affiche adresse livraison dans commentaire
960       If rstBC.Fields("AffichageInstructions") = True Then
965         Call rstConfig.Requery
    
970         DR_Commande.Sections("Section3").Controls("lblCommentaire").Caption = "Shipping Address:" & vbNewLine & DR_Commande.Sections("Section3").Controls("lblCommentaire").Caption
975       End If
    rstBCPiece.MoveFirst
   Do While rstBCPiece.EOF = False
    
        If rstBCPiece.Fields("NoItem") <> vbNull Then
           If Len(rstBCPiece.Fields("NoItem")) > 26 Then
            DR_Commande.Sections("section1").Controls("text2").Font.SIZE = 8
           End If
        End If
        rstBCPiece.MoveNext
    Loop
980       Set DR_Commande.DataSource = rstBCPiece
    
985       DR_Commande.Orientation = rptOrientLandscape
    
990       Call DR_Commande.Show(vbModal)

995       If m_eForm <> I_RETOUR_MARCHANDISE Then
1000        If UCase(rstFournisseur.Fields("NomFournisseur")) = "SOLUTION GRB INC." Then
1005          DR_Commande_recu.Orientation = rptOrientLandscape

1010          If m_eForm = I_PROJET_ELEC Or m_eForm = I_PROJET_MEC Then
1015            Call rstBCPiece.Close

1020            If m_eForm = I_PROJET_ELEC Then
1025              Call rstBCPiece.Open("SELECT * FROM GRB_BonsCommandes_Pieces LEFT JOIN GRB_InventaireElec ON GRB_BonsCommandes_Pieces.NoItem = GRB_InventaireElec.NoItem WHERE NoBonCommande = '" & sNoBC & "' AND NoFournisseur = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)
1030            Else
1035              Call rstBCPiece.Open("SELECT * FROM GRB_BonsCommandes_Pieces LEFT JOIN GRB_InventaireMec ON GRB_BonsCommandes_Pieces.NoItem = GRB_InventaireMec.NoItem WHERE NoBonCommande = '" & sNoBC & "' AND NoFournisseur = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)
1040            End If

1045            DR_Commande_recu.Sections("Section1").Controls("txtNoItem").DataField = "GRB_BonsCommandes_Pieces.NoItem"
1050            DR_Commande_recu.Sections("Section1").Controls("txtDescription").DataField = "GRB_BonsCommandes_Pieces.Description"
1055          Else
1060            If m_eForm = I_ACHAT_ELEC Or m_eForm = I_ACHAT_MEC Then
1065              Call rstBCPiece.Close

1070              If m_eForm = I_ACHAT_ELEC Then
1075                Call rstBCPiece.Open("SELECT * FROM GRB_BonsCommandes_Pieces LEFT JOIN GRB_InventaireElec ON GRB_BonsCommandes_Pieces.NoItem = GRB_InventaireElec.NoItem WHERE NoBonCommande = '" & sNoBC & "' AND NoFournisseur = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)
1080              Else
1085                Call rstBCPiece.Open("SELECT * FROM GRB_BonsCommandes_Pieces LEFT JOIN GRB_InventaireMec ON GRB_BonsCommandes_Pieces.NoItem = GRB_InventaireMec.NoItem WHERE NoBonCommande = '" & sNoBC & "' AND NoFournisseur = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)
1090              End If
                Dim testgll As String
                testgll = "GRB_BonsCommandes_Pieces.NoItem"
1095              DR_Commande_recu.Sections("Section1").Controls("txtNoItem").DataField = "GRB_BonsCommandes_Pieces.NoItem"
1100              DR_Commande_recu.Sections("Section1").Controls("txtDescription").DataField = "GRB_BonsCommandes_Pieces.Description"

1105            End If
1110          End If

1115          Set DR_Commande_recu.DataSource = rstBCPiece

1120          DR_Commande_recu.Sections("Section2").Controls("lblfournisseur").Caption = rstFournisseur.Fields("NomFournisseur")
1125          DR_Commande_recu.Sections("Section2").Controls("lblprojet").Caption = rstBC.Fields("NoProjet")
1130          DR_Commande_recu.Sections("Section5").Controls("lbldatereq").Caption = rstBC.Fields("DateRequise")

1135          Call DR_Commande_recu.Show(vbModal)
1140        End If
1145      Else
1150        If m_eForm = I_RETOUR_MARCHANDISE Then
1155          If m_eImpRetour = MODE_RETOUR Then
1160            Call ImprimerRetour(rstBC.Fields("NoBonCommande"), rstBC.Fields("NoFournisseur"), rstBC.Fields("VotreNoSoum"))
1165            Call ImprimerRetourDossier(rstBC.Fields("NoBonCommande"), rstBC.Fields("NoFournisseur"))
1170          End If
1175        End If
1180      End If

          'Prochain enregistrement
1185      Call rstBC.MoveNext
    
1190      Call rstBCPiece.Close
    
1195      Call rstConfig.Close

1200      Call rstFournisseur.Close
1205    Loop

1210    Set rstFournisseur = Nothing
1215    Set rstConfig = Nothing
1220    Set rstBCPiece = Nothing

1225    Call g_connData.Execute("DELETE * FROM GRB_BonsCommandes_Pieces WHERE NoBonCommande = '" & sNoBC & "'")
1230    Call g_connData.Execute("DELETE * FROM GRB_BonsCommandes WHERE NoBonCommande = '" & sNoBC & "'")

1235    Call Unload(Me)
    
1240    Screen.MousePointer = vbDefault

1245    Exit Sub

AfficherErreur:

1250  woups "frmBonCommande", "cmdImprimer_Click", Err, Erl
End Sub

Private Sub ImprimerRetour(ByVal sNoRetour As String, ByVal iNoFRS As Integer, ByVal sNoRMA As String)
  
5       On Error GoTo AfficherErreur

10      Dim rstBCPiece As ADODB.Recordset
15      Dim rstFRS     As ADODB.Recordset
  
20      Set rstBCPiece = New ADODB.Recordset
  
25      Call rstBCPiece.Open("SELECT * FROM GRB_BonsCommandes_Pieces WHERE NoBonCommande = '" & sNoRetour & "' AND NoFournisseur = " & iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)
    
30      Set DR_Retour.DataSource = rstBCPiece
  
35      DR_Retour.Orientation = rptOrientLandscape
  
40      Set rstFRS = New ADODB.Recordset
  
45      Call rstFRS.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)
  
50      DR_Retour.Sections("Section2").Controls("lblFournisseur").Caption = rstFRS.Fields("NomFournisseur")
  
55      Call rstFRS.Close
60      Set rstFRS = Nothing
  
65      DR_Retour.Sections("Section2").Controls("lblNoProjet").Caption = sNoRetour
70      DR_Retour.Sections("Section2").Controls("lblNoRMA").Caption = sNoRMA
75      DR_Retour.Sections("Section2").Controls("lblDate").Caption = ConvertDate(Date)
  
80      Call DR_Retour.Show(vbModal)
  
85      Call rstBCPiece.Close
90      Set rstBCPiece = Nothing

95      Exit Sub

AfficherErreur:

100     woups "frmBonCommande", "ImprimerRetour", Err, Erl
End Sub

Private Sub ImprimerRetourDossier(ByVal sNoRetour As String, ByVal iNoFRS As Integer)

5       On Error GoTo AfficherErreur

10      Dim rstBC      As ADODB.Recordset
15      Dim rstBCPiece As ADODB.Recordset
20      Dim rstFRS     As ADODB.Recordset

25      Set rstBC = New ADODB.Recordset
30      Set rstBCPiece = New ADODB.Recordset

35      Call rstBC.Open("SELECT * FROM GRB_BonsCommandes WHERE NoBonCommande = '" & sNoRetour & "' AND NoFournisseur = " & iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)

40      Call rstBCPiece.Open("SELECT * FROM GRB_BonsCommandes_Pieces WHERE NoBonCommande = '" & sNoRetour & "' AND NoFournisseur = " & iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)
        
45      Set DR_Retour.DataSource = rstBCPiece
  
50      DR_Retour.Orientation = rptOrientLandscape
  
55      Set rstFRS = New ADODB.Recordset
  
60      Call rstFRS.Open("SELECT * FROM GRB_Fournisseur WHERE IDFRS = " & iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)
     
65      DR_Commande.Orientation = rptOrientLandscape

70      DR_Commande.Caption = "Retour de marchandise"
  
75      DR_Commande.Sections("Section2").Controls("lblTitreBC").Caption = "# RMA"

80      DR_Commande.Sections("Section2").Controls("lblTitreCommande").Caption = "RETOUR DE MARCHANDISE"

85      DR_Commande.Sections("Section2").Controls("lblTitreNoSoum").Caption = "Notre #"

90      If Not IsNull(rstBC.Fields("VotreNoSoum")) Then
95        DR_Commande.Sections("Section2").Controls("lblNoBC").Caption = rstBC.Fields("VotreNoSoum")
100     Else
105       DR_Commande.Sections("Section2").Controls("lblNoBC").Caption = vbNullString
110     End If
    
115     If Not IsNull(rstBC.Fields("Commentaire")) Then
120       DR_Commande.Sections("Section3").Controls("lblCommentaire").Caption = rstBC.Fields("Commentaire")
125     Else
130       DR_Commande.Sections("Section3").Controls("lblCommentaire").Caption = vbNullString
135     End If
    
140     DR_Commande.Sections("section2").Controls("lblCommandePar").Caption = rstBC.Fields("CommandePar")

145     If Not IsNull(rstBC.Fields("Attention")) Then
150       DR_Commande.Sections("section2").Controls("lblContact").Caption = rstBC.Fields("Attention")
155     Else
160       DR_Commande.Sections("Section2").Controls("lblContact").Caption = vbNullString
165     End If

170     DR_Commande.Sections("section2").Controls("lblDate").Caption = rstBC.Fields("DateCommande")
    
175     If Not IsNull(rstBC.Fields("DateRequise")) Then
180       DR_Commande.Sections("section2").Controls("lblDateRequise").Caption = rstBC.Fields("DateRequise")
185     Else
190       DR_Commande.Sections("section2").Controls("lblDateRequise").Caption = vbNullString
195     End If
    
200     DR_Commande.Sections("section2").Controls("lblFax").Caption = rstFRS.Fields("Fax")
205     DR_Commande.Sections("section2").Controls("lblFournisseur").Caption = rstFRS.Fields("NomFournisseur")
210     DR_Commande.Sections("section2").Controls("lblTel").Caption = rstFRS.Fields("telephonne")
   
215     DR_Commande.Sections("Section2").Controls("lblNoSoum").Caption = rstBC.Fields("NoBonCommande")
   
220     DR_Commande.Sections("Section3").Controls("lblTotalFin").Caption = Conversion(rstBC.Fields("total"), MODE_ARGENT)
    
225     If Not IsNull(rstBC.Fields("Transport")) Then
230       DR_Commande.Sections("section2").Controls("lblTransport").Caption = rstBC.Fields("Transport")
235     Else
240       DR_Commande.Sections("section2").Controls("lblTransport").Caption = " "
245     End If
    
250     DR_Commande.Sections("Section3").Controls("lblPiedPage").Visible = False
    
260     Set DR_Commande.DataSource = rstBCPiece
   
265     Call DR_Commande.Show(vbModal)
  
270     Call rstFRS.Close
275     Set rstFRS = Nothing

280     Call rstBC.Close
285     Set rstBC = Nothing

290     Call rstBCPiece.Close
295     Set rstBCPiece = Nothing
  
300     Screen.MousePointer = vbDefault

305     Exit Sub

AfficherErreur:

310     woups "frmBonCommande", "ImprimerRetourDossier", Err, Erl
End Sub

Private Sub cmdInstructions_Click()

5       On Error GoTo AfficherErreur

10      Call OuvrirForm(FrmBonCommande_Instruction, True)

15      Exit Sub

AfficherErreur:

20      woups "frmBonCommande", "cmdInstructions_Click", Err, Erl
End Sub



Private Sub mvwDate_DateClick(ByVal DateClicked As Date)

5       On Error GoTo AfficherErreur

10      txtDateRequise.Text = ConvertDate(DateClicked)
  
15      mvwDate.Visible = False

20      Exit Sub

AfficherErreur:

25      woups "frmBonCommande", "mvwDate_DateClick", Err, Erl
End Sub

Private Sub mvwDate_LostFocus()

5       On Error GoTo AfficherErreur

10      mvwDate.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmBonCommande", "mvwDate_LostFocus", Err, Erl
End Sub

Private Sub EnregistrerBonCommandeRetourMarchandiseProjet()

5       On Error GoTo AfficherErreur

10      Dim rstBC         As ADODB.Recordset
15      Dim rstBCPiece    As ADODB.Recordset
20      Dim rstPiece      As ADODB.Recordset
25      Dim rstFRS        As ADODB.Recordset
30      Dim iCompteur     As Integer
35      Dim dblTotal      As Double
40      Dim sWhere        As String
45      Dim sWherePiece   As String
50      Dim sWhereNoLigne As String
55      Dim sEscompte     As String
    
        'Recordset source
60      sWhere = "(IDProjet = '" & m_sNoProjet & "')"
        
65      sWherePiece = "GRB_Projet_Pieces.NumItem In ("
70      sWhereNoLigne = "GRB_Projet_Pieces.NuméroLigne In ("
        
75      For iCompteur = 1 To m_collPieces.count
80        If iCompteur <> m_collPieces.count Then
85          sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "', "
90          sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ", "
95        Else
100         sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "')"
105         sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ")"
110       End If
115     Next

120     Set rstFRS = New ADODB.Recordset
125     Set rstBC = New ADODB.Recordset
130     Set rstPiece = New ADODB.Recordset
135     Set rstBCPiece = New ADODB.Recordset
  
140     sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne
  
145     Call rstFRS.Open("SELECT DISTINCT GRB_Projet_Pieces.IDFRS, GRB_Fournisseur.CondTransport FROM GRB_Projet_Pieces INNER JOIN GRB_Fournisseur ON GRB_Projet_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
    
        'Recordsets destinations
150     Call rstBC.Open("SELECT * FROM GRB_BonsCommandes", g_connData, adOpenDynamic, adLockOptimistic)
  
155     Do While Not rstFRS.EOF
160       Call rstBC.AddNew
    
          'Enregistrement du bon
165       rstBC.Fields("NoBonCommande") = txtVotreNoSoum.Text
170       rstBC.Fields("NoFournisseur") = rstFRS.Fields("IDFRS")
175       rstBC.Fields("NoProjet") = m_sNoProjet
180       rstBC.Fields("Attention") = ""
    
185       If Not IsNull(rstFRS.Fields("CondTransport")) And rstFRS.Fields("CondTransport") <> vbNullString Then
190         rstBC.Fields("Transport") = rstFRS.Fields("CondTransport")
195       Else
200         rstBC.Fields("Transport") = "Votre camion"
205       End If
    
210       rstBC.Fields("DateRequise") = ConvertDate(Date)
215       rstBC.Fields("DateCommande") = ConvertDate(Date)

220       If m_eForm = I_RETOUR_MARCHANDISE Then
225         rstBC.Fields("CommandePar") = m_sEmploye
230       Else
235         rstBC.Fields("CommandePar") = g_sEmploye
240       End If

245       rstBC.Fields("LangueImpression") = "Français"
           
250       sWhere = "(IDProjet = '" & m_sNoProjet & "' AND IDFRS = " & rstFRS.Fields("IDFRS") & ")"

255       sWherePiece = "NumItem In ("
260       sWhereNoLigne = "NuméroLigne In ("
           
265       For iCompteur = 1 To m_collPieces.count
270         If iCompteur <> m_collPieces.count Then
275           sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "', "
280           sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ", "
285         Else
290           sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "')"
295           sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ")"
300         End If
305       Next
           
310       sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne
           
315       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
                   
320       dblTotal = 0
           
          'Enregistrement des pièces
325       Do While Not rstPiece.EOF
330         Call rstBCPiece.Open("SELECT * FROM GRB_BonsCommandes_Pieces WHERE NoItem = '" & Replace(rstPiece.Fields("NumItem"), "'", "''") & "' AND NoFournisseur = " & rstPiece.Fields("IDFRS") & " AND NoBonCommande = '" & txtVotreNoSoum.Text & "' AND Prix = '" & rstPiece.Fields("PrixOrigine") & "'", g_connData, adOpenDynamic, adLockOptimistic)
      
335         If Trim(rstPiece.Fields("ESCOMPTE")) <> "" Then
340           sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
            
345           Do While CDbl(sEscompte) > 1
350             sEscompte = CDbl(sEscompte) / 100
355           Loop

360           dblTotal = dblTotal + CDbl(Conversion(CStr((Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * (1 - CDbl(sEscompte))) * CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString))), MODE_DECIMAL))
365         Else
370           dblTotal = dblTotal + CDbl(Conversion(CStr(Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * Replace(rstPiece.Fields("Qté"), "-", vbNullString)), MODE_DECIMAL))
375         End If
      
380         If rstBCPiece.EOF Then
385           Call rstBCPiece.AddNew

390           rstBCPiece.Fields("NoBonCommande") = txtVotreNoSoum.Text
395           rstBCPiece.Fields("NoFournisseur") = rstPiece.Fields("IDFRS")
400           rstBCPiece.Fields("Qté") = Replace(rstPiece.Fields("Qté"), "-", vbNullString)
        
405           rstBCPiece.Fields("NoItem") = rstPiece.Fields("NumItem")

410           rstBCPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne")
        
415           rstBCPiece.Fields("Description") = rstPiece.Fields("Desc_fr")
420           rstBCPiece.Fields("Manufact") = rstPiece.Fields("Manufact")
        
425           rstBCPiece.Fields("Prix") = rstPiece.Fields("PrixOrigine")
        
430           If rstPiece.Fields("Escompte") <> vbNullString Then
435             rstBCPiece.Fields("Escompte") = rstPiece.Fields("Escompte")
440           Else
445             rstBCPiece.Fields("Escompte") = "0"
450           End If
      
455           If Trim(rstPiece.Fields("Escompte")) <> "" Then
460             sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
            
465             Do While CDbl(sEscompte) > 1
470               sEscompte = CDbl(sEscompte) / 100
475             Loop

480             rstBCPiece.Fields("Total") = Conversion(CStr((Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * (1 - CDbl(sEscompte))) * CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString))), MODE_DECIMAL)
485           Else
490             rstBCPiece.Fields("Total") = Conversion(CStr(Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * Replace(rstPiece.Fields("Qté"), "-", vbNullString)), MODE_DECIMAL)
495           End If
500         Else
505           rstBCPiece.Fields("Qté") = CDbl(rstBCPiece.Fields("Qté")) + CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString))

510           rstBCPiece.Fields("NuméroLigne") = rstBCPiece.Fields("NuméroLigne") & ", " & rstPiece.Fields("NuméroLigne")

515           If Trim(rstPiece.Fields("Escompte")) <> "" Then
520             sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
            
525             Do While CDbl(sEscompte) > 1
530               sEscompte = CDbl(sEscompte) / 100
535             Loop

540             rstBCPiece.Fields("Total") = Replace(rstBCPiece.Fields("Total"), ".", ",") + (CDbl(Replace(rstPiece.Fields("PrixOrigine"), ".", ",")) * (1 - CDbl(sEscompte)) * CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString)))
545           Else
550             rstBCPiece.Fields("Total") = Replace(rstBCPiece.Fields("Total"), ".", ",") + (CDbl(Replace(rstPiece.Fields("PrixOrigine"), ".", ",")) * CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString)))
555           End If
560         End If
   
565         Call rstBCPiece.Update
      
570         Call rstBCPiece.Close
      
575         Call rstPiece.MoveNext
580       Loop

585       Call rstPiece.Close
    
590       rstBC.Fields("Total") = Conversion(CStr(dblTotal), MODE_DECIMAL)
    
595       Call rstBC.Update
   
600       Call rstFRS.MoveNext
605     Loop
  
610     Call rstFRS.Close
615     Set rstFRS = Nothing
    
620     Call rstBC.Close
625     Set rstBC = Nothing

630     Set rstPiece = Nothing
635     Set rstBCPiece = Nothing

640     Exit Sub

AfficherErreur:

645     woups "frmBonCommande", "EnregistrerBonCommandeRetourMarchandiseProjet", Err, Erl
End Sub

Private Sub EnregistrerBonCommandeRetourMarchandiseAchat()

5       On Error GoTo AfficherErreur

10      Dim rstBC         As ADODB.Recordset
15      Dim rstBCPiece    As ADODB.Recordset
20      Dim rstPiece      As ADODB.Recordset
25      Dim rstFRS        As ADODB.Recordset
30      Dim iCompteur     As Integer
35      Dim dblTotal      As Double
40      Dim sWhere        As String
45      Dim sWherePiece   As String
50      Dim sWhereNoLigne As String
55      Dim sEscompte     As String
    
        'Recordset source
60      sWhere = "(IDAchat = '" & m_sNoAchat & "' AND IndexAchat = " & m_iIndexAchat & ")"

65      sWherePiece = "GRB_Achat_Pieces.PIECE In ("
70      sWhereNoLigne = "GRB_Achat_Pieces.NuméroLigne In ("
        
75      For iCompteur = 1 To m_collPieces.count
80        If iCompteur <> m_collPieces.count Then
85          sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "', "
90          sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ", "
95        Else
100         sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "')"
105         sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ")"
110       End If
115     Next
  
120     Set rstFRS = New ADODB.Recordset
125     Set rstBC = New ADODB.Recordset
130     Set rstPiece = New ADODB.Recordset
135     Set rstBCPiece = New ADODB.Recordset
  
140     sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne
  
145     Call rstFRS.Open("SELECT DISTINCT GRB_Achat_Pieces.IDFRS, GRB_Fournisseur.CondTransport FROM GRB_Achat_Pieces INNER JOIN GRB_Fournisseur ON GRB_Achat_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
    
        'Recordsets destinations
150     Call rstBC.Open("SELECT * FROM GRB_BonsCommandes", g_connData, adOpenDynamic, adLockOptimistic)
  
155     Do While Not rstFRS.EOF
160       Call rstBC.AddNew
   
          'Enregistrement du bon
165       rstBC.Fields("NoBonCommande") = txtVotreNoSoum.Text
170       rstBC.Fields("NoFournisseur") = rstFRS.Fields("IDFRS")
175       rstBC.Fields("NoProjet") = m_sNoAchat & " - " & m_iIndexAchat
180       rstBC.Fields("Attention") = ""
    
185       If Not IsNull(rstFRS.Fields("CondTransport")) And rstFRS.Fields("CondTransport") <> vbNullString Then
190         rstBC.Fields("Transport") = rstFRS.Fields("CondTransport")
195       Else
200         rstBC.Fields("Transport") = "Votre camion"
205       End If
    
210       rstBC.Fields("DateRequise") = ConvertDate(Date)
215       rstBC.Fields("DateCommande") = ConvertDate(Date)

220       If m_eForm = I_RETOUR_MARCHANDISE Then
225         rstBC.Fields("CommandePar") = m_sEmploye
230       Else
235         rstBC.Fields("CommandePar") = g_sEmploye
240       End If

245       rstBC.Fields("LangueImpression") = "Français"
           
250       sWhere = "(IDAchat = '" & m_sNoAchat & "' AND IndexAchat = " & m_iIndexAchat & " AND IDFRS = " & rstFRS.Fields("IDFRS") & ")"
          
255       sWherePiece = "PIECE In ("
260       sWhereNoLigne = "NuméroLigne In ("
           
265       For iCompteur = 1 To m_collPieces.count
270         If iCompteur <> m_collPieces.count Then
275           sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "', "
280           sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ", "
285         Else
290           sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "')"
295           sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ")"
300         End If
305       Next
           
310       sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne
           
315       Call rstPiece.Open("SELECT * FROM GRB_Achat_Pieces WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
                   
320       dblTotal = 0
           
          'Enregistrement des pièces
325       Do While Not rstPiece.EOF
330         Call rstBCPiece.Open("SELECT * FROM GRB_BonsCommandes_Pieces WHERE NoItem = '" & Replace(rstPiece.Fields("PIECE"), "'", "''") & "' AND NoFournisseur = " & rstPiece.Fields("IDFRS") & " AND NoBonCommande = '" & txtVotreNoSoum.Text & "' AND Prix = '" & rstPiece.Fields("PrixOrigine") & "'", g_connData, adOpenDynamic, adLockOptimistic)
      
335         If Trim(rstPiece.Fields("ESCOMPTE")) <> "" Then
340           sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
            
345           Do While CDbl(sEscompte) > 1
350             sEscompte = CDbl(sEscompte) / 100
355           Loop

360           dblTotal = dblTotal + CDbl(Conversion(CStr((Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * (1 - CDbl(sEscompte))) * CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString))), MODE_DECIMAL))
365         Else
370           dblTotal = dblTotal + CDbl(Conversion(CStr(Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * Replace(rstPiece.Fields("Qté"), "-", vbNullString)), MODE_DECIMAL))
375         End If
      
380         If rstBCPiece.EOF Then
385           Call rstBCPiece.AddNew

390           rstBCPiece.Fields("NoBonCommande") = txtVotreNoSoum.Text
395           rstBCPiece.Fields("NoFournisseur") = rstPiece.Fields("IDFRS")
400           rstBCPiece.Fields("Qté") = Replace(rstPiece.Fields("Qté"), "-", vbNullString)
        
405           rstBCPiece.Fields("NoItem") = rstPiece.Fields("PIECE")

410           rstBCPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne")
        
415           rstBCPiece.Fields("Description") = rstPiece.Fields("Desc_fr")
420           rstBCPiece.Fields("Manufact") = rstPiece.Fields("Manufact")
        
425           rstBCPiece.Fields("Prix") = rstPiece.Fields("PrixOrigine")
        
430           If rstPiece.Fields("Escompte") <> vbNullString Then
435             rstBCPiece.Fields("Escompte") = rstPiece.Fields("Escompte")
440           Else
445             rstBCPiece.Fields("Escompte") = "0"
450           End If
      
455           If Trim(rstPiece.Fields("Escompte")) <> "" Then
460             sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
            
465             Do While CDbl(sEscompte) > 1
470               sEscompte = CDbl(sEscompte) / 100
475             Loop

480             rstBCPiece.Fields("Total") = Conversion(CStr((Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * (1 - CDbl(sEscompte))) * CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString))), MODE_DECIMAL)
485           Else
490             rstBCPiece.Fields("Total") = Conversion(CStr(Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * Replace(rstPiece.Fields("Qté"), "-", vbNullString)), MODE_DECIMAL)
495           End If
500         Else
505           rstBCPiece.Fields("Qté") = CDbl(rstBCPiece.Fields("Qté")) + CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString))

510           rstBCPiece.Fields("NuméroLigne") = rstBCPiece.Fields("NuméroLigne") & ", " & rstPiece.Fields("NuméroLigne")

515           If Trim(rstPiece.Fields("Escompte")) <> "" Then
520             sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
            
525             Do While CDbl(sEscompte) > 1
530               sEscompte = CDbl(sEscompte) / 100
535             Loop

540             rstBCPiece.Fields("Total") = Replace(rstBCPiece.Fields("Total"), ".", ",") + (CDbl(Replace(rstPiece.Fields("PrixOrigine"), ".", ",")) * (1 - CDbl(sEscompte)) * CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString)))
545           Else
550             rstBCPiece.Fields("Total") = Replace(rstBCPiece.Fields("Total"), ".", ",") + (CDbl(Replace(rstPiece.Fields("PrixOrigine"), ".", ",")) * CDbl(Replace(rstPiece.Fields("Qté"), "-", vbNullString)))
555           End If
560         End If
     
565         Call rstBCPiece.Update
      
570         Call rstBCPiece.Close
      
575         Call rstPiece.MoveNext
580       Loop

585       Call rstPiece.Close
    
590       rstBC.Fields("Total") = Conversion(CStr(dblTotal), MODE_DECIMAL)
    
595       Call rstBC.Update
    
600       Call rstFRS.MoveNext
605     Loop

610     Call rstFRS.Close
615     Set rstFRS = Nothing
    
620     Call rstBC.Close
625     Set rstBC = Nothing

630     Set rstPiece = Nothing
635     Set rstBCPiece = Nothing

640     Exit Sub

AfficherErreur:

645     woups "frmBonCommande", "EnregistrerBonCommandeRetourMarchandiseAchat", Err, Erl
End Sub

Private Sub EnregistrerBonCommandeProjet()

5       On Error GoTo AfficherErreur

10      Dim rstBC         As ADODB.Recordset
15      Dim rstBCPiece    As ADODB.Recordset
20      Dim rstPiece      As ADODB.Recordset
25      Dim rstFRS        As ADODB.Recordset
30      Dim iCompteur     As Integer
35      Dim dblTotal      As Double
40      Dim sType         As String
45      Dim sWhere        As String
50      Dim sWherePiece   As String
55      Dim sWhereNoLigne As String
60      Dim sEscompte     As String
     
65      If m_eForm = I_PROJET_ELEC Then
70        sType = "E"
75      Else
80        sType = "M"
85      End If
     
        'Recordset source
90      sWhere = "(IDProjet = '" & m_sNoProjet & "' AND Type = '" & sType & "')"

95      sWherePiece = "GRB_Projet_Pieces.NumItem In ("
100     sWhereNoLigne = "GRB_Projet_Pieces.NuméroLigne In ("
        
105     For iCompteur = 1 To m_collPieces.count
110       If iCompteur <> m_collPieces.count Then
115         sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "', "
120         sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ", "
125       Else
130         sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "')"
135         sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ")"
140       End If
145     Next

150     Set rstBC = New ADODB.Recordset
155     Set rstFRS = New ADODB.Recordset
160     Set rstPiece = New ADODB.Recordset
165     Set rstBCPiece = New ADODB.Recordset

170     sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne
  
175     Call rstFRS.Open("SELECT DISTINCT GRB_Projet_Pieces.IDFRS, GRB_Fournisseur.CondTransport FROM GRB_Projet_Pieces INNER JOIN GRB_Fournisseur ON GRB_Projet_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
    
        'Recordsets destinations
180     Call rstBC.Open("SELECT * FROM GRB_BonsCommandes", g_connData, adOpenDynamic, adLockOptimistic)
  
185     Do While Not rstFRS.EOF
190       Call rstBC.AddNew
    
          'Enregistrement du bon
195       rstBC.Fields("NoBonCommande") = txtNoBC.Text
200       rstBC.Fields("NoFournisseur") = rstFRS.Fields("IDFRS")
205       rstBC.Fields("NoProjet") = m_sNoProjet
210       rstBC.Fields("Attention") = ""
    
215       If Not IsNull(rstFRS.Fields("CondTransport")) And rstFRS.Fields("CondTransport") <> vbNullString Then
220         rstBC.Fields("Transport") = rstFRS.Fields("CondTransport")
225       Else
230         rstBC.Fields("Transport") = "Votre camion"
235       End If
    
240       If m_eForm = I_PROJET_ELEC Then
245         rstBC.Fields("DateRequise") = ""
250       Else
255         rstBC.Fields("DateRequise") = ConvertDate(Date)
260       End If

265       rstBC.Fields("DateCommande") = ConvertDate(Date)
270       rstBC.Fields("CommandePar") = g_sEmploye

275       If m_eLangage = FRANCAIS Then
280         rstBC.Fields("LangueImpression") = "Français"
285       Else
290         rstBC.Fields("LangueImpression") = "Anglais"
295       End If
    
300       rstBC.Fields("Type") = sType

305       sWhere = "(IDProjet = '" & m_sNoProjet & "' AND IDFRS = " & rstFRS.Fields("IDFRS") & ")"
           
310       sWherePiece = "NumItem In ("
315       sWhereNoLigne = "NuméroLigne In ("
          
320       For iCompteur = 1 To m_collPieces.count
325         If iCompteur <> m_collPieces.count Then
330           sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "', "
335           sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ", "
340         Else
345           sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "')"
350           sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ")"
355         End If
360       Next

365       sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne
           
370       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
                   
375       dblTotal = 0
           
          'Enregistrement des pièces
380       Do While Not rstPiece.EOF
385         Call rstBCPiece.Open("SELECT * FROM GRB_BonsCommandes_Pieces WHERE NoItem = '" & Replace(rstPiece.Fields("NumItem"), "'", "''") & "' AND NoFournisseur = " & rstPiece.Fields("IDFRS") & " AND NoBonCommande = '" & txtNoBC.Text & "' AND Prix = '" & rstPiece.Fields("PrixOrigine") & "'", g_connData, adOpenDynamic, adLockOptimistic)
            
390         If Trim(rstPiece.Fields("ESCOMPTE")) <> "" Then
395           sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
            
400           Do While CDbl(sEscompte) > 1
405             sEscompte = CDbl(sEscompte) / 100
410           Loop

415           dblTotal = dblTotal + CDbl(Conversion(CStr((Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * (1 - CDbl(sEscompte))) * CDbl(rstPiece.Fields("Qté"))), MODE_DECIMAL))
420         Else
425           dblTotal = dblTotal + CDbl(Conversion(CStr(Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * rstPiece.Fields("Qté")), MODE_DECIMAL))
430         End If

            'Si la pièce n'existe pas, on l'ajoute
            'sinon, on change la quantité et le total
435         If rstBCPiece.EOF Then
440           Call rstBCPiece.AddNew
                
445           rstBCPiece.Fields("Type") = sType
             
450           rstBCPiece.Fields("NoBonCommande") = txtNoBC.Text
455           rstBCPiece.Fields("NoFournisseur") = rstPiece.Fields("IDFRS")
460           rstBCPiece.Fields("Qté") = rstPiece.Fields("Qté")
        
465           rstBCPiece.Fields("NoItem") = rstPiece.Fields("NumItem")

470           rstBCPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne")
        
475           If rstBC.Fields("LangueImpression") = "Français" Then
480             rstBCPiece.Fields("Description") = rstPiece.Fields("DESC_FR")
485           Else
490             rstBCPiece.Fields("Description") = rstPiece.Fields("DESC_EN")
495           End If

500           rstBCPiece.Fields("Manufact") = rstPiece.Fields("Manufact")
        
505           rstBCPiece.Fields("Prix") = rstPiece.Fields("PrixOrigine")
        
510           If rstPiece.Fields("Escompte") <> vbNullString Then
515             rstBCPiece.Fields("Escompte") = rstPiece.Fields("Escompte")
520           Else
525             rstBCPiece.Fields("Escompte") = "0"
530           End If
      
535           If Trim(rstPiece.Fields("Escompte")) <> "" Then
540             sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
            
545             Do While CDbl(sEscompte) > 1
550               sEscompte = CDbl(sEscompte) / 100
555             Loop

560             rstBCPiece.Fields("Total") = Conversion(CStr((Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * (1 - CDbl(sEscompte))) * CDbl(rstPiece.Fields("Qté"))), MODE_DECIMAL)
565           Else
570             rstBCPiece.Fields("Total") = Conversion(CStr(Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * rstPiece.Fields("Qté")), MODE_DECIMAL)
575           End If
580         Else
585           rstBCPiece.Fields("Qté") = CDbl(rstBCPiece.Fields("Qté")) + CDbl(rstPiece.Fields("Qté"))

590           rstBCPiece.Fields("NuméroLigne") = rstBCPiece.Fields("NuméroLigne") & ", " & rstPiece.Fields("NuméroLigne")

595           If Trim(rstPiece.Fields("Escompte")) <> "" Then
600             sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
            
605             Do While CDbl(sEscompte) > 1
610               sEscompte = CDbl(sEscompte) / 100
615             Loop

620             rstBCPiece.Fields("Total") = Replace(rstBCPiece.Fields("Total"), ".", ",") + (CDbl(Replace(rstPiece.Fields("PrixOrigine"), ".", ",")) * (1 - CDbl(sEscompte)) * CDbl(rstPiece.Fields("Qté")))
625           Else
630             rstBCPiece.Fields("Total") = Replace(rstBCPiece.Fields("Total"), ".", ",") + (CDbl(Replace(rstPiece.Fields("PrixOrigine"), ".", ",")) * CDbl(rstPiece.Fields("Qté")))
635           End If
640         End If
     
645         Call rstBCPiece.Update
      
650         Call rstBCPiece.Close
      
655         Call rstPiece.MoveNext
660       Loop

665       Call rstPiece.Close
    
670       rstBC.Fields("Total") = Conversion(CStr(dblTotal), MODE_DECIMAL)
    
675       Call rstBC.Update
    
680       Call rstFRS.MoveNext
685     Loop
  
690     Call rstFRS.Close
695     Set rstFRS = Nothing
  
700     Call rstBC.Close
705     Set rstBC = Nothing

710     Set rstBCPiece = Nothing
715     Set rstPiece = Nothing

720     Exit Sub

AfficherErreur:

725     woups "frmBonCommande", "EnregistrerBonCommandeProjet", Err, Erl
End Sub

Private Sub EnregistrerBonCommandeAchat()

5       On Error GoTo AfficherErreur

10      Dim rstBC         As ADODB.Recordset
15      Dim rstBCPiece    As ADODB.Recordset
20      Dim rstPiece      As ADODB.Recordset
25      Dim rstFRS        As ADODB.Recordset
30      Dim iCompteur     As Integer
35      Dim dblTotal      As Double
40      Dim sType         As String
45      Dim sWhere        As String
50      Dim sWherePiece   As String
55      Dim sWhereNoLigne As String
60      Dim sEscompte     As String
     
65      If m_eForm = I_ACHAT_ELEC Then
70        sType = "E"
75      Else
80        sType = "M"
85      End If
     
        'Recordset source
90      sWhere = "(IDAchat = '" & m_sNoAchat & "' AND IndexAchat = " & m_iIndexAchat & ")"
        
95      sWherePiece = "GRB_Achat_Pieces.PIECE In ("
100     sWhereNoLigne = "GRB_Achat_Pieces.NuméroLigne In ("
        
105     For iCompteur = 1 To m_collPieces.count
110       If iCompteur <> m_collPieces.count Then
115         sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "', "
120         sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ", "
125       Else
130         sWherePiece = sWherePiece & "'" & Replace(m_collPieces(iCompteur), "'", "''") & "')"
135         sWhereNoLigne = sWhereNoLigne & m_collNoLigne(iCompteur) & ")"
140       End If
145     Next

150     Set rstFRS = New ADODB.Recordset
155     Set rstBC = New ADODB.Recordset
160     Set rstPiece = New ADODB.Recordset
165     Set rstBCPiece = New ADODB.Recordset

170     sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne
       
        'Recordset source
175     Call rstFRS.Open("SELECT DISTINCT GRB_Achat_Pieces.IDFRS, GRB_Fournisseur.CondTransport FROM GRB_Achat_Pieces INNER JOIN GRB_Fournisseur ON GRB_Achat_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)

        'Recordsets destinations
180     Call rstBC.Open("SELECT * FROM GRB_BonsCommandes", g_connData, adOpenDynamic, adLockOptimistic)
  
185     Do While Not rstFRS.EOF
190       Call rstBC.AddNew
    
          'Enregistrement du bon
195       rstBC.Fields("NoBonCommande") = txtNoBC.Text
200       rstBC.Fields("NoFournisseur") = rstFRS.Fields("IDFRS")
205       rstBC.Fields("NoProjet") = m_sNoAchat
210       rstBC.Fields("Attention") = ""
    
215       If Not IsNull(rstFRS.Fields("CondTransport")) And rstFRS.Fields("CondTransport") <> vbNullString Then
220         rstBC.Fields("Transport") = rstFRS.Fields("CondTransport")
225       Else
230         rstBC.Fields("Transport") = "Votre camion"
235       End If

240       If m_eForm = I_ACHAT_ELEC Then
245         rstBC.Fields("DateRequise") = ""
250       Else
255         rstBC.Fields("DateRequise") = ConvertDate(Date)
260       End If

265       rstBC.Fields("DateCommande") = ConvertDate(Date)
270       rstBC.Fields("CommandePar") = g_sEmploye
275       rstBC.Fields("LangueImpression") = "Français"
    
280       rstBC.Fields("Type") = sType
            
285       Call rstPiece.Open("SELECT * FROM GRB_Achat_Pieces WHERE " & sWhere & " AND IDFRS = " & rstFRS.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
                    
290       dblTotal = 0
           
          'Enregistrement des pièces
295       Do While Not rstPiece.EOF
300         Call rstBCPiece.Open("SELECT * FROM GRB_BonsCommandes_Pieces WHERE NoItem = '" & Replace(rstPiece.Fields("PIECE"), "'", "''") & "' AND NoFournisseur = " & rstPiece.Fields("IDFRS") & " AND NoBonCommande = '" & txtNoBC.Text & "' AND Prix = '" & rstPiece.Fields("PrixOrigine") & "'", g_connData, adOpenDynamic, adLockOptimistic)
      
305         If Trim(rstPiece.Fields("ESCOMPTE")) <> "" Then
310           sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
            
315           Do While CDbl(sEscompte) > 1
320             sEscompte = CDbl(sEscompte) / 100
325           Loop

330           If Not IsNull(rstPiece.Fields("PrixOrigine")) Then
335             dblTotal = dblTotal + CDbl(Conversion(CStr((Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * (1 - CDbl(sEscompte))) * CDbl(rstPiece.Fields("Qté"))), MODE_DECIMAL))
340           End If
345         Else
350           If Not IsNull(rstPiece.Fields("PrixOrigine")) Then
355             dblTotal = dblTotal + CDbl(Conversion(CStr(Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * rstPiece.Fields("Qté")), MODE_DECIMAL))
360           End If
365         End If
          
            'Si la pièce n'existe pas, on l'ajoute
            'sinon, on change la quantité et le total
370         If rstBCPiece.EOF Then
375           Call rstBCPiece.AddNew
                 
380           rstBCPiece.Fields("Type") = sType
                 
385           rstBCPiece.Fields("NoBonCommande") = txtNoBC.Text
390           rstBCPiece.Fields("NoFournisseur") = rstPiece.Fields("IDFRS")
395           rstBCPiece.Fields("Qté") = rstPiece.Fields("Qté")

400           rstBCPiece.Fields("NoItem") = rstPiece.Fields("PIECE")

405           rstBCPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne")

410           rstBCPiece.Fields("Description") = rstPiece.Fields("Desc_fr")
415           rstBCPiece.Fields("Manufact") = rstPiece.Fields("Manufact")

420           rstBCPiece.Fields("Prix") = rstPiece.Fields("PrixOrigine")

425           If Not IsNull(rstPiece.Fields("Escompte")) And rstPiece.Fields("Escompte") <> vbNullString Then
430             rstBCPiece.Fields("Escompte") = rstPiece.Fields("Escompte")
435           Else
440             rstBCPiece.Fields("Escompte") = "0"
445           End If

450           If Trim(rstPiece.Fields("Escompte")) <> "" Then
455             sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
            
460             Do While CDbl(sEscompte) > 1
465               sEscompte = CDbl(sEscompte) / 100
470             Loop

475             If Not IsNull(rstPiece.Fields("PrixOrigine")) Then
480               rstBCPiece.Fields("Total") = Conversion(CStr((Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * (1 - CDbl(sEscompte))) * CDbl(rstPiece.Fields("Qté"))), MODE_DECIMAL)
485             End If
490           Else
495             If Not IsNull(rstPiece.Fields("PrixOrigine")) Then
500               rstBCPiece.Fields("Total") = Conversion(CStr(Replace(rstPiece.Fields("PrixOrigine"), ".", ",") * rstPiece.Fields("Qté")), MODE_DECIMAL)
505             End If
510           End If
515         Else
520           rstBCPiece.Fields("Qté") = CDbl(rstBCPiece.Fields("Qté")) + CDbl(rstPiece.Fields("Qté"))

525           rstBCPiece.Fields("NuméroLigne") = rstBCPiece.Fields("NuméroLigne") & ", " & rstPiece.Fields("NuméroLigne")

530           If Trim(rstPiece.Fields("Escompte")) <> "" Then
535             sEscompte = Replace(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), "%", "")
            
540             Do While CDbl(sEscompte) > 1
545               sEscompte = CDbl(sEscompte) / 100
550             Loop

555             rstBCPiece.Fields("Total") = Replace(rstBCPiece.Fields("Total"), ".", ",") + (CDbl(Replace(rstPiece.Fields("PrixOrigine"), ".", ",")) * (1 - CDbl(sEscompte)) * CDbl(rstPiece.Fields("Qté")))
560           Else
565             rstBCPiece.Fields("Total") = Replace(rstBCPiece.Fields("Total"), ".", ",") + (CDbl(Replace(rstPiece.Fields("PrixOrigine"), ".", ",")) * CDbl(rstPiece.Fields("Qté")))
570           End If
575         End If

580         Call rstBCPiece.Update

585         Call rstBCPiece.Close

590         Call rstPiece.MoveNext
595       Loop

600       rstBC.Fields("Total") = Conversion(CStr(dblTotal), MODE_DECIMAL)

605       Call rstBC.Update

610       Call rstFRS.MoveNext

615       Call rstPiece.Close
620     Loop
  
625     Call rstFRS.Close
630     Set rstFRS = Nothing

635     Call rstBC.Close
640     Set rstBC = Nothing

645     Set rstPiece = Nothing
650     Set rstBCPiece = Nothing

655     Exit Sub

AfficherErreur:

660     woups "frmBonCommande", "EnregistrerBonCommande", Err, Erl
End Sub

Private Sub EnregistrerModifFournisseur()

5       On Error GoTo AfficherErreur

10      Dim rstBC      As ADODB.Recordset
15      Dim rstBCPiece As ADODB.Recordset
20      Dim iCompteur  As Integer
25      Dim itmBC      As ListItem
30      Dim sNoBC      As String

35      If m_eForm = I_RETOUR_MARCHANDISE Then
40        sNoBC = txtVotreNoSoum.Text
45      Else
50        sNoBC = txtNoBC.Text
55      End If

60      Set rstBC = New ADODB.Recordset

65      Call rstBC.Open("SELECT * FROM GRB_BonsCommandes WHERE NoBonCommande = '" & sNoBC & "' AND NoFournisseur = " & m_iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)
  
        'Enregistre le bon de commande
70      rstBC.Fields("Attention") = cmbContact.Text
75      rstBC.Fields("Transport") = txtTransport.Text
80      rstBC.Fields("DateRequise") = txtDateRequise.Text

85      If m_eForm = I_RETOUR_MARCHANDISE Then
90        rstBC.Fields("VotreNoSoum") = txtNoBC.Text
95      Else
100       rstBC.Fields("VotreNoSoum") = txtVotreNoSoum.Text
105     End If

110     rstBC.Fields("Commentaire") = txtcommentaire.Text
115     rstBC.Fields("Total") = Conversion(txtTotal.Text, MODE_PAS_FORMAT)
120     rstBC.Fields("AffichageInstructions") = chkAfficherInstructions.Value
  
125     If optImpression(I_IMP_FRANCAIS).Value = True Then
130       rstBC.Fields("LangueImpression") = "Français"
135     Else
140       rstBC.Fields("LangueImpression") = "Anglais"
145     End If
  
150     Call rstBC.Update
  
155     Call rstBC.Close
160     Set rstBC = Nothing
        
165     Set rstBCPiece = New ADODB.Recordset
        
170     If m_eForm <> I_PROJET_ELEC And m_eForm <> I_PROJET_MEC Then
175       For iCompteur = 1 To lvwBonCommande.ListItems.count
180         Set itmBC = lvwBonCommande.ListItems(iCompteur)
    
            'Enregistre les pièces
185         Call rstBCPiece.Open("SELECT * FROM GRB_BonsCommandes_Pieces WHERE NoBonCommande = '" & sNoBC & "' AND NoFournisseur = " & m_iNoFRS & " AND NoItem = '" & Replace(itmBC.SubItems(I_COL_NO_ITEM), "'", "''") & "' AND NuméroLigne = '" & itmBC.Tag & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
190         If Not rstBCPiece.EOF Then
195           rstBCPiece.Fields("Qté") = itmBC.Text
200           rstBCPiece.Fields("Total") = itmBC.SubItems(I_COL_TOTAL)
    
205           Call rstBCPiece.Update
210         Else
215           Call MsgBox("Impossible d'enregistrer le bon de commande!", vbOKOnly, "Erreur")
220         End If
    
225         Call rstBCPiece.Close
230       Next

235       Set rstBCPiece = Nothing
240     End If

245     Exit Sub

AfficherErreur:

250     woups "frmBonCommande", "EnregistrerModifFournisseur", Err, Erl
End Sub

Private Sub txtDateRequise_KeyPress(KeyAscii As Integer)

5       On Error GoTo AfficherErreur

10      If KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
15        KeyAscii = 0
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmBonCommande", "txtDateRequise_KeyPress", Err, Erl
End Sub
