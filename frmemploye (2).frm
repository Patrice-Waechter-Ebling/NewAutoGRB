VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmemploye 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employés"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   Icon            =   "frmemploye.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   5850
   Begin VB.ComboBox cmbFamille 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmemploye.frx":0442
      Left            =   1800
      List            =   "frmemploye.frx":0444
      Locked          =   -1  'True
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox txtFamille 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame fraEmploye 
      BackColor       =   &H00000000&
      Caption         =   "Ajout d'employé"
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
      Height          =   1575
      Left            =   1560
      TabIndex        =   32
      Top             =   3360
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton cmdAnnulEmploye 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   2760
         TabIndex        =   17
         Top             =   1080
         Width           =   975
      End
      Begin VB.ComboBox cmbAjoutEmploye 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Employé"
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
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSupprimePunch 
      Caption         =   "Supprimer"
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdAjoutPunch 
      Caption         =   "Ajouter"
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CheckBox chkActif 
      BackColor       =   &H00000000&
      Caption         =   "Actif"
      Enabled         =   0   'False
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
      Left            =   3480
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox cmbEmployePunch 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4080
      Width           =   2415
   End
   Begin MSMask.MaskEdBox mskPagette 
      Height          =   285
      Left            =   1800
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdConfig 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Configuration"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ComboBox cmbGroupe 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmemploye.frx":0446
      Left            =   1800
      List            =   "frmemploye.frx":0448
      Locked          =   -1  'True
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txtpage 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdmodifier 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modifier"
      Height          =   495
      Left            =   1560
      TabIndex        =   24
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdannuler 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Annuler"
      Height          =   495
      Left            =   1560
      TabIndex        =   23
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtinitiale 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdFermer 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fermer"
      Height          =   495
      Left            =   4440
      TabIndex        =   26
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtconfirme 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdsupprimer 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Supprimer"
      Height          =   495
      Left            =   3000
      TabIndex        =   25
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdajouter 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ajouter"
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtpasswd 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox txtuserid 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3000
      Width           =   2415
   End
   Begin VB.ComboBox cmbEmploye 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Text            =   "cmbEmploye"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton cmdenregistré 
      Caption         =   "Enregistrer"
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtemployé 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtGroupe 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSMask.MaskEdBox mskTelephone 
      Height          =   285
      Left            =   1800
      TabIndex        =   19
      Top             =   4920
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskCellulaire 
      Height          =   285
      Left            =   1800
      TabIndex        =   20
      Top             =   5280
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtCell 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox txtTel 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Famille"
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
      Left            =   600
      TabIndex        =   42
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblPunchPour 
      BackStyle       =   0  'Transparent
      Caption         =   "Puncher pour :"
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
      Left            =   600
      TabIndex        =   35
      Top             =   4140
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Groupe"
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
      Left            =   600
      TabIndex        =   29
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Cellulaire"
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
      Left            =   600
      TabIndex        =   40
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Téléphone"
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
      Left            =   600
      TabIndex        =   38
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pagette"
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
      Left            =   600
      TabIndex        =   36
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Initiale"
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
      Left            =   600
      TabIndex        =   28
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblconfirme 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmation"
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
      Left            =   600
      TabIndex        =   34
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Passwd"
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
      Left            =   600
      TabIndex        =   31
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "User id"
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
      Left            =   600
      TabIndex        =   30
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employé"
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
      Left            =   600
      TabIndex        =   27
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "frmemploye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enumMode
 MODE_AJOUT = 0
 MODE_MODIF = 1
 MODE_MODIF_NON_AUTORISE = 2
 MODE_INACTIF = 3
End Enum

Private m_bModeAjouter As Boolean 'Mode ajouter ou non(modifié)
Private m_iNoEmploye As Integer
Private m_bModifEmploye As Boolean

Private Sub cmbemploye_Click()

 On Error GoTo Oups
 
 ''''''''''''''''''''''''''''
 'affiche employé sélectioné'
 ''''''''''''''''''''''''''''
 
 'Set du numéro d'employé
 m_iNoEmploye = cmbEmploye.ItemData(cmbEmploye.ListIndex)
 
 'Affiche l'employé sélectionné
 Call AfficherEmploye
 
 'Si l'employé sélectionné est le même que celui qui s'est loggé
 'la modification est permise
 If UCase(txtuserid.Text) = UCase(g_sUserID) Then
 cmdModifier.Enabled = True
 Else
 'Sinon, elle ne l'est pas, il faut ré-activer les boutons selon le groupe
 Call ActiverBoutonsGroupe
 End If
 
 txtemployé.Text = cmbEmploye.Text

 Exit Sub

Oups:

 wOups "frmemploye", "cmbemploye_Click", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboGroupe()

 On Error GoTo Oups

 'Rempli le combo des groupes de sécurité
 Dim rstGroupe As ADODB.Recordset
 Dim iCompteur As Integer
 
 'Il faut vider le groupe avant de le remplir
 Call cmbGroupe.Clear
 
 'Ouverture de la table GrbGroupes
 Set rstGroupe = New ADODB.Recordset
 
 Call rstGroupe.Open("SELECT * FROM GrbGroupes", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstGroupe.EOF
 'Ajout du nom du groupe dans le combo
 Call cmbGroupe.AddItem(rstGroupe.Fields("NomGroupe"))
 
 'Ajout du numéro de groupe dans l'ItemData du combo
 cmbGroupe.ItemData(cmbGroupe.newIndex) = rstGroupe.Fields("IDGroupe")
 
 'Si c'est en mode ajout
 If m_bModeAjouter = True Then
 'On sélectionne le groupe "Par défaut".. par défaut
 If cmbGroupe.LIST(cmbGroupe.newIndex) = S_GROUPE_DEFAUT Then
  cmbGroupe.ListIndex = cmbGroupe.newIndex
  End If
  End If
 
  Call rstGroupe.MoveNext
  Loop
 
  Call rstGroupe.Close
  Set rstGroupe = Nothing

  Exit Sub

Oups:

10 wOups "frmemploye", "RemplirComboGroupe", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboFamille()

 On Error GoTo Oups

 'Rempli le combo des Familles de sécurité
 Dim rstFamille As ADODB.Recordset
 Dim iCompteur As Integer
 
 'Il faut vider le Famille avant de le remplir
 Call cmbFamille.Clear
 
 'Ouverture de la table GrbFamilles
 Set rstFamille = New ADODB.Recordset
 
 Call rstFamille.Open("SELECT * FROM GrbFamille ORDER BY Famille", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstFamille.EOF
 'Ajout du nom du Famille dans le combo
 Call cmbFamille.AddItem(rstFamille.Fields("Famille"))
 
 'Ajout du numéro de Famille dans l'ItemData du combo
 cmbFamille.ItemData(cmbFamille.newIndex) = rstFamille.Fields("IDFamille")
 
 Call rstFamille.MoveNext
 Loop
 
  Call rstFamille.Close
  Set rstFamille = Nothing

  Exit Sub

Oups:

  wOups "frmemploye", "RemplirComboFamille", Err, Err.number, Err.Description
End Sub

Private Sub Cmdajouter_Click()

 On Error GoTo Oups
 
 '''''''''''''''''''''
 'met en mode ajouter
 '''''''''''''''''''''
 m_bModeAjouter = True
 
 Call MontrerControles(MODE_AJOUT)
 
 Call RemplirComboGroupe
 
 Call LockedChamps(MODE_AJOUT)
 
 Call ViderChamps
 
 Call AfficherMasque

 txtemployé.SetFocus

 Exit Sub

Oups:

 wOups "frmemploye", "cmdAjouter_Click", Err, Err.number, Err.Description
End Sub

Private Sub AfficherMasque()

 On Error GoTo Oups

 mskPagette.Text = txtPage.Text
 mskTelephone.Text = txtTel.Text
 mskCellulaire.Text = txtCell.Text
 
 txtPage.Visible = False
 mskPagette.Visible = True
 
 txtTel.Visible = False
 mskTelephone.Visible = True
 
 txtCell.Visible = False
 mskCellulaire.Visible = True

 Exit Sub

Oups:

  wOups "frmemploye", "AfficherMasque", Err, Err.number, Err.Description
End Sub

Private Sub CacherMasque()

 On Error GoTo Oups

 txtPage.Text = mskPagette.Text
 txtTel.Text = mskTelephone.Text
 txtCell.Text = mskCellulaire.Text
 
 txtPage.Visible = True
 mskPagette.Visible = False
 
 txtTel.Visible = True
 mskTelephone.Visible = False
 
 txtCell.Visible = True
 mskCellulaire.Visible = False

 Exit Sub

Oups:

  wOups "frmemploye", "CacherMasque", Err, Err.number, Err.Description
End Sub

Private Sub ViderChamps()

 On Error GoTo Oups
 
 'Vide les champs
 txtemployé.Text = vbNullString
 txtuserid.Text = vbNullString
 txtpasswd.Text = vbNullString
 txtconfirme.Text = vbNullString
 txtinitiale.Text = vbNullString
 txtCell.Text = vbNullString
 txtPage.Text = vbNullString
 txtTel.Text = vbNullString

 cmbFamille.ListIndex = -1

 Exit Sub

Oups:

  wOups "frmemploye", "ViderChamps", Err, Err.number, Err.Description
End Sub

Private Sub LockedChamps(ByVal eMode As enumMode)

 On Error GoTo Oups
 
 'On barre les champs par rapport à un mode
 Dim bIDPassTel As Boolean 'Contient le userID, le password ainsi que Cell, Tel et pagette
 Dim bNomGroupe As Boolean 'Contient le nom et le groupe
 Dim bFamille As Boolean
 Dim bInitiales As Boolean 'Contient les initiales
 Dim bPunch As Boolean
 Dim bChkPunch As Boolean
 Dim bChkActif As Boolean
 
 Select Case eMode
 Case MODE_MODIF:
 bInitiales = True
 bPunch = True
 
  Case MODE_AJOUT:
  bPunch = True
 
  Case MODE_INACTIF:
  bIDPassTel = True
  bNomGroupe = True
  bFamille = True
  bInitiales = True
  bPunch = True
 bChkPunch = True
bChkActif = True
 
 'Si le user a le droit de modifié ses infos seulement
 Case MODE_MODIF_NON_AUTORISE:
 bNomGroupe = True
 bInitiales = True
 bPunch = True
 bChkPunch = False
 bChkActif = False
End Select

txtCell.Locked = bIDPassTel
txtinitiale.Locked = bInitiales
txtPage.Locked = bIDPassTel
1  txtpasswd.Locked = bIDPassTel
txtTel.Locked = bIDPassTel
 txtuserid.Locked = bIDPassTel
cmbGroupe.Locked = bNomGroupe
 cmbFamille.Locked = bFamille
txtemployé.Locked = bNomGroupe
 txtGroupe.Locked = bNomGroupe
1  chkActif.Enabled = Not bChkActif

 Exit Sub

Oups:

 wOups "frmemploye", "LockedChamps", Err, Err.number, Err.Description
End Sub

Private Sub MontrerControles(ByVal eMode As enumMode)

 On Error GoTo Oups
 
 'Met les controles visible/invisible
 Dim bCmbGroupe As Boolean
 Dim bCmbFamille As Boolean
 Dim bCmbEmploye As Boolean
 Dim bTxtGroupe As Boolean
 Dim bTxtFamille As Boolean
 Dim bTxtEmploye As Boolean
 Dim bAjouter As Boolean
 Dim bModifier As Boolean
 Dim bSupprimer As Boolean
 Dim bEnregistrer As Boolean
  Dim bAnnuler As Boolean
  Dim bQuitter As Boolean
  Dim bGroupe As Boolean
  Dim bConfirmPwd As Boolean
  Dim bPunchPour As Boolean
 
  Select Case eMode
 Case MODE_AJOUT, MODE_MODIF:
  bTxtEmploye = True
  bCmbGroupe = True
 bCmbFamille = True
bEnregistrer = True
 bAnnuler = True
 bConfirmPwd = True
 
 Case MODE_MODIF_NON_AUTORISE:
 bTxtGroupe = True
 bTxtFamille = True
 bTxtEmploye = True
 bEnregistrer = True
 bAnnuler = True
 bConfirmPwd = True
 bPunchPour = True
 
 Case MODE_INACTIF:
 bCmbEmploye = True
 bTxtGroupe = True
 bTxtFamille = True
 bAjouter = True
 bModifier = True
 bSupprimer = True
 bQuitter = True
 bGroupe = True
1  bPunchPour = True
 End Select
 
 txtemployé.Visible = bTxtEmploye
cmbEmploye.Visible = bCmbEmploye

txtGroupe.Visible = bTxtGroupe
cmbGroupe.Visible = bCmbGroupe

txtFamille.Visible = bTxtFamille
cmbFamille.Visible = bCmbFamille

Cmdajouter.Visible = bAjouter
cmdModifier.Visible = bModifier
cmdsupprimer.Visible = bSupprimer
Cmdfermer.Visible = bQuitter
cmdenregistré.Visible = bEnregistrer
2  cmdAnnuler.Visible = bAnnuler
txtconfirme.Visible = bConfirmPwd

2  cmdConfig.Enabled = bGroupe

lblPunchPour.Visible = bPunchPour
2  cmbEmployePunch.Visible = bPunchPour
cmdAjoutPunch.Visible = bPunchPour
2  cmdSupprimePunch.Visible = bPunchPour

Exit Sub

Oups:

30 wOups "frmemploye", "MontrerControles", Err, Err.number, Err.Description
End Sub

Private Sub cmdAjoutPunch_Click()

 On Error GoTo Oups

 Call RemplirComboEmployeActif
 
 fraEmploye.Visible = True

 Exit Sub

Oups:

 wOups "frmemploye", "cmdAjoutPunch_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulEmploye_Click()

 On Error GoTo Oups

 fraEmploye.Visible = False

 Exit Sub

Oups:

 wOups "frmemploye", "cmdAnnulEmploye_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 Call MontrerControles(MODE_INACTIF)
 
 Call LockedChamps(MODE_INACTIF)
 
 Call CacherMasque
 
 txtemployé.Text = cmbEmploye.Text
 
 Call AfficherEmploye
 
 Call ActiverBoutonsGroupe

 Exit Sub

Oups:

 wOups "frmemploye", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdConfig_Click()

 On Error GoTo Oups

 'Affiche le form pour la configuration des groupes
 Call OuvrirForm(frmGroupes, True)
 
 Call AfficherEmploye

 Exit Sub

Oups:

 wOups "frmemploye", "cmdConfig_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdenregistré_Click()

 On Error GoTo Oups
 
 ''''''''''''''''''''''''''''''
 ' enregistre nouveau employé '
 ''''''''''''''''''''''''''''''
 Dim rstEmploye As ADODB.Recordset
 Dim rstUserId As ADODB.Recordset
 Dim sEmploye As String
 Dim iCompteur As Integer
 
 sEmploye = txtemployé.Text
 
 'si le nom de l'employé, le User ID, le Password et la confirmation du password ne sont pas vide
 If Trim(txtemployé.Text) <> vbNullString And Trim(txtuserid.Text) <> vbNullString And Trim(txtpasswd.Text) <> vbNullString And Trim(txtconfirme.Text) <> vbNullString And Trim(txtinitiale.Text) <> vbNullString And cmbFamille.ListIndex <> -1 Then
 'Si le password et la confirmation sont pareils
 If txtpasswd.Text = txtconfirme.Text Then
 'Ouverture de la connection
 Screen.MousePointer = vbHourglass

 Set rstEmploye = New ADODB.Recordset

 'Si en mode ajouter
 If m_bModeAjouter = True Then
 'Si le nom de l'employé ne se trouve pas dans le combo
  If ComboContient(cmbEmploye, txtemployé.Text) = False Then
 'Ouverture du recordset sur la table GrbEmployé
  Call rstEmploye.Open("SELECT * FROM Grbemployés", g_connData, adOpenDynamic, adLockOptimistic)
 
 'tant que c'est pas la fin du recordset
  Do While Not rstEmploye.EOF
 'Si les initiales sont les meme que l'employé ajouté
  If rstEmploye.Fields("Initiale") = txtinitiale.Text Then
  Call MsgBox("Ces initiales existent déjà!")
 
  Screen.MousePointer = vbDefault
 
  Exit Sub
  End If
 
 'Si le user id existe déjà
 If UCase(rstEmploye.Fields("loginname")) = UCase(txtuserid.Text) Then
 Call MsgBox("User ID existant!")
 
 Screen.MousePointer = vbDefault

 Exit Sub
 End If
 
 Call rstEmploye.MoveNext
 Loop
 
 Call rstEmploye.AddNew
 Else
 Call MsgBox("Cet employé existe déjà!")
 
 Exit Sub
 End If
 Else
 'Si ce n'est pas un ajout
 
 'Ouverture du recordset sur la table GrbEmployés ou le no d'employe est = à la variable m_iNoEmploye
 Call rstEmploye.Open("SELECT * FROM Grbemployés WHERE noemploye = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)
 
 'Si le contenu de txtuserid est différent au contenu du champs loginname
 If txtuserid.Text <> rstEmploye.Fields("loginname") Then
 'Si le contenu de txtuserid est différent de g_sUserID
 If txtuserid.Text <> g_sUserID Then
 Set rstUserId = New ADODB.Recordset

 Call rstUserId.Open("SELECT * FROM Grbemployés", g_connData, adOpenDynamic, adLockOptimistic)
 
 Do While Not rstUserId.EOF
 'Si le loginname du recordest est égal au txtuserid
1  If UCase(rstUserId.Fields("loginname")) = UCase(txtuserid.Text) Then
 Call MsgBox("User ID existant!")

 Call rstUserId.Close
 Set rstUserId = Nothing

 Call rstEmploye.Close
 Set rstEmploye = Nothing
 
 Exit Sub
 End If
 
 Call rstUserId.MoveNext
 Loop

 Call rstUserId.Close
 Set rstUserId = Nothing
 End If
 End If
 End If
 
 rstEmploye.Fields("employe").Value = txtemployé.Text
 
 'Si l'employé fait des modif sur lui-même
 If g_sUserID = rstEmploye.Fields("loginname") Then
 g_sUserID = txtuserid.Text
 End If
 
 rstEmploye.Fields("loginname").Value = txtuserid.Text
 rstEmploye.Fields("passwd").Value = txtpasswd.Text
 rstEmploye.Fields("initiale").Value = txtinitiale.Text
rstEmploye.Fields("Actif").Value = chkActif.Value

 If m_bModeAjouter = False Then
 If chkActif.Value = vbUnchecked Then
 Call g_connData.Execute("DELETE * FROM GrbAutorisationPunch WHERE NoEmploye = " & m_iNoEmploye & " OR AutoriserPar = " & m_iNoEmploye)
 End If
 End If

 If mskTelephone.Text = vbNullString Then
 rstEmploye.Fields("tel").Value = " "
 Else
 rstEmploye.Fields("tel").Value = mskTelephone.Text
 End If
 
 If mskCellulaire.Text = vbNullString Then
 rstEmploye.Fields("cell").Value = " "
 Else
 rstEmploye.Fields("cell").Value = mskCellulaire.Text
 End If
 
 If mskPagette.Text = vbNullString Then
 rstEmploye.Fields("page").Value = " "
 Else
 rstEmploye.Fields("page").Value = mskPagette.Text
4 End If
 
 'Celà veut dire que l'utilisateur a le droit de modifier le groupe
4 If cmbGroupe.Visible = True Then
4 rstEmploye.Fields("groupe").Value = cmbGroupe.ItemData(cmbGroupe.ListIndex)
4 End If

4 If cmbFamille.Visible = True Then
4 If cmbFamille.ListIndex <> -1 Then
4 rstEmploye.Fields("Famille").Value = cmbFamille.ItemData(cmbFamille.ListIndex)
4 End If
4 End If
 
4 Call rstEmploye.Update
 
4 Call rstEmploye.Close
4  Set rstEmploye = Nothing
 
4  Screen.MousePointer = vbDefault
 
4  Call MontrerControles(MODE_INACTIF)
 
4  Call LockedChamps(MODE_INACTIF)
 
4  Call ActiverBoutonsGroupe
 
 'remplis combo
4  Call RemplirComboEmploye
 
4  Call RemplirComboEmployeActif
 
4  For iCompteur = 0 To cmbEmploye.ListCount
50 If cmbEmploye.LIST(iCompteur) = sEmploye Then
 cmbEmploye.ListIndex = iCompteur
 
 Exit For
 End If
 Next

 m_bModeAjouter = False
 Else
 Call MsgBox("Le mot de passe est incorrect!")
 End If
 
 Call CacherMasque
 Else
 Call MsgBox("Champs vide!")
5  End If

5  Exit Sub

Oups:

5  wOups "frmemploye", "cmdenregistré_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdModifier_Click()

 On Error GoTo Oups
 
 '''''''''''''''''''''
 'met en mode modifier
 '''''''''''''''''''''
 Call AfficherMasque
 
 'Si le user a le droit de modifier les autres user
 If m_bModifEmploye = True Then
 Call MontrerControles(MODE_MODIF)
 Call RemplirComboGroupe

 If txtGroupe.Text <> "" Then
 cmbGroupe.Text = txtGroupe.Text
 Else
 cmbGroupe.ListIndex = -1
 End If

 If txtFamille.Text <> "" Then
  cmbFamille.Text = txtFamille.Text
  Else
  cmbFamille.ListIndex = -1
  End If

  Call LockedChamps(MODE_MODIF)
  Else
  Call MontrerControles(MODE_MODIF_NON_AUTORISE)
  Call LockedChamps(MODE_MODIF_NON_AUTORISE)
10 End If
 
txtconfirme.Text = txtpasswd.Text
m_bModeAjouter = False

Exit Sub

Oups:

wOups "frmemploye", "cmdModifier_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmemploye", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdOK_Click()

 On Error GoTo Oups

 Dim rstAutorisation As ADODB.Recordset
 
 Set rstAutorisation = New ADODB.Recordset
 
 Call rstAutorisation.Open("SELECT * FROM GrbAutorisationPunch", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call rstAutorisation.AddNew
 
 rstAutorisation.Fields("NoEmploye") = cmbAjoutEmploye.ItemData(cmbAjoutEmploye.ListIndex)
 rstAutorisation.Fields("AutoriserPar") = cmbEmploye.ItemData(cmbEmploye.ListIndex)
 
 Call rstAutorisation.Update
 
 Call rstAutorisation.Close
 Set rstAutorisation = Nothing
 
 Call RemplirComboEmployePunch
 
  fraEmploye.Visible = False

  Exit Sub

Oups:

  wOups "frmemploye", "cmdOK_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdSupprimePunch_Click()

 On Error GoTo Oups

 Dim iEmploye As Integer
 Dim iPunch As Integer
 
 If cmbEmployePunch.ListCount > 0 Then
 iEmploye = cmbEmploye.ItemData(cmbEmploye.ListIndex)
 iPunch = cmbEmployePunch.ItemData(cmbEmployePunch.ListIndex)
 
 If cmbEmployePunch.ListIndex > -1 Then
 If MsgBox("Êtes vous sûr de vouloir supprimer cet employé?", vbYesNo) = vbYes Then
 Call g_connData.Execute("DELETE * FROM GrbAutorisationPunch WHERE NoEmploye = " & iPunch & " AND AutoriserPar = " & iEmploye)
 End If
 
 Call RemplirComboEmployePunch
  End If
  End If

  Exit Sub

Oups:

  wOups "frmemploye", "cmdSupprimePunch_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdsupprimer_Click()

 On Error GoTo Oups

 '''''''''''''''''''''''''''''''''''
 'supprime employé
 ''''''''''''''''''''''''''''''
 Dim rstProjSoum As ADODB.Recordset
 Dim rstFT As ADODB.Recordset
 Dim sTampon As String

 If cmbEmploye.ListCount > 0 Then
 'si on veut supprimer
 If MsgBox("Etes-vous sur de supprimer cet enregistrement?", vbYesNo, "Supprimer") = vbYes Then
 Set rstFT = New ADODB.Recordset
 
 Call rstFT.Open("SELECT * FROM GrbPunch WHERE NoEmploye = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstFT.EOF Then
 Set rstProjSoum = New ADODB.Recordset

 Call rstProjSoum.Open("SELECT * FROM GrbSoumission_Modif WHERE NoEmployé = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)
 
  If rstProjSoum.EOF Then
  Call rstProjSoum.Close
 
  Call rstProjSoum.Open("SELECT * FROM GrbProjet_Modif WHERE NoEmployé = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)
 
  If rstProjSoum.EOF Then
 'delete employe
  Call g_connData.Execute("DELETE * FROM Grbemployés WHERE noemploye = " & m_iNoEmploye)
  
  Call rstProjSoum.Close
  Set rstProjSoum = Nothing
 
  Call rstFT.Close
 Set rstFT = Nothing
 
 Call RemplirComboEmploye
 
 If cmbEmploye.ListCount > 0 Then
 cmbEmploye.ListIndex = 0
 End If
 Else
 Call MsgBox("Impossible d'effacer cet employé, il est utilisé dans le projet " & rstProjSoum.Fields("IDProjet") & "!", vbOKOnly, "Erreur")
 
 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 End If
 Else
 Call MsgBox("Impossible d'effacer cet employé, il est utilisé dans la soumission " & rstProjSoum.Fields("IDSoumission") & "!", vbOKOnly, "Erreur")
 
 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 End If
 Else
 Call MsgBox("Impossible d'effacer cet employé, il est utilisé dans les feuilles de temps pour le projet " & rstFT.Fields("NoProjet") & "!", vbOKOnly, "Erreur")
 
 Call rstFT.Close
 Set rstFT = Nothing
1  End If
 End If
 End If

Exit Sub

Oups:

wOups "frmemploye", "cmdSupprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups
 
 'remplis combo employe
 Call RemplirComboEmploye
 
 Call RemplirComboFamille
 
 Call MontrerControles(MODE_INACTIF)
 
 Call LockedChamps(MODE_INACTIF)
 
 Call ActiverBoutonsGroupe

 'selectionne dans combo employe
 If cmbEmploye.ListCount >= 0 Then
 cmbEmploye.ListIndex = 0
 End If

 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

  wOups "frmemploye", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub ActiverBoutonsGroupe()

 On Error GoTo Oups
 
 'Activation des boutons selon le groupe
 m_bModifEmploye = g_bModificationEmployes
 
 Cmdajouter.Enabled = m_bModifEmploye
 cmdModifier.Enabled = m_bModifEmploye
 cmdsupprimer.Enabled = m_bModifEmploye
 
 cmdConfig.Enabled = g_bModificationGroupes
 
 cmdAjoutPunch.Enabled = g_bModificationPunchEmployes
 cmdSupprimePunch.Enabled = g_bModificationPunchEmployes
 
 Exit Sub

Oups:

 wOups "frmemploye", "ActiverBoutonsGroupe", Err, Err.number, Err.Description
End Sub

Private Sub AfficherEmploye()

 On Error GoTo Oups
 
 ''''''''''''''''''''''''''''''''''''''''''
 ' affiche donne de l'employé selectionné '
 ''''''''''''''''''''''''''''''''''''''''''
 Dim rstEmploye As ADODB.Recordset
 Dim rstGroupe As ADODB.Recordset
 Dim rstFamille As ADODB.Recordset
 
 Set rstEmploye = New ADODB.Recordset
 
 Call rstEmploye.Open("SELECT * FROM Grbemployés WHERE NoEmploye = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)
 
 'REMPLIS LES CHAMPS
 If Not rstEmploye.EOF Then
 txtpasswd.Text = rstEmploye.Fields("passwd")
 txtuserid.Text = rstEmploye.Fields("loginname")
 txtinitiale.Text = rstEmploye.Fields("initiale")
 
 If IsNull(rstEmploye.Fields("groupe")) Then
  txtGroupe.Text = vbNullString
  Else
  Set rstGroupe = New ADODB.Recordset

  Call rstGroupe.Open("SELECT * FROM GrbGroupes WHERE IDGroupe = " & rstEmploye.Fields("Groupe"), g_connData, adOpenDynamic, adLockOptimistic)
 
  txtGroupe.Text = rstGroupe.Fields("NomGroupe")
 
  Call rstGroupe.Close
  Set rstGroupe = Nothing
  End If

If IsNull(rstEmploye.Fields("Famille")) Then
txtFamille.Text = vbNullString
 Else
 Set rstFamille = New ADODB.Recordset

 Call rstFamille.Open("SELECT * FROM GrbFamille WHERE IDFamille = " & rstEmploye.Fields("Famille"), g_connData, adOpenDynamic, adLockOptimistic)
 
 txtFamille.Text = rstFamille.Fields("Famille")
 
 Call rstFamille.Close
 Set rstFamille = Nothing
 End If
 
 If IsNull(rstEmploye.Fields("cell")) Then
 txtCell.Text = vbNullString
 Else
 txtCell.Text = Trim(rstEmploye.Fields("cell"))
 End If
 
 If IsNull(rstEmploye.Fields("Page")) Then
 txtPage.Text = vbNullString
 Else
 txtPage.Text = Trim(rstEmploye.Fields("Page"))
 End If
 
1  If IsNull(rstEmploye.Fields("tel")) Then
 txtTel.Text = vbNullString
 Else
 txtTel.Text = Trim(rstEmploye.Fields("tel"))
 End If
 
 chkActif.Value = Abs(CInt(rstEmploye.Fields("Actif")))
 
 Call RemplirComboEmployePunch
End If
 
Call rstEmploye.Close
Set rstEmploye = Nothing

Exit Sub

Oups:

wOups "frmemploye", "AfficherEmploye", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboEmployePunch()

 On Error GoTo Oups

 Dim rstEmployePunch As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset

 Call cmbEmployePunch.Clear

 Set rstEmployePunch = New ADODB.Recordset

 Call rstEmployePunch.Open("SELECT * FROM GrbAutorisationPunch WHERE AutoriserPar = " & cmbEmploye.ItemData(cmbEmploye.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
 
 Set rstEmploye = New ADODB.Recordset
 
 Do While Not rstEmployePunch.EOF
 Call rstEmploye.Open("SELECT Employe, NoEmploye FROM GrbEmployés WHERE NoEmploye = " & rstEmployePunch.Fields("NoEmploye"), g_connData, adOpenDynamic, adLockOptimistic)
 
 Call cmbEmployePunch.AddItem(rstEmploye.Fields("Employe"))
 
 cmbEmployePunch.ItemData(cmbEmployePunch.newIndex) = rstEmploye.Fields("NoEmploye")
 
  Call rstEmploye.Close
 
  Call rstEmployePunch.MoveNext
  Loop

  Set rstEmploye = Nothing
 
  If cmbEmployePunch.ListCount > 0 Then
  cmbEmployePunch.ListIndex = 0
  End If
 
  Call rstEmployePunch.Close
10 Set rstEmployePunch = Nothing

Exit Sub

Oups:

wOups "frmemploye", "RemplirComboEmployePunch", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboEmploye()

 On Error GoTo Oups
 
 '''''''''''''''''''''''''
 'rempli le combo employé'
 '''''''''''''''''''''''''
 Dim rstEmploye As ADODB.Recordset
 
 Set rstEmploye = New ADODB.Recordset
 
 Call rstEmploye.Open("SELECT * FROM Grbemployés ORDER BY employe", g_connData, adOpenDynamic, adLockOptimistic)

 Call cmbEmploye.Clear
 
 'remplis le combo employé
 Do While Not rstEmploye.EOF
 Call cmbEmploye.AddItem(rstEmploye.Fields("employe"))
 
 cmbEmploye.ItemData(cmbEmploye.newIndex) = rstEmploye.Fields("noEmploye")
 
 Call rstEmploye.MoveNext
 Loop
 
 Call rstEmploye.Close
  Set rstEmploye = Nothing

  Exit Sub

Oups:

  wOups "frmemploye", "RemplirComboEmploye", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboEmployeActif()

 On Error GoTo Oups

 Dim rstEmploye As ADODB.Recordset
 Dim iCompteur As Integer
 Dim iCompteur2 As Integer
 Dim bSupprimer As Boolean
 
 Set rstEmploye = New ADODB.Recordset
 
 Call rstEmploye.Open("SELECT * FROM Grbemployés WHERE Actif = true ORDER BY employe", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call cmbAjoutEmploye.Clear
 
 'rempli le combo employé
 Do While Not rstEmploye.EOF
 Call cmbAjoutEmploye.AddItem(rstEmploye.Fields("employe"))
 
 cmbAjoutEmploye.ItemData(cmbAjoutEmploye.newIndex) = rstEmploye.Fields("noEmploye")
 
  Call rstEmploye.MoveNext
  Loop
 
  Call rstEmploye.Close
  Set rstEmploye = Nothing
 
  iCompteur = 0
 
 'Il faut enlever les employés déjà dans le combo et l'employé en cours
  Do While iCompteur <= cmbAjoutEmploye.ListCount - 1
  bSupprimer = False
 
 'Si c'est l'employé en cours
  If cmbAjoutEmploye.LIST(iCompteur) = cmbEmploye.Text Then
 bSupprimer = True
1 Else
 iCompteur2 = 0
 
 'Si c'est les employés dans le combo
 Do While iCompteur2 <= cmbEmployePunch.ListCount - 1
 If cmbEmployePunch.LIST(iCompteur2) = cmbAjoutEmploye.LIST(iCompteur) Then
 bSupprimer = True
 End If
 
 iCompteur2 = iCompteur2 + 1
 Loop
 End If
 
 If bSupprimer = True Then
 Call cmbAjoutEmploye.RemoveItem(iCompteur)
Else
 iCompteur = iCompteur + 1
 End If
Loop
 
 If cmbAjoutEmploye.ListCount > 0 Then
 cmbAjoutEmploye.ListIndex = 0
 End If

1  Exit Sub

Oups:

 wOups "frmemploye", "RemplirComboEmployeActif", Err, Err.number, Err.Description
End Sub

Private Sub mskCellulaire_GotFocus()

 On Error GoTo Oups

 mskCellulaire.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmemploye", "mskCellulaire_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskPagette_GotFocus()

 On Error GoTo Oups

 mskPagette.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmemploye", "mskPagette_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskTelephone_GotFocus()

 On Error GoTo Oups

 mskTelephone.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmemploye", "mskTelephone_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskCellulaire_LostFocus()

 On Error GoTo Oups

 mskCellulaire.mask = vbNullString
 
 If mskCellulaire.Text = "(___) ___-____" Then
 mskCellulaire.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmemploye", "mskCellulaire_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskPagette_LostFocus()

 On Error GoTo Oups

 mskPagette.mask = vbNullString
 
 If mskPagette.Text = "(___) ___-____" Then
 mskPagette.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmemploye", "mskPagette_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskTelephone_LostFocus()

 On Error GoTo Oups

 mskTelephone.mask = vbNullString
 
 If mskTelephone.Text = "(___) ___-____" Then
 mskTelephone.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmemploye", "mskTelephone_LostFocus", Err, Err.number, Err.Description
End Sub
