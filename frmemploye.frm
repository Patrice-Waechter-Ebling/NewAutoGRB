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
   Picture         =   "frmemploye.frx":0442
   ScaleHeight     =   6390
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbFamille 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmemploye.frx":334F
      Left            =   1800
      List            =   "frmemploye.frx":3351
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
      ItemData        =   "frmemploye.frx":3353
      Left            =   1800
      List            =   "frmemploye.frx":3355
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

Private m_bModeAjouter  As Boolean  'Mode ajouter ou non(modifié)
Private m_iNoEmploye    As Integer
Private m_bModifEmploye As Boolean

Private Sub cmbemploye_Click()

5       On Error GoTo AfficherErreur
        
        ''''''''''''''''''''''''''''
        'affiche employé sélectioné'
        ''''''''''''''''''''''''''''
  
        'Set du numéro d'employé
10      m_iNoEmploye = cmbEmploye.ItemData(cmbEmploye.ListIndex)
  
        'Affiche l'employé sélectionné
15      Call AfficherEmploye
    
        'Si l'employé sélectionné est le même que celui qui s'est loggé
        'la modification est permise
20      If UCase(txtuserid.Text) = UCase(g_sUserID) Then
25        cmdModifier.Enabled = True
30      Else
          'Sinon, elle ne l'est pas, il faut ré-activer les boutons selon le groupe
35        Call ActiverBoutonsGroupe
40      End If
  
45      txtemployé.Text = cmbEmploye.Text

50      Exit Sub

AfficherErreur:

55      woups "frmemploye", "cmbemploye_Click", Err, Erl
End Sub

Private Sub RemplirComboGroupe()

5       On Error GoTo AfficherErreur

        'Rempli le combo des groupes de sécurité
10      Dim rstGroupe As ADODB.Recordset
15      Dim iCompteur As Integer
                
        'Il faut vider le groupe avant de le remplir
20      Call cmbGroupe.Clear
    
        'Ouverture de la table GRB_Groupes
25      Set rstGroupe = New ADODB.Recordset
        
30      Call rstGroupe.Open("SELECT * FROM GRB_Groupes", g_connData, adOpenDynamic, adLockOptimistic)
  
        'Tant que ce n'est pas la fin des enregistrements
35      Do While Not rstGroupe.EOF
          'Ajout du nom du groupe dans le combo
40        Call cmbGroupe.AddItem(rstGroupe.Fields("NomGroupe"))
      
          'Ajout du numéro de groupe dans l'ItemData du combo
45        cmbGroupe.ItemData(cmbGroupe.newIndex) = rstGroupe.Fields("IDGroupe")
      
          'Si c'est en mode ajout
50        If m_bModeAjouter = True Then
            'On sélectionne le groupe "Par défaut".. par défaut
55          If cmbGroupe.LIST(cmbGroupe.newIndex) = S_GROUPE_DEFAUT Then
60            cmbGroupe.ListIndex = cmbGroupe.newIndex
65          End If
70        End If
      
75        Call rstGroupe.MoveNext
80      Loop
  
85      Call rstGroupe.Close
90      Set rstGroupe = Nothing

95      Exit Sub

AfficherErreur:

100     woups "frmemploye", "RemplirComboGroupe", Err, Erl
End Sub

Private Sub RemplirComboFamille()

5       On Error GoTo AfficherErreur

        'Rempli le combo des Familles de sécurité
10      Dim rstFamille As ADODB.Recordset
15      Dim iCompteur  As Integer
                
        'Il faut vider le Famille avant de le remplir
20      Call cmbFamille.Clear
    
        'Ouverture de la table GRB_Familles
25      Set rstFamille = New ADODB.Recordset
        
30      Call rstFamille.Open("SELECT * FROM GRB_Famille ORDER BY Famille", g_connData, adOpenDynamic, adLockOptimistic)
  
        'Tant que ce n'est pas la fin des enregistrements
35      Do While Not rstFamille.EOF
          'Ajout du nom du Famille dans le combo
40        Call cmbFamille.AddItem(rstFamille.Fields("Famille"))
      
          'Ajout du numéro de Famille dans l'ItemData du combo
45        cmbFamille.ItemData(cmbFamille.newIndex) = rstFamille.Fields("IDFamille")
            
50        Call rstFamille.MoveNext
55      Loop
  
60      Call rstFamille.Close
65      Set rstFamille = Nothing

70      Exit Sub

AfficherErreur:

75      woups "frmemploye", "RemplirComboFamille", Err, Erl
End Sub

Private Sub Cmdajouter_Click()

5       On Error GoTo AfficherErreur
        
        '''''''''''''''''''''
        'met en mode ajouter
        '''''''''''''''''''''
10      m_bModeAjouter = True
    
15      Call MontrerControles(MODE_AJOUT)
  
20      Call RemplirComboGroupe
    
25      Call LockedChamps(MODE_AJOUT)
  
30      Call ViderChamps
  
35      Call AfficherMasque

40      txtemployé.SetFocus

45      Exit Sub

AfficherErreur:

50      woups "frmemploye", "cmdAjouter_Click", Err, Erl
End Sub

Private Sub AfficherMasque()

5      On Error GoTo AfficherErreur

10      mskPagette.Text = txtPage.Text
15      mskTelephone.Text = txtTel.Text
20      mskCellulaire.Text = txtCell.Text
  
25      txtPage.Visible = False
30      mskPagette.Visible = True
  
35      txtTel.Visible = False
40      mskTelephone.Visible = True
  
45      txtCell.Visible = False
50      mskCellulaire.Visible = True

55      Exit Sub

AfficherErreur:

60      woups "frmemploye", "AfficherMasque", Err, Erl
End Sub

Private Sub CacherMasque()

5       On Error GoTo AfficherErreur

10      txtPage.Text = mskPagette.Text
15      txtTel.Text = mskTelephone.Text
20      txtCell.Text = mskCellulaire.Text
  
25      txtPage.Visible = True
30      mskPagette.Visible = False
  
35      txtTel.Visible = True
40      mskTelephone.Visible = False
  
45      txtCell.Visible = True
50      mskCellulaire.Visible = False

55      Exit Sub

AfficherErreur:

60      woups "frmemploye", "CacherMasque", Err, Erl
End Sub

Private Sub ViderChamps()

5       On Error GoTo AfficherErreur
        
        'Vide les champs
10      txtemployé.Text = vbNullString
15      txtuserid.Text = vbNullString
20      txtpasswd.Text = vbNullString
25      txtconfirme.Text = vbNullString
30      txtinitiale.Text = vbNullString
35      txtCell.Text = vbNullString
40      txtPage.Text = vbNullString
45      txtTel.Text = vbNullString

50      cmbFamille.ListIndex = -1

55      Exit Sub

AfficherErreur:

60      woups "frmemploye", "ViderChamps", Err, Erl
End Sub

Private Sub LockedChamps(ByVal eMode As enumMode)

5       On Error GoTo AfficherErreur
        
        'On barre les champs par rapport à un mode
10      Dim bIDPassTel    As Boolean 'Contient le userID, le password ainsi que Cell, Tel et pagette
15      Dim bNomGroupe    As Boolean 'Contient le nom et le groupe
20      Dim bFamille      As Boolean
25      Dim bInitiales    As Boolean 'Contient les initiales
30      Dim bPunch        As Boolean
35      Dim bChkPunch     As Boolean
40      Dim bChkActif     As Boolean
  
45      Select Case eMode
          Case MODE_MODIF:
50          bInitiales = True
55          bPunch = True
    
60        Case MODE_AJOUT:
65          bPunch = True
            
70       Case MODE_INACTIF:
75          bIDPassTel = True
80          bNomGroupe = True
85          bFamille = True
90          bInitiales = True
95          bPunch = True
100         bChkPunch = True
105         bChkActif = True
    
          'Si le user a le droit de modifié ses infos seulement
110       Case MODE_MODIF_NON_AUTORISE:
115         bNomGroupe = True
120         bInitiales = True
125         bPunch = True
130         bChkPunch = False
135         bChkActif = False
140     End Select

145     txtCell.Locked = bIDPassTel
150     txtinitiale.Locked = bInitiales
155     txtPage.Locked = bIDPassTel
160     txtpasswd.Locked = bIDPassTel
165     txtTel.Locked = bIDPassTel
170     txtuserid.Locked = bIDPassTel
175     cmbGroupe.Locked = bNomGroupe
180     cmbFamille.Locked = bFamille
185     txtemployé.Locked = bNomGroupe
190     txtGroupe.Locked = bNomGroupe
195     chkActif.Enabled = Not bChkActif

200     Exit Sub

AfficherErreur:

205     woups "frmemploye", "LockedChamps", Err, Erl
End Sub

Private Sub MontrerControles(ByVal eMode As enumMode)

5       On Error GoTo AfficherErreur
        
        'Met les controles visible/invisible
10      Dim bCmbGroupe    As Boolean
15      Dim bCmbFamille   As Boolean
20      Dim bCmbEmploye   As Boolean
25      Dim bTxtGroupe    As Boolean
30      Dim bTxtFamille   As Boolean
35      Dim bTxtEmploye   As Boolean
40      Dim bAjouter      As Boolean
45      Dim bModifier     As Boolean
50      Dim bSupprimer    As Boolean
55      Dim bEnregistrer  As Boolean
60      Dim bAnnuler      As Boolean
65      Dim bQuitter      As Boolean
70      Dim bGroupe       As Boolean
75      Dim bConfirmPwd   As Boolean
80      Dim bPunchPour    As Boolean
    
85      Select Case eMode
          Case MODE_AJOUT, MODE_MODIF:
90          bTxtEmploye = True
95          bCmbGroupe = True
100         bCmbFamille = True
105         bEnregistrer = True
110         bAnnuler = True
115         bConfirmPwd = True
          
          Case MODE_MODIF_NON_AUTORISE:
120         bTxtGroupe = True
125         bTxtFamille = True
130         bTxtEmploye = True
135         bEnregistrer = True
140         bAnnuler = True
145         bConfirmPwd = True
150         bPunchPour = True
    
          Case MODE_INACTIF:
155         bCmbEmploye = True
160         bTxtGroupe = True
165         bTxtFamille = True
170         bAjouter = True
175         bModifier = True
180         bSupprimer = True
185         bQuitter = True
190         bGroupe = True
195         bPunchPour = True
200     End Select
  
205     txtemployé.Visible = bTxtEmploye
210     cmbEmploye.Visible = bCmbEmploye

215     txtGroupe.Visible = bTxtGroupe
220     cmbGroupe.Visible = bCmbGroupe

225     txtFamille.Visible = bTxtFamille
230     cmbFamille.Visible = bCmbFamille

235     Cmdajouter.Visible = bAjouter
240     cmdModifier.Visible = bModifier
245     cmdsupprimer.Visible = bSupprimer
250     Cmdfermer.Visible = bQuitter
255     cmdenregistré.Visible = bEnregistrer
260     cmdAnnuler.Visible = bAnnuler
265     txtconfirme.Visible = bConfirmPwd

270     cmdConfig.Enabled = bGroupe

275     lblPunchPour.Visible = bPunchPour
280     cmbEmployePunch.Visible = bPunchPour
285     cmdAjoutPunch.Visible = bPunchPour
290     cmdSupprimePunch.Visible = bPunchPour

295     Exit Sub

AfficherErreur:

300     woups "frmemploye", "MontrerControles", Err, Erl
End Sub

Private Sub cmdAjoutPunch_Click()

5       On Error GoTo AfficherErreur

10      Call RemplirComboEmployeActif
  
15      fraEmploye.Visible = True

20      Exit Sub

AfficherErreur:

25      woups "frmemploye", "cmdAjoutPunch_Click", Err, Erl
End Sub

Private Sub cmdAnnulEmploye_Click()

5       On Error GoTo AfficherErreur

10      fraEmploye.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmemploye", "cmdAnnulEmploye_Click", Err, Erl
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      Call MontrerControles(MODE_INACTIF)
  
15      Call LockedChamps(MODE_INACTIF)
                    
20      Call CacherMasque
  
25      txtemployé.Text = cmbEmploye.Text
  
30      Call AfficherEmploye
  
35      Call ActiverBoutonsGroupe

40      Exit Sub

AfficherErreur:

45      woups "frmemploye", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdConfig_Click()

5       On Error GoTo AfficherErreur

        'Affiche le form pour la configuration des groupes
10      Call OuvrirForm(frmGroupes, True)
    
15      Call AfficherEmploye

20      Exit Sub

AfficherErreur:

25      woups "frmemploye", "cmdConfig_Click", Err, Erl
End Sub

Private Sub cmdenregistré_Click()

5       On Error GoTo AfficherErreur
        
        ''''''''''''''''''''''''''''''
        ' enregistre nouveau employé '
        ''''''''''''''''''''''''''''''
10      Dim rstEmploye As ADODB.Recordset
15      Dim rstUserId  As ADODB.Recordset
20      Dim sEmploye   As String
25      Dim iCompteur  As Integer
    
30      sEmploye = txtemployé.Text
  
        'si le nom de l'employé, le User ID, le Password et la confirmation du password ne sont pas vide
35      If Trim(txtemployé.Text) <> vbNullString And Trim(txtuserid.Text) <> vbNullString And Trim(txtpasswd.Text) <> vbNullString And Trim(txtconfirme.Text) <> vbNullString And Trim(txtinitiale.Text) <> vbNullString And cmbFamille.ListIndex <> -1 Then
          'Si le password et la confirmation sont pareils
40        If txtpasswd.Text = txtconfirme.Text Then
            'Ouverture de la connection
45          Screen.MousePointer = vbHourglass

50          Set rstEmploye = New ADODB.Recordset

            'Si en mode ajouter
55          If m_bModeAjouter = True Then
              'Si le nom de l'employé ne se trouve pas dans le combo
60            If ComboContient(cmbEmploye, txtemployé.Text) = False Then
                'Ouverture du recordset sur la table GRB_Employé
65              Call rstEmploye.Open("SELECT * FROM GRB_employés", g_connData, adOpenDynamic, adLockOptimistic)
            
                'tant que c'est pas la fin du recordset
70              Do While Not rstEmploye.EOF
                  'Si les initiales sont les meme que l'employé ajouté
75                If rstEmploye.Fields("Initiale") = txtinitiale.Text Then
80                  Call MsgBox("Ces initiales existent déjà!")
                  
85                  Screen.MousePointer = vbDefault
                
90                  Exit Sub
95                End If
              
                  'Si le user id existe déjà
100               If UCase(rstEmploye.Fields("loginname")) = UCase(txtuserid.Text) Then
105                 Call MsgBox("User ID existant!")
               
110                 Screen.MousePointer = vbDefault

115                 Exit Sub
120               End If
             
125               Call rstEmploye.MoveNext
130             Loop
            
135             Call rstEmploye.AddNew
140           Else
145             Call MsgBox("Cet employé existe déjà!")
            
150             Exit Sub
155           End If
160         Else
              'Si ce n'est pas un ajout
        
              'Ouverture du recordset sur la table GRB_Employés ou le no d'employe est = à la variable m_iNoEmploye
165           Call rstEmploye.Open("SELECT * FROM GRB_employés WHERE noemploye = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)
          
              'Si le contenu de txtuserid est différent au contenu du champs loginname
170           If txtuserid.Text <> rstEmploye.Fields("loginname") Then
                'Si le contenu de txtuserid est différent de g_sUserID
175             If txtuserid.Text <> g_sUserID Then
180               Set rstUserId = New ADODB.Recordset

185               Call rstUserId.Open("SELECT * FROM GRB_employés", g_connData, adOpenDynamic, adLockOptimistic)
              
190               Do While Not rstUserId.EOF
                    'Si le loginname du recordest est égal au txtuserid
195                 If UCase(rstUserId.Fields("loginname")) = UCase(txtuserid.Text) Then
200                   Call MsgBox("User ID existant!")

205                   Call rstUserId.Close
210                   Set rstUserId = Nothing

215                   Call rstEmploye.Close
220                   Set rstEmploye = Nothing
               
225                   Exit Sub
230                 End If
              
235                 Call rstUserId.MoveNext
240               Loop

245               Call rstUserId.Close
250               Set rstUserId = Nothing
255             End If
260           End If
265         End If
        
270         rstEmploye.Fields("employe").Value = txtemployé.Text
        
            'Si l'employé fait des modif sur lui-même
275         If g_sUserID = rstEmploye.Fields("loginname") Then
280           g_sUserID = txtuserid.Text
285         End If
                          
290         rstEmploye.Fields("loginname").Value = txtuserid.Text
295         rstEmploye.Fields("passwd").Value = txtpasswd.Text
300         rstEmploye.Fields("initiale").Value = txtinitiale.Text
305         rstEmploye.Fields("Actif").Value = chkActif.Value

310         If m_bModeAjouter = False Then
315           If chkActif.Value = vbUnchecked Then
320             Call g_connData.Execute("DELETE * FROM GRB_AutorisationPunch WHERE NoEmploye = " & m_iNoEmploye & " OR AutoriserPar = " & m_iNoEmploye)
325           End If
330         End If

335         If mskTelephone.Text = vbNullString Then
340           rstEmploye.Fields("tel").Value = " "
345         Else
350           rstEmploye.Fields("tel").Value = mskTelephone.Text
355         End If
        
360         If mskCellulaire.Text = vbNullString Then
365           rstEmploye.Fields("cell").Value = " "
370         Else
375           rstEmploye.Fields("cell").Value = mskCellulaire.Text
380         End If
        
385         If mskPagette.Text = vbNullString Then
390           rstEmploye.Fields("page").Value = " "
395         Else
400           rstEmploye.Fields("page").Value = mskPagette.Text
405         End If
        
            'Celà veut dire que l'utilisateur a le droit de modifier le groupe
410         If cmbGroupe.Visible = True Then
415           rstEmploye.Fields("groupe").Value = cmbGroupe.ItemData(cmbGroupe.ListIndex)
420         End If

425         If cmbFamille.Visible = True Then
430           If cmbFamille.ListIndex <> -1 Then
435             rstEmploye.Fields("Famille").Value = cmbFamille.ItemData(cmbFamille.ListIndex)
440           End If
445         End If
                  
450         Call rstEmploye.Update
      
455         Call rstEmploye.Close
460         Set rstEmploye = Nothing
           
465         Screen.MousePointer = vbDefault
              
470         Call MontrerControles(MODE_INACTIF)
              
475         Call LockedChamps(MODE_INACTIF)
        
480         Call ActiverBoutonsGroupe
         
            'remplis combo
485         Call RemplirComboEmploye
      
490         Call RemplirComboEmployeActif
      
495         For iCompteur = 0 To cmbEmploye.ListCount
500           If cmbEmploye.LIST(iCompteur) = sEmploye Then
505             cmbEmploye.ListIndex = iCompteur
        
510             Exit For
515           End If
520         Next

525         m_bModeAjouter = False
530       Else
535         Call MsgBox("Le mot de passe est incorrect!")
540       End If
   
545       Call CacherMasque
550     Else
555       Call MsgBox("Champs vide!")
560     End If

565     Exit Sub

AfficherErreur:

570     woups "frmemploye", "cmdenregistré_Click", Err, Erl
End Sub

Private Sub cmdModifier_Click()

5       On Error GoTo AfficherErreur
        
        '''''''''''''''''''''
        'met en mode modifier
        '''''''''''''''''''''
10      Call AfficherMasque
  
        'Si le user a le droit de modifier les autres user
15      If m_bModifEmploye = True Then
20        Call MontrerControles(MODE_MODIF)
25        Call RemplirComboGroupe

30        If txtGroupe.Text <> "" Then
35          cmbGroupe.Text = txtGroupe.Text
40        Else
45          cmbGroupe.ListIndex = -1
50        End If

55        If txtFamille.Text <> "" Then
60          cmbFamille.Text = txtFamille.Text
65        Else
70          cmbFamille.ListIndex = -1
75        End If

80        Call LockedChamps(MODE_MODIF)
85      Else
90        Call MontrerControles(MODE_MODIF_NON_AUTORISE)
95        Call LockedChamps(MODE_MODIF_NON_AUTORISE)
100     End If
    
105     txtconfirme.Text = txtpasswd.Text
110     m_bModeAjouter = False

115     Exit Sub

AfficherErreur:

120     woups "frmemploye", "cmdModifier_Click", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmemploye", "cmdFermer_Click", Err, Erl
End Sub

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur

10      Dim rstAutorisation As ADODB.Recordset
  
15      Set rstAutorisation = New ADODB.Recordset
  
20      Call rstAutorisation.Open("SELECT * FROM GRB_AutorisationPunch", g_connData, adOpenDynamic, adLockOptimistic)
  
25      Call rstAutorisation.AddNew
    
30      rstAutorisation.Fields("NoEmploye") = cmbAjoutEmploye.ItemData(cmbAjoutEmploye.ListIndex)
35      rstAutorisation.Fields("AutoriserPar") = cmbEmploye.ItemData(cmbEmploye.ListIndex)
    
40      Call rstAutorisation.Update
    
45      Call rstAutorisation.Close
50      Set rstAutorisation = Nothing
  
55      Call RemplirComboEmployePunch
  
60      fraEmploye.Visible = False

65      Exit Sub

AfficherErreur:

70      woups "frmemploye", "cmdOK_Click", Err, Erl
End Sub

Private Sub cmdSupprimePunch_Click()

5       On Error GoTo AfficherErreur

10      Dim iEmploye As Integer
15      Dim iPunch   As Integer
  
20      If cmbEmployePunch.ListCount > 0 Then
25        iEmploye = cmbEmploye.ItemData(cmbEmploye.ListIndex)
30        iPunch = cmbEmployePunch.ItemData(cmbEmployePunch.ListIndex)
  
35        If cmbEmployePunch.ListIndex > -1 Then
40          If MsgBox("Êtes vous sûr de vouloir supprimer cet employé?", vbYesNo) = vbYes Then
45            Call g_connData.Execute("DELETE * FROM GRB_AutorisationPunch WHERE NoEmploye = " & iPunch & " AND AutoriserPar = " & iEmploye)
50          End If
      
55          Call RemplirComboEmployePunch
60        End If
65      End If

70      Exit Sub

AfficherErreur:

75      woups "frmemploye", "cmdSupprimePunch_Click", Err, Erl
End Sub

Private Sub cmdsupprimer_Click()

5       On Error GoTo AfficherErreur

        '''''''''''''''''''''''''''''''''''
        'supprime employé
        ''''''''''''''''''''''''''''''
10      Dim rstProjSoum As ADODB.Recordset
15      Dim rstFT       As ADODB.Recordset
20      Dim sTampon     As String

25      If cmbEmploye.ListCount > 0 Then
          'si on veut supprimer
30        If MsgBox("Etes-vous sur de supprimer cet enregistrement?", vbYesNo, "Supprimer") = vbYes Then
35          Set rstFT = New ADODB.Recordset
                    
40          Call rstFT.Open("SELECT * FROM GRB_Punch WHERE NoEmploye = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)
        
45          If rstFT.EOF Then
50            Set rstProjSoum = New ADODB.Recordset

55            Call rstProjSoum.Open("SELECT * FROM GRB_Soumission_Modif WHERE NoEmployé = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)
        
60            If rstProjSoum.EOF Then
65              Call rstProjSoum.Close
          
70              Call rstProjSoum.Open("SELECT * FROM GRB_Projet_Modif WHERE NoEmployé = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)
          
75              If rstProjSoum.EOF Then
                  'delete employe
80                Call g_connData.Execute("DELETE * FROM GRB_employés WHERE noemploye = " & m_iNoEmploye)
                                    
85                Call rstProjSoum.Close
90                Set rstProjSoum = Nothing
              
95                Call rstFT.Close
100               Set rstFT = Nothing
            
105               Call RemplirComboEmploye
              
110               If cmbEmploye.ListCount > 0 Then
115                 cmbEmploye.ListIndex = 0
120               End If
125             Else
130               Call MsgBox("Impossible d'effacer cet employé, il est utilisé dans le projet " & rstProjSoum.Fields("IDProjet") & "!", vbOKOnly, "Erreur")
             
135               Call rstProjSoum.Close
140               Set rstProjSoum = Nothing
145             End If
150           Else
155             Call MsgBox("Impossible d'effacer cet employé, il est utilisé dans la soumission " & rstProjSoum.Fields("IDSoumission") & "!", vbOKOnly, "Erreur")
            
160             Call rstProjSoum.Close
165             Set rstProjSoum = Nothing
170           End If
175         Else
180           Call MsgBox("Impossible d'effacer cet employé, il est utilisé dans les feuilles de temps pour le projet " & rstFT.Fields("NoProjet") & "!", vbOKOnly, "Erreur")
          
185           Call rstFT.Close
190           Set rstFT = Nothing
195         End If
200       End If
205     End If

210     Exit Sub

AfficherErreur:

215     woups "frmemploye", "cmdSupprimer_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur
        
        'remplis combo employe
10      Call RemplirComboEmploye
  
15      Call RemplirComboFamille
  
20      Call MontrerControles(MODE_INACTIF)
  
25      Call LockedChamps(MODE_INACTIF)
  
30      Call ActiverBoutonsGroupe

        'selectionne dans combo employe
35      If cmbEmploye.ListCount >= 0 Then
40        cmbEmploye.ListIndex = 0
45      End If

50      Screen.MousePointer = vbDefault

55      Exit Sub

AfficherErreur:

60      woups "frmemploye", "Form_Load", Err, Erl
End Sub

Private Sub ActiverBoutonsGroupe()

5       On Error GoTo AfficherErreur
        
        'Activation des boutons selon le groupe
10      m_bModifEmploye = g_bModificationEmployes
  
15      Cmdajouter.Enabled = m_bModifEmploye
20      cmdModifier.Enabled = m_bModifEmploye
25      cmdsupprimer.Enabled = m_bModifEmploye
  
30      cmdConfig.Enabled = g_bModificationGroupes
  
35      cmdAjoutPunch.Enabled = g_bModificationPunchEmployes
40      cmdSupprimePunch.Enabled = g_bModificationPunchEmployes
  
45      Exit Sub

AfficherErreur:

50      woups "frmemploye", "ActiverBoutonsGroupe", Err, Erl
End Sub

Private Sub AfficherEmploye()

5       On Error GoTo AfficherErreur
        
        ''''''''''''''''''''''''''''''''''''''''''
        ' affiche donne de l'employé selectionné '
        ''''''''''''''''''''''''''''''''''''''''''
10      Dim rstEmploye As ADODB.Recordset
15      Dim rstGroupe  As ADODB.Recordset
20      Dim rstFamille As ADODB.Recordset
  
25      Set rstEmploye = New ADODB.Recordset
  
30      Call rstEmploye.Open("SELECT * FROM GRB_employés WHERE NoEmploye = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)
        
        'REMPLIS LES CHAMPS
35      If Not rstEmploye.EOF Then
40        txtpasswd.Text = rstEmploye.Fields("passwd")
45        txtuserid.Text = rstEmploye.Fields("loginname")
50        txtinitiale.Text = rstEmploye.Fields("initiale")
      
55        If IsNull(rstEmploye.Fields("groupe")) Then
60          txtGroupe.Text = vbNullString
65        Else
70          Set rstGroupe = New ADODB.Recordset

75          Call rstGroupe.Open("SELECT * FROM GRB_Groupes WHERE IDGroupe = " & rstEmploye.Fields("Groupe"), g_connData, adOpenDynamic, adLockOptimistic)
        
80          txtGroupe.Text = rstGroupe.Fields("NomGroupe")
        
85          Call rstGroupe.Close
90          Set rstGroupe = Nothing
95        End If

100       If IsNull(rstEmploye.Fields("Famille")) Then
105         txtFamille.Text = vbNullString
110       Else
115         Set rstFamille = New ADODB.Recordset

120         Call rstFamille.Open("SELECT * FROM GRB_Famille WHERE IDFamille = " & rstEmploye.Fields("Famille"), g_connData, adOpenDynamic, adLockOptimistic)
        
125         txtFamille.Text = rstFamille.Fields("Famille")
        
130         Call rstFamille.Close
135         Set rstFamille = Nothing
140       End If
        
145       If IsNull(rstEmploye.Fields("cell")) Then
150         txtCell.Text = vbNullString
155       Else
160         txtCell.Text = Trim(rstEmploye.Fields("cell"))
165       End If
        
170       If IsNull(rstEmploye.Fields("Page")) Then
175         txtPage.Text = vbNullString
180       Else
185         txtPage.Text = Trim(rstEmploye.Fields("Page"))
190       End If
        
195       If IsNull(rstEmploye.Fields("tel")) Then
200         txtTel.Text = vbNullString
205       Else
210         txtTel.Text = Trim(rstEmploye.Fields("tel"))
215       End If
      
220       chkActif.Value = Abs(CInt(rstEmploye.Fields("Actif")))
    
225       Call RemplirComboEmployePunch
230     End If
  
235     Call rstEmploye.Close
240     Set rstEmploye = Nothing

245     Exit Sub

AfficherErreur:

250     woups "frmemploye", "AfficherEmploye", Err, Erl
End Sub

Private Sub RemplirComboEmployePunch()

5       On Error GoTo AfficherErreur

10      Dim rstEmployePunch As ADODB.Recordset
15      Dim rstEmploye      As ADODB.Recordset

20      Call cmbEmployePunch.Clear

25      Set rstEmployePunch = New ADODB.Recordset

30      Call rstEmployePunch.Open("SELECT * FROM GRB_AutorisationPunch WHERE AutoriserPar = " & cmbEmploye.ItemData(cmbEmploye.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
       
35      Set rstEmploye = New ADODB.Recordset
       
40      Do While Not rstEmployePunch.EOF
45        Call rstEmploye.Open("SELECT Employe, NoEmploye FROM GRB_Employés WHERE NoEmploye = " & rstEmployePunch.Fields("NoEmploye"), g_connData, adOpenDynamic, adLockOptimistic)
      
50        Call cmbEmployePunch.AddItem(rstEmploye.Fields("Employe"))
      
55        cmbEmployePunch.ItemData(cmbEmployePunch.newIndex) = rstEmploye.Fields("NoEmploye")
      
60        Call rstEmploye.Close
        
65        Call rstEmployePunch.MoveNext
70      Loop

75      Set rstEmploye = Nothing
        
80      If cmbEmployePunch.ListCount > 0 Then
85        cmbEmployePunch.ListIndex = 0
90      End If
    
95      Call rstEmployePunch.Close
100     Set rstEmployePunch = Nothing

105     Exit Sub

AfficherErreur:

110     woups "frmemploye", "RemplirComboEmployePunch", Err, Erl
End Sub

Private Sub RemplirComboEmploye()

5       On Error GoTo AfficherErreur
        
        '''''''''''''''''''''''''
        'rempli le combo employé'
        '''''''''''''''''''''''''
10      Dim rstEmploye As ADODB.Recordset
  
15      Set rstEmploye = New ADODB.Recordset
  
20      Call rstEmploye.Open("SELECT * FROM GRB_employés ORDER BY employe", g_connData, adOpenDynamic, adLockOptimistic)

25      Call cmbEmploye.Clear
  
        'remplis le combo employé
30      Do While Not rstEmploye.EOF
35        Call cmbEmploye.AddItem(rstEmploye.Fields("employe"))
        
40        cmbEmploye.ItemData(cmbEmploye.newIndex) = rstEmploye.Fields("noEmploye")
      
45        Call rstEmploye.MoveNext
50      Loop
  
55      Call rstEmploye.Close
60      Set rstEmploye = Nothing

65      Exit Sub

AfficherErreur:

70      woups "frmemploye", "RemplirComboEmploye", Err, Erl
End Sub

Private Sub RemplirComboEmployeActif()

5       On Error GoTo AfficherErreur

10      Dim rstEmploye As ADODB.Recordset
15      Dim iCompteur  As Integer
20      Dim iCompteur2 As Integer
25      Dim bSupprimer As Boolean
  
30      Set rstEmploye = New ADODB.Recordset
  
35      Call rstEmploye.Open("SELECT * FROM GRB_employés WHERE Actif = true ORDER BY employe", g_connData, adOpenDynamic, adLockOptimistic)
    
40      Call cmbAjoutEmploye.Clear
    
        'rempli le combo employé
45      Do While Not rstEmploye.EOF
50        Call cmbAjoutEmploye.AddItem(rstEmploye.Fields("employe"))
      
55        cmbAjoutEmploye.ItemData(cmbAjoutEmploye.newIndex) = rstEmploye.Fields("noEmploye")
      
60        Call rstEmploye.MoveNext
65      Loop
    
70      Call rstEmploye.Close
75      Set rstEmploye = Nothing
  
80      iCompteur = 0
  
        'Il faut enlever les employés déjà dans le combo et l'employé en cours
85      Do While iCompteur <= cmbAjoutEmploye.ListCount - 1
90        bSupprimer = False
    
          'Si c'est l'employé en cours
95       If cmbAjoutEmploye.LIST(iCompteur) = cmbEmploye.Text Then
100         bSupprimer = True
105       Else
110         iCompteur2 = 0
    
            'Si c'est les employés dans le combo
115         Do While iCompteur2 <= cmbEmployePunch.ListCount - 1
120           If cmbEmployePunch.LIST(iCompteur2) = cmbAjoutEmploye.LIST(iCompteur) Then
125             bSupprimer = True
130           End If
        
135           iCompteur2 = iCompteur2 + 1
140         Loop
145       End If
    
150       If bSupprimer = True Then
155         Call cmbAjoutEmploye.RemoveItem(iCompteur)
160       Else
165         iCompteur = iCompteur + 1
170       End If
175     Loop
  
180     If cmbAjoutEmploye.ListCount > 0 Then
185       cmbAjoutEmploye.ListIndex = 0
190     End If

195     Exit Sub

AfficherErreur:

200     woups "frmemploye", "RemplirComboEmployeActif", Err, Erl
End Sub

Private Sub mskCellulaire_GotFocus()

5       On Error GoTo AfficherErreur

10      mskCellulaire.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmemploye", "mskCellulaire_GotFocus", Err, Erl
End Sub

Private Sub mskPagette_GotFocus()

5       On Error GoTo AfficherErreur

10      mskPagette.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmemploye", "mskPagette_GotFocus", Err, Erl
End Sub

Private Sub mskTelephone_GotFocus()

5       On Error GoTo AfficherErreur

10      mskTelephone.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmemploye", "mskTelephone_GotFocus", Err, Erl
End Sub

Private Sub mskCellulaire_LostFocus()

5       On Error GoTo AfficherErreur

10      mskCellulaire.mask = vbNullString
  
15      If mskCellulaire.Text = "(___) ___-____" Then
20        mskCellulaire.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmemploye", "mskCellulaire_LostFocus", Err, Erl
End Sub

Private Sub mskPagette_LostFocus()

5       On Error GoTo AfficherErreur

10      mskPagette.mask = vbNullString
  
15      If mskPagette.Text = "(___) ___-____" Then
20        mskPagette.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmemploye", "mskPagette_LostFocus", Err, Erl
End Sub

Private Sub mskTelephone_LostFocus()

5       On Error GoTo AfficherErreur

10      mskTelephone.mask = vbNullString
  
15      If mskTelephone.Text = "(___) ___-____" Then
20        mskTelephone.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmemploye", "mskTelephone_LostFocus", Err, Erl
End Sub
