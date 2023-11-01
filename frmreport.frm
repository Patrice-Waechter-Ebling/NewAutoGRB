VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmreport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rapports"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "frmreport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbFournisseur2 
      Height          =   315
      Left            =   4200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   67
      Top             =   7680
      Width           =   4095
   End
   Begin VB.CommandButton cmdRechercherFRS2 
      Caption         =   "..."
      Height          =   315
      Left            =   8400
      TabIndex        =   66
      Top             =   7680
      Width           =   375
   End
   Begin VB.Frame fraChoixRapport 
      Caption         =   "LISTE DE RAPPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtMsg 
         Height          =   3615
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.CheckBox chkProblemes 
         Caption         =   "PROBLÈMES"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox txtObjet 
         Height          =   285
         Left            =   960
         TabIndex        =   19
         Top             =   5400
         Width           =   2895
      End
      Begin VB.CommandButton cmdMsg 
         Caption         =   "Message"
         Height          =   375
         Left            =   2760
         TabIndex        =   14
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtDe 
         Height          =   285
         Left            =   960
         TabIndex        =   17
         Top             =   5040
         Width           =   2175
      End
      Begin VB.TextBox txtPage 
         Height          =   285
         Left            =   960
         TabIndex        =   16
         Top             =   4680
         Width           =   735
      End
      Begin VB.CheckBox chkFaxAnglais 
         Caption         =   "Fax Anglais"
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CheckBox chkFaxFrancais 
         Caption         =   "Fax Français"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CheckBox chkBonLivraison 
         Caption         =   "BON DE LIVRAISON"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Sélectionner tout"
         Height          =   495
         Left            =   960
         TabIndex        =   21
         Top             =   5880
         Width           =   2055
      End
      Begin VB.CheckBox ChkFinFab 
         Caption         =   "FINS DE FABRICATION"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   3240
         Width           =   2055
      End
      Begin VB.CheckBox ChkFabFerm 
         Caption         =   "FABRICATION - FERMETURE"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2880
         Width           =   3495
      End
      Begin VB.CheckBox ChkFabFermMéca 
         Caption         =   "FABRICATION - FERMETURE MÉCANIQUE"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Width           =   3615
      End
      Begin VB.CheckBox ChkProg 
         Caption         =   "PROGRAMMATION"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CheckBox ChkConcept 
         Caption         =   "CONCEPTION"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CheckBox ChkFourn 
         Caption         =   "FOURNISSEUR"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox ChkClient 
         Caption         =   "CLIENT"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox ChkBonTravail 
         Caption         =   "BON DE TRAVAIL"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblObjet 
         Caption         =   "Objet:"
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   5400
         Width           =   495
      End
      Begin VB.Label lblDe 
         Caption         =   "De:"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   5040
         Width           =   495
      End
      Begin VB.Label lblPage 
         Caption         =   "Pages"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   4680
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbContactFRS 
      Height          =   315
      ItemData        =   "frmreport.frx":0442
      Left            =   4200
      List            =   "frmreport.frx":0444
      TabIndex        =   64
      Top             =   6960
      Width           =   4095
   End
   Begin VB.CommandButton cmdRechercherFRS 
      Caption         =   "..."
      Height          =   315
      Left            =   8400
      TabIndex        =   60
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton cmdRechercherClient2 
      Caption         =   "..."
      Height          =   315
      Left            =   8400
      TabIndex        =   53
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton cmdRechercherClient 
      Caption         =   "..."
      Height          =   315
      Left            =   8400
      TabIndex        =   23
      Top             =   240
      Width           =   375
   End
   Begin MSMask.MaskEdBox mskDateTravaux 
      Height          =   255
      Left            =   4200
      TabIndex        =   47
      Top             =   4560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtProjetClient 
      Height          =   285
      Left            =   5880
      TabIndex        =   39
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton cmdReport 
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
      Height          =   615
      Left            =   720
      TabIndex        =   61
      Top             =   6960
      Width           =   1335
   End
   Begin VB.ComboBox cmbFournisseur 
      Height          =   315
      Left            =   4200
      TabIndex        =   59
      Text            =   "cmbFournisseur"
      Top             =   6360
      Width           =   4095
   End
   Begin MSMask.MaskEdBox mskDateLivraison 
      Height          =   315
      Left            =   6960
      TabIndex        =   57
      Top             =   5760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Left            =   4200
      TabIndex        =   56
      Top             =   5760
      Width           =   2655
   End
   Begin VB.TextBox txtNomProjet 
      Height          =   285
      Left            =   5880
      TabIndex        =   35
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox txtNoCommande 
      Height          =   255
      Left            =   6720
      TabIndex        =   44
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox txtNoProjet 
      Height          =   285
      Left            =   4200
      TabIndex        =   34
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtNoSoumission 
      Height          =   285
      Left            =   4200
      TabIndex        =   30
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   62
      Top             =   6960
      Width           =   1335
   End
   Begin VB.ComboBox cmbGRB 
      Height          =   315
      Left            =   4200
      TabIndex        =   27
      Top             =   1440
      Width           =   4095
   End
   Begin VB.ComboBox cmbContact 
      Height          =   315
      ItemData        =   "frmreport.frx":0446
      Left            =   4200
      List            =   "frmreport.frx":0448
      TabIndex        =   25
      Top             =   840
      Width           =   4095
   End
   Begin VB.ComboBox cmbClient 
      Height          =   315
      ItemData        =   "frmreport.frx":044A
      Left            =   4200
      List            =   "frmreport.frx":044C
      Sorted          =   -1  'True
      TabIndex        =   22
      Text            =   "cmbClient"
      Top             =   240
      Width           =   4095
   End
   Begin VB.ComboBox cmbClient2 
      Height          =   315
      Left            =   4200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   5160
      Width           =   4095
   End
   Begin MSMask.MaskEdBox mskDateCommande 
      Height          =   255
      Left            =   4200
      TabIndex        =   42
      Top             =   3960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDate 
      Height          =   255
      Left            =   4200
      TabIndex        =   38
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskHeureTravaux 
      Height          =   255
      Left            =   6720
      TabIndex        =   49
      Top             =   4560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDateDue 
      Height          =   255
      Left            =   5880
      TabIndex        =   31
      Top             =   2160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Label lblFournisseur2 
      Caption         =   "Fournisseur Expédié à"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   68
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Label lblDateDue 
      Caption         =   "Date dûe (AA-MM-JJ)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   29
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblFormatHeurePrevue 
      Caption         =   "HH:MM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   50
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label lblFormatDateTravaux 
      Caption         =   "AA-MM-JJ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   48
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label lblFormatDateCommande 
      Caption         =   "AA-MM-JJ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   43
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblContactFRS 
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   63
      Top             =   6720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblHeureTravaux 
      Caption         =   "Heure prévue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   46
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblProjetClient 
      Caption         =   "Projet Client"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   37
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label lblFournisseur 
      Caption         =   "Fournisseur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   58
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label lblDateLivraison 
      Caption         =   "Date livraison (AA-MM-JJ)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   55
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label lblTransport 
      Caption         =   "Transport"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   54
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label lblClient2 
      Caption         =   "Client Expédié à"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   51
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblClient 
      Caption         =   "Client"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Label lblDate 
      Caption         =   "Date (AA-MM-JJ)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   36
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblDateTravaux 
      Caption         =   "Date travaux"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   45
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblNomProjet 
      Caption         =   "Nom Projet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   33
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label lblDateCommande 
      Caption         =   "Date Commande"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   40
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblNoCommande 
      Caption         =   "No. Commande Client"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   41
      Top             =   3720
      Width           =   1830
   End
   Begin VB.Label lblNoProjet 
      Caption         =   "No. Projet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   32
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblNoSoumission 
      Caption         =   "No. Soumission"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   28
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblGRB 
      Caption         =   "Représentant GRB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   26
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblContact 
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   24
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frmreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const S_SELECT_ALL   As String = "Sélectionner tout"
Private Const S_UNSELECT_ALL As String = "Désélectionner tout"

Private Enum enumLangueFax
  FAX_FRANCAIS = 0
  FAX_ANGLAIS = 1
End Enum

Public Enum enumForm
  FRM_CLIENTS = 0
  FRM_FRS = 1
  FRM_CONTACTS = 2
End Enum

Private m_iNoClient     As Integer
Private m_iNoClient2    As Integer
Private m_iNoContact    As Integer
Private m_iNoGRB        As Integer
Private m_iNoFRS        As Integer
Private m_iNoFRS2       As Integer
Private m_bSelectAll    As Boolean

Private m_sFaxClientFRS As String
Private m_sFaxContact   As String

Private m_sTelClientFRS As String
Private m_sTelContact   As String

Public Sub Afficher(ByVal iNoClientFRS, iNoContact, eForm As enumForm)

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      cmdSelect.Caption = S_UNSELECT_ALL

20      Call cmdselect_Click

25      chkFaxFrancais.Value = vbChecked
  
30      cmbclient.ListIndex = -1
35      cmbContact.ListIndex = -1

40      cmbFournisseur.ListIndex = -1
45      cmbContactFRS.ListIndex = -1
  
50      Select Case eForm
          Case FRM_CLIENTS:
55          For iCompteur = 0 To cmbclient.ListCount - 1
60            If cmbclient.ItemData(iCompteur) = iNoClientFRS Then
65              cmbclient.ListIndex = iCompteur

70              Exit For
75            End If
80          Next

85          For iCompteur = 0 To cmbContact.ListCount - 1
90            If cmbContact.ItemData(iCompteur) = iNoContact Then
95              cmbContact.ListIndex = iCompteur

100             Exit For
105           End If
110         Next

          Case FRM_FRS:
115         For iCompteur = 0 To cmbFournisseur.ListCount - 1
120           If cmbFournisseur.ItemData(iCompteur) = iNoClientFRS Then
125             cmbFournisseur.ListIndex = iCompteur

130             Exit For
135           End If
140         Next

145         For iCompteur = 0 To cmbContactFRS.ListCount - 1
150           If cmbContactFRS.ItemData(iCompteur) = iNoContact Then
155             cmbContactFRS.ListIndex = iCompteur

160             Exit For
165           End If
170         Next

          Case FRM_CONTACTS:
175         For iCompteur = 0 To cmbContact.ListCount - 1
180           If cmbContact.ItemData(iCompteur) = iNoContact Then
185             cmbContact.ListIndex = iCompteur

190             Exit For
195           End If
200         Next
205     End Select

210     txtMsg.Visible = True

215     Call Me.Show

220     Exit Sub

AfficherErreur:

225     woups "frmreport", "Afficher", Err, Erl
End Sub

Private Sub chkBonLivraison_Click()

5       On Error GoTo AfficherErreur

10      If m_bSelectAll = False Then
15        Call AfficherControles
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmreport", "chkBonLivraison_Click", Err, Erl
End Sub

Private Sub chkBonTravail_Click()

5       On Error GoTo AfficherErreur

10      If m_bSelectAll = False Then
15        Call AfficherControles
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmreport", "chkBonTravail_Click", Err, Erl
End Sub

Private Sub chkClient_Click()

5       On Error GoTo AfficherErreur

10      If m_bSelectAll = False Then
15        Call AfficherControles
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmreport", "chkClient_Click", Err, Erl
End Sub

Private Sub chkFourn_Click()

5       On Error GoTo AfficherErreur

10      If m_bSelectAll = False Then
15        Call AfficherControles
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmreport", "chkFourn_Click", Err, Erl
End Sub

Private Sub chkConcept_Click()

5       On Error GoTo AfficherErreur

10      If m_bSelectAll = False Then
15        Call AfficherControles
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmreport", "chkConcept_Click", Err, Erl
End Sub

Private Sub chkProblemes_Click()

5       On Error GoTo AfficherErreur

10      If m_bSelectAll = False Then
15        Call AfficherControles
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmreport", "chkProblemes_Click", Err, Erl

End Sub

Private Sub chkProg_Click()

5       On Error GoTo AfficherErreur

10      If m_bSelectAll = False Then
15        Call AfficherControles
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmreport", "chkProg_Click", Err, Erl
End Sub

Private Sub chkFabFermMéca_Click()

5       On Error GoTo AfficherErreur

10      If m_bSelectAll = False Then
15        Call AfficherControles
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmreport", "chkFabFermMéca_Click", Err, Erl
End Sub

Private Sub chkFabFerm_Click()

5       On Error GoTo AfficherErreur

10      If m_bSelectAll = False Then
15        Call AfficherControles
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmreport", "chkFabFerm_Click", Err, Erl
End Sub

Private Sub chkFaxFrancais_Click()

5       On Error GoTo AfficherErreur

10      If m_bSelectAll = False Then
15        Call AfficherControles
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmreport", "chkFaxFrancais_Click", Err, Erl
End Sub

Private Sub chkFaxAnglais_Click()

5       On Error GoTo AfficherErreur

10      If m_bSelectAll = False Then
15        Call AfficherControles
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmreport", "chkFaxAnglais_Click", Err, Erl
End Sub

Private Sub cmbclient_Click()

5       On Error GoTo AfficherErreur

10      Dim rstClient As ADODB.Recordset

15      m_sTelClientFRS = ""
20      m_sTelContact = ""

25      m_sFaxClientFRS = ""
30      m_sTelContact = ""
  
        'affiche client selectionné dans textbox
35      If cmbclient.ListIndex <> -1 Then
40        Set rstClient = New ADODB.Recordset

45        Call rstClient.Open("SELECT Fax, Telephonne FROM GRB_Client WHERE IDClient = " & cmbclient.ItemData(cmbclient.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
    
50        If Not rstClient.EOF Then
55          If Not IsNull(rstClient.Fields("Fax")) Then
60            m_sFaxClientFRS = rstClient.Fields("Fax")
65          Else
70            m_sFaxClientFRS = vbNullString
75          End If
      
80          If Not IsNull(rstClient.Fields("Telephonne")) Then
85            m_sTelClientFRS = rstClient.Fields("Telephonne")
90          Else
95            m_sTelClientFRS = vbNullString
100         End If
105       End If
    
          'ferme la table
110       Call rstClient.Close
115       Set rstClient = Nothing
     
120       If cmbclient.ListIndex > -1 Then
125         m_iNoClient = cmbclient.ItemData(cmbclient.ListIndex)
130       End If

135       Call RemplirComboContact
140     End If

145     Exit Sub

AfficherErreur:

150     woups "frmreport", "cmbclient_Click", Err, Erl
End Sub

Private Sub AfficherControles()

5       On Error GoTo AfficherErreur
        
        'Affichage des contrôles selon le rapport choisi
   
        'Met tous les contrôles invisible
10      cmbclient.Visible = False
15      cmbClient2.Visible = False
20      cmbContact.Visible = False
25      cmbContactFRS.Visible = False
30      cmbFournisseur.Visible = False
35      cmbFournisseur2.Visible = False
40      cmbGRB.Visible = False

45      cmdMsg.Visible = False
50      cmdRechercherClient.Visible = False
55      cmdRechercherClient2.Visible = False
60      cmdRechercherFRS.Visible = False
65      cmdRechercherFRS2.Visible = False

70      lblClient.Visible = False
75      lblClient2.Visible = False
80      lblContact.Visible = False
85      lblContactFRS.Visible = False
90      lblDate.Visible = False
95      lblDateCommande.Visible = False
100     lblDateDue.Visible = False
105     lblDateLivraison.Visible = False
110     lblDateTravaux.Visible = False
115     lblDe.Visible = False
120     lblFormatDateCommande.Visible = False
125     lblFormatDateTravaux.Visible = False
130     lblFormatHeurePrevue.Visible = False
135     lblFournisseur.Visible = False
140     lblFournisseur2.Visible = False
145     lblGRB.Visible = False
150     lblHeureTravaux.Visible = False
155     lblNoCommande.Visible = False
160     lblNomProjet.Visible = False
165     lblNoProjet.Visible = False
170     lblNoSoumission.Visible = False
175     lblObjet.Visible = False
180     lblPage.Visible = False
185     lblProjetClient.Visible = False
190     lbltransport.Visible = False

195     mskDate.Visible = False
200     mskDateCommande.Visible = False
205     mskDateDue.Visible = False
210     mskDateLivraison.Visible = False
215     mskDateTravaux.Visible = False
220     mskHeureTravaux.Visible = False

225     txtDe.Visible = False
230     txtMsg.Visible = False
235     txtNoCommande.Visible = False
240     txtNomProjet.Visible = False
245     txtnoprojet.Visible = False
250     txtNoSoumission.Visible = False
255     txtObjet.Visible = False
260     txtPage.Visible = False
265     txtProjetClient.Visible = False
270     txtTransport.Visible = False

        'Si c'est le rapport de problèmes
275     If chkProblemes.Value = vbChecked Then
280       lblGRB.Visible = True
285       cmbGRB.Visible = True

290       lblNoProjet.Visible = True
295       txtnoprojet.Visible = True

300       lblNoSoumission.Visible = True
305       txtNoSoumission.Visible = True
310     End If
  
        'Si c'est client ou fournisseur
315     If ChkClient.Value = vbChecked Or ChkFourn.Value = vbChecked Then
320       lblDateDue.Visible = True
325       mskDateDue.Visible = True
330     End If
  
        'Si c'est bon de travail, bon de livraison, client, fournisseur,
        'conception, programmation, Fermeture mécanique, fermeture, fax
335     If ChkBonTravail.Value = vbChecked Or chkBonLivraison.Value = vbChecked Or _
           ChkClient.Value = vbChecked Or ChkFourn.Value = vbChecked Or _
           ChkConcept.Value = vbChecked Or ChkProg.Value = vbChecked Or _
           ChkFabFermMéca.Value = vbChecked Or ChkFabFerm.Value = vbChecked Or _
           chkFaxFrancais.Value = vbChecked Or chkFaxAnglais.Value = vbChecked Then
    
340       cmbclient.Visible = True
345       lblClient.Visible = True
    
350       cmdRechercherClient.Visible = True

355       cmbContact.Visible = True
360       lblContact.Visible = True
    
365       txtnoprojet.Visible = True
370       lblNoProjet.Visible = True
375     End If
  
        'Si c'est bon de travail ou bon de livraison
380     If ChkBonTravail.Value = vbChecked Or chkBonLivraison.Value = vbChecked Then
385       txtNoCommande.Visible = True
390       lblNoCommande.Visible = True
395     End If
  
        'Si c'est bon de travail
400     If ChkBonTravail.Value = vbChecked Then
405       cmbGRB.Visible = True
410       lblGRB.Visible = True
    
415       mskDateCommande.Visible = True
420       lblDateCommande.Visible = True
425       lblFormatDateCommande.Visible = True
    
430       mskDateTravaux.Visible = True
435       lblDateTravaux.Visible = True
440       lblFormatDateTravaux.Visible = True
    
445       mskHeureTravaux.Visible = True
450       lblHeureTravaux.Visible = True
455       lblFormatHeurePrevue.Visible = True
460     End If
  
        'Si c'est bon de livraison ou fax
465     If chkBonLivraison.Value = vbChecked Or _
          chkFaxFrancais.Value = vbChecked Or chkFaxAnglais.Value = vbChecked Then
    
470       cmbFournisseur.Visible = True
475       lblFournisseur.Visible = True
    
480       cmbContactFRS.Visible = True
485       lblContactFRS.Visible = True
    
490       cmdRechercherFRS.Visible = True
495     End If

        'Si c'est bon de livraison
500     If chkBonLivraison.Value = vbChecked Then
505       cmbClient2.Visible = True
510       lblClient2.Visible = True

515       cmbFournisseur2.Visible = True
520       lblFournisseur2.Visible = True
    
525       cmdRechercherClient2.Visible = True
530       cmdRechercherFRS2.Visible = True
    
535       txtTransport.Visible = True
540       lbltransport.Visible = True
    
545       mskDateLivraison.Visible = True
550       lblDateLivraison.Visible = True
555     End If
  
        'Si c'est client, fournisseur, conception, programmation, fermeture mécan,
        'fermeture ou fax
560     If ChkClient.Value = vbChecked Or ChkFourn.Value = vbChecked Or _
           ChkConcept.Value = vbChecked Or ChkProg.Value = vbChecked Or _
           ChkFabFermMéca.Value = vbChecked Or ChkFabFerm.Value = vbChecked Or _
           chkFaxFrancais.Value = vbChecked Or chkFaxAnglais.Value = vbChecked Then
    
565       txtNoSoumission.Visible = True
570       lblNoSoumission.Visible = True
575     End If
  
        'Si c'est client, fournisseur, conception, programmation, fermeture mécan,
        'fermeture
580     If ChkClient.Value = vbChecked Or ChkFourn.Value = vbChecked Or _
           ChkConcept.Value = vbChecked Or ChkProg.Value = vbChecked Or _
           ChkFabFermMéca.Value = vbChecked Or ChkFabFerm.Value = vbChecked Then
    
585       txtNomProjet.Visible = True
590       lblNomProjet.Visible = True
    
595       mskDate.Visible = True
600       lblDate.Visible = True
    
605       txtProjetClient.Visible = True
610       lblProjetClient.Visible = True
615     End If
  
        'Si c'est fax
620     If chkFaxFrancais.Value = vbChecked Or chkFaxAnglais.Value = vbChecked Then
625       txtPage.Visible = True
630       lblPage.Visible = True
   
635       txtDe.Visible = True
640       lblDe.Visible = True
    
645       cmdMsg.Visible = True
    
650       txtObjet.Visible = True
655       lblObjet.Visible = True
660     End If

665     Exit Sub

AfficherErreur:

670     woups "frmreport", "AfficherControles", Err, Erl
End Sub

Private Sub cmdRechercherClient_Click()

5       On Error GoTo AfficherErreur

10      Dim sRecherche As String
  
15      sRecherche = InputBox("Entrez le texte à rechercher.")
  
20      Call RemplirComboClient(sRecherche)
  
25      m_iNoClient = 0
  
30      Call RemplirComboContact
  
35      Call cmbclient.SetFocus

40      Exit Sub

AfficherErreur:

45      woups "frmreport", "cmdRechercherClient_Click", Err, Erl
End Sub
  
Private Sub cmdRechercherClient2_Click()

5       On Error GoTo AfficherErreur

10      Dim sRecherche As String
  
15      sRecherche = InputBox("Entrez le texte à rechercher.")
  
20      Call RemplirComboClient2(sRecherche)

25      m_iNoClient2 = 0
  
30      Call cmbClient2.SetFocus

35      Exit Sub

AfficherErreur:

40      woups "frmreport", "cmdRechercherClient2_Click", Err, Erl
End Sub

Private Sub cmdRechercherFRS_Click()

5       On Error GoTo AfficherErreur

10      Dim sRecherche As String
  
15      sRecherche = InputBox("Entrez le texte à rechercher.")
  
20      Call RemplirComboFRS(sRecherche)

25      m_iNoFRS = 0
  
30      Call cmbFournisseur.SetFocus

35      Exit Sub

AfficherErreur:

40      woups "frmreport", "cmdRechercherFRS_Click", Err, Erl
End Sub

Private Sub cmdRechercherFRS2_Click()

5       On Error GoTo AfficherErreur

10      Dim sRecherche As String

15      sRecherche = InputBox("Entrez le texte à rechercher.")

20      Call RemplirComboFRS2(sRecherche)

25      m_iNoFRS2 = 0

30      Call cmbFournisseur2.SetFocus

35      Exit Sub

AfficherErreur:

40      woups "frmreport", "cmdRechercherFRS2_Click", Err, Erl
End Sub

Private Sub cmbclient2_Click()

5       On Error GoTo AfficherErreur

        'Affiche client2 selectionné dans textbox
10      If cmbClient2.ListIndex > -1 Then
15        m_iNoClient2 = cmbClient2.ItemData(cmbClient2.ListIndex)
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmreport", "cmbclient2_Click", Err, Erl
End Sub

Private Sub cmbFournisseur2_Click()

5       On Error GoTo AfficherErreur

        'Affiche client2 selectionné dans textbox
10      If cmbFournisseur2.ListIndex > -1 Then
15        m_iNoFRS2 = cmbFournisseur2.ItemData(cmbFournisseur2.ListIndex)
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmreport", "cmbFournisseur2_Click", Err, Erl
End Sub

Private Sub cmbFournisseur_Click()

5       On Error GoTo AfficherErreur

10      Dim rstFRS As ADODB.Recordset
        'affiche fournisseur selectionné dans textbox
  
15      m_sFaxContact = ""
20      m_sFaxClientFRS = ""
        
25      m_sTelContact = ""
30      m_sTelClientFRS = ""
  
35      If cmbFournisseur.ListIndex <> -1 Then
40        Set rstFRS = New ADODB.Recordset

45        Call rstFRS.Open("SELECT Fax, Telephonne FROM GRB_Fournisseur WHERE IDFRS = " & cmbFournisseur.ItemData(cmbFournisseur.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
  
50        If Not rstFRS.EOF Then
55          If Not IsNull(rstFRS.Fields("Fax")) Then
60            m_sFaxClientFRS = rstFRS.Fields("Fax")
65          Else
70            m_sFaxClientFRS = vbNullString
75          End If
    
80          If Not IsNull(rstFRS.Fields("Telephonne")) Then
85            m_sTelClientFRS = rstFRS.Fields("Telephonne")
90          Else
95            m_sTelClientFRS = vbNullString
100          End If
105       End If
  
          'ferme la table
110       Call rstFRS.Close
115       Set rstFRS = Nothing
120     End If

125     If cmbFournisseur.ListIndex > -1 Then
130       m_iNoFRS = cmbFournisseur.ItemData(cmbFournisseur.ListIndex)
135     End If

140     Call RemplirComboContactFRS

145     Exit Sub

AfficherErreur:

150     woups "frmreport", "cmbFournisseur_Click", Err, Erl
End Sub

Private Sub cmbContact_Click()

5       On Error GoTo AfficherErreur

10      Dim rstContact As ADODB.Recordset
15      Dim sTampon    As String

20      If cmbContact.ListIndex <> -1 Then
25        m_iNoContact = cmbContact.ItemData(cmbContact.ListIndex)
30      End If

35      If cmbContact.ListIndex <> -1 Then
40        sTampon = cmbContact.ItemData(cmbContact.ListIndex)
    
45        Set rstContact = New ADODB.Recordset
    
50        Call rstContact.Open("SELECT Telephonne, Fax FROM GRB_contact WHERE IDContact = " & sTampon, g_connData, adOpenDynamic, adLockOptimistic)
  
          'remplis le champ text, telephone et fax
55        If Not rstContact.EOF Then
60          If IsNull(rstContact.Fields("telephonne")) Then
65            m_sTelContact = vbNullString
70          Else
75            m_sTelContact = rstContact.Fields("telephonne")
80          End If
      
85          If IsNull(rstContact.Fields("fax")) Then
90            m_sFaxContact = vbNullString
95          Else
100           m_sFaxContact = rstContact.Fields("fax")
105         End If
110       End If
  
115       Call rstContact.Close
120       Set rstContact = Nothing
125     End If

130     Exit Sub

AfficherErreur:

135     woups "frmreport", "cmbContact_Click", Err, Erl
End Sub

Private Sub cmbContactFRS_Click()

5       On Error GoTo AfficherErreur

10      Dim rstContact As ADODB.Recordset
  
30      If cmbContactFRS.ListIndex <> -1 Then
          'Si il y a un client de choisi, le fax sera pour le client
35        If cmbContact.ListIndex = -1 Then
40          Set rstContact = New ADODB.Recordset

45          Call rstContact.Open("SELECT Telephonne, Fax FROM GRB_Contact WHERE IDContact = " & cmbContactFRS.ItemData(cmbContactFRS.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
  
            'remplis le champ text, telephone et fax
50          If Not rstContact.EOF Then
55            If Not IsNull(rstContact.Fields("telephonne")) Then
60              m_sTelContact = rstContact.Fields("telephonne")
65            Else
70              m_sTelContact = vbNullString
75            End If
      
80            If Not IsNull(rstContact.Fields("fax")) Then
85              m_sFaxContact = rstContact.Fields("fax")
90            Else
95              m_sFaxContact = vbNullString
100           End If
105         End If
  
110         Call rstContact.Close
115         Set rstContact = Nothing
120       End If
125     End If

130     Exit Sub

AfficherErreur:

135     woups "frmreport", "cmbContactFRS_Click", Err, Erl
End Sub

Private Sub cmbgrb_Click()

5       On Error GoTo AfficherErreur
        
        'affiche contactgrb selectionné dans textbox
10      If cmbGRB.ListIndex > -1 Then
15        m_iNoGRB = cmbGRB.ItemData(cmbGRB.ListIndex)
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmreport", "cmbgrb_Click", Err, Erl
End Sub

Private Sub cmdmsg_Click()

5       On Error GoTo AfficherErreur

10      If txtMsg.Visible = True Then
15        txtMsg.Visible = False
20      Else
25        txtMsg.Visible = True

30        Call txtMsg.SetFocus
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmreport", "cmdmsg_Click", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmreport", "cmdFermer_Click", Err, Erl
End Sub

Private Sub ImprimerBonTravail()

5       On Error GoTo AfficherErreur

10      Dim rstBonTravail As ADODB.Recordset
  
15      Set rstBonTravail = New ADODB.Recordset
    
20      Call rstBonTravail.Open("SELECT * FROM GRB_impression_bonlivraison", g_connData, adOpenDynamic, adLockOptimistic)

        'Set le rapport
25      Set DR_BonTravail.DataSource = rstBonTravail

        'Contenu label
30      DR_BonTravail.Sections(1).Controls("lblClient").Caption = cmbclient.Text
35      DR_BonTravail.Sections(1).Controls("lblContact").Caption = cmbContact.Text
40      DR_BonTravail.Sections(1).Controls("lblTelephone").Caption = m_sTelContact
45      DR_BonTravail.Sections(1).Controls("lblFax").Caption = m_sFaxContact
50      DR_BonTravail.Sections(1).Controls("lblRepresentantGRB").Caption = cmbGRB.Text
55      DR_BonTravail.Sections(1).Controls("lblBonTravail").Caption = txtnoprojet.Text
60      DR_BonTravail.Sections(1).Controls("lblNoCommandeClient").Caption = txtNoCommande.Text
65      DR_BonTravail.Sections(1).Controls("lblDateCommande").Caption = mskDateCommande.Text
70      DR_BonTravail.Sections(1).Controls("lblDateHeure").Caption = mskDateTravaux.Text & " " & mskHeureTravaux.Text
          
        'Affiche rapport
75      Call DR_BonTravail.Show(vbModal)
  
80      Call rstBonTravail.Close
85      Set rstBonTravail = Nothing

90      Exit Sub

AfficherErreur:

95      woups "frmreport", "ImprimerBonTravail", Err, Erl
End Sub

Private Sub ImprimerBonLivraison()

5       On Error GoTo AfficherErreur

10      Dim rstBonLivraison As ADODB.Recordset
15      Dim rstFournisseur  As ADODB.Recordset
20      Dim rstFournisseur2 As ADODB.Recordset
25      Dim rstClient       As ADODB.Recordset
30      Dim rstClient2      As ADODB.Recordset
35      Dim rstContact      As ADODB.Recordset
40      Dim sTampon         As String

        'ouvre fenetre bon livraison
        'pour entrer qte
45      Call OuvrirForm(frmbonlivraison, True)
          
50      Screen.MousePointer = vbHourglass
          
55      Set rstBonLivraison = New ADODB.Recordset
60      Set rstFournisseur = New ADODB.Recordset
65      Set rstFournisseur2 = New ADODB.Recordset
70      Set rstClient = New ADODB.Recordset
75      Set rstClient2 = New ADODB.Recordset
          
80      Call rstBonLivraison.Open("SELECT * FROM GRB_impression_bonlivraison WHERE user = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
85      Call rstFournisseur.Open("SELECT * FROM GRB_Fournisseur WHERE IDFRS = " & m_iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)
90      Call rstFournisseur2.Open("SELECT * FROM GRB_Fournisseur WHERE IDFRS = " & m_iNoFRS2, g_connData, adOpenDynamic, adLockOptimistic)
95      Call rstClient.Open("SELECT * FROM GRB_client WHERE IDClient = " & m_iNoClient, g_connData, adOpenDynamic, adLockOptimistic)
100     Call rstClient2.Open("SELECT * FROM GRB_client WHERE IDClient = " & m_iNoClient2, g_connData, adOpenDynamic, adLockOptimistic)
      
        'set le rapport
105     Set DR_BonLivraison.DataSource = rstBonLivraison
      
        'contenu label
110     DR_BonLivraison.Sections(1).Controls(34).Caption = txtnoprojet.Text
115     DR_BonLivraison.Sections(1).Controls(35).Caption = ConvertDate(Date)
          
        'si fournisseur
120     If cmbFournisseur.ListIndex <> -1 Then
125       DR_BonLivraison.Sections(1).Controls(36).Caption = cmbFournisseur.Text
        
130       If rstFournisseur.EOF Then
135         DR_BonLivraison.Sections(1).Controls(37).Caption = vbNullString
140       Else
145         sTampon = vbNullString
            ''''''''''''''''''''''''''''''''''''''''''''
            'rempli adresse pay ville prov et codepostal si pas vide
            '''''''''''''''''''''''''''''''''''''''''''''
            'adresse
150         If Not IsNull(rstFournisseur.Fields("adresse")) Then
155           sTampon = rstFournisseur.Fields("adresse")
160         End If
           
165         DR_BonLivraison.Sections(1).Controls(37).Caption = sTampon
          
170         sTampon = vbNullString
              
            'ville
175         If Not IsNull(rstFournisseur.Fields("ville")) Then
180           sTampon = rstFournisseur.Fields("ville")
185         End If
            
            'province
190         If Not IsNull(rstFournisseur.Fields("Prov/Etat")) Then
195           If rstFournisseur.Fields("Prov/Etat") <> vbNullString Then
200             sTampon = sTampon + ", " + rstFournisseur.Fields("Prov/Etat")
205           End If
210         End If
              
215         DR_BonLivraison.Sections(1).Controls(44).Caption = sTampon
         
220         sTampon = vbNullString
              
            'Pays
225         If Not IsNull(rstFournisseur.Fields("pays")) Then
230           sTampon = rstFournisseur.Fields("pays")
235         End If
            
            'codepostal
240         If Not IsNull(rstFournisseur.Fields("codepostal")) Then
245           If rstFournisseur.Fields("CodePostal") <> vbNullString Then
250             sTampon = sTampon + ", " + rstFournisseur.Fields("codepostal")
255           End If
260         End If
              
265         DR_BonLivraison.Sections(1).Controls(45).Caption = sTampon
270       End If
          
275       DR_BonLivraison.Sections(1).Controls(38).Caption = txtNoCommande.Text
280       DR_BonLivraison.Sections(1).Controls(39).Caption = cmbFournisseur2.Text
        
285       If rstFournisseur2.EOF Then
290         DR_BonLivraison.Sections(1).Controls(40).Caption = vbNullString
295       Else
300         sTampon = vbNullString
            ''''''''''''''''''''''''''''''''''''''''''''
            'rempli adresse pay ville prov et codepostal si pas vide
            '''''''''''''''''''''''''''''''''''''''''''''
            'adresse
305         If Not IsNull(rstFournisseur2.Fields("adresse")) Then
310           sTampon = sTampon + rstFournisseur2.Fields("adresse")
315         End If
            
320         DR_BonLivraison.Sections(1).Controls(40).Caption = sTampon

325         sTampon = vbNullString
           
            'ville
330         If Not IsNull(rstFournisseur2.Fields("ville")) Then
335           sTampon = rstFournisseur2.Fields("ville")
340         End If
            
            'prov
345         If Not IsNull(rstFournisseur2.Fields("prov/etat")) Then
350           If rstFournisseur2.Fields("prov/etat") <> vbNullString Then
355             sTampon = sTampon + ", " + rstFournisseur2.Fields("prov/etat")
360           End If
365         End If
              
370         DR_BonLivraison.Sections(1).Controls(46).Caption = sTampon

375         sTampon = vbNullString
            
            'pays
380         If Not IsNull(rstFournisseur2.Fields("pays")) Then
385           sTampon = rstFournisseur2.Fields("pays")
390         End If
              
            'codepostal
395         If Not IsNull(rstFournisseur2.Fields("codepostal")) Then
400           If rstFournisseur2.Fields("CodePostal") <> vbNullString Then
405             sTampon = sTampon + ", " + rstFournisseur2.Fields("codepostal")
410           End If
415         End If
              
420         DR_BonLivraison.Sections(1).Controls(47).Caption = sTampon
425       End If
      
430       DR_BonLivraison.Sections(1).Controls(41).Caption = cmbContactFRS.Text
435     Else
          'si client
440       DR_BonLivraison.Sections(1).Controls(36).Caption = cmbclient.Text
        
445       If rstClient.EOF Then
450         DR_BonLivraison.Sections(1).Controls(37).Caption = vbNullString
455       Else
460         sTampon = vbNullString
            ''''''''''''''''''''''''''''''''''''''''''''
            'rempli adresse pay ville prov et codepostal si pas vide
            '''''''''''''''''''''''''''''''''''''''''''''
            'adresse
465         If Not IsNull(rstClient.Fields("adresseliv")) Then
470           sTampon = sTampon + rstClient.Fields("adresseliv")
475         End If
              
480         DR_BonLivraison.Sections(1).Controls(37).Caption = sTampon

485         sTampon = vbNullString
           
            'ville
490         If Not IsNull(rstClient.Fields("villeliv")) Then
495           sTampon = rstClient.Fields("villeliv")
500         End If
          
            'pays
505         If Not IsNull(rstClient.Fields("paysliv")) Then
510           If rstClient.Fields("PaysLiv") <> vbNullString Then
515             sTampon = sTampon + ", " + rstClient.Fields("paysliv")
520           End If
525         End If
           
530         DR_BonLivraison.Sections(1).Controls(44).Caption = sTampon
535         sTampon = vbNullString
            
            'province
540         If Not IsNull(rstClient.Fields("prov/etatliv")) Then
545           sTampon = rstClient.Fields("prov/etatliv")
550         End If
              
            'codepostal
555         If Not IsNull(rstClient.Fields("codepostalliv")) Then
560           If rstClient.Fields("CodePostalLiv") <> vbNullString Then
565             sTampon = sTampon + ", " + rstClient.Fields("codepostalliv")
570           End If
575         End If
           
580         DR_BonLivraison.Sections(1).Controls(45).Caption = sTampon
585       End If
          
590       DR_BonLivraison.Sections(1).Controls(38).Caption = txtNoCommande.Text
595       DR_BonLivraison.Sections(1).Controls(39).Caption = cmbClient2.Text
         
600       If rstClient2.EOF Then
605         DR_BonLivraison.Sections(1).Controls(40).Caption = vbNullString
610       Else
615         sTampon = vbNullString
            ''''''''''''''''''''''''''''''''''''''''''''
            'rempli adresse pay ville prov et codepostal si pas vide
            '''''''''''''''''''''''''''''''''''''''''''''
            'adresse
620         If Not IsNull(rstClient2.Fields("adresseliv")) Then
625           sTampon = rstClient2.Fields("adresseliv")
630         End If
           
635         DR_BonLivraison.Sections(1).Controls(40).Caption = sTampon
640         sTampon = vbNullString
        
            'ville
645         If Not IsNull(rstClient2.Fields("villeliv")) Then
650           sTampon = rstClient2.Fields("villeliv")
655         End If
            
            'pays
660         If Not IsNull(rstClient2.Fields("paysliv")) Then
665           If rstClient2.Fields("PaysLiv") <> vbNullString Then
670             sTampon = sTampon + ", " + rstClient2.Fields("paysliv")
675           End If
680         End If
              
685         DR_BonLivraison.Sections(1).Controls(46).Caption = sTampon
690         sTampon = vbNullString
          
            'province
695         If Not IsNull(rstClient2.Fields("prov/etatliv")) Then
700           sTampon = rstClient2.Fields("prov/etatliv")
705         End If
          
            'codepostal
710         If Not IsNull(rstClient2.Fields("codepostalliv")) Then
715           If rstClient2.Fields("CodePostalLiv") <> vbNullString Then
720             sTampon = sTampon + ", " + rstClient2.Fields("codepostalliv")
725           End If
730         End If
            
735         DR_BonLivraison.Sections(1).Controls(47).Caption = sTampon
740       End If
      
745       DR_BonLivraison.Sections(1).Controls(41).Caption = cmbContact.Text
750     End If
    
755     DR_BonLivraison.Sections(1).Controls(42).Caption = txtTransport.Text
760     DR_BonLivraison.Sections(1).Controls(43).Caption = mskDateLivraison.Text
      
              'affiche rapport
765     Call DR_BonLivraison.Show(vbModal)
    
770     Call rstBonLivraison.Close
775     Set rstBonLivraison = Nothing
    
780     Call rstFournisseur.Close
785     Set rstFournisseur = Nothing
    
790     Call rstClient.Close
795     Set rstClient = Nothing
    
800     Call rstClient2.Close
805     Set rstClient2 = Nothing

810     Exit Sub

AfficherErreur:

815     woups "frmreport", "ImprimerBonLivraison", Err, Erl
End Sub

Private Sub ImprimerClient()

5       On Error GoTo AfficherErreur

10      Dim rstBonLivraison As ADODB.Recordset
  
15      Set rstBonLivraison = New ADODB.Recordset
  
20      Call rstBonLivraison.Open("SELECT * FROM GRB_impression_bonlivraison", g_connData, adOpenDynamic, adLockOptimistic)
  
        'set le rapport
25      Set DR_Client.DataSource = rstBonLivraison
    
        'contenu label
30      DR_Client.Sections("Section4").Controls("lblClient").Caption = cmbclient.Text
35      DR_Client.Sections("Section4").Controls("lblContact").Caption = cmbContact.Text
40      DR_Client.Sections("Section4").Controls("lblTel").Caption = m_sTelContact
45      DR_Client.Sections("Section4").Controls("lblFax").Caption = m_sFaxContact
50      DR_Client.Sections("Section4").Controls("lblNoSoum").Caption = txtNoSoumission.Text
55      DR_Client.Sections("Section4").Controls("lblNoProj").Caption = txtnoprojet.Text
60      DR_Client.Sections("Section4").Controls("lblProjetNom").Caption = txtNomProjet.Text
65      DR_Client.Sections("Section4").Controls("lblDate").Caption = mskDate.Text
70      DR_Client.Sections("Section4").Controls("lblDateOuverture").Caption = ConvertDate(Date)
75      DR_Client.Sections("Section4").Controls("lblDateDue").Caption = mskDateDue.Text
80      DR_Client.Sections("Section4").Controls("lblProjet").Caption = txtProjetClient.Text
    
        'affiche rapport
85      Call DR_Client.Show(vbModal)
  
90      Call rstBonLivraison.Close
95      Set rstBonLivraison = Nothing

100     Exit Sub

AfficherErreur:

105     woups "frmreport", "ImprimerClient", Err, Erl
End Sub

Private Sub ImprimerFournisseur()

5       On Error GoTo AfficherErreur

10      Dim rstBonLivraison As ADODB.Recordset
  
15      Set rstBonLivraison = New ADODB.Recordset
  
20      Call rstBonLivraison.Open("SELECT * FROM GRB_impression_bonlivraison", g_connData, adOpenDynamic, adLockOptimistic)
  
        'set le rapport
25      Set DR_Fournisseur.DataSource = rstBonLivraison
    
        'contenu label
30      DR_Fournisseur.Sections("Section4").Controls("lblClient").Caption = cmbclient.Text
35      DR_Fournisseur.Sections("Section4").Controls("lblContact").Caption = cmbContact.Text
40      DR_Fournisseur.Sections("Section4").Controls("lblTel").Caption = m_sTelContact
45      DR_Fournisseur.Sections("Section4").Controls("lblFax").Caption = m_sFaxContact
50      DR_Fournisseur.Sections("Section4").Controls("lblNoSoum").Caption = txtNoSoumission.Text
55      DR_Fournisseur.Sections("Section4").Controls("lblNoProj").Caption = txtnoprojet.Text
60      DR_Fournisseur.Sections("Section4").Controls("lblProjetNom").Caption = txtNomProjet.Text
65      DR_Fournisseur.Sections("Section4").Controls("lblDate").Caption = mskDate.Text
70      DR_Fournisseur.Sections("Section4").Controls("lblDateOuverture").Caption = ConvertDate(Date)
75      DR_Fournisseur.Sections("Section4").Controls("lblDateDue").Caption = mskDateDue.Text
80      DR_Fournisseur.Sections("Section4").Controls("lblProjet").Caption = txtProjetClient.Text
        
        'affiche rapport
85      Call DR_Fournisseur.Show(vbModal)
  
90      Call rstBonLivraison.Close
95      Set rstBonLivraison = Nothing

100     Exit Sub

AfficherErreur:

105     woups "frmreport", "ImprimerFournisseur", Err, Erl
End Sub

Private Sub ImprimerConception()

5       On Error GoTo AfficherErreur

10      Dim rstBonLivraison As ADODB.Recordset

15      Set rstBonLivraison = New ADODB.Recordset

20      Call rstBonLivraison.Open("SELECT * FROM GRB_impression_bonlivraison", g_connData, adOpenDynamic, adLockOptimistic)
    
        'set le rapport
25      Set DR_Conception.DataSource = rstBonLivraison
      
        'contenu label
30      DR_Conception.Sections("Section4").Controls(90).Caption = cmbclient.Text
35      DR_Conception.Sections("Section4").Controls(91).Caption = cmbContact.Text
40      DR_Conception.Sections("Section4").Controls(92).Caption = m_sTelContact
45      DR_Conception.Sections("Section4").Controls(93).Caption = m_sFaxContact
50      DR_Conception.Sections("Section4").Controls(94).Caption = txtNoSoumission.Text
55      DR_Conception.Sections("Section4").Controls(95).Caption = txtnoprojet.Text
60      DR_Conception.Sections("Section4").Controls(96).Caption = txtNomProjet.Text
65      DR_Conception.Sections("Section4").Controls(97).Caption = mskDate.Text
70      DR_Conception.Sections("Section4").Controls(98).Caption = Trim$(Right$(CStr(Year(Date)), 2) + "-" + CStr(Month(Date)) + "-" + CStr(Day(Date)))
75      DR_Conception.Sections("Section4").Controls(99).Caption = txtProjetClient.Text
     
        'affiche rapport
80      Call DR_Conception.Show(vbModal)
   
85      Call rstBonLivraison.Close
90      Set rstBonLivraison = Nothing

95      Exit Sub

AfficherErreur:

100     woups "frmreport", "ImprimerConception", Err, Erl
End Sub

Private Sub ImprimerProgrammation()

5       On Error GoTo AfficherErreur

10      Dim rstBonLivraison As ADODB.Recordset

15      Set rstBonLivraison = New ADODB.Recordset

20      Call rstBonLivraison.Open("SELECT * FROM GRB_impression_bonlivraison", g_connData, adOpenDynamic, adLockOptimistic)
  
        'Set le rapport
25      Set DR_Programmation.DataSource = rstBonLivraison
     
        'Contenu label
30      DR_Programmation.Sections("Section4").Controls("lblClient").Caption = cmbclient.Text
35      DR_Programmation.Sections("Section4").Controls("lblContact").Caption = cmbContact.Text
40      DR_Programmation.Sections("Section4").Controls("lblTelephone").Caption = m_sTelContact
45      DR_Programmation.Sections("Section4").Controls("lblFax").Caption = m_sFaxContact
50      DR_Programmation.Sections("Section4").Controls("lblNoSoum").Caption = txtNoSoumission.Text
55      DR_Programmation.Sections("Section4").Controls("lblNoProj").Caption = txtnoprojet.Text
65      DR_Programmation.Sections("Section4").Controls("lblProjetNom").Caption = txtNomProjet.Text
70      DR_Programmation.Sections("Section4").Controls("lblDate").Caption = mskDate.Text
75      DR_Programmation.Sections("Section4").Controls("lblProjetClient").Caption = txtProjetClient.Text
      
        'Affiche rapport
80      Call DR_Programmation.Show(vbModal)

85      Call rstBonLivraison.Close
90      Set rstBonLivraison = Nothing

95      Exit Sub

AfficherErreur:

100     woups "frmreport", "ImprimerProgrammation", Err, Erl
End Sub

Private Sub ImprimerFermetureMecanique()

5       On Error GoTo AfficherErreur

10      Dim rstBonLivraison As ADODB.Recordset

15      Set rstBonLivraison = New ADODB.Recordset

20      Call rstBonLivraison.Open("SELECT * FROM GRB_impression_bonlivraison", g_connData, adOpenDynamic, adLockOptimistic)
    
        'set le rapport
25      Set DR_FermeMeca.DataSource = rstBonLivraison
      
        'contenu label
30      DR_FermeMeca.Sections("Section4").Controls("lblClient").Caption = cmbclient.Text
35      DR_FermeMeca.Sections("Section4").Controls("lblContact").Caption = cmbContact.Text
40      DR_FermeMeca.Sections("Section4").Controls("lblTel").Caption = m_sTelContact
45      DR_FermeMeca.Sections("Section4").Controls("lblFax").Caption = m_sFaxContact
50      DR_FermeMeca.Sections("Section4").Controls("lblSoum").Caption = txtNoSoumission.Text
55      DR_FermeMeca.Sections("Section4").Controls("lblProj").Caption = txtnoprojet.Text
60      DR_FermeMeca.Sections("Section4").Controls("lblProjetNom").Caption = txtNomProjet.Text
65      DR_FermeMeca.Sections("Section4").Controls("lblDate").Caption = mskDate.Text
70      DR_FermeMeca.Sections("Section4").Controls("lblDateOuverture").Caption = ConvertDate(Date)
75      DR_FermeMeca.Sections("Section4").Controls("lblProjet").Caption = txtProjetClient.Text
      
        'affiche rapport
80      Call DR_FermeMeca.Show(vbModal)
  
85      Call rstBonLivraison.Close
90      Set rstBonLivraison = Nothing

95      Exit Sub

AfficherErreur:

100     woups "frmreport", "ImprimerFermetureMecanique", Err, Erl
End Sub

Private Sub ImprimerFermeture()

5       On Error GoTo AfficherErreur

10      Dim rstBonLivraison As ADODB.Recordset
  
15      Set rstBonLivraison = New ADODB.Recordset
  
20      Call rstBonLivraison.Open("SELECT * FROM GRB_impression_bonlivraison", g_connData, adOpenDynamic, adLockOptimistic)

        'set le rapport
25      Set DR_Fermeture.DataSource = rstBonLivraison
      
        'contenu label
30      DR_Fermeture.Sections(1).Controls(117).Caption = cmbclient.Text
35      DR_Fermeture.Sections(1).Controls(118).Caption = cmbContact.Text
40      DR_Fermeture.Sections(1).Controls(119).Caption = m_sTelContact
45      DR_Fermeture.Sections(1).Controls(120).Caption = m_sFaxContact
50      DR_Fermeture.Sections(1).Controls(121).Caption = txtNoSoumission.Text
55      DR_Fermeture.Sections(1).Controls(122).Caption = txtnoprojet.Text
60      DR_Fermeture.Sections(1).Controls(123).Caption = txtNomProjet.Text
65      DR_Fermeture.Sections(1).Controls(124).Caption = mskDate.Text
70      DR_Fermeture.Sections(1).Controls(125).Caption = ConvertDate(Date)
75      DR_Fermeture.Sections(1).Controls(126).Caption = txtProjetClient.Text
    
        'affiche rapport
80      Call DR_Fermeture.Show(vbModal)
  
85      Call rstBonLivraison.Close
90      Set rstBonLivraison = Nothing

95      Exit Sub

AfficherErreur:

100     woups "frmreport", "ImprimerFermeture", Err, Erl
End Sub

Private Sub ImprimerFinFabrication()

5       On Error GoTo AfficherErreur

10      Dim rstBonLivraison As ADODB.Recordset
  
15      Set rstBonLivraison = New ADODB.Recordset
  
20      Call rstBonLivraison.Open("SELECT * FROM GRB_impression_bonlivraison", g_connData, adOpenDynamic, adLockOptimistic)
    
        'set le rapport
25      Set DR_FinFab.DataSource = rstBonLivraison
      
        'affiche rapport
30      Call DR_FinFab.Show(vbModal)
    
35      Call rstBonLivraison.Close
40      Set rstBonLivraison = Nothing

45      Exit Sub

AfficherErreur:

50      woups "frmreport", "ImprimerFinFabrication", Err, Erl
End Sub

Private Sub ImprimerFax(ByVal eLangue As enumLangueFax)

5       On Error GoTo AfficherErreur

10      Dim drFax                   As DataReport
15      Dim rstBonLivraison         As ADODB.Recordset
20      Dim bClient                 As Boolean
25      Dim bClientTexte            As Boolean
30      Dim bClientListIndex        As Boolean
35      Dim bFournisseur            As Boolean
40      Dim bFournisseurTexte       As Boolean
45      Dim bFournisseurListIndex   As Boolean
50      Dim bContactClient          As Boolean
55      Dim bContactClientTexte     As Boolean
60      Dim bContactClientListIndex As Boolean
65      Dim bContactFRS             As Boolean
70      Dim bContactFRSTexte        As Boolean
75      Dim bContactFRSListIndex    As Boolean
80      Dim sMessage                As String
  
85      If eLangue = FAX_ANGLAIS Then
90        Set drFax = DR_FaxAnglais
95      Else
100       Set drFax = DR_FaxFrancais
105     End If
      
110     If cmbclient.ListIndex <> -1 Or cmbclient.Text <> "" Then
115       bClient = True

120       If cmbclient.ListIndex <> -1 Then
125         bClientListIndex = True
130       Else
135         bClientTexte = True
140       End If
145     End If
      
150     If cmbFournisseur.ListIndex <> -1 Or cmbFournisseur.Text <> "" Then
155       bFournisseur = True

160       If cmbFournisseur.ListIndex <> -1 Then
165         bFournisseurListIndex = True
170       Else
175         bFournisseurTexte = True
180       End If
185     End If
    
190     If cmbContact.ListIndex <> -1 Or cmbContact.Text <> "" Then
195       bContactClient = True

200       If cmbContact.ListIndex <> -1 Then
205         bContactClientListIndex = True
210       Else
215         bContactClientTexte = True
220       End If
225     End If

230     If cmbContactFRS.ListIndex <> -1 Or cmbContactFRS.Text <> "" Then
235       bContactFRS = True

240       If cmbContactFRS.ListIndex <> -1 Then
245         bContactFRSListIndex = True
250       Else
255         bContactFRSTexte = True
260       End If
265     End If

270     If bClient = False And bFournisseur = False And bContactClient = False And bContactFRS = False Then
275       If MsgBox("Voulez-vous choisir un destinataire?", vbYesNo) = vbYes Then
280         Exit Sub
285       End If
290     End If
  
        'Ce recordset ne sert à rien, il est utilisé uniquement pour le DataSource
        'du DataReport. Un DataReport ne peut être ouvert s'il n'a pas de DataSource
295     Set rstBonLivraison = New ADODB.Recordset
     
300     Call rstBonLivraison.Open("SELECT * FROM GRB_Impression_BonLivraison", g_connData, adOpenDynamic, adLockOptimistic)
    
305     Set drFax.DataSource = rstBonLivraison
      
        'Contenu label
310     drFax.Sections(1).Controls("lblDate").Caption = ConvertDate(Date)
        
315     If bClient = True Then
320       drFax.Sections(1).Controls("lblAttention").Caption = cmbContact.Text
325     Else
330       If bFournisseur = True Then
335         drFax.Sections(1).Controls("lblAttention").Caption = cmbContactFRS.Text
340       Else
345         If bContactClient = True Then
350           drFax.Sections(1).Controls("lblAttention").Caption = cmbContact.Text
355         Else
360           If bContactFRS = True Then
365             drFax.Sections(1).Controls("lblAttention").Caption = cmbContactFRS.Text
370           Else
375             drFax.Sections(1).Controls("lblAttention").Caption = ""
380           End If
385         End If
390       End If
395     End If
        
400     If bClient = True Then
405       drFax.Sections(1).Controls("lblEntreprise").Caption = cmbclient.Text
410     Else
415       If bFournisseur = True Then
420         drFax.Sections(1).Controls("lblEntreprise").Caption = cmbFournisseur.Text
425       Else
430         drFax.Sections(1).Controls("lblEntreprise").Caption = ""
435       End If
440     End If
                    
445     If bClientListIndex = True And bContactClientListIndex = True Then
450       sMessage = "Voulez-vous afficher le numéro de fax du client?" & vbNewLine & _
                     "Oui - Fax du client" & vbNewLine & _
                     "Non - Fax du contact"
455     Else
460       If bFournisseurListIndex = True And bContactFRSListIndex = True Then
465         sMessage = "Voulez-vous afficher le numéro de fax du fournisseur?" & vbNewLine & _
                       "Oui - Fax du fournisseur" & vbNewLine & _
                       "Non - Fax du contact"
470       End If
475     End If
                     
480     If sMessage = vbNullString Then
485       If bFournisseurListIndex = True Or bClientListIndex = True Then
490         drFax.Sections(1).Controls("lblFax").Caption = m_sFaxClientFRS
495       Else
500         drFax.Sections(1).Controls("lblFax").Caption = m_sFaxContact
505       End If
510     Else
515       If MsgBox(sMessage, vbYesNo) = vbYes Then
520         drFax.Sections(1).Controls("lblFax").Caption = m_sFaxClientFRS
525       Else
530         drFax.Sections(1).Controls("lblFax").Caption = m_sFaxContact
535       End If
540     End If
            
545     If txtnoprojet.Text <> vbNullString Then
550       drFax.Sections(1).Controls("lblNoProjetSoum").Caption = "# Projet:"
555       drFax.Sections(1).Controls("lblProjet").Caption = txtnoprojet.Text
560     Else
565       drFax.Sections(1).Controls("lblNoProjetSoum").Caption = "# Soumission:"
570       drFax.Sections(1).Controls("lblProjet").Caption = txtNoSoumission.Text
575     End If
      
580     drFax.Sections(1).Controls("lblPage").Caption = txtPage.Text
585     drFax.Sections(1).Controls("lblDe").Caption = txtDe.Text
590     drFax.Sections(1).Controls("lblMessage").Caption = txtMsg.Text
595     drFax.Sections(1).Controls("lblSujet").Caption = txtObjet.Text
  
        'Affiche rapport
600     drFax.Orientation = rptOrientPortrait
        
605     If eLangue = FAX_ANGLAIS Then
610       Call DR_FaxAnglais.Show(vbModal)
615     Else
620       Call DR_FaxFrancais.Show(vbModal)
625     End If

630     Call rstBonLivraison.Close
635     Set rstBonLivraison = Nothing

640     Exit Sub

AfficherErreur:

645     woups "frmreport", "ImprimerFax", Err, Erl
End Sub

Private Sub ImprimerProblemes()
  
5       On Error GoTo AfficherErreur

10      Dim rstBonLivraison As ADODB.Recordset
      
        'Ce recordset ne sert à rien, il est utilisé uniquement pour le DataSource
        'du DataReport. Un DataReport ne peut être ouvert s'il n'a pas de DataSource
15      Set rstBonLivraison = New ADODB.Recordset
     
20      Call rstBonLivraison.Open("SELECT * FROM GRB_Impression_BonLivraison", g_connData, adOpenDynamic, adLockOptimistic)
    
25      Set DR_Probleme.DataSource = rstBonLivraison
      
        'Contenu label
30      If txtNoSoumission.Text <> "" Then
35        DR_Probleme.Sections("Section4").Controls("lblTitreProjSoum").Caption = "# Soum :"
40        DR_Probleme.Sections("Section4").Controls("lblNoProjSoum").Caption = txtNoSoumission.Text
45      Else
50        DR_Probleme.Sections("Section4").Controls("lblTitreProjSoum").Caption = "# Projet :"
55        DR_Probleme.Sections("Section4").Controls("lblNoProjSoum").Caption = txtnoprojet.Text
60      End If

65      DR_Probleme.Sections("Section4").Controls("lblNomEmploye").Caption = cmbGRB.Text
        
        'Affiche rapport
70      DR_Probleme.Orientation = rptOrientLandscape
        
75      Call DR_Probleme.Show(vbModal)

80      Call rstBonLivraison.Close
85      Set rstBonLivraison = Nothing

90      Exit Sub

AfficherErreur:

95      woups "frmreport", "ImprimerProblemes", Err, Erl
End Sub

Private Sub cmdreport_Click()

5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbHourglass
  
        'rapport bon de travail
15      If ChkBonTravail.Value = vbChecked Then
20        Call ImprimerBonTravail
25      End If
  
        'rapport bon de livraison
30      If chkBonLivraison.Value = vbChecked Then
35        Call ImprimerBonLivraison
40      End If
  
        'rapport client
45      If ChkClient.Value = vbChecked Then
50        Call ImprimerClient
55      End If
      
        'rapport fournisseur
60      If ChkFourn.Value = vbChecked Then
65        Call ImprimerFournisseur
70      End If
  
        'rapport conception
75      If ChkConcept.Value = vbChecked Then
80        Call ImprimerConception
85      End If
  
        'rapport programmation
90      If ChkProg.Value = vbChecked Then
95        Call ImprimerProgrammation
100     End If
  
        'rapport fabrication - fermeture mécanique
105     If ChkFabFermMéca.Value = vbChecked Then
110       Call ImprimerFermetureMecanique
115     End If
  
        'rapport fabrication - fermeture
120     If ChkFabFerm.Value = vbChecked Then
125       Call ImprimerFermeture
130     End If
  
        'rapport fin fabrication
135     If ChkFinFab.Value = vbChecked Then
140       Call ImprimerFinFabrication
145     End If
  
        'rapport de problèmes
150     If chkProblemes.Value = vbChecked Then
155       Call ImprimerProblemes
160     End If
  
        'rapport fax francais
165     If chkFaxFrancais.Value = vbChecked Then
170       Call ImprimerFax(FAX_FRANCAIS)
175     End If
  
        'rapport fax anglais
180     If chkFaxAnglais.Value = vbChecked Then
185       Call ImprimerFax(FAX_ANGLAIS)
190     End If
  
195     Screen.MousePointer = vbDefault

200     Exit Sub

AfficherErreur:

205     woups "frmreport", "cmdreport_Click", Err, Erl
End Sub

Private Sub cmdselect_Click()

5       On Error GoTo AfficherErreur

10      Dim iValue As Integer
  
15      If cmdSelect.Caption = S_SELECT_ALL Then
20        iValue = vbChecked
    
25        cmdSelect.Caption = S_UNSELECT_ALL
30      Else
35        iValue = vbUnchecked
    
40        cmdSelect.Caption = S_SELECT_ALL
45      End If

50      m_bSelectAll = True
  
55      ChkBonTravail.Value = iValue
60      ChkClient.Value = iValue
65      ChkConcept.Value = iValue
70      ChkFabFerm.Value = iValue
75      ChkFabFermMéca.Value = iValue
80      ChkFinFab.Value = iValue
85      ChkFourn.Value = iValue
90      ChkProg.Value = iValue
95      chkBonLivraison.Value = iValue
100     chkProblemes.Value = iValue
105     chkFaxFrancais.Value = iValue
110     chkFaxAnglais.Value = iValue
  
115     Call AfficherControles
  
120     m_bSelectAll = False

125     Exit Sub

AfficherErreur:

130     woups "frmreport", "cmdselect_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      m_iNoClient = 0
15      m_iNoClient2 = 0
20      m_iNoContact = 0
25      m_iNoFRS = 0
30      m_iNoGRB = 0

        'rempli les combo
35      Call RemplirComboClient(vbNullString)
40      Call RemplirComboClient2(vbNullString)
45      Call RemplirComboContact
50      Call RemplirComboGRB
55      Call RemplirComboFRS(vbNullString)
60      Call RemplirComboFRS2(vbNullString)
  
65      Call AfficherControles
  
70      Screen.MousePointer = vbDefault

        'rempli
75      mskDate.Text = Year(Date) & "-" & Right$("0" & Month(Date), 2) & "-" & Right$("0" & Day(Date), 2)

80      Exit Sub

AfficherErreur:

85      woups "frmreport", "Form_Load", Err, Erl
End Sub

Private Sub RemplirComboFRS(ByVal sRecherche As String)

5       On Error GoTo AfficherErreur

10      Dim rstFournisseur As ADODB.Recordset

15      Set rstFournisseur = New ADODB.Recordset

20      Call rstFournisseur.Open("SELECT NomFournisseur, IDFRS FROM GRB_Fournisseur WHERE Instr(1, NomFournisseur, '" & Replace(sRecherche, "'", "''") & "') > 0 AND Supprimé = False ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)

        'vide combo
25      Call cmbFournisseur.Clear
 
        'rempli les combo tant que pas fin d'enregistrement
30      Do While Not rstFournisseur.EOF
35        Call cmbFournisseur.AddItem(rstFournisseur.Fields("NomFournisseur"))
40        cmbFournisseur.ItemData(cmbFournisseur.newIndex) = rstFournisseur.Fields("IDFRS")

    
45        Call rstFournisseur.MoveNext
50      Loop

55      Call rstFournisseur.Close
60      Set rstFournisseur = Nothing
 
65      Exit Sub

AfficherErreur:

70      woups "frmreport", "RemplirComboFRS", Err, Erl
End Sub

Private Sub RemplirComboFRS2(ByVal sRecherche As String)

5       On Error GoTo AfficherErreur

10      Dim rstFournisseur As ADODB.Recordset

15      Set rstFournisseur = New ADODB.Recordset

20      Call rstFournisseur.Open("SELECT NomFournisseur, IDFRS FROM GRB_Fournisseur WHERE Instr(1, NomFournisseur, '" & Replace(sRecherche, "'", "''") & "') > 0 AND Supprimé = False ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)

        'vide combo
25      Call cmbFournisseur2.Clear
 
        'rempli les combo tant que pas fin d'enregistrement
30      Do While Not rstFournisseur.EOF
35        Call cmbFournisseur2.AddItem(rstFournisseur.Fields("NomFournisseur"))
40        cmbFournisseur2.ItemData(cmbFournisseur2.newIndex) = rstFournisseur.Fields("IDFRS")

    
45        Call rstFournisseur.MoveNext
50      Loop

55      Call rstFournisseur.Close
60      Set rstFournisseur = Nothing
 
65      Exit Sub

AfficherErreur:

70      woups "frmreport", "RemplirComboFRS2", Err, Erl
End Sub

Private Sub RemplirComboClient(ByVal sRecherche As String)

5       On Error GoTo AfficherErreur

10      Dim rstClient As ADODB.Recordset
        
        'set les tables
15      Set rstClient = New ADODB.Recordset
        
20      Call rstClient.Open("SELECT NomClient, IDClient FROM GRB_client WHERE Instr(1, NomClient, '" & Replace(sRecherche, "'", "''") & "') > 0 AND Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)

        'vide combo
25      Call cmbclient.Clear
  
        'rempli les combo tant que pas fin d'enregistrement
30      Do While Not rstClient.EOF
35        Call cmbclient.AddItem(rstClient.Fields("nomclient"))
40        cmbclient.ItemData(cmbclient.newIndex) = rstClient.Fields("idclient")
    
45        Call rstClient.MoveNext
50      Loop
  
55      Call rstClient.Close
60      Set rstClient = Nothing

65      Exit Sub

AfficherErreur:

70      woups "frmreport", "RemplirComboClient", Err, Erl
End Sub

Private Sub RemplirComboClient2(ByVal sRecherche As String)

5       On Error GoTo AfficherErreur

10      Dim rstClient As ADODB.Recordset
  
15      Set rstClient = New ADODB.Recordset
  
20      Call rstClient.Open("SELECT NomClient, IDClient FROM GRB_client WHERE Instr(1, NomClient, '" & Replace(sRecherche, "'", "''") & "') > 0 AND Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)

        'vide combo
25      Call cmbClient2.Clear
  
        'rempli les combo tant que pas fin d'enregistrement
30      Do While Not rstClient.EOF
35        Call cmbClient2.AddItem(rstClient.Fields("nomclient"))
40        cmbClient2.ItemData(cmbClient2.newIndex) = rstClient.Fields("idclient")
  
      
45        Call rstClient.MoveNext
50      Loop
  
55      Call rstClient.Close
60      Set rstClient = Nothing

65      Exit Sub

AfficherErreur:

70      woups "frmreport", "RemplirComboClient2", Err, Erl
End Sub

Private Sub RemplirComboContact()

5       On Error GoTo AfficherErreur

10      Dim rstContact As ADODB.Recordset

        'si client de selectionné, remplis les liste contact pour le client
        'sinon met tout les contact
15      Set rstContact = New ADODB.Recordset
        
20      If m_iNoClient > 0 Then
25        Call rstContact.Open("SELECT GRB_Contact.IDContact, GRB_Contact.NomContact, GRB_ContactClient.NoClient FROM GRB_Contact INNER JOIN GRB_ContactClient ON GRB_Contact.IDContact = GRB_ContactClient.NoContact WHERE CStr(GRB_ContactClient.noclient) = CStr('" & m_iNoClient & "') ORDER BY GRB_contact.NomContact", g_connData, adOpenDynamic, adLockOptimistic)
30      Else
35        Call rstContact.Open("SELECT NomContact, IDContact FROM GRB_Contact WHERE Supprimé = False ORDER BY Nomcontact", g_connData, adOpenDynamic, adLockOptimistic)
40      End If

        'vide combo
45      Call cmbContact.Clear

        'rempli les combo tant que pas fin d'enregistrement
50      Do While Not rstContact.EOF
          'si trouve le text dans le nom du contact, ajoute dans combo
55        If Not IsNull(rstContact.Fields("NomContact")) Then
60          Call cmbContact.AddItem(rstContact.Fields("NomContact"))
65          cmbContact.ItemData(cmbContact.newIndex) = rstContact.Fields("IDContact")
         
70          Call rstContact.MoveNext
75        End If
80      Loop

85      Call rstContact.Close
90      Set rstContact = Nothing

95      Exit Sub

AfficherErreur:

100     woups "frmreport", "RemplirComboContact", Err, Erl
End Sub

Private Sub RemplirComboContactFRS()

5       On Error GoTo AfficherErreur

10      Dim rstContactFRS As ADODB.Recordset
15      Dim rstContact    As ADODB.Recordset
  
        'si fournisseur de selectionné, remplis les liste contact pour le client
        'sinon met tout les contact
20      If m_iNoFRS > 0 Then
25        Set rstContactFRS = New ADODB.Recordset

30        Call rstContactFRS.Open("SELECT * FROM GRB_ContactFRS WHERE NoFRS = " & m_iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)
35      Else
40        Exit Sub
45      End If

        'vide combo
50      Call cmbContactFRS.Clear
    
        'rempli les combo tant que pas fin d'enregistrement
55      Set rstContact = New ADODB.Recordset
        
60      Do While Not rstContactFRS.EOF
65        Call rstContact.Open("SELECT NomContact, IDContact FROM GRB_Contact WHERE IDContact = " & rstContactFRS.Fields("NoContact"), g_connData, adOpenDynamic, adLockOptimistic)
        
70        If Not rstContact.EOF Then
75          Call cmbContactFRS.AddItem(rstContact.Fields("NomContact"))
80          cmbContactFRS.ItemData(cmbContactFRS.newIndex) = rstContact.Fields("IDContact")
85        End If
    
90        Call rstContact.Close

95        Call rstContactFRS.MoveNext
100     Loop
    
105     Set rstContact = Nothing
    
110     Call rstContactFRS.Close
115     Set rstContactFRS = Nothing

120     Exit Sub

AfficherErreur:

125     woups "frmreport", "RemplirComboContactFRS", Err, Erl
End Sub

Private Sub RemplirComboGRB()

5       On Error GoTo AfficherErreur

10      Dim rstContactGRB As ADODB.Recordset

15      Set rstContactGRB = New ADODB.Recordset
  
20      Call rstContactGRB.Open("SELECT employe, noEmploye FROM GRB_employés WHERE Actif = true ORDER BY employe", g_connData, adOpenDynamic, adLockOptimistic)

        'vide combo
25      Call cmbGRB.Clear

        'rempli les combo tant que pas fin d'enregistrement
30      Do While Not rstContactGRB.EOF
35        Call cmbGRB.AddItem(rstContactGRB.Fields("Employe"))
40        cmbGRB.ItemData(cmbGRB.newIndex) = rstContactGRB.Fields("noEmploye")
          
45        Call rstContactGRB.MoveNext
50      Loop
  
55      Call rstContactGRB.Close
60      Set rstContactGRB = Nothing

65      Exit Sub

AfficherErreur:

70      woups "frmreport", "RemplirComboGRB", Err, Erl
End Sub

Private Sub mskDate_GotFocus()

5       On Error GoTo AfficherErreur

10      If Len(mskDate.Text) = 10 Then
15        mskDate.Text = Right$(mskDate.Text, 8)
20      End If
  
25      mskDate.mask = "##-##-##"

30      Exit Sub

AfficherErreur:

35      woups "frmreport", "mskDate_GotFocus", Err, Erl
End Sub

Private Sub mskDate_LostFocus()

5       On Error GoTo AfficherErreur

10      mskDate.mask = vbNullString
  
15      If mskDate.Text = "__-__-__" Then
20        mskDate.Text = vbNullString
25      Else
30        If Len(mskDate.Text) = 8 Then
35          If IsDate(mskDate.Text) Then
40            mskDate.Text = Year(DateSerial(Left$(mskDate.Text, 2), Mid$(mskDate.Text, 4, 2), Right$(mskDate.Text, 2))) & Mid$(mskDate.Text, 3, 8)
45          End If
50        End If
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmreport", "mskDate_LostFocus", Err, Erl
End Sub

Private Sub mskDateDue_GotFocus()

5       On Error GoTo AfficherErreur

10      If Len(mskDateDue.Text) = 10 Then
15        mskDateDue.Text = Right$(mskDateDue.Text, 8)
20      End If
  
25      mskDateDue.mask = "##-##-##"

30      Exit Sub

AfficherErreur:

35      woups "frmreport", "mskDateDue_GotFocus", Err, Erl
End Sub

Private Sub mskDateDue_LostFocus()

5       On Error GoTo AfficherErreur

10      mskDateDue.mask = vbNullString
  
15      If mskDateDue.Text = "__-__-__" Then
20        mskDateDue.Text = vbNullString
25      Else
30        If Len(mskDateDue.Text) = 8 Then
35          If IsDate(mskDateDue.Text) Then
40            mskDateDue.Text = Year(DateSerial(Left$(mskDateDue.Text, 2), Mid$(mskDateDue.Text, 4, 2), Right$(mskDateDue.Text, 2))) & Mid$(mskDateDue.Text, 3, 8)
45          End If
50        End If
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmreport", "mskDateDue_LostFocus", Err, Erl
End Sub


Private Sub mskdatecommande_GotFocus()

5       On Error GoTo AfficherErreur

10      If Len(mskDateCommande.Text) = 10 Then
15        mskDateCommande.Text = Right$(mskDateCommande.Text, 8)
20      End If
  
25      mskDateCommande.mask = "##-##-##"

30      Exit Sub

AfficherErreur:

35      woups "frmreport", "mskdatecommande_GotFocus", Err, Erl
End Sub

Private Sub mskdatecommande_LostFocus()

5       On Error GoTo AfficherErreur

10      mskDateCommande.mask = vbNullString
  
15      If mskDateCommande.Text = "__-__-__" Then
20        mskDateCommande.Text = vbNullString
25      Else
30        If Len(mskDateCommande.Text) = 8 Then
35          If IsDate(mskDateCommande.Text) Then
40            mskDateCommande.Text = Year(DateSerial(Left$(mskDateCommande.Text, 2), Mid$(mskDateCommande.Text, 4, 2), Right$(mskDateCommande, 2))) & Mid$(mskDateCommande.Text, 3, 8)
45          End If
50        End If
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmreport", "mskdatecommande_LostFocus", Err, Erl
End Sub

Private Sub mskdatelivraison_GotFocus()

5       On Error GoTo AfficherErreur

10      If Len(mskDateLivraison.Text) = 10 Then
15        mskDateLivraison.Text = Right$(mskDateCommande.Text, 8)
20      End If
  
25      mskDateLivraison.mask = "##-##-##"

30      Exit Sub

AfficherErreur:

35      woups "frmreport", "mskdatelivraison_GotFocus", Err, Erl
End Sub

Private Sub mskdatelivraison_LostFocus()

5       On Error GoTo AfficherErreur

10      mskDateLivraison.mask = vbNullString
  
15      If mskDateLivraison.Text = "__-__-__" Then
20        mskDateLivraison.Text = vbNullString
25      Else
30        If Len(mskDateLivraison.Text) = 8 Then
35          If IsDate(mskDateLivraison.Text) Then
40            mskDateLivraison.Text = Year(DateSerial(Left$(mskDateLivraison.Text, 2), Mid$(mskDateLivraison.Text, 4, 2), Right$(mskDateLivraison.Text, 2))) & Mid$(mskDateLivraison.Text, 3, 8)
45          End If
50        End If
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmreport", "mskdatelivraison_LostFocus", Err, Erl
End Sub

Private Sub mskDateTravaux_GotFocus()

5       On Error GoTo AfficherErreur

10      If Len(mskDateTravaux.Text) = 10 Then
15        mskDateTravaux.Text = Right$(mskDateTravaux.Text, 8)
20      End If
  
25      mskDateTravaux.mask = "##-##-##"

30      Exit Sub

AfficherErreur:

35      woups "frmreport", "mskDateTravaux_GotFocus", Err, Erl
End Sub

Private Sub mskDateTravaux_LostFocus()

5       On Error GoTo AfficherErreur

10      mskDateTravaux.mask = vbNullString
  
15      If mskDateTravaux.Text = "__-__-__" Then
20        mskDateTravaux.Text = vbNullString
25      Else
30        If Len(mskDateTravaux.Text) = 8 Then
35          If IsDate(mskDateTravaux.Text) Then
40            mskDateTravaux.Text = Year(DateSerial(Left$(mskDateTravaux.Text, 2), Mid$(mskDateTravaux.Text, 4, 2), Right$(mskDateTravaux, 2))) & Mid$(mskDateTravaux.Text, 3, 8)
45          End If
50        End If
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmreport", "mskDateTravaux_LostFocus", Err, Erl
End Sub

Private Sub mskHeureTravaux_GotFocus()

5       On Error GoTo AfficherErreur

10      mskHeureTravaux.mask = "##:##"

15      Exit Sub

AfficherErreur:

20      woups "frmreport", "mskHeureTravaux_GotFocus", Err, Erl
End Sub

Private Sub mskHeureTravaux_LostFocus()

5       On Error GoTo AfficherErreur

10      mskHeureTravaux.mask = vbNullString
      
15      If mskHeureTravaux.Text = "__:__" Then
20        mskHeureTravaux.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmreport", "mskHeureTravaux_LostFocus", Err, Erl
End Sub
