VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmreport 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rapports"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "frmreport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   8895
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
      BackColor       =   &H00FFFFFF&
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
      BackStyle       =   0  'Transparent
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
      BackStyle       =   0  'Transparent
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
      BackStyle       =   0  'Transparent
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
      BackStyle       =   0  'Transparent
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
      BackStyle       =   0  'Transparent
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
      BackStyle       =   0  'Transparent
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
      BackStyle       =   0  'Transparent
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
      BackStyle       =   0  'Transparent
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00FFFFFF&
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
      BackStyle       =   0  'Transparent
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
      BackStyle       =   0  'Transparent
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
      BackStyle       =   0  'Transparent
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00FFFFFF&
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

Private Const S_SELECT_ALL As String = "Sélectionner tout"
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

Private m_iNoClient As Integer
Private m_iNoContact As Integer
Private m_iNoGRB As Integer
Private m_iNoFRS As Integer
Private m_bSelectAll As Boolean

Private m_sFaxClientFRS As String
Private m_sFaxContact As String

Private m_sTelClientFRS As String
Private m_sTelContact As String

Public Sub Afficher(ByVal iNoClientFRS, iNoContact, eForm As enumForm)

 On Error GoTo Oups

 Dim iCompteur As Integer

 cmdSelect.Caption = S_UNSELECT_ALL

 Call cmdselect_Click

 chkFaxFrancais.Value = vbChecked
 
 cmbclient.ListIndex = -1
 cmbContact.ListIndex = -1

 cmbFournisseur.ListIndex = -1
 cmbContactFRS.ListIndex = -1
 
 Select Case eForm
 Case FRM_CLIENTS:
 For iCompteur = 0 To cmbclient.ListCount - 1
  If cmbclient.ItemData(iCompteur) = iNoClientFRS Then
  cmbclient.ListIndex = iCompteur

  Exit For
  End If
  Next

  For iCompteur = 0 To cmbContact.ListCount - 1
  If cmbContact.ItemData(iCompteur) = iNoContact Then
  cmbContact.ListIndex = iCompteur

 Exit For
 End If
 Next

 Case FRM_FRS:
 For iCompteur = 0 To cmbFournisseur.ListCount - 1
 If cmbFournisseur.ItemData(iCompteur) = iNoClientFRS Then
 cmbFournisseur.ListIndex = iCompteur

 Exit For
 End If
 Next

 For iCompteur = 0 To cmbContactFRS.ListCount - 1
 If cmbContactFRS.ItemData(iCompteur) = iNoContact Then
 cmbContactFRS.ListIndex = iCompteur

 Exit For
 End If
 Next

 Case FRM_CONTACTS:
 For iCompteur = 0 To cmbContact.ListCount - 1
 If cmbContact.ItemData(iCompteur) = iNoContact Then
 cmbContact.ListIndex = iCompteur

 Exit For
1  End If
 Next
 End Select

txtMsg.Visible = True

Call Me.Show

Exit Sub

Oups:

wOups "frmreport", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub chkBonLivraison_Click()

 On Error GoTo Oups

 If m_bSelectAll = False Then
 Call AfficherControles
 End If

 Exit Sub

Oups:

 wOups "frmreport", "chkBonLivraison_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkBonTravail_Click()

 On Error GoTo Oups

 If m_bSelectAll = False Then
 Call AfficherControles
 End If

 Exit Sub

Oups:

 wOups "frmreport", "chkBonTravail_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkClient_Click()

 On Error GoTo Oups

 If m_bSelectAll = False Then
 Call AfficherControles
 End If

 Exit Sub

Oups:

 wOups "frmreport", "chkClient_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkFourn_Click()

 On Error GoTo Oups

 If m_bSelectAll = False Then
 Call AfficherControles
 End If

 Exit Sub

Oups:

 wOups "frmreport", "chkFourn_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkConcept_Click()

 On Error GoTo Oups

 If m_bSelectAll = False Then
 Call AfficherControles
 End If

 Exit Sub

Oups:

 wOups "frmreport", "chkConcept_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkProblemes_Click()

 On Error GoTo Oups

 If m_bSelectAll = False Then
 Call AfficherControles
 End If

 Exit Sub

Oups:

 wOups "frmreport", "chkProblemes_Click", Err, Err.number, Err.Description

End Sub

Private Sub chkProg_Click()

 On Error GoTo Oups

 If m_bSelectAll = False Then
 Call AfficherControles
 End If

 Exit Sub

Oups:

 wOups "frmreport", "chkProg_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkFabFermMéca_Click()

 On Error GoTo Oups

 If m_bSelectAll = False Then
 Call AfficherControles
 End If

 Exit Sub

Oups:

 wOups "frmreport", "chkFabFermMéca_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkFabFerm_Click()

 On Error GoTo Oups

 If m_bSelectAll = False Then
 Call AfficherControles
 End If

 Exit Sub

Oups:

 wOups "frmreport", "chkFabFerm_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkFaxFrancais_Click()

 On Error GoTo Oups

 If m_bSelectAll = False Then
 Call AfficherControles
 End If

 Exit Sub

Oups:

 wOups "frmreport", "chkFaxFrancais_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkFaxAnglais_Click()

 On Error GoTo Oups

 If m_bSelectAll = False Then
 Call AfficherControles
 End If

 Exit Sub

Oups:

 wOups "frmreport", "chkFaxAnglais_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbclient_Click()

 On Error GoTo Oups

 Dim rstClient As ADODB.Recordset

 m_sTelClientFRS = ""
 m_sTelContact = ""

 m_sFaxClientFRS = ""
 m_sTelContact = ""
 
 'affiche client selectionné dans textbox
 If cmbclient.ListIndex <> -1 Then
 Set rstClient = New ADODB.Recordset

 Call rstClient.Open("SELECT Fax, Telephonne FROM GrbClient WHERE IDClient = " & cmbclient.ItemData(cmbclient.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstClient.EOF Then
 If Not IsNull(rstClient.Fields("Fax")) Then
  m_sFaxClientFRS = rstClient.Fields("Fax")
  Else
  m_sFaxClientFRS = vbNullString
  End If
 
  If Not IsNull(rstClient.Fields("Telephonne")) Then
  m_sTelClientFRS = rstClient.Fields("Telephonne")
  Else
  m_sTelClientFRS = vbNullString
 End If
1 End If
 
 'ferme la table
 Call rstClient.Close
 Set rstClient = Nothing
 
 If cmbclient.ListIndex > -1 Then
 m_iNoClient = cmbclient.ItemData(cmbclient.ListIndex)
 End If

 Call RemplirComboContact
End If

Exit Sub

Oups:

wOups "frmreport", "cmbclient_Click", Err, Err.number, Err.Description
End Sub

Private Sub AfficherControles()

 On Error GoTo Oups
 
 'Affichage des contrôles selon le rapport choisi
 
 'Met tous les contrôles invisible
 cmbclient.Visible = False
 cmbClient2.Visible = False
 cmbContact.Visible = False
 cmbContactFRS.Visible = False
 cmbFournisseur.Visible = False
 cmbFournisseur2.Visible = False
 cmbGRB.Visible = False

 cmdMsg.Visible = False
 cmdRechercherClient.Visible = False
 cmdRechercherClient2.Visible = False
  cmdRechercherFRS.Visible = False
  cmdRechercherFRS2.Visible = False

  lblClient.Visible = False
  lblClient2.Visible = False
  lblContact.Visible = False
  lblContactFRS.Visible = False
  lblDate.Visible = False
  lblDateCommande.Visible = False
10 lblDateDue.Visible = False
lblDateLivraison.Visible = False
lblDateTravaux.Visible = False
lblDe.Visible = False
lblFormatDateCommande.Visible = False
lblFormatDateTravaux.Visible = False
lblFormatHeurePrevue.Visible = False
lblFournisseur.Visible = False
lblFournisseur2.Visible = False
lblGRB.Visible = False
lblHeureTravaux.Visible = False
lblNoCommande.Visible = False
1  lblNomProjet.Visible = False
lblNoProjet.Visible = False
 lblNoSoumission.Visible = False
lblObjet.Visible = False
 lblPage.Visible = False
lblProjetClient.Visible = False
 lbltransport.Visible = False

1  mskDate.Visible = False
 mskDateCommande.Visible = False
 mskDateDue.Visible = False
mskDateLivraison.Visible = False
mskDateTravaux.Visible = False
mskHeureTravaux.Visible = False

txtDe.Visible = False
txtMsg.Visible = False
txtNoCommande.Visible = False
txtNomProjet.Visible = False
txtnoprojet.Visible = False
txtNoSoumission.Visible = False
txtObjet.Visible = False
2  txtPage.Visible = False
txtProjetClient.Visible = False
2  txtTransport.Visible = False

 'Si c'est le rapport de problèmes
If chkProblemes.Value = vbChecked Then
lblGRB.Visible = True
 cmbGRB.Visible = True

lblNoProjet.Visible = True
 txtnoprojet.Visible = True

lblNoSoumission.Visible = True
3 txtNoSoumission.Visible = True
End If
 
 'Si c'est client ou fournisseur
If ChkClient.Value = vbChecked Or ChkFourn.Value = vbChecked Then
 lblDateDue.Visible = True
 mskDateDue.Visible = True
End If
 
 'Si c'est bon de travail, bon de livraison, client, fournisseur,
 'conception, programmation, Fermeture mécanique, fermeture, fax
If ChkBonTravail.Value = vbChecked Or chkBonLivraison.Value = vbChecked Or _
 ChkClient.Value = vbChecked Or ChkFourn.Value = vbChecked Or _
 ChkConcept.Value = vbChecked Or ChkProg.Value = vbChecked Or _
 ChkFabFermMéca.Value = vbChecked Or ChkFabFerm.Value = vbChecked Or _
 chkFaxFrancais.Value = vbChecked Or chkFaxAnglais.Value = vbChecked Then
 
 cmbclient.Visible = True
 lblClient.Visible = True
 
 cmdRechercherClient.Visible = True

 cmbContact.Visible = True
lblContact.Visible = True
 
 txtnoprojet.Visible = True
lblNoProjet.Visible = True
End If
 
 'Si c'est bon de travail ou bon de livraison
3  If ChkBonTravail.Value = vbChecked Or chkBonLivraison.Value = vbChecked Then
 txtNoCommande.Visible = True
 lblNoCommande.Visible = True
 End If
 
 'Si c'est bon de travail
40 If ChkBonTravail.Value = vbChecked Then
4 cmbGRB.Visible = True
4 lblGRB.Visible = True
 
4 mskDateCommande.Visible = True
4 lblDateCommande.Visible = True
4 lblFormatDateCommande.Visible = True
 
4 mskDateTravaux.Visible = True
4 lblDateTravaux.Visible = True
4 lblFormatDateTravaux.Visible = True
 
4 mskHeureTravaux.Visible = True
4 lblHeureTravaux.Visible = True
4 lblFormatHeurePrevue.Visible = True
4  End If
 
 'Si c'est bon de livraison ou fax
4  If chkBonLivraison.Value = vbChecked Or _
 chkFaxFrancais.Value = vbChecked Or chkFaxAnglais.Value = vbChecked Then
 
4  cmbFournisseur.Visible = True
4  lblFournisseur.Visible = True
 
4  cmbContactFRS.Visible = True
4  lblContactFRS.Visible = True
 
4  cmdRechercherFRS.Visible = True
4  End If

 'Si c'est bon de livraison
50 If chkBonLivraison.Value = vbChecked Then
5 cmbClient2.Visible = True
 lblClient2.Visible = True

 cmbFournisseur2.Visible = True
 lblFournisseur2.Visible = True
 
 cmdRechercherClient2.Visible = True
 cmdRechercherFRS2.Visible = True
 
 txtTransport.Visible = True
 lbltransport.Visible = True
 
 mskDateLivraison.Visible = True
 lblDateLivraison.Visible = True
 End If
 
 'Si c'est client, fournisseur, conception, programmation, fermeture mécan,
 'fermeture ou fax
5  If ChkClient.Value = vbChecked Or ChkFourn.Value = vbChecked Or _
 ChkConcept.Value = vbChecked Or ChkProg.Value = vbChecked Or _
 ChkFabFermMéca.Value = vbChecked Or ChkFabFerm.Value = vbChecked Or _
 chkFaxFrancais.Value = vbChecked Or chkFaxAnglais.Value = vbChecked Then
 
5  txtNoSoumission.Visible = True
5  lblNoSoumission.Visible = True
5  End If
 
 'Si c'est client, fournisseur, conception, programmation, fermeture mécan,
 'fermeture
5  If ChkClient.Value = vbChecked Or ChkFourn.Value = vbChecked Or _
 ChkConcept.Value = vbChecked Or ChkProg.Value = vbChecked Or _
 ChkFabFermMéca.Value = vbChecked Or ChkFabFerm.Value = vbChecked Then
 
5  txtNomProjet.Visible = True
5  lblNomProjet.Visible = True
 
5  mskDate.Visible = True
60 lblDate.Visible = True
 
  txtProjetClient.Visible = True
  lblProjetClient.Visible = True
  End If
 
 'Si c'est fax
  If chkFaxFrancais.Value = vbChecked Or chkFaxAnglais.Value = vbChecked Then
  txtPage.Visible = True
  lblPage.Visible = True
 
  txtDe.Visible = True
  lblDe.Visible = True
 
  cmdMsg.Visible = True
 
  txtObjet.Visible = True
  lblObjet.Visible = True
6  End If

6  Exit Sub

Oups:

6  wOups "frmreport", "AfficherControles", Err, Err.number, Err.Description
End Sub

Private Sub cmdRechercherClient_Click()

 On Error GoTo Oups

 Dim sRecherche As String
 
 sRecherche = InputBox("Entrez le texte à rechercher.")
 
 Call RemplirComboClient(sRecherche)
 
 m_iNoClient = 0
 
 Call RemplirComboContact
 
 Call cmbclient.SetFocus

 Exit Sub

Oups:

 wOups "frmreport", "cmdRechercherClient_Click", Err, Err.number, Err.Description
End Sub
 
Private Sub cmdRechercherClient2_Click()

 On Error GoTo Oups

 Dim sRecherche As String
 
 sRecherche = InputBox("Entrez le texte à rechercher.")
 
 Call RemplirComboClient2(sRecherche)

 m_iNoClient2 = 0
 
 Call cmbClient2.SetFocus

 Exit Sub

Oups:

 wOups "frmreport", "cmdRechercherClient2_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRechercherFRS_Click()

 On Error GoTo Oups

 Dim sRecherche As String
 
 sRecherche = InputBox("Entrez le texte à rechercher.")
 
 Call RemplirComboFRS(sRecherche)

 m_iNoFRS = 0
 
 Call cmbFournisseur.SetFocus

 Exit Sub

Oups:

 wOups "frmreport", "cmdRechercherFRS_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRechercherFRS2_Click()

 On Error GoTo Oups

 Dim sRecherche As String

 sRecherche = InputBox("Entrez le texte à rechercher.")

 Call RemplirComboFRS2(sRecherche)

 m_iNoFRS2 = 0

 Call cmbFournisseur2.SetFocus

 Exit Sub

Oups:

 wOups "frmreport", "cmdRechercherFRS2_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbclient2_Click()

 On Error GoTo Oups

 'Affiche client2 selectionné dans textbox
 If cmbClient2.ListIndex > -1 Then
 m_iNoClient2 = cmbClient2.ItemData(cmbClient2.ListIndex)
 End If

 Exit Sub

Oups:

 wOups "frmreport", "cmbclient2_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbFournisseur2_Click()

 On Error GoTo Oups

 'Affiche client2 selectionné dans textbox
 If cmbFournisseur2.ListIndex > -1 Then
 m_iNoFRS2 = cmbFournisseur2.ItemData(cmbFournisseur2.ListIndex)
 End If

 Exit Sub

Oups:

 wOups "frmreport", "cmbFournisseur2_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbFournisseur_Click()

 On Error GoTo Oups

 Dim rstFRS As ADODB.Recordset
 'affiche fournisseur selectionné dans textbox
 
 m_sFaxContact = ""
 m_sFaxClientFRS = ""
 
 m_sTelContact = ""
 m_sTelClientFRS = ""
 
 If cmbFournisseur.ListIndex <> -1 Then
 Set rstFRS = New ADODB.Recordset

 Call rstFRS.Open("SELECT Fax, Telephonne FROM GrbFournisseur WHERE IDFRS = " & cmbFournisseur.ItemData(cmbFournisseur.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstFRS.EOF Then
 If Not IsNull(rstFRS.Fields("Fax")) Then
  m_sFaxClientFRS = rstFRS.Fields("Fax")
  Else
  m_sFaxClientFRS = vbNullString
  End If
 
  If Not IsNull(rstFRS.Fields("Telephonne")) Then
  m_sTelClientFRS = rstFRS.Fields("Telephonne")
  Else
  m_sTelClientFRS = vbNullString
 End If
1 End If
 
 'ferme la table
 Call rstFRS.Close
 Set rstFRS = Nothing
End If

If cmbFournisseur.ListIndex > -1 Then
 m_iNoFRS = cmbFournisseur.ItemData(cmbFournisseur.ListIndex)
End If

Call RemplirComboContactFRS

Exit Sub

Oups:

wOups "frmreport", "cmbFournisseur_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbContact_Click()

 On Error GoTo Oups

 Dim rstContact As ADODB.Recordset
 Dim sTampon As String

 If cmbContact.ListIndex <> -1 Then
 m_iNoContact = cmbContact.ItemData(cmbContact.ListIndex)
 End If

 If cmbContact.ListIndex <> -1 Then
 sTampon = cmbContact.ItemData(cmbContact.ListIndex)
 
 Set rstContact = New ADODB.Recordset
 
 Call rstContact.Open("SELECT Telephonne, Fax FROM Grbcontact WHERE IDContact = " & sTampon, g_connData, adOpenDynamic, adLockOptimistic)
 
 'remplis le champ text, telephone et fax
 If Not rstContact.EOF Then
  If IsNull(rstContact.Fields("telephonne")) Then
  m_sTelContact = vbNullString
  Else
  m_sTelContact = rstContact.Fields("telephonne")
  End If
 
  If IsNull(rstContact.Fields("fax")) Then
  m_sFaxContact = vbNullString
  Else
 m_sFaxContact = rstContact.Fields("fax")
End If
 End If
 
 Call rstContact.Close
 Set rstContact = Nothing
End If

Exit Sub

Oups:

wOups "frmreport", "cmbContact_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbContactFRS_Click()

 On Error GoTo Oups

 Dim rstContact As ADODB.Recordset
 
 If cmbContactFRS.ListIndex <> -1 Then
 'Si il y a un client de choisi, le fax sera pour le client
 If cmbContact.ListIndex = -1 Then
 Set rstContact = New ADODB.Recordset

 Call rstContact.Open("SELECT Telephonne, Fax FROM GrbContact WHERE IDContact = " & cmbContactFRS.ItemData(cmbContactFRS.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
 
 'remplis le champ text, telephone et fax
 If Not rstContact.EOF Then
 If Not IsNull(rstContact.Fields("telephonne")) Then
  m_sTelContact = rstContact.Fields("telephonne")
  Else
  m_sTelContact = vbNullString
  End If
 
  If Not IsNull(rstContact.Fields("fax")) Then
  m_sFaxContact = rstContact.Fields("fax")
  Else
  m_sFaxContact = vbNullString
 End If
End If
 
 Call rstContact.Close
 Set rstContact = Nothing
 End If
End If

Exit Sub

Oups:

wOups "frmreport", "cmbContactFRS_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbGrbClick()

 On Error GoTo Oups
 
 'affiche contactgrb selectionné dans textbox
 If cmbGRB.ListIndex > -1 Then
 m_iNoGRB = cmbGRB.ItemData(cmbGRB.ListIndex)
 End If

 Exit Sub

Oups:

 wOups "frmreport", "cmbGrbClick", Err, Err.number, Err.Description
End Sub

Private Sub cmdmsg_Click()

 On Error GoTo Oups

 If txtMsg.Visible = True Then
 txtMsg.Visible = False
 Else
 txtMsg.Visible = True

 Call txtMsg.SetFocus
 End If

 Exit Sub

Oups:

 wOups "frmreport", "cmdmsg_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmreport", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerBonTravail()

 On Error GoTo Oups

 Dim rstBonTravail As ADODB.Recordset
 
 Set rstBonTravail = New ADODB.Recordset
 
 Call rstBonTravail.Open("SELECT * FROM Grbimpression_bonlivraison", g_connData, adOpenDynamic, adLockOptimistic)

 'Set le rapport
 Set DR_BonTravail.DataSource = rstBonTravail

 'Contenu label
 DR_BonTravail.Sections(1).Controls("lblClient").Caption = cmbclient.Text
 DR_BonTravail.Sections(1).Controls("lblContact").Caption = cmbContact.Text
 DR_BonTravail.Sections(1).Controls("lblTelephone").Caption = m_sTelContact
 DR_BonTravail.Sections(1).Controls("lblFax").Caption = m_sFaxContact
 DR_BonTravail.Sections(1).Controls("lblRepresentantGRB").Caption = cmbGRB.Text
 DR_BonTravail.Sections(1).Controls("lblBonTravail").Caption = txtnoprojet.Text
  DR_BonTravail.Sections(1).Controls("lblNoCommandeClient").Caption = txtNoCommande.Text
  DR_BonTravail.Sections(1).Controls("lblDateCommande").Caption = mskDateCommande.Text
  DR_BonTravail.Sections(1).Controls("lblDateHeure").Caption = mskDateTravaux.Text & " " & mskHeureTravaux.Text
 
 'Affiche rapport
  Call DR_BonTravail.Show(vbModal)
 
  Call rstBonTravail.Close
  Set rstBonTravail = Nothing

  Exit Sub

Oups:

  wOups "frmreport", "ImprimerBonTravail", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerBonLivraison()

 On Error GoTo Oups

 Dim rstBonLivraison As ADODB.Recordset
 Dim rstFournisseur As ADODB.Recordset
 Dim rstFournisseur2 As ADODB.Recordset
 Dim rstClient As ADODB.Recordset
 Dim rstClient As ADODB.Recordset
 Dim rstContact As ADODB.Recordset
 Dim sTampon As String

 'ouvre fenetre bon livraison
 'pour entrer qte
 Call OuvrirForm(frmbonlivraison, True)
 
 Screen.MousePointer = vbHourglass
 
 Set rstBonLivraison = New ADODB.Recordset
  Set rstFournisseur = New ADODB.Recordset
  Set rstFournisseur2 = New ADODB.Recordset
  Set rstClient = New ADODB.Recordset
  Set rstClient2 = New ADODB.Recordset
 
  Call rstBonLivraison.Open("SELECT * FROM Grbimpression_bonlivraison WHERE user = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
  Call rstFournisseur.Open("SELECT * FROM GrbFournisseur WHERE IDFRS = " & m_iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)
  Call rstFournisseur2.Open("SELECT * FROM GrbFournisseur WHERE IDFRS = " & m_iNoFRS2, g_connData, adOpenDynamic, adLockOptimistic)
  Call rstClient.Open("SELECT * FROM Grbclient WHERE IDClient = " & m_iNoClient, g_connData, adOpenDynamic, adLockOptimistic)
10 Call rstClient2.Open("SELECT * FROM Grbclient WHERE IDClient = " & m_iNoClient2, g_connData, adOpenDynamic, adLockOptimistic)
 
 'set le rapport
Set DR_BonLivraison.DataSource = rstBonLivraison
 
 'contenu label
DR_BonLivraison.Sections(1).Controls(34).Caption = txtnoprojet.Text
DR_BonLivraison.Sections(1).Controls(35).Caption = ConvertDate(Date)
 
 'si fournisseur
If cmbFournisseur.ListIndex <> -1 Then
 DR_BonLivraison.Sections(1).Controls(36).Caption = cmbFournisseur.Text
 
 If rstFournisseur.EOF Then
 DR_BonLivraison.Sections(1).Controls(37).Caption = vbNullString
 Else
 sTampon = vbNullString
 ''''''''''''''''''''''''''''''''''''''''''''
 'rempli adresse pay ville prov et codepostal si pas vide
 '''''''''''''''''''''''''''''''''''''''''''''
 'adresse
 If Not IsNull(rstFournisseur.Fields("adresse")) Then
 sTampon = rstFournisseur.Fields("adresse")
 End If
 
 DR_BonLivraison.Sections(1).Controls(37).Caption = sTampon
 
 sTampon = vbNullString
 
 'ville
 If Not IsNull(rstFournisseur.Fields("ville")) Then
 sTampon = rstFournisseur.Fields("ville")
 End If
 
 'province
 If Not IsNull(rstFournisseur.Fields("Prov/Etat")) Then
1  If rstFournisseur.Fields("Prov/Etat") <> vbNullString Then
 sTampon = sTampon + ", " + rstFournisseur.Fields("Prov/Etat")
 End If
 End If
 
 DR_BonLivraison.Sections(1).Controls(44).Caption = sTampon
 
 sTampon = vbNullString
 
 'Pays
 If Not IsNull(rstFournisseur.Fields("pays")) Then
 sTampon = rstFournisseur.Fields("pays")
 End If
 
 'codepostal
 If Not IsNull(rstFournisseur.Fields("codepostal")) Then
 If rstFournisseur.Fields("CodePostal") <> vbNullString Then
 sTampon = sTampon + ", " + rstFournisseur.Fields("codepostal")
 End If
 End If
 
 DR_BonLivraison.Sections(1).Controls(45).Caption = sTampon
End If
 
 DR_BonLivraison.Sections(1).Controls(38).Caption = txtNoCommande.Text
DR_BonLivraison.Sections(1).Controls(39).Caption = cmbFournisseur2.Text
 
 If rstFournisseur2.EOF Then
 DR_BonLivraison.Sections(1).Controls(40).Caption = vbNullString
 Else
 sTampon = vbNullString
 ''''''''''''''''''''''''''''''''''''''''''''
 'rempli adresse pay ville prov et codepostal si pas vide
 '''''''''''''''''''''''''''''''''''''''''''''
 'adresse
If Not IsNull(rstFournisseur2.Fields("adresse")) Then
 sTampon = sTampon + rstFournisseur2.Fields("adresse")
 End If
 
 DR_BonLivraison.Sections(1).Controls(40).Caption = sTampon

 sTampon = vbNullString
 
 'ville
 If Not IsNull(rstFournisseur2.Fields("ville")) Then
 sTampon = rstFournisseur2.Fields("ville")
 End If
 
 'prov
 If Not IsNull(rstFournisseur2.Fields("prov/etat")) Then
 If rstFournisseur2.Fields("prov/etat") <> vbNullString Then
 sTampon = sTampon + ", " + rstFournisseur2.Fields("prov/etat")
 End If
 End If
 
 DR_BonLivraison.Sections(1).Controls(46).Caption = sTampon

 sTampon = vbNullString
 
 'pays
 If Not IsNull(rstFournisseur2.Fields("pays")) Then
 sTampon = rstFournisseur2.Fields("pays")
 End If
 
 'codepostal
 If Not IsNull(rstFournisseur2.Fields("codepostal")) Then
 If rstFournisseur2.Fields("CodePostal") <> vbNullString Then
4 sTampon = sTampon + ", " + rstFournisseur2.Fields("codepostal")
4 End If
4 End If
 
4 DR_BonLivraison.Sections(1).Controls(47).Caption = sTampon
4 End If
 
4 DR_BonLivraison.Sections(1).Controls(41).Caption = cmbContactFRS.Text
4 Else
 'si client
4 DR_BonLivraison.Sections(1).Controls(36).Caption = cmbclient.Text
 
4 If rstClient.EOF Then
4 DR_BonLivraison.Sections(1).Controls(37).Caption = vbNullString
4 Else
4  sTampon = vbNullString
 ''''''''''''''''''''''''''''''''''''''''''''
 'rempli adresse pay ville prov et codepostal si pas vide
 '''''''''''''''''''''''''''''''''''''''''''''
 'adresse
4  If Not IsNull(rstClient.Fields("adresseliv")) Then
4  sTampon = sTampon + rstClient.Fields("adresseliv")
4  End If
 
4  DR_BonLivraison.Sections(1).Controls(37).Caption = sTampon

4  sTampon = vbNullString
 
 'ville
4  If Not IsNull(rstClient.Fields("villeliv")) Then
4  sTampon = rstClient.Fields("villeliv")
50 End If
 
 'pays
If Not IsNull(rstClient.Fields("paysliv")) Then
 If rstClient.Fields("PaysLiv") <> vbNullString Then
 sTampon = sTampon + ", " + rstClient.Fields("paysliv")
 End If
 End If
 
 DR_BonLivraison.Sections(1).Controls(44).Caption = sTampon
 sTampon = vbNullString
 
 'province
 If Not IsNull(rstClient.Fields("prov/etatliv")) Then
 sTampon = rstClient.Fields("prov/etatliv")
 End If
 
 'codepostal
 If Not IsNull(rstClient.Fields("codepostalliv")) Then
5  If rstClient.Fields("CodePostalLiv") <> vbNullString Then
5  sTampon = sTampon + ", " + rstClient.Fields("codepostalliv")
5  End If
5  End If
 
5  DR_BonLivraison.Sections(1).Controls(45).Caption = sTampon
5  End If
 
5  DR_BonLivraison.Sections(1).Controls(38).Caption = txtNoCommande.Text
5  DR_BonLivraison.Sections(1).Controls(39).Caption = cmbClient2.Text
 
60 If rstClient2.EOF Then
  DR_BonLivraison.Sections(1).Controls(40).Caption = vbNullString
  Else
  sTampon = vbNullString
 ''''''''''''''''''''''''''''''''''''''''''''
 'rempli adresse pay ville prov et codepostal si pas vide
 '''''''''''''''''''''''''''''''''''''''''''''
 'adresse
  If Not IsNull(rstClient2.Fields("adresseliv")) Then
  sTampon = rstClient2.Fields("adresseliv")
  End If
 
  DR_BonLivraison.Sections(1).Controls(40).Caption = sTampon
  sTampon = vbNullString
 
 'ville
  If Not IsNull(rstClient2.Fields("villeliv")) Then
  sTampon = rstClient2.Fields("villeliv")
  End If
 
 'pays
6  If Not IsNull(rstClient2.Fields("paysliv")) Then
6  If rstClient2.Fields("PaysLiv") <> vbNullString Then
6  sTampon = sTampon + ", " + rstClient2.Fields("paysliv")
6  End If
6  End If
 
6  DR_BonLivraison.Sections(1).Controls(46).Caption = sTampon
6  sTampon = vbNullString
 
 'province
6  If Not IsNull(rstClient2.Fields("prov/etatliv")) Then
70 sTampon = rstClient2.Fields("prov/etatliv")
  End If
 
 'codepostal
  If Not IsNull(rstClient2.Fields("codepostalliv")) Then
  If rstClient2.Fields("CodePostalLiv") <> vbNullString Then
  sTampon = sTampon + ", " + rstClient2.Fields("codepostalliv")
  End If
  End If
 
  DR_BonLivraison.Sections(1).Controls(47).Caption = sTampon
  End If
 
  DR_BonLivraison.Sections(1).Controls(41).Caption = cmbContact.Text
  End If
 
  DR_BonLivraison.Sections(1).Controls(42).Caption = txtTransport.Text
   DR_BonLivraison.Sections(1).Controls(43).Caption = mskDateLivraison.Text
 
 'affiche rapport
   Call DR_BonLivraison.Show(vbModal)
 
7  Call rstBonLivraison.Close
7  Set rstBonLivraison = Nothing
 
7  Call rstFournisseur.Close
7  Set rstFournisseur = Nothing
 
7  Call rstClient.Close
7  Set rstClient = Nothing
 
80 Call rstClient2.Close
80 Set rstClient2 = Nothing

  Exit Sub

Oups:

  wOups "frmreport", "ImprimerBonLivraison", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerClient()

 On Error GoTo Oups

 Dim rstBonLivraison As ADODB.Recordset
 
 Set rstBonLivraison = New ADODB.Recordset
 
 Call rstBonLivraison.Open("SELECT * FROM Grbimpression_bonlivraison", g_connData, adOpenDynamic, adLockOptimistic)
 
 'set le rapport
 Set DR_Client.DataSource = rstBonLivraison
 
 'contenu label
 DR_Client.Sections("Section4").Controls("lblClient").Caption = cmbclient.Text
 DR_Client.Sections("Section4").Controls("lblContact").Caption = cmbContact.Text
 DR_Client.Sections("Section4").Controls("lblTel").Caption = m_sTelContact
 DR_Client.Sections("Section4").Controls("lblFax").Caption = m_sFaxContact
 DR_Client.Sections("Section4").Controls("lblNoSoum").Caption = txtNoSoumission.Text
 DR_Client.Sections("Section4").Controls("lblNoProj").Caption = txtnoprojet.Text
  DR_Client.Sections("Section4").Controls("lblProjetNom").Caption = txtNomProjet.Text
  DR_Client.Sections("Section4").Controls("lblDate").Caption = mskDate.Text
  DR_Client.Sections("Section4").Controls("lblDateOuverture").Caption = ConvertDate(Date)
  DR_Client.Sections("Section4").Controls("lblDateDue").Caption = mskDateDue.Text
  DR_Client.Sections("Section4").Controls("lblProjet").Caption = txtProjetClient.Text
 
 'affiche rapport
  Call DR_Client.Show(vbModal)
 
  Call rstBonLivraison.Close
  Set rstBonLivraison = Nothing

10 Exit Sub

Oups:

wOups "frmreport", "ImprimerClient", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerFournisseur()

 On Error GoTo Oups

 Dim rstBonLivraison As ADODB.Recordset
 
 Set rstBonLivraison = New ADODB.Recordset
 
 Call rstBonLivraison.Open("SELECT * FROM Grbimpression_bonlivraison", g_connData, adOpenDynamic, adLockOptimistic)
 
 'set le rapport
 Set DR_Fournisseur.DataSource = rstBonLivraison
 
 'contenu label
 DR_Fournisseur.Sections("Section4").Controls("lblClient").Caption = cmbclient.Text
 DR_Fournisseur.Sections("Section4").Controls("lblContact").Caption = cmbContact.Text
 DR_Fournisseur.Sections("Section4").Controls("lblTel").Caption = m_sTelContact
 DR_Fournisseur.Sections("Section4").Controls("lblFax").Caption = m_sFaxContact
 DR_Fournisseur.Sections("Section4").Controls("lblNoSoum").Caption = txtNoSoumission.Text
 DR_Fournisseur.Sections("Section4").Controls("lblNoProj").Caption = txtnoprojet.Text
  DR_Fournisseur.Sections("Section4").Controls("lblProjetNom").Caption = txtNomProjet.Text
  DR_Fournisseur.Sections("Section4").Controls("lblDate").Caption = mskDate.Text
  DR_Fournisseur.Sections("Section4").Controls("lblDateOuverture").Caption = ConvertDate(Date)
  DR_Fournisseur.Sections("Section4").Controls("lblDateDue").Caption = mskDateDue.Text
  DR_Fournisseur.Sections("Section4").Controls("lblProjet").Caption = txtProjetClient.Text
 
 'affiche rapport
  Call DR_Fournisseur.Show(vbModal)
 
  Call rstBonLivraison.Close
  Set rstBonLivraison = Nothing

10 Exit Sub

Oups:

wOups "frmreport", "ImprimerFournisseur", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerConception()

 On Error GoTo Oups

 Dim rstBonLivraison As ADODB.Recordset

 Set rstBonLivraison = New ADODB.Recordset

 Call rstBonLivraison.Open("SELECT * FROM Grbimpression_bonlivraison", g_connData, adOpenDynamic, adLockOptimistic)
 
 'set le rapport
 Set DR_Conception.DataSource = rstBonLivraison
 
 'contenu label
 DR_Conception.Sections("Section4").Controls(90).Caption = cmbclient.Text
 DR_Conception.Sections("Section4").Controls(91).Caption = cmbContact.Text
 DR_Conception.Sections("Section4").Controls(92).Caption = m_sTelContact
 DR_Conception.Sections("Section4").Controls(93).Caption = m_sFaxContact
 DR_Conception.Sections("Section4").Controls(94).Caption = txtNoSoumission.Text
 DR_Conception.Sections("Section4").Controls(95).Caption = txtnoprojet.Text
  DR_Conception.Sections("Section4").Controls(96).Caption = txtNomProjet.Text
  DR_Conception.Sections("Section4").Controls(97).Caption = mskDate.Text
  DR_Conception.Sections("Section4").Controls(98).Caption = Trim$(Right$(CStr(Year(Date)), 2) + "-" + CStr(Month(Date)) + "-" + CStr(Day(Date)))
  DR_Conception.Sections("Section4").Controls(99).Caption = txtProjetClient.Text
 
 'affiche rapport
  Call DR_Conception.Show(vbModal)
 
  Call rstBonLivraison.Close
  Set rstBonLivraison = Nothing

  Exit Sub

Oups:

10 wOups "frmreport", "ImprimerConception", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerProgrammation()

 On Error GoTo Oups

 Dim rstBonLivraison As ADODB.Recordset

 Set rstBonLivraison = New ADODB.Recordset

 Call rstBonLivraison.Open("SELECT * FROM Grbimpression_bonlivraison", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Set le rapport
 Set DR_Programmation.DataSource = rstBonLivraison
 
 'Contenu label
 DR_Programmation.Sections("Section4").Controls("lblClient").Caption = cmbclient.Text
 DR_Programmation.Sections("Section4").Controls("lblContact").Caption = cmbContact.Text
 DR_Programmation.Sections("Section4").Controls("lblTelephone").Caption = m_sTelContact
 DR_Programmation.Sections("Section4").Controls("lblFax").Caption = m_sFaxContact
 DR_Programmation.Sections("Section4").Controls("lblNoSoum").Caption = txtNoSoumission.Text
 DR_Programmation.Sections("Section4").Controls("lblNoProj").Caption = txtnoprojet.Text
  DR_Programmation.Sections("Section4").Controls("lblProjetNom").Caption = txtNomProjet.Text
  DR_Programmation.Sections("Section4").Controls("lblDate").Caption = mskDate.Text
  DR_Programmation.Sections("Section4").Controls("lblProjetClient").Caption = txtProjetClient.Text
 
 'Affiche rapport
  Call DR_Programmation.Show(vbModal)

  Call rstBonLivraison.Close
  Set rstBonLivraison = Nothing

  Exit Sub

Oups:

10 wOups "frmreport", "ImprimerProgrammation", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerFermetureMecanique()

 On Error GoTo Oups

 Dim rstBonLivraison As ADODB.Recordset

 Set rstBonLivraison = New ADODB.Recordset

 Call rstBonLivraison.Open("SELECT * FROM Grbimpression_bonlivraison", g_connData, adOpenDynamic, adLockOptimistic)
 
 'set le rapport
 Set DR_FermeMeca.DataSource = rstBonLivraison
 
 'contenu label
 DR_FermeMeca.Sections("Section4").Controls("lblClient").Caption = cmbclient.Text
 DR_FermeMeca.Sections("Section4").Controls("lblContact").Caption = cmbContact.Text
 DR_FermeMeca.Sections("Section4").Controls("lblTel").Caption = m_sTelContact
 DR_FermeMeca.Sections("Section4").Controls("lblFax").Caption = m_sFaxContact
 DR_FermeMeca.Sections("Section4").Controls("lblSoum").Caption = txtNoSoumission.Text
 DR_FermeMeca.Sections("Section4").Controls("lblProj").Caption = txtnoprojet.Text
  DR_FermeMeca.Sections("Section4").Controls("lblProjetNom").Caption = txtNomProjet.Text
  DR_FermeMeca.Sections("Section4").Controls("lblDate").Caption = mskDate.Text
  DR_FermeMeca.Sections("Section4").Controls("lblDateOuverture").Caption = ConvertDate(Date)
  DR_FermeMeca.Sections("Section4").Controls("lblProjet").Caption = txtProjetClient.Text
 
 'affiche rapport
  Call DR_FermeMeca.Show(vbModal)
 
  Call rstBonLivraison.Close
  Set rstBonLivraison = Nothing

  Exit Sub

Oups:

10 wOups "frmreport", "ImprimerFermetureMecanique", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerFermeture()

 On Error GoTo Oups

 Dim rstBonLivraison As ADODB.Recordset
 
 Set rstBonLivraison = New ADODB.Recordset
 
 Call rstBonLivraison.Open("SELECT * FROM Grbimpression_bonlivraison", g_connData, adOpenDynamic, adLockOptimistic)

 'set le rapport
 Set DR_Fermeture.DataSource = rstBonLivraison
 
 'contenu label
 DR_Fermeture.Sections(1).Controls(117).Caption = cmbclient.Text
 DR_Fermeture.Sections(1).Controls(118).Caption = cmbContact.Text
 DR_Fermeture.Sections(1).Controls(119).Caption = m_sTelContact
 DR_Fermeture.Sections(1).Controls(120).Caption = m_sFaxContact
 DR_Fermeture.Sections(1).Controls(121).Caption = txtNoSoumission.Text
 DR_Fermeture.Sections(1).Controls(122).Caption = txtnoprojet.Text
  DR_Fermeture.Sections(1).Controls(123).Caption = txtNomProjet.Text
  DR_Fermeture.Sections(1).Controls(124).Caption = mskDate.Text
  DR_Fermeture.Sections(1).Controls(125).Caption = ConvertDate(Date)
  DR_Fermeture.Sections(1).Controls(126).Caption = txtProjetClient.Text
 
 'affiche rapport
  Call DR_Fermeture.Show(vbModal)
 
  Call rstBonLivraison.Close
  Set rstBonLivraison = Nothing

  Exit Sub

Oups:

10 wOups "frmreport", "ImprimerFermeture", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerFinFabrication()

 On Error GoTo Oups

 Dim rstBonLivraison As ADODB.Recordset
 
 Set rstBonLivraison = New ADODB.Recordset
 
 Call rstBonLivraison.Open("SELECT * FROM Grbimpression_bonlivraison", g_connData, adOpenDynamic, adLockOptimistic)
 
 'set le rapport
 Set DR_FinFab.DataSource = rstBonLivraison
 
 'affiche rapport
 Call DR_FinFab.Show(vbModal)
 
 Call rstBonLivraison.Close
 Set rstBonLivraison = Nothing

 Exit Sub

Oups:

 wOups "frmreport", "ImprimerFinFabrication", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerFax(ByVal eLangue As enumLangueFax)

 On Error GoTo Oups

 Dim drFax As DataReport
 Dim rstBonLivraison As ADODB.Recordset
 Dim bClient As Boolean
 Dim bClientTexte As Boolean
 Dim bClientListIndex As Boolean
 Dim bFournisseur As Boolean
 Dim bFournisseurTexte As Boolean
 Dim bFournisseurListIndex As Boolean
 Dim bContactClient As Boolean
 Dim bContactClientTexte As Boolean
  Dim bContactClientListIndex As Boolean
  Dim bContactFRS As Boolean
  Dim bContactFRSTexte As Boolean
  Dim bContactFRSListIndex As Boolean
  Dim sMessage As String
 
  If eLangue = FAX_ANGLAIS Then
  Set drFax = DR_FaxAnglais
  Else
Set drFax = DR_FaxFrancais
End If
 
If cmbclient.ListIndex <> -1 Or cmbclient.Text <> "" Then
 bClient = True

 If cmbclient.ListIndex <> -1 Then
 bClientListIndex = True
 Else
 bClientTexte = True
 End If
End If
 
If cmbFournisseur.ListIndex <> -1 Or cmbFournisseur.Text <> "" Then
 bFournisseur = True

If cmbFournisseur.ListIndex <> -1 Then
 bFournisseurListIndex = True
 Else
 bFournisseurTexte = True
 End If
End If
 
 If cmbContact.ListIndex <> -1 Or cmbContact.Text <> "" Then
1  bContactClient = True

 If cmbContact.ListIndex <> -1 Then
 bContactClientListIndex = True
 Else
 bContactClientTexte = True
 End If
End If

If cmbContactFRS.ListIndex <> -1 Or cmbContactFRS.Text <> "" Then
 bContactFRS = True

 If cmbContactFRS.ListIndex <> -1 Then
 bContactFRSListIndex = True
 Else
 bContactFRSTexte = True
End If
End If

2  If bClient = False And bFournisseur = False And bContactClient = False And bContactFRS = False Then
 If MsgBox("Voulez-vous choisir un destinataire?", vbYesNo) = vbYes Then
 Exit Sub
 End If
2  End If
 
 'Ce recordset ne sert à rien, il est utilisé uniquement pour le DataSource
 'du DataReport. Un DataReport ne peut être ouvert s'il n'a pas de DataSource
Set rstBonLivraison = New ADODB.Recordset
 
30 Call rstBonLivraison.Open("SELECT * FROM GrbImpression_BonLivraison", g_connData, adOpenDynamic, adLockOptimistic)
 
Set drFax.DataSource = rstBonLivraison
 
 'Contenu label
drFax.Sections(1).Controls("lblDate").Caption = ConvertDate(Date)
 
If bClient = True Then
 drFax.Sections(1).Controls("lblAttention").Caption = cmbContact.Text
Else
 If bFournisseur = True Then
 drFax.Sections(1).Controls("lblAttention").Caption = cmbContactFRS.Text
 Else
 If bContactClient = True Then
 drFax.Sections(1).Controls("lblAttention").Caption = cmbContact.Text
 Else
 If bContactFRS = True Then
 drFax.Sections(1).Controls("lblAttention").Caption = cmbContactFRS.Text
 Else
 drFax.Sections(1).Controls("lblAttention").Caption = ""
 End If
 End If
 End If
 End If
 
40 If bClient = True Then
4 drFax.Sections(1).Controls("lblEntreprise").Caption = cmbclient.Text
4 Else
4 If bFournisseur = True Then
4 drFax.Sections(1).Controls("lblEntreprise").Caption = cmbFournisseur.Text
4 Else
4 drFax.Sections(1).Controls("lblEntreprise").Caption = ""
4 End If
4 End If
 
4 If bClientListIndex = True And bContactClientListIndex = True Then
4 sMessage = "Voulez-vous afficher le numéro de fax du client?" & vbNewLine & _
 "Oui - Fax du client" & vbNewLine & _
 "Non - Fax du contact"
4 Else
4  If bFournisseurListIndex = True And bContactFRSListIndex = True Then
4  sMessage = "Voulez-vous afficher le numéro de fax du fournisseur?" & vbNewLine & _
 "Oui - Fax du fournisseur" & vbNewLine & _
 "Non - Fax du contact"
4  End If
4  End If
 
4  If sMessage = vbNullString Then
4  If bFournisseurListIndex = True Or bClientListIndex = True Then
4  drFax.Sections(1).Controls("lblFax").Caption = m_sFaxClientFRS
4  Else
50 drFax.Sections(1).Controls("lblFax").Caption = m_sFaxContact
5 End If
 Else
 If MsgBox(sMessage, vbYesNo) = vbYes Then
 drFax.Sections(1).Controls("lblFax").Caption = m_sFaxClientFRS
 Else
 drFax.Sections(1).Controls("lblFax").Caption = m_sFaxContact
 End If
 End If
 
 If txtnoprojet.Text <> vbNullString Then
 drFax.Sections(1).Controls("lblNoProjetSoum").Caption = "# Projet:"
 drFax.Sections(1).Controls("lblProjet").Caption = txtnoprojet.Text
5  Else
5  drFax.Sections(1).Controls("lblNoProjetSoum").Caption = "# Soumission:"
5  drFax.Sections(1).Controls("lblProjet").Caption = txtNoSoumission.Text
5  End If
 
5  drFax.Sections(1).Controls("lblPage").Caption = txtPage.Text
5  drFax.Sections(1).Controls("lblDe").Caption = txtDe.Text
5  drFax.Sections(1).Controls("lblMessage").Caption = txtMsg.Text
5  drFax.Sections(1).Controls("lblSujet").Caption = txtObjet.Text
 
 'Affiche rapport
60 drFax.Orientation = rptOrientPortrait
 
60 If eLangue = FAX_ANGLAIS Then
  Call DR_FaxAnglais.Show(vbModal)
  Else
  Call DR_FaxFrancais.Show(vbModal)
  End If

  Call rstBonLivraison.Close
  Set rstBonLivraison = Nothing

  Exit Sub

Oups:

  wOups "frmreport", "ImprimerFax", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerProblemes()
 
 On Error GoTo Oups

 Dim rstBonLivraison As ADODB.Recordset
 
 'Ce recordset ne sert à rien, il est utilisé uniquement pour le DataSource
 'du DataReport. Un DataReport ne peut être ouvert s'il n'a pas de DataSource
 Set rstBonLivraison = New ADODB.Recordset
 
 Call rstBonLivraison.Open("SELECT * FROM GrbImpression_BonLivraison", g_connData, adOpenDynamic, adLockOptimistic)
 
 Set DR_Probleme.DataSource = rstBonLivraison
 
 'Contenu label
 If txtNoSoumission.Text <> "" Then
 DR_Probleme.Sections("Section4").Controls("lblTitreProjSoum").Caption = "# Soum :"
 DR_Probleme.Sections("Section4").Controls("lblNoProjSoum").Caption = txtNoSoumission.Text
 Else
 DR_Probleme.Sections("Section4").Controls("lblTitreProjSoum").Caption = "# Projet :"
 DR_Probleme.Sections("Section4").Controls("lblNoProjSoum").Caption = txtnoprojet.Text
  End If

  DR_Probleme.Sections("Section4").Controls("lblNomEmploye").Caption = cmbGRB.Text
 
 'Affiche rapport
  DR_Probleme.Orientation = rptOrientLandscape
 
  Call DR_Probleme.Show(vbModal)

  Call rstBonLivraison.Close
  Set rstBonLivraison = Nothing

  Exit Sub

Oups:

  wOups "frmreport", "ImprimerProblemes", Err, Err.number, Err.Description
End Sub

Private Sub cmdreport_Click()

 On Error GoTo Oups

 Screen.MousePointer = vbHourglass
 
 'rapport bon de travail
 If ChkBonTravail.Value = vbChecked Then
 Call ImprimerBonTravail
 End If
 
 'rapport bon de livraison
 If chkBonLivraison.Value = vbChecked Then
 Call ImprimerBonLivraison
 End If
 
 'rapport client
 If ChkClient.Value = vbChecked Then
 Call ImprimerClient
 End If
 
 'rapport fournisseur
  If ChkFourn.Value = vbChecked Then
  Call ImprimerFournisseur
  End If
 
 'rapport conception
  If ChkConcept.Value = vbChecked Then
  Call ImprimerConception
  End If
 
 'rapport programmation
  If ChkProg.Value = vbChecked Then
  Call ImprimerProgrammation
10 End If
 
 'rapport fabrication - fermeture mécanique
If ChkFabFermMéca.Value = vbChecked Then
 Call ImprimerFermetureMecanique
End If
 
 'rapport fabrication - fermeture
If ChkFabFerm.Value = vbChecked Then
 Call ImprimerFermeture
End If
 
 'rapport fin fabrication
If ChkFinFab.Value = vbChecked Then
 Call ImprimerFinFabrication
End If
 
 'rapport de problèmes
If chkProblemes.Value = vbChecked Then
 Call ImprimerProblemes
1  End If
 
 'rapport fax francais
If chkFaxFrancais.Value = vbChecked Then
 Call ImprimerFax(FAX_FRANCAIS)
End If
 
 'rapport fax anglais
 If chkFaxAnglais.Value = vbChecked Then
 Call ImprimerFax(FAX_ANGLAIS)
 End If
 
1  Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmreport", "cmdreport_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdselect_Click()

 On Error GoTo Oups

 Dim iValue As Integer
 
 If cmdSelect.Caption = S_SELECT_ALL Then
 iValue = vbChecked
 
 cmdSelect.Caption = S_UNSELECT_ALL
 Else
 iValue = vbUnchecked
 
 cmdSelect.Caption = S_SELECT_ALL
 End If

 m_bSelectAll = True
 
 ChkBonTravail.Value = iValue
  ChkClient.Value = iValue
  ChkConcept.Value = iValue
  ChkFabFerm.Value = iValue
  ChkFabFermMéca.Value = iValue
  ChkFinFab.Value = iValue
  ChkFourn.Value = iValue
  ChkProg.Value = iValue
  chkBonLivraison.Value = iValue
10 chkProblemes.Value = iValue
chkFaxFrancais.Value = iValue
chkFaxAnglais.Value = iValue
 
Call AfficherControles
 
m_bSelectAll = False

Exit Sub

Oups:

wOups "frmreport", "cmdselect_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 m_iNoClient = 0
 m_iNoClient2 = 0
 m_iNoContact = 0
 m_iNoFRS = 0
 m_iNoGRB = 0

 'rempli les combo
 Call RemplirComboClient(vbNullString)
 Call RemplirComboClient2(vbNullString)
 Call RemplirComboContact
 Call RemplirComboGRB
 Call RemplirComboFRS(vbNullString)
  Call RemplirComboFRS2(vbNullString)
 
  Call AfficherControles
 
  Screen.MousePointer = vbDefault

 'rempli
  mskDate.Text = Year(Date) & "-" & Right$("0" & Month(Date), 2) & "-" & Right$("0" & Day(Date), 2)

  Exit Sub

Oups:

  wOups "frmreport", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboFRS(ByVal sRecherche As String)

 On Error GoTo Oups

 Dim rstFournisseur As ADODB.Recordset

 Set rstFournisseur = New ADODB.Recordset

 Call rstFournisseur.Open("SELECT NomFournisseur, IDFRS FROM GrbFournisseur WHERE Instr(1, NomFournisseur, '" & Replace(sRecherche, "'", "''") & "') > 0 AND Supprimé = False ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)

 'vide combo
 Call cmbFournisseur.Clear
 
 'rempli les combo tant que pas fin d'enregistrement
 Do While Not rstFournisseur.EOF
 Call cmbFournisseur.AddItem(rstFournisseur.Fields("NomFournisseur"))
 cmbFournisseur.ItemData(cmbFournisseur.newIndex) = rstFournisseur.Fields("IDFRS")

 
 Call rstFournisseur.MoveNext
 Loop

 Call rstFournisseur.Close
  Set rstFournisseur = Nothing
 
  Exit Sub

Oups:

  wOups "frmreport", "RemplirComboFRS", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboFRS2(ByVal sRecherche As String)

 On Error GoTo Oups

 Dim rstFournisseur As ADODB.Recordset

 Set rstFournisseur = New ADODB.Recordset

 Call rstFournisseur.Open("SELECT NomFournisseur, IDFRS FROM GrbFournisseur WHERE Instr(1, NomFournisseur, '" & Replace(sRecherche, "'", "''") & "') > 0 AND Supprimé = False ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)

 'vide combo
 Call cmbFournisseur2.Clear
 
 'rempli les combo tant que pas fin d'enregistrement
 Do While Not rstFournisseur.EOF
 Call cmbFournisseur2.AddItem(rstFournisseur.Fields("NomFournisseur"))
 cmbFournisseur2.ItemData(cmbFournisseur2.newIndex) = rstFournisseur.Fields("IDFRS")

 
 Call rstFournisseur.MoveNext
 Loop

 Call rstFournisseur.Close
  Set rstFournisseur = Nothing
 
  Exit Sub

Oups:

  wOups "frmreport", "RemplirComboFRS2", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboClient(ByVal sRecherche As String)

 On Error GoTo Oups

 Dim rstClient As ADODB.Recordset
 
 'set les tables
 Set rstClient = New ADODB.Recordset
 
 Call rstClient.Open("SELECT NomClient, IDClient FROM Grbclient WHERE Instr(1, NomClient, '" & Replace(sRecherche, "'", "''") & "') > 0 AND Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)

 'vide combo
 Call cmbclient.Clear
 
 'rempli les combo tant que pas fin d'enregistrement
 Do While Not rstClient.EOF
 Call cmbclient.AddItem(rstClient.Fields("nomclient"))
 cmbclient.ItemData(cmbclient.newIndex) = rstClient.Fields("idclient")
 
 Call rstClient.MoveNext
 Loop
 
 Call rstClient.Close
  Set rstClient = Nothing

  Exit Sub

Oups:

  wOups "frmreport", "RemplirComboClient", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboClient2(ByVal sRecherche As String)

 On Error GoTo Oups

 Dim rstClient As ADODB.Recordset
 
 Set rstClient = New ADODB.Recordset
 
 Call rstClient.Open("SELECT NomClient, IDClient FROM Grbclient WHERE Instr(1, NomClient, '" & Replace(sRecherche, "'", "''") & "') > 0 AND Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)

 'vide combo
 Call cmbClient2.Clear
 
 'rempli les combo tant que pas fin d'enregistrement
 Do While Not rstClient.EOF
 Call cmbClient2.AddItem(rstClient.Fields("nomclient"))
 cmbClient2.ItemData(cmbClient2.newIndex) = rstClient.Fields("idclient")
 
 
 Call rstClient.MoveNext
 Loop
 
 Call rstClient.Close
  Set rstClient = Nothing

  Exit Sub

Oups:

  wOups "frmreport", "RemplirComboClient2", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboContact()

 On Error GoTo Oups

 Dim rstContact As ADODB.Recordset

 'si client de selectionné, remplis les liste contact pour le client
 'sinon met tout les contact
 Set rstContact = New ADODB.Recordset
 
 If m_iNoClient > 0 Then
 Call rstContact.Open("SELECT GrbContact.IDContact, GrbContact.NomContact, GrbContactClient.NoClient FROM GrbContact INNER JOIN GrbContactClient ON GrbContact.IDContact = GrbContactClient.NoContact WHERE CStr(GrbContactClient.noclient) = CStr('" & m_iNoClient & "') ORDER BY Grbcontact.NomContact", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstContact.Open("SELECT NomContact, IDContact FROM GrbContact WHERE Supprimé = False ORDER BY Nomcontact", g_connData, adOpenDynamic, adLockOptimistic)
 End If

 'vide combo
 Call cmbContact.Clear

 'rempli les combo tant que pas fin d'enregistrement
 Do While Not rstContact.EOF
 'si trouve le text dans le nom du contact, ajoute dans combo
 If Not IsNull(rstContact.Fields("NomContact")) Then
  Call cmbContact.AddItem(rstContact.Fields("NomContact"))
  cmbContact.ItemData(cmbContact.newIndex) = rstContact.Fields("IDContact")
 
  Call rstContact.MoveNext
  End If
  Loop

  Call rstContact.Close
  Set rstContact = Nothing

  Exit Sub

Oups:

10 wOups "frmreport", "RemplirComboContact", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboContactFRS()

 On Error GoTo Oups

 Dim rstContactFRS As ADODB.Recordset
 Dim rstContact As ADODB.Recordset
 
 'si fournisseur de selectionné, remplis les liste contact pour le client
 'sinon met tout les contact
 If m_iNoFRS > 0 Then
 Set rstContactFRS = New ADODB.Recordset

 Call rstContactFRS.Open("SELECT * FROM GrbContactFRS WHERE NoFRS = " & m_iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Exit Sub
 End If

 'vide combo
 Call cmbContactFRS.Clear
 
 'rempli les combo tant que pas fin d'enregistrement
 Set rstContact = New ADODB.Recordset
 
  Do While Not rstContactFRS.EOF
  Call rstContact.Open("SELECT NomContact, IDContact FROM GrbContact WHERE IDContact = " & rstContactFRS.Fields("NoContact"), g_connData, adOpenDynamic, adLockOptimistic)
 
  If Not rstContact.EOF Then
  Call cmbContactFRS.AddItem(rstContact.Fields("NomContact"))
  cmbContactFRS.ItemData(cmbContactFRS.newIndex) = rstContact.Fields("IDContact")
  End If
 
  Call rstContact.Close

  Call rstContactFRS.MoveNext
10 Loop
 
Set rstContact = Nothing
 
Call rstContactFRS.Close
Set rstContactFRS = Nothing

Exit Sub

Oups:

wOups "frmreport", "RemplirComboContactFRS", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboGRB()

 On Error GoTo Oups

 Dim rstContactGRB As ADODB.Recordset

 Set rstContactGRB = New ADODB.Recordset
 
 Call rstContactGRB.Open("SELECT employe, noEmploye FROM Grbemployés WHERE Actif = true ORDER BY employe", g_connData, adOpenDynamic, adLockOptimistic)

 'vide combo
 Call cmbGRB.Clear

 'rempli les combo tant que pas fin d'enregistrement
 Do While Not rstContactGRB.EOF
 Call cmbGRB.AddItem(rstContactGRB.Fields("Employe"))
 cmbGRB.ItemData(cmbGRB.newIndex) = rstContactGRB.Fields("noEmploye")
 
 Call rstContactGRB.MoveNext
 Loop
 
 Call rstContactGRB.Close
  Set rstContactGRB = Nothing

  Exit Sub

Oups:

  wOups "frmreport", "RemplirComboGRB", Err, Err.number, Err.Description
End Sub

Private Sub mskDate_GotFocus()

 On Error GoTo Oups

 If Len(mskDate.Text) = 10 Then
 mskDate.Text = Right$(mskDate.Text, 8)
 End If
 
 mskDate.mask = "##-##-##"

 Exit Sub

Oups:

 wOups "frmreport", "mskDate_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDate_LostFocus()

 On Error GoTo Oups

 mskDate.mask = vbNullString
 
 If mskDate.Text = "__-__-__" Then
 mskDate.Text = vbNullString
 Else
 If Len(mskDate.Text) =   Then
 If IsDate(mskDate.Text) Then
 mskDate.Text = Year(DateSerial(Left$(mskDate.Text, 2), Mid$(mskDate.Text, 4, 2), Right$(mskDate.Text, 2))) & Mid$(mskDate.Text, 3, 8)
 End If
 End If
 End If

  Exit Sub

Oups:

  wOups "frmreport", "mskDate_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDateDue_GotFocus()

 On Error GoTo Oups

 If Len(mskDateDue.Text) = 10 Then
 mskDateDue.Text = Right$(mskDateDue.Text, 8)
 End If
 
 mskDateDue.mask = "##-##-##"

 Exit Sub

Oups:

 wOups "frmreport", "mskDateDue_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDateDue_LostFocus()

 On Error GoTo Oups

 mskDateDue.mask = vbNullString
 
 If mskDateDue.Text = "__-__-__" Then
 mskDateDue.Text = vbNullString
 Else
 If Len(mskDateDue.Text) =   Then
 If IsDate(mskDateDue.Text) Then
 mskDateDue.Text = Year(DateSerial(Left$(mskDateDue.Text, 2), Mid$(mskDateDue.Text, 4, 2), Right$(mskDateDue.Text, 2))) & Mid$(mskDateDue.Text, 3, 8)
 End If
 End If
 End If

  Exit Sub

Oups:

  wOups "frmreport", "mskDateDue_LostFocus", Err, Err.number, Err.Description
End Sub


Private Sub mskdatecommande_GotFocus()

 On Error GoTo Oups

 If Len(mskDateCommande.Text) = 10 Then
 mskDateCommande.Text = Right$(mskDateCommande.Text, 8)
 End If
 
 mskDateCommande.mask = "##-##-##"

 Exit Sub

Oups:

 wOups "frmreport", "mskdatecommande_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskdatecommande_LostFocus()

 On Error GoTo Oups

 mskDateCommande.mask = vbNullString
 
 If mskDateCommande.Text = "__-__-__" Then
 mskDateCommande.Text = vbNullString
 Else
 If Len(mskDateCommande.Text) =   Then
 If IsDate(mskDateCommande.Text) Then
 mskDateCommande.Text = Year(DateSerial(Left$(mskDateCommande.Text, 2), Mid$(mskDateCommande.Text, 4, 2), Right$(mskDateCommande, 2))) & Mid$(mskDateCommande.Text, 3, 8)
 End If
 End If
 End If

  Exit Sub

Oups:

  wOups "frmreport", "mskdatecommande_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskdatelivraison_GotFocus()

 On Error GoTo Oups

 If Len(mskDateLivraison.Text) = 10 Then
 mskDateLivraison.Text = Right$(mskDateCommande.Text, 8)
 End If
 
 mskDateLivraison.mask = "##-##-##"

 Exit Sub

Oups:

 wOups "frmreport", "mskdatelivraison_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskdatelivraison_LostFocus()

 On Error GoTo Oups

 mskDateLivraison.mask = vbNullString
 
 If mskDateLivraison.Text = "__-__-__" Then
 mskDateLivraison.Text = vbNullString
 Else
 If Len(mskDateLivraison.Text) =   Then
 If IsDate(mskDateLivraison.Text) Then
 mskDateLivraison.Text = Year(DateSerial(Left$(mskDateLivraison.Text, 2), Mid$(mskDateLivraison.Text, 4, 2), Right$(mskDateLivraison.Text, 2))) & Mid$(mskDateLivraison.Text, 3, 8)
 End If
 End If
 End If

  Exit Sub

Oups:

  wOups "frmreport", "mskdatelivraison_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDateTravaux_GotFocus()

 On Error GoTo Oups

 If Len(mskDateTravaux.Text) = 10 Then
 mskDateTravaux.Text = Right$(mskDateTravaux.Text, 8)
 End If
 
 mskDateTravaux.mask = "##-##-##"

 Exit Sub

Oups:

 wOups "frmreport", "mskDateTravaux_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDateTravaux_LostFocus()

 On Error GoTo Oups

 mskDateTravaux.mask = vbNullString
 
 If mskDateTravaux.Text = "__-__-__" Then
 mskDateTravaux.Text = vbNullString
 Else
 If Len(mskDateTravaux.Text) =   Then
 If IsDate(mskDateTravaux.Text) Then
 mskDateTravaux.Text = Year(DateSerial(Left$(mskDateTravaux.Text, 2), Mid$(mskDateTravaux.Text, 4, 2), Right$(mskDateTravaux, 2))) & Mid$(mskDateTravaux.Text, 3, 8)
 End If
 End If
 End If

  Exit Sub

Oups:

  wOups "frmreport", "mskDateTravaux_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskHeureTravaux_GotFocus()

 On Error GoTo Oups

 mskHeureTravaux.mask = "##:##"

 Exit Sub

Oups:

 wOups "frmreport", "mskHeureTravaux_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskHeureTravaux_LostFocus()

 On Error GoTo Oups

 mskHeureTravaux.mask = vbNullString
 
 If mskHeureTravaux.Text = "__:__" Then
 mskHeureTravaux.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmreport", "mskHeureTravaux_LostFocus", Err, Err.number, Err.Description
End Sub
