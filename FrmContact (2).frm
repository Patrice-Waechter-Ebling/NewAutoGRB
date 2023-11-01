VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmContact 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contacts"
   ClientHeight    =   7545
   ClientLeft      =   3330
   ClientTop       =   2670
   ClientWidth     =   9540
   Icon            =   "FrmContact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   9540
   Begin VB.CommandButton cmdMailList 
      Caption         =   "Ajouter au mailing list"
      Height          =   495
      Left            =   5520
      TabIndex        =   49
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Frame fraEtatOutlook 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   480
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   8295
      Begin VB.Label lblEtatOutlook 
         Alignment       =   2  'Center
         Caption         =   "Recherche du client dans Outlook ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   6495
      End
   End
   Begin VB.TextBox txtNomContact 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Compagnie"
      DataSource      =   "datContact"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1320
      Width           =   6135
   End
   Begin VB.ComboBox cmbContact 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1320
      Width           =   6375
   End
   Begin VB.TextBox txtTitre 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Compagnie"
      DataSource      =   "datContact"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2400
      Width           =   5895
   End
   Begin VB.TextBox txtCommentaire 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Compagnie"
      DataSource      =   "datContact"
      Height          =   645
      Left            =   1440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   33
      Top             =   5280
      Width           =   5895
   End
   Begin VB.CommandButton cmdFax 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Envoyer Fax"
      Height          =   495
      Left            =   7320
      TabIndex        =   47
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopier 
      Caption         =   "&Copier"
      Height          =   495
      Left            =   1440
      TabIndex        =   41
      Top             =   6960
      Width           =   975
   End
   Begin VB.TextBox txtTelDomicile 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Teldomicile"
      DataSource      =   "datContact"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   3480
      Width           =   5895
   End
   Begin VB.CommandButton cmdRafraichir 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rafraîchir"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdRechercher 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rechercher"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdreport 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Impression"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   40
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox txtPoste 
      BackColor       =   &H00FFFFFF&
      DataField       =   "noposte"
      DataSource      =   "datContact"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3120
      Width           =   5895
   End
   Begin VB.TextBox txtRechercher 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton CmdModif 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Modifier"
      Height          =   495
      Left            =   4560
      TabIndex        =   46
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H00FFFFFF&
      DataField       =   "E-mail"
      DataSource      =   "datContact"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   4920
      Width           =   5895
   End
   Begin VB.TextBox txtPagette 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Pagette"
      DataSource      =   "datContact"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   4560
      Width           =   5895
   End
   Begin VB.TextBox txtFax 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Fax"
      DataSource      =   "datContact"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   4200
      Width           =   5895
   End
   Begin VB.TextBox txtCellulaire 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Cellulaire"
      DataSource      =   "datContact"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   3840
      Width           =   5895
   End
   Begin VB.TextBox txtTelephone 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Telephonne"
      DataSource      =   "datContact"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2760
      Width           =   5895
   End
   Begin VB.TextBox txtCompagnie 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Compagnie"
      DataSource      =   "datContact"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2040
      Width           =   5895
   End
   Begin VB.CommandButton CmdSupp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Supprimer"
      Height          =   495
      Left            =   3480
      TabIndex        =   45
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton CmdAdd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Ajouter"
      Height          =   495
      Left            =   2520
      TabIndex        =   44
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Fermer"
      Height          =   495
      Left            =   8520
      TabIndex        =   48
      Top             =   6960
      Width           =   975
   End
   Begin MSMask.MaskEdBox mskTelephone 
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskCellulaire 
      Height          =   285
      Left            =   1440
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskFax 
      Height          =   285
      Left            =   1440
      TabIndex        =   24
      Top             =   4200
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskPagette 
      Height          =   285
      Left            =   1440
      TabIndex        =   27
      Top             =   4560
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskTelDomicile 
      Height          =   285
      Left            =   1440
      TabIndex        =   19
      Top             =   3480
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton CmdAnul 
      Caption         =   "A&nnuler"
      Height          =   495
      Left            =   2520
      TabIndex        =   43
      Top             =   6960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CmdEnr 
      Caption         =   "&Enregistrer"
      Height          =   495
      Left            =   1440
      TabIndex        =   42
      Top             =   6960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Titre :"
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
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Commentaire :"
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
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   32
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label lblUserModification 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Par :"
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
      Height          =   255
      Left            =   2520
      TabIndex        =   39
      Top             =   6375
      Width           =   1335
   End
   Begin VB.Label lblUserCreation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Par :"
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
      Height          =   255
      Left            =   2520
      TabIndex        =   36
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label lblDateModification 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2004-09-14"
      Height          =   285
      Left            =   1440
      TabIndex        =   38
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label lblDateCreation 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2004-09-14"
      Height          =   285
      Left            =   1440
      TabIndex        =   35
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Modification :"
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
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   37
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Création :"
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
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   34
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Domicile :"
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
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   30
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Poste :"
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
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   16
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblRechercher 
      BackStyle       =   0  'Transparent
      Caption         =   "Rechercher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact :"
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
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail :"
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
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   29
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pagette :"
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
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   26
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax :"
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
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   23
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cellulaire :"
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
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   20
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Téléphone :"
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
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Compagnie :"
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
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   4680
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "FrmContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enumMode
 MODE_AJOUT_MODIF = 0
 MODE_INACTIF = 1
End Enum

Private m_bModeAjout As Boolean
Private m_bRenommer As Boolean
Private m_iNoContact As Integer

Public m_bAnnulerDistList As Boolean
'Public m_otlDistList As Outlook.DistListItem

Private Sub RemplirComboContact()

 On Error GoTo Oups

 'Rempli le combo des contacts
 Dim rstContact As ADODB.Recordset
 
 'Ouverture de la table
 Set rstContact = New ADODB.Recordset
 
 Call rstContact.Open("SELECT NomContact, Compagnie, IDContact FROM GrbContact WHERE Supprimé = False", g_connData, adOpenDynamic, adLockOptimistic)

 'Il faut vider le combo avant de le remplir
 Call cmbContact.Clear

 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstContact.EOF
 'Ajout du nom du contact dans le combo
 Call cmbContact.AddItem(rstContact.Fields("NomContact") & " - " & rstContact.Fields("Compagnie"))
 
 'Ajout du numéro du contact dans l'itemData du combo
 cmbContact.ItemData(cmbContact.newIndex) = rstContact.Fields("IDContact")
 
 Call rstContact.MoveNext
 Loop

 Call rstContact.Close
  Set rstContact = Nothing
 
  If cmbContact.ListCount > 0 Then
  cmbContact.ListIndex = 0
  End If

  Exit Sub

Oups:

  wOups "frmContact", "RemplirComboContact", Err, Err.number, Err.Description
End Sub

Private Sub EnregistrerContact()

 On Error GoTo Oups

 Dim rstContact As ADODB.Recordset

 Set rstContact = New ADODB.Recordset
 
 If m_bModeAjout = True Then
 Call rstContact.Open("SELECT * FROM GrbContact", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call rstContact.AddNew

 rstContact.Fields("DateCréation") = ConvertDate(Date)
 rstContact.Fields("UserCréation") = g_sInitiale
 Else
 Call rstContact.Open("SELECT * FROM Grbcontact WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)

 rstContact.Fields("DateModification") = ConvertDate(Date)
  rstContact.Fields("UserModification") = g_sInitiale
  End If
 
 'Enregistrement du contact
  rstContact.Fields("NomContact") = txtNomContact.Text
  rstContact.Fields("Compagnie") = txtCompagnie.Text
  rstContact.Fields("Titre") = txtTitre.Text
  rstContact.Fields("Telephonne") = mskTelephone.Text
  rstContact.Fields("Fax") = mskFax.Text
  rstContact.Fields("Pagette") = mskPagette.Text
10 rstContact.Fields("Cellulaire") = mskCellulaire.Text
rstContact.Fields("E-mail") = txtEmail.Text
rstContact.Fields("NoPoste") = txtPoste.Text
rstContact.Fields("TelDomicile") = mskTelDomicile.Text
rstContact.Fields("Commentaire") = txtcommentaire.Text

rstContact.Fields("EntryIDOutlook") = ModifierContactExchange(rstContact.Fields("IDContact"))

If m_bModeAjout = True Then
 m_bModeAjout = False
End If

Call rstContact.Update
 
Call rstContact.Close
Set rstContact = Nothing

1  Exit Sub

Oups:

wOups "frmContact", "EnregistrerContact", Err, Err.number, Err.Description
End Sub

Private Function ModifierContactExchange(ByVal iContactID As Integer) As String

 On Error GoTo Oups

 Dim otlApp As Outlook.Application
 Dim otlContact As Outlook.ContactItem
 Dim folContact As Outlook.MAPIFolder
 Dim bDejaOuvert As Boolean
 Dim sNom() As String

 If m_bModeAjout = True Then
 lblEtatOutlook.Caption = "Ajout du contact dans Outlook ..."
 Else
 lblEtatOutlook.Caption = "Modification du contact dans Outlook ..."
 End If

  fraEtatOutlook.Visible = True

  Set otlApp = OuvrirOutlook(bDejaOuvert)

  Set folContact = GetFolder(otlApp, "Contacts GRB")

  If m_bModeAjout = True Then
  Set otlContact = folContact.Items.Add(olContactItem)

  otlContact.User1 = iContactID
  Else
  Set otlContact = folContact.Items.Find("[User1] = " & iContactID)
10 End If

If otlContact Is Nothing Then
 Call MsgBox("Le contact " & txtNomContact.Text & " n'a pas été trouvé pour la mise à jour Exchange!", vbOKOnly, "Erreur")

 fraEtatOutlook.Visible = False

 DoEvents

 Exit Function
End If

sNom = Split(Trim$(txtNomContact.Text), " ")

Select Case UBound(sNom)
 Case 0:
 otlContact.FirstName = sNom(0)
 
 Case 1:
 otlContact.FirstName = sNom(0)
 otlContact.LastName = sNom(1)

 Case 2
 otlContact.FirstName = sNom(0)
 otlContact.MiddleName = sNom(1)
 otlContact.LastName = sNom(2)
End Select
 
 otlContact.Title = ""

otlContact.CompanyName = txtCompagnie.Text
 otlContact.JobTitle = txtTitre.Text

1  If mskTelephone.Text <> "(___) ___-____" Then
 If Trim$(txtPoste.Text) <> "" Then
 otlContact.BusinessTelephoneNumber = mskTelephone.Text & " Ext : " & txtPoste.Text
 Else
 otlContact.BusinessTelephoneNumber = mskTelephone.Text
 End If
End If

If mskFax.Text <> "(___) ___-____" Then
 otlContact.BusinessFaxNumber = mskFax.Text
End If

If mskCellulaire.Text <> "(___) ___-____" Then
 otlContact.MobileTelephoneNumber = mskCellulaire.Text
End If

2  If mskPagette.Text <> "(___) ___-____" Then
 otlContact.PagerNumber = mskPagette.Text
2  End If

otlContact.Email1Address = txtEmail.Text

2  If mskTelDomicile.Text <> "(___) ___-____" Then
 otlContact.HomeTelephoneNumber = mskTelDomicile.Text
2  End If

If txtcommentaire.Text <> "" Then
otlContact.Body = txtcommentaire.Text
End If

Call otlContact.Save

ModifierContactExchange = otlContact.EntryID

If bDejaOuvert = False Then
 Call otlApp.Quit
End If

Set otlApp = Nothing

fraEtatOutlook.Visible = False

DoEvents

Exit Function

Oups:

woups"frmContact", "ModifierContactExchange", Err, Erl, "iContactID = " & iContactID)

3  fraEtatOutlook.Visible = False
End Function

Private Sub SupprimerContactExchange(ByVal iContactID As Integer)

 On Error GoTo Oups

 Dim otlApp As Outlook.Application
 Dim otlContact As Outlook.ContactItem
 Dim folContact As MAPIFolder
 Dim bDejaOuvert As Boolean

 lblEtatOutlook.Caption = "Suppression du contact dans Outlook ..."
 fraEtatOutlook.Visible = True

 Set otlApp = OuvrirOutlook(bDejaOuvert)

 Set folContact = GetFolder(otlApp, "Contacts GRB")

 Set otlContact = folContact.Items.Find("[User1] = " & iContactID)

 If Not otlContact Is Nothing Then
  Call otlContact.Delete
  End If

  If bDejaOuvert = False Then
  Call otlApp.Quit
  End If

  Set otlApp = Nothing

  fraEtatOutlook.Visible = False

  DoEvents

10 Exit Sub

Oups:

wOups "frmContact", "SupprimerContactExchange", Err, Err.number, Err.Description

fraEtatOutlook.Visible = False
End Sub

Private Sub cmdCopier_Click()

 On Error GoTo Oups

 'Copie un contact
 Dim sName As String
 Dim bAjouter As Boolean
 
 'On procede a la saisie du nom et du contact
 sName = InputBox("Veuillez entrer le nom du contact", "SAISIE DU NOM", "Nom du contact")
 
 If sName <> vbNullString Then
 If ExisteDansBD(sName) = True Then
 If MsgBox("Il y a déjà un contact portant ce nom!" & vbNewLine & "Voulez-vous l'ajouter quand même?", vbYesNo) = vbYes Then
 bAjouter = True
 Else
 bAjouter = False
 End If
  Else
  If ContientCaracteresIncorrects(sName) = True Then
  Call MsgBox("Il y a des caractères non permis!", vbOKOnly, "Erreur")

  bAjouter = False
  Else
  bAjouter = True
  End If
  End If
10 Else
1 bAjouter = False
End If

If bAjouter = True Then
 Screen.MousePointer = vbHourglass
 
 'On montre seulement les boutton pour enregistrer
 Call AfficherControles(MODE_AJOUT_MODIF)
 
 'On montre les maskEdBox
 Call HideEdMask(False)
 
 m_bModeAjout = True
 
 'On affiche le nom du nouveau client dans le textbox
 'pour éviter le ScrollDown durant l'ajout
 
 txtNomContact.Text = sName
 
 Call ViderBarrerChamps(False, False)
 
 Call txtCompagnie.SetFocus
 
 Screen.MousePointer = vbDefault
1  End If

Exit Sub

Oups:

 wOups "frmContact", "cmdCopier_Click", Err, Err.number, Err.Description
End Sub

Private Function ExisteDansBD(ByVal sName As String) As Boolean
 
 On Error GoTo Oups

 Dim rstContact As ADODB.Recordset

 Set rstContact = New ADODB.Recordset

 Call rstContact.Open("SELECT NomContact FROM GrbContact WHERE NomContact = '" & Replace(sName, "'", "''") & "' AND Supprimé = False", g_connData, adOpenForwardOnly, adLockReadOnly)

 If rstContact.EOF Then
 ExisteDansBD = False
 Else
 ExisteDansBD = True
 End If

 Call rstContact.Close
 Set rstContact = Nothing

  Exit Function

Oups:

  wOups "frmContact", "ExisteDansBD", Err, Err.number, Err.Description
End Function

Private Function ContientCaracteresIncorrects(ByVal sName As String) As Boolean

 On Error GoTo Oups

 If InStr(1, sName, ",") > 0 Or InStr(1, sName, ";") > 0 Or InStr(1, sName, ":") > 0 Or InStr(1, sName, "(") > 0 Or InStr(1, sName, ")") > 0 Then
 ContientCaracteresIncorrects = True
 Else
 ContientCaracteresIncorrects = False
 End If

 Exit Function

Oups:

 wOups "frmContact", "ContientCaracteresIncorrects", Err, Err.number, Err.Description
End Function

Private Sub cmdFax_Click()

 On Error GoTo Oups

 If cmbContact.ListCount > 0 Then
 Call frmreport.Afficher(0, m_iNoContact, FRM_CONTACTS)
 End If

 Exit Sub

Oups:

 wOups "frmContact", "cmdFax_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdMailList_Click()

 On Error GoTo Oups

 Dim otlApp As Outlook.Application
 Dim folContact As Outlook.MAPIFolder
 Dim itmContact As Outlook.ContactItem
 Dim otlRecipient As Outlook.Recipient
 Dim bDejaOuvert As Boolean
 Dim rstContact As ADODB.Recordset
 Dim sIDContact As String
 Dim sContact As String

 If cmbContact.ListIndex <> -1 Then
 If Trim$(txtEmail.Text) <> "" Then
  sIDContact = cmbContact.ItemData(cmbContact.ListIndex)
  sContact = cmbContact.Text
  End If
 
  If sIDContact <> "" Then
  Set otlApp = OuvrirOutlook(bDejaOuvert)

  lblEtatOutlook.Caption = "Recherche des listes de distribution..."

  fraEtatOutlook.Visible = True

  Call frmChoixMailList.Afficher(Me, otlApp)

 If m_bAnnulerDistList = False Then
 lblEtatOutlook.Caption = "Ajout du contact dans la liste de distribution..."
 
 fraEtatOutlook.Visible = True

 Set folContact = GetFolder(otlApp, "Contacts GRB")

 Set itmContact = folContact.Items.Find("[User1] = " & sIDContact)

 If Not itmContact Is Nothing Then
 Set otlRecipient = otlApp.Session.CreateRecipient(itmContact.Email1DisplayName)

 If otlRecipient.Resolve = True Then
 Call m_otlDistList.AddMember(otlRecipient)
 
 Call m_otlDistList.Save
 Else
 Call MsgBox("Impossible de trouver le contact '" & sContact & "' !", vbOKOnly, "Erreur")
 End If
 Else
 Call MsgBox("Contact '" & sContact & "' introuvable!", vbOKOnly, "Erreur")
 End If
 End If

 If bDejaOuvert = False Then
 Call otlApp.Quit
1  End If

 Set otlApp = Nothing

 fraEtatOutlook.Visible = False
 Else
 Call MsgBox("Le ou les contacts n'ont pas d'email!", vbOKOnly, "Erreur")
 End If
Else
 Call MsgBox("Aucun contact sélectionné!", vbOKOnly, "Erreur")
End If

Exit Sub

Oups:

If Err.number = 2 And Erl = 145 Then
 Call MsgBox("La liste de distribution est pleine!", vbOKOnly, "Erreur")
Else
wOups "frmContact", "cmdMailList_Click", Err, Err.number, Err.Description
End If

2  fraEtatOutlook.Visible = False
End Sub

Private Sub cmdRafraichir_Click()

 On Error GoTo Oups

 'Rempli la liste avec tous les contacts après avoir fait une recherche
 Screen.MousePointer = vbHourglass

 Call RemplirComboContact

 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmContact", "cmdRafraichir_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdreport_Click()

 On Error GoTo Oups

 'Impression de la liste des contacts
 Dim rstContact As ADODB.Recordset

 Set rstContact = New ADODB.Recordset

 If MsgBox("Voulez-vous imprimer ce contact uniquement?", vbYesNo) = vbYes Then
 Call rstContact.Open("SELECT * FROM GrbContact WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)
 Else
 If MsgBox("Voulez-vous filtrer par la compagnie '" & txtCompagnie.Text & "'?", vbYesNo) = vbYes Then
 Call rstContact.Open("SELECT * FROM GrbContact WHERE Compagnie = '" & Replace(txtCompagnie.Text, "'", "''") & "' AND Supprimé = False ORDER BY NomContact", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstContact.Open("SELECT * FROM GrbContact WHERE Supprimé = False ORDER BY NomContact", g_connData, adOpenDynamic, adLockOptimistic)
 End If
  End If

  Screen.MousePointer = vbHourglass
 
 'set le rapport
  Set DR_ListeContact.DataSource = rstContact
 
  DR_ListeContact.Orientation = rptOrientPortrait

  Call DR_ListeContact.Show(vbModal)
 
  Call rstContact.Close
  Set rstContact = Nothing
 
  Screen.MousePointer = vbDefault

10 Exit Sub

Oups:

wOups "frmContact", "cmdreport_Click", Err, Err.number, Err.Description
End Sub

Private Sub AfficherControles(ByVal eMode As enumMode)

 On Error GoTo Oups

 'Proc qui fait le switch bouton visible/invible et enabled/disabled
 Dim bCmbContact As Boolean
 Dim bTxtNomContact As Boolean
 Dim bTxtRechercher As Boolean
 Dim bCmdAdd As Boolean
 Dim bCmdEnr As Boolean
 Dim bCmdModif As Boolean
 Dim bCmdSupp As Boolean
 Dim bCmdAnul As Boolean
 Dim bCmdQuit As Boolean
 Dim bCmdRenommer As Boolean
  Dim bCmdRafraichir As Boolean
  Dim bCmdRechercher As Boolean
  Dim bCmdImprimer As Boolean
  Dim bCmdCopier As Boolean
  Dim bFax As Boolean
  Dim bCmdMailList As Boolean
 
  Select Case eMode
 Case MODE_AJOUT_MODIF:
  bCmdEnr = True
 bCmdAnul = True
bTxtNomContact = True
 
 Case MODE_INACTIF:
 bCmbContact = True
 bTxtRechercher = True
 bCmdAdd = True
 bCmdModif = True
 bCmdSupp = True
 bCmdQuit = True
 bCmdRenommer = True
 bCmdRafraichir = True
 bCmdImprimer = True
 bCmdCopier = True
 bFax = True
 bCmdMailList = True
 
 If Len(Trim$(txtRechercher.Text)) > 0 Then
 bCmdRechercher = True
 End If
End Select
 
 cmbContact.Visible = bCmbContact
1  txtNomContact.Visible = bTxtNomContact
 txtRechercher.Enabled = bTxtRechercher
 CmdAdd.Visible = bCmdAdd
CmdEnr.Visible = bCmdEnr
CmdModif.Visible = bCmdModif
CmdSupp.Visible = bCmdSupp
CmdAnul.Visible = bCmdAnul
CmdQuit.Visible = bCmdQuit
cmdRechercher.Enabled = bCmdRechercher
cmdReport.Visible = bCmdImprimer
cmdCopier.Visible = bCmdCopier
cmdFax.Visible = bFax
cmdMailList.Visible = bCmdMailList

2  Exit Sub

Oups:

wOups "frmContact", "AfficherControles", Err, Err.number, Err.Description
End Sub

Private Sub CmdAdd_Click()

 On Error GoTo Oups

 'proc qui permet d'ajouter un contact à la BD
 Dim sName As String
 Dim bAjouter As Boolean
 
 'On procede a la saisie du nom et du contact
 sName = InputBox("Veuillez entrer le nom du contact" & vbNewLine & _
 vbNewLine & _
 "SVP, respectez le bon orthographe!", "SAISIE DU NOM", "Nom du contact")
 
 If sName <> vbNullString Then
 If ExisteDansBD(sName) = True Then
 If MsgBox("Il y a déjà un contact portant ce nom!" & vbNewLine & "Voulez-vous l'ajouter quand même?", vbYesNo) = vbYes Then
 bAjouter = True
 Else
 bAjouter = False
 End If
  Else
  If ContientCaracteresIncorrects(sName) = True Then
  Call MsgBox("Il y a des caractères non permis!", vbOKOnly, "Erreur")

  bAjouter = False
  Else
  bAjouter = True
  End If
  End If
10 Else
1 bAjouter = False
End If

If bAjouter = True Then
 Screen.MousePointer = vbHourglass
 
 m_bModeAjout = True
 
 'On montre seulement les boutton pour enregistrer
 Call AfficherControles(MODE_AJOUT_MODIF)
 
 'On montre les maskEdBox
 Call HideEdMask(False)
 
 'On affiche le nom du nouveau client dans le textbox
 'pour éviter le ScrollDown durant l'ajout
 
 txtNomContact.Text = sName
 
 Call ViderBarrerChamps(False, True)
 
 Call txtCompagnie.SetFocus
 
 Screen.MousePointer = vbDefault
1  End If

Exit Sub

Oups:

 wOups "frmContact", "CmdAdd_Click", Err, Err.number, Err.Description
End Sub

Private Sub ViderBarrerChamps(ByVal bLocked As Boolean, ByVal bVider As Boolean)
 
 On Error GoTo Oups
 
 If bVider = True Then
 txtCompagnie.Text = vbNullString
 txtTelephone.Text = vbNullString
 txtTitre.Text = vbNullString
 txtPoste.Text = vbNullString
 txtFax.Text = vbNullString
 txtPagette.Text = vbNullString
 txtCellulaire.Text = vbNullString
 txtEmail.Text = vbNullString
 txtTelDomicile.Text = vbNullString
  txtcommentaire.Text = vbNullString
  lblDateCreation.Caption = vbNullString
  lblUserCreation.Caption = vbNullString
  lblDateModification.Caption = vbNullString
  lblUserModification.Caption = vbNullString
  End If
 
  txtNomContact.Locked = bLocked
  txtCompagnie.Locked = bLocked
10 txtTelephone.Locked = bLocked
txtTitre.Locked = bLocked
txtPoste.Locked = bLocked
txtFax.Locked = bLocked
txtPagette.Locked = bLocked
txtCellulaire.Locked = bLocked
txtEmail.Locked = bLocked
txtTelDomicile.Locked = bLocked
txtcommentaire.Locked = bLocked

Exit Sub

Oups:

wOups "frmContact", "ViderBarrerChamps", Err, Err.number, Err.Description
End Sub

Private Sub AfficherContact()

 On Error GoTo Oups
 
 'Affiche le contact sélectionné
 Dim rstContact As ADODB.Recordset
 
 Set rstContact = New ADODB.Recordset
 
 Call rstContact.Open("SELECT * FROM GrbContact WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)
 
 'Compagnie
 If Not IsNull(rstContact.Fields("Compagnie")) Then
 txtCompagnie.Text = rstContact.Fields("Compagnie")
 Else
 txtCompagnie.Text = vbNullString
 End If

 'Titre
 If Not IsNull(rstContact.Fields("Titre")) Then
 txtTitre.Text = rstContact.Fields("Titre")
  Else
  txtTitre.Text = vbNullString
  End If

 'Telephonne
  If Not IsNull(rstContact.Fields("Telephonne")) Then
  txtTelephone.Text = rstContact.Fields("Telephonne")
  Else
  txtTelephone.Text = vbNullString
  End If

 'noPoste
10 If Not IsNull(rstContact.Fields("noPoste")) Then
1 txtPoste.Text = rstContact.Fields("noPoste")
Else
 txtPoste.Text = vbNullString
End If
 
 'Fax
If Not IsNull(rstContact.Fields("Fax")) Then
 txtFax.Text = rstContact.Fields("Fax")
Else
 txtFax.Text = vbNullString
End If
 
 'Pagette
If Not IsNull(rstContact.Fields("Pagette")) Then
 txtPagette.Text = rstContact.Fields("Pagette")
1  Else
 txtPagette.Text = vbNullString
 End If

 'Cellulaire
If Not IsNull(rstContact.Fields("Cellulaire")) Then
 txtCellulaire.Text = rstContact.Fields("Cellulaire")
Else
 txtCellulaire.Text = vbNullString
1  End If

 'teldomicile
 If Not IsNull(rstContact.Fields("teldomicile")) Then
 txtTelDomicile.Text = rstContact.Fields("teldomicile")
Else
 txtTelDomicile.Text = vbNullString
End If

 'E-mail
If Not IsNull(rstContact.Fields("E-mail")) Then
 txtEmail.Text = rstContact.Fields("E-mail")
Else
 txtEmail.Text = vbNullString
End If

 'Commentaire
If Not IsNull(rstContact.Fields("Commentaire")) Then
 txtcommentaire.Text = rstContact.Fields("Commentaire")
2  Else
 txtcommentaire.Text = vbNullString
2  End If

 'Création
If Not IsNull(rstContact.Fields("DateCréation")) Then
lblDateCreation.Caption = rstContact.Fields("DateCréation")
Else
lblDateCreation.Caption = vbNullString
End If

 'User Création
30 If Not IsNull(rstContact.Fields("UserCréation")) Then
3 lblUserCreation.Caption = "Par : " & rstContact.Fields("UserCréation")
Else
 lblUserCreation.Caption = vbNullString
End If

 'Modification
If Not IsNull(rstContact.Fields("DateModification")) Then
 lblDateModification.Caption = rstContact.Fields("DateModification")
Else
 lblDateModification.Caption = vbNullString
End If

 'User Modification
If Not IsNull(rstContact.Fields("UserModification")) Then
 lblUserModification.Caption = "Par : " & rstContact.Fields("UserModification")
3  Else
 lblUserModification.Caption = vbNullString
3  End If

Call rstContact.Close
3  Set rstContact = Nothing

Exit Sub

Oups:

3  wOups "frmContact", "AfficherContact", Err, Err.number, Err.Description
End Sub

Private Sub HideEdMask(ByVal bVisible As Boolean)

 On Error GoTo Oups
 
 'proc qui rend visible/ou non les maskEdBox
 'On en profite pour les nettoyer du dernier Enregistrement
 'et on fait l'inverse avec les textBox
 If m_bModeAjout = True Then
 txtTelephone.Text = vbNullString
 txtCellulaire.Text = vbNullString
 txtPagette.Text = vbNullString
 txtFax.Text = vbNullString
 txtTelDomicile = vbNullString
 
 mskTelephone.Text = vbNullString
 mskCellulaire.Text = vbNullString
 mskPagette.Text = vbNullString
 mskFax.Text = vbNullString
  mskTelDomicile.Text = vbNullString
  Else
  mskTelephone.Text = txtTelephone.Text
  mskCellulaire.Text = txtCellulaire.Text
  mskPagette.Text = txtPagette.Text
  mskFax.Text = txtFax.Text
  mskTelDomicile.Text = txtTelDomicile.Text
  End If
 
10 mskTelephone.Visible = Not bVisible
mskCellulaire.Visible = Not bVisible
mskPagette.Visible = Not bVisible
mskFax.Visible = Not bVisible
mskTelDomicile.Visible = Not bVisible
 
txtTelephone.Visible = bVisible
txtCellulaire.Visible = bVisible
txtPagette.Visible = bVisible
txtFax.Visible = bVisible
txtTelDomicile.Visible = bVisible

Exit Sub

Oups:

wOups "frmContact", "HideEdMask", Err, Err.number, Err.Description
End Sub

Private Sub CmdAnul_Click()

 On Error GoTo Oups
 
 'Annule l'ajout ou la modif
 Screen.MousePointer = vbHourglass
 
 'On cache le maskEdBox
 Call HideEdMask(True)
 
 'commentaire unlock
 'txtNomClient.Visible = False
 m_bModeAjout = False

 'on retablis les bouttons
 Call AfficherControles(MODE_INACTIF)

 'jfc 15oct
 'on affiche les donnée du premier enreg
 Call cmbContact_Click
 
 Call ViderBarrerChamps(True, True)
 
 Call cmbContact_Click
 
 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmContact", "CmdAnul_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdEnr_Click()

 On Error GoTo Oups
 
 'Enregistrement d'un contact dnas la BD
 Dim iCompteur As Integer
 Dim sContact As String
 Dim bSave As Boolean
 
 'Nom du contact
 sContact = txtNomContact.Text
 
 If Trim$(txtNomContact.Text) = "" Or Trim$(txtCompagnie.Text) = "" Then
 Call MsgBox("Le nom du contact et la compagnie sont obligatoires!", vbOKOnly, "Erreur")

 bSave = False
 Else
 If m_bModeAjout = True Then
 bSave = True
  Else
  If Trim$(Left$(cmbContact.Text, InStr(1, cmbContact.Text, " - ") - 1)) = txtNomContact.Text Then
  bSave = True
  Else
  If ExisteDansBD(txtNomContact.Text) = True Then
  If MsgBox("Il y a déjà un contact portant ce nom!" & vbNewLine & "Voulez-vous l'ajouter quand même?", vbYesNo) = vbYes Then
  bSave = True
  Else
 bSave = False
 End If
 Else
 bSave = True
 End If
 End If
 End If
End If
 
If bSave = True Then
 Screen.MousePointer = vbHourglass
 
 Call EnregistrerContact
 
 'On cache les MaskEdBox
 Call HideEdMask(True)
 
 'On met a jour le combo
Call RemplirComboContact
 
 'Retablir les boutons
 Call AfficherControles(MODE_INACTIF)
 
 For iCompteur = 0 To cmbContact.ListCount - 1
 If Trim$(Left$(cmbContact.LIST(iCompteur), InStr(1, cmbContact.LIST(iCompteur), "-") - 1)) = sContact Then
 cmbContact.ListIndex = iCompteur
 
 Exit For
 End If
1  Next
 
 Call cmbContact.SetFocus

 Call ViderBarrerChamps(True, False)
 
 Screen.MousePointer = vbDefault
End If

Exit Sub

Oups:

wOups "frmContact", "CmdEnr_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdModif_Click()

 On Error GoTo Oups
 
 'Pour modifier l'enregistrement courant
 If cmbContact.ListCount > 0 Then
 Screen.MousePointer = vbHourglass
 
 Call HideEdMask(False)
 Call AfficherControles(MODE_AJOUT_MODIF)
 Call ViderBarrerChamps(False, False)
 Call txtCompagnie.SetFocus
 
 Screen.MousePointer = vbDefault
 Else
 Call MsgBox("Aucun enregistrement sélectionné!")
 End If

  Exit Sub

Oups:

  wOups "frmContact", "CmdModif_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdquit_Click()

 On Error GoTo Oups
 'Fermeture de la fenêtre
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmContact", "cmdquit_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdSupp_Click()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim rstContact As ADODB.Recordset
 Dim rstLiaison As ADODB.Recordset
 Dim bPeutEffacer As Boolean
 
 'fonction qui supprime lenregistrement courant
 If cmbContact.ListCount > 0 Then
 If MsgBox("Etes-vous sur de supprimer cette enregistrement?", vbYesNo, "Supprimer") = vbYes Then
 Screen.MousePointer = vbHourglass
 
 'open table
 Set rstProjSoum = New ADODB.Recordset
 
 Call rstProjSoum.Open("SELECT * FROM GrbSoumissionMec WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)

 'Si existe pas dans soumission, on peut le deleté
 If rstProjSoum.EOF Then
  Call rstProjSoum.Close
 
  Call rstProjSoum.Open("SELECT * FROM GrbProjetMec WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)
 
  If rstProjSoum.EOF Then
  Call rstProjSoum.Close
 
  Call rstProjSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)
 
  If rstProjSoum.EOF Then
  Call rstProjSoum.Close
 
  Call rstProjSoum.Open("SELECT * FROM GrbProjetElec WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstProjSoum.EOF Then
 bPeutEffacer = True

 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 Else
 bPeutEffacer = False
 
 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 End If
 Else
 bPeutEffacer = False
 
 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 End If
 Else
 bPeutEffacer = False
 
 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 End If
1  Else
 bPeutEffacer = False
 
 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 End If
 End If
 
 Call SupprimerContactExchange(m_iNoContact)

 If bPeutEffacer = True Then
 Call g_connData.Execute("DELETE * FROM GrbContact WHERE IDContact = " & m_iNoContact)
 Else
 Set rstContact = New ADODB.Recordset

 Call rstContact.Open("SELECT * FROM GrbContact WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)

 rstContact.Fields("Supprimé") = True

 Call rstContact.Update

 Call rstContact.Close
 Set rstContact = Nothing
 End If

Set rstLiaison = New ADODB.Recordset

 Call rstLiaison.Open("SELECT * FROM GrbContactClient WHERE NoContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)

If Not rstLiaison.EOF Then
 Do While Not rstLiaison.EOF
 Call rstLiaison.Delete

 Call rstLiaison.MoveNext
 Loop
 End If

 Call rstLiaison.Close

 Call rstLiaison.Open("SELECT * FROM GrbContactFRS WHERE NoContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)

 If Not rstLiaison.EOF Then
 Do While Not rstLiaison.EOF
 Call rstLiaison.Delete

 Call rstLiaison.MoveNext
 Loop
 End If

Call rstLiaison.Close
 Set rstLiaison = Nothing
 
Call RemplirComboContact
 
 Screen.MousePointer = vbDefault
3  Else
 Call MsgBox("Aucun enregistrement sélectionné!")
3  End If

 Exit Sub

Oups:

40 wOups "frmContact", "CmdSupp_Click", Err, Err.number, Err.Description
End Sub

Private Sub LierContactClient(ByVal iContactID As Integer, ByVal iClientID As Integer)

 On Error GoTo Oups

 Dim otlApp As Outlook.Application
 Dim itmContact As Outlook.ContactItem
 Dim itmClient As Outlook.ContactItem
 Dim folClient As MAPIFolder
 Dim folContact As MAPIFolder
 Dim rstContactClient As ADODB.Recordset
 Dim rstClient As ADODB.Recordset
 Dim bDejaOuvert As Boolean
 Dim iCompteur As Integer

 Set otlApp = OuvrirOutlook(bDejaOuvert)

  Set folClient = GetFolder(otlApp, "Clients GRB")
  Set folContact = GetFolder(otlApp, "Contacts GRB")

  Set rstClient = New ADODB.Recordset

  Call rstClient.Open("SELECT EntryIDOutlook FROM GrbClient WHERE IDClient = " & iClientID, g_connData, adOpenForwardOnly, adLockReadOnly)

  Set itmClient = folClient.Items.Find("[User1] = " & iClientID)

  If Not itmClient Is Nothing Then
  Do While itmClient.Links.count > 0
 Set itmContact = folContact.Items.Find("[User1] = " & itmClient.Links.Item(1).Item.User1)

 For iCompteur = 1 To itmContact.Links.count
 If itmContact.Links.Item(1).Item.User1 = itmClient.User1 Then
 Call itmContact.Links.Remove(iCompteur)

 Call itmContact.Save

 Exit For
 End If
 Next

 Call itmClient.Links.Remove(1)
 Loop

 Call itmClient.Save

 Call rstClient.Close
Set rstClient = Nothing

 Set rstContactClient = New ADODB.Recordset

 Call rstContactClient.Open("SELECT * FROM GrbContactClient WHERE NoClient = " & iClientID, g_connData, adOpenForwardOnly, adLockReadOnly)

 Do While Not rstContactClient.EOF
 If rstContactClient.Fields("NoContact") <> iContactID Then
 Set itmContact = folContact.Items.Find("[User1] = " & rstContactClient.Fields("NoContact"))

 If Not itmContact Is Nothing Then
1  Call itmClient.Links.Add(itmContact)

 Call itmClient.Save

 Call itmContact.Links.Add(itmClient)

 Call itmContact.Save
 End If
 End If

 Call rstContactClient.MoveNext
 Loop

 Call rstContactClient.Close
 Set rstContactClient = Nothing
Else
 Call MsgBox("Client introuvable!", vbOKOnly, "Erreur")

 Call rstClient.Close
Set rstClient = Nothing
End If

2  If bDejaOuvert = False Then
 Call otlApp.Quit
2  End If

Set otlApp = Nothing

2  DoEvents

Exit Sub

Oups:

30 wOups "frmClient", "LierContactClient", Err, Err.number, Err.Description

fraEtatOutlook.Visible = False
End Sub

Private Sub LierContactFournisseur(ByVal iContactID As Integer, ByVal iFournisseurID As Integer)

 On Error GoTo Oups

 Dim otlApp As Outlook.Application
 Dim itmContact As Outlook.ContactItem
 Dim itmFRS As Outlook.ContactItem
 Dim folFRS As MAPIFolder
 Dim folContact As MAPIFolder
 Dim rstContactFRS As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim bDejaOuvert As Boolean
 Dim iCompteur As Integer

 fraEtatOutlook.Visible = True

  Set otlApp = OuvrirOutlook(bDejaOuvert)

  Set folFRS = GetFolder(otlApp, "Fournisseurs GRB")
  Set folContact = GetFolder(otlApp, "Contacts GRB")

  Set rstFRS = New ADODB.Recordset

  Call rstFRS.Open("SELECT EntryIDOutlook FROM GrbFournisseur WHERE IDFRS = " & iFournisseurID, g_connData, adOpenForwardOnly, adLockReadOnly)

  Set itmFRS = folFRS.Items.Find("[User1] = " & iFournisseurID)

  If Not itmFRS Is Nothing Then
  Do While itmFRS.Links.count > 0
 Set itmContact = folContact.Items.Find("[User1] = " & itmFRS.Links.Item(1).Item.User1)

For iCompteur = 1 To itmContact.Links.count
 If itmContact.Links.Item(1).Item.User1 = itmFRS.User1 Then
 Call itmContact.Links.Remove(iCompteur)

 Call itmContact.Save

 Exit For
 End If
 Next

 Call itmFRS.Links.Remove(1)
 Loop

 Call itmFRS.Save

 Call rstFRS.Close
Set rstFRS = Nothing

 Set rstContactFRS = New ADODB.Recordset

 Call rstContactFRS.Open("SELECT * FROM GrbContactFRS WHERE NoFRS = " & iFournisseurID, g_connData, adOpenForwardOnly, adLockReadOnly)

 Do While Not rstContactFRS.EOF
 If rstContactFRS.Fields("NoContact") <> iContactID Then
 Set itmContact = folContact.Items.Find("[User1] = " & rstContactFRS.Fields("NoContact"))

 If Not itmContact Is Nothing Then
1  Call itmFRS.Links.Add(itmContact)

 Call itmFRS.Save

 Call itmContact.Links.Add(itmFRS)

 Call itmContact.Save
 End If
 End If

 Call rstContactFRS.MoveNext
 Loop

 Call rstContactFRS.Close
 Set rstContactFRS = Nothing
Else
 Call MsgBox("Fournisseur introuvable!", vbOKOnly, "Erreur")

 Call rstFRS.Close
Set rstFRS = Nothing
End If

2  If bDejaOuvert = False Then
 Call otlApp.Quit
2  End If

Set otlApp = Nothing

2  fraEtatOutlook.Visible = False

DoEvents

30 Exit Sub

Oups:

wOups "frmFRS", "LierContactFournisseur", Err, Err.number, Err.Description

fraEtatOutlook.Visible = False
End Sub


Public Sub cmbContact_Click()

 On Error GoTo Oups
 
 'Quand le user selectionne un enregistrement on se posotionne dessus
 If cmbContact.Text <> vbNullString Then
 txtNomContact.Text = Trim$(Left$(cmbContact.Text, InStr(1, cmbContact.Text, " - ") - 1))
 Else
 cmbContact.Text = txtNomContact.Text
 End If
 
 If cmbContact.ListIndex > -1 Then
 If m_bRenommer = False And m_bModeAjout = False Then
 m_iNoContact = cmbContact.ItemData(cmbContact.ListIndex)
 End If
 End If
 
 'remplis le combo dépendant le contact sélectionné
  Call AfficherContact

  Exit Sub

Oups:

  wOups "frmContact", "cmbContact_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Call RemplirComboContact

 Call HideEdMask(True)

 Call AfficherControles(MODE_INACTIF)
 
 Call ActiverBoutonsGroupe

 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmContact", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub ActiverBoutonsGroupe()

 On Error GoTo Oups

 CmdAdd.Enabled = g_bModificationContacts
 CmdModif.Enabled = g_bModificationContacts
 CmdSupp.Enabled = g_bModificationContacts
 cmdCopier.Enabled = g_bModificationContacts
 cmdMailList.Enabled = g_bModificationListeDistribution

 Exit Sub

Oups:

 wOups "frmContact", "ActiverBoutonsGroupe", Err, Err.number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)

 On Error GoTo Oups

 Set FrmContact = Nothing

 Exit Sub

Oups:

 wOups "frmContact", "Form_Unload", Err, Err.number, Err.Description
End Sub
Private Sub mskTelephone_GotFocus()

 On Error GoTo Oups

 mskTelephone.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmContact", "mskTelephone_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskTelephone_LostFocus()

 On Error GoTo Oups

 mskTelephone.mask = vbNullString

 If mskTelephone.Text = "(___) ___-____" Then
 mskTelephone.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmContact", "mskTelephone_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskCellulaire_GotFocus()

 On Error GoTo Oups

 mskCellulaire.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmContact", "mskCellulaire_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskCellulaire_LostFocus()

 On Error GoTo Oups

 mskCellulaire.mask = vbNullString

 If mskCellulaire.Text = "(___) ___-____" Then
 mskCellulaire.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmContact", "mskCellulaire_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskFax_GotFocus()

 On Error GoTo Oups

 mskFax.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmContact", "mskFax_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskFax_LostFocus()

 On Error GoTo Oups

 mskFax.mask = vbNullString

 If mskFax.Text = "(___) ___-____" Then
 mskFax.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmContact", "mskFax_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskPagette_GotFocus()

 On Error GoTo Oups

 mskPagette.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmContact", "mskPagette_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskPagette_LostFocus()

 On Error GoTo Oups

 mskPagette.mask = vbNullString

 If mskPagette.Text = "(___) ___-____" Then
 mskPagette.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmContact", "mskPagette_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskTelDomicile_GotFocus()

 On Error GoTo Oups

 mskTelDomicile.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmContact", "mskTelDomicile_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskTelDomicile_LostFocus()

 On Error GoTo Oups

 mskTelDomicile.mask = vbNullString

 If mskTelDomicile.Text = "(___) ___-____" Then
 mskTelDomicile.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmContact", "mskTelDomicile_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub cmdRechercher_Click()

 On Error GoTo Oups

 Dim rstContact As ADODB.Recordset
 Dim sSearch As String
 
 sSearch = txtRechercher.Text
 
 Screen.MousePointer = vbHourglass
 
 'vide les champs
 Call ViderBarrerChamps(True, True)
 
 'Filtre pour selection des Nomcontact
 Set rstContact = New ADODB.Recordset
 
 Call rstContact.Open("SELECT NomContact, Compagnie, IDContact FROM GrbContact WHERE Instr(1, NomContact,'" & Replace(sSearch, "'", "''") & "') > 0 Or Instr(1, Compagnie, '" & Replace(sSearch, "'", "''") & "') > 0 AND Supprimé = False ORDER BY NomContact", g_connData, adOpenDynamic, adLockOptimistic)
 
 'vide combo
 Call cmbContact.Clear
 
 Do While Not rstContact.EOF
 Call cmbContact.AddItem(rstContact.Fields("NomContact") & " - " & rstContact.Fields("Compagnie"))
  cmbContact.ItemData(cmbContact.newIndex) = rstContact.Fields("IDContact")
 
  Call rstContact.MoveNext
  Loop
 
  Call rstContact.Close
  Set rstContact = Nothing
 
  Screen.MousePointer = vbDefault
 
  If cmbContact.ListCount > 0 Then
  cmbContact.ListIndex = 0
10 End If

Exit Sub

Oups:

wOups "frmContact", "cmdRechercher_Click", Err, Err.number, Err.Description
End Sub

Private Sub txtRechercher_Change()

 On Error GoTo Oups

 If Len(Trim$(txtRechercher.Text)) > 0 Then
 cmdRechercher.Enabled = True
 Else
 cmdRechercher.Enabled = False
 End If

 Exit Sub

Oups:

 wOups "frmContact", "txtRechercher_Change", Err, Err.number, Err.Description
End Sub
