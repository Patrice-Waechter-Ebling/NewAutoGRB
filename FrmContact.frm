VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmContact 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contacts"
   ClientHeight    =   7545
   ClientLeft      =   3330
   ClientTop       =   2670
   ClientWidth     =   9540
   Icon            =   "FrmContact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "FrmContact.frx":0442
   ScaleHeight     =   7545
   ScaleWidth      =   9540
   StartUpPosition =   2  'CenterScreen
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
      Left            =   600
      TabIndex        =   14
      Top             =   2880
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   32
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label lblUserModification 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   39
      Top             =   6375
      Width           =   1335
   End
   Begin VB.Label lblUserCreation 
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
      ForeColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   37
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   34
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   30
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   29
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   26
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   23
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   20
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
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

Private m_bModeAjout      As Boolean
Private m_bRenommer       As Boolean
Private m_iNoContact      As Integer

Public m_bAnnulerDistList As Boolean
Public m_otlDistList      As Outlook.DistListItem

Private Sub RemplirComboContact()

5       On Error GoTo AfficherErreur

        'Rempli le combo des contacts
10      Dim rstContact As ADODB.Recordset
  
        'Ouverture de la table
15      Set rstContact = New ADODB.Recordset
        
20      Call rstContact.Open("SELECT NomContact, Compagnie, IDContact FROM GRB_Contact WHERE Supprimé = False", g_connData, adOpenDynamic, adLockOptimistic)

        'Il faut vider le combo avant de le remplir
25      Call cmbContact.Clear

        'Tant que ce n'est pas la fin des enregistrements
30      Do While Not rstContact.EOF
          'Ajout du nom du contact dans le combo
35        Call cmbContact.AddItem(rstContact.Fields("NomContact") & " - " & rstContact.Fields("Compagnie"))
    
          'Ajout du numéro du contact dans l'itemData du combo
40        cmbContact.ItemData(cmbContact.newIndex) = rstContact.Fields("IDContact")
  
45        Call rstContact.MoveNext
50      Loop

55      Call rstContact.Close
60      Set rstContact = Nothing
  
65      If cmbContact.ListCount > 0 Then
70        cmbContact.ListIndex = 0
75      End If

80      Exit Sub

AfficherErreur:

85      woups "frmContact", "RemplirComboContact", Err, Erl
End Sub

Private Sub EnregistrerContact()

5       On Error GoTo AfficherErreur

10      Dim rstContact As ADODB.Recordset

15      Set rstContact = New ADODB.Recordset
    
20      If m_bModeAjout = True Then
25        Call rstContact.Open("SELECT * FROM GRB_Contact", g_connData, adOpenDynamic, adLockOptimistic)
    
30        Call rstContact.AddNew

35        rstContact.Fields("DateCréation") = ConvertDate(Date)
40        rstContact.Fields("UserCréation") = g_sInitiale
45      Else
50        Call rstContact.Open("SELECT * FROM GRB_contact WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)

55        rstContact.Fields("DateModification") = ConvertDate(Date)
60        rstContact.Fields("UserModification") = g_sInitiale
65      End If
  
        'Enregistrement du contact
70      rstContact.Fields("NomContact") = txtNomContact.Text
75      rstContact.Fields("Compagnie") = txtCompagnie.Text
80      rstContact.Fields("Titre") = txtTitre.Text
85      rstContact.Fields("Telephonne") = mskTelephone.Text
90      rstContact.Fields("Fax") = mskFax.Text
95      rstContact.Fields("Pagette") = mskPagette.Text
100     rstContact.Fields("Cellulaire") = mskCellulaire.Text
105     rstContact.Fields("E-mail") = txtEmail.Text
110     rstContact.Fields("NoPoste") = txtPoste.Text
115     rstContact.Fields("TelDomicile") = mskTelDomicile.Text
120     rstContact.Fields("Commentaire") = txtcommentaire.Text

125     rstContact.Fields("EntryIDOutlook") = ModifierContactExchange(rstContact.Fields("IDContact"))

130     If m_bModeAjout = True Then
135       m_bModeAjout = False
140     End If

145     Call rstContact.Update
    
150     Call rstContact.Close
155     Set rstContact = Nothing

160     Exit Sub

AfficherErreur:

165     woups "frmContact", "EnregistrerContact", Err, Erl
End Sub

Private Function ModifierContactExchange(ByVal iContactID As Integer) As String

5       On Error GoTo AfficherErreur

10      Dim otlApp      As Outlook.Application
15      Dim otlContact  As Outlook.ContactItem
20      Dim folContact  As Outlook.MAPIFolder
25      Dim bDejaOuvert As Boolean
30      Dim sNom()      As String

35      If m_bModeAjout = True Then
40        lblEtatOutlook.Caption = "Ajout du contact dans Outlook ..."
45      Else
50        lblEtatOutlook.Caption = "Modification du contact dans Outlook ..."
55      End If

60      fraEtatOutlook.Visible = True

65      Set otlApp = OuvrirOutlook(bDejaOuvert)

70      Set folContact = GetFolder(otlApp, "Contacts GRB")

75      If m_bModeAjout = True Then
80        Set otlContact = folContact.Items.Add(olContactItem)

85        otlContact.User1 = iContactID
90      Else
95        Set otlContact = folContact.Items.Find("[User1] = " & iContactID)
100     End If

105     If otlContact Is Nothing Then
110       Call MsgBox("Le contact " & txtNomContact.Text & " n'a pas été trouvé pour la mise à jour Exchange!", vbOKOnly, "Erreur")

115       fraEtatOutlook.Visible = False

120       DoEvents

125       Exit Function
130     End If

135     sNom = Split(Trim$(txtNomContact.Text), " ")

140     Select Case UBound(sNom)
          Case 0:
145         otlContact.FirstName = sNom(0)
  
          Case 1:
150         otlContact.FirstName = sNom(0)
155         otlContact.LastName = sNom(1)

          Case 2
160         otlContact.FirstName = sNom(0)
165         otlContact.MiddleName = sNom(1)
170         otlContact.LastName = sNom(2)
175     End Select
        
180     otlContact.Title = ""

185     otlContact.CompanyName = txtCompagnie.Text
190     otlContact.JobTitle = txtTitre.Text

195     If mskTelephone.Text <> "(___) ___-____" Then
200       If Trim$(txtPoste.Text) <> "" Then
205         otlContact.BusinessTelephoneNumber = mskTelephone.Text & " Ext : " & txtPoste.Text
210       Else
215         otlContact.BusinessTelephoneNumber = mskTelephone.Text
220       End If
225     End If

230     If mskFax.Text <> "(___) ___-____" Then
235       otlContact.BusinessFaxNumber = mskFax.Text
240     End If

245     If mskCellulaire.Text <> "(___) ___-____" Then
250       otlContact.MobileTelephoneNumber = mskCellulaire.Text
255     End If

260     If mskPagette.Text <> "(___) ___-____" Then
265       otlContact.PagerNumber = mskPagette.Text
270     End If

275     otlContact.Email1Address = txtEmail.Text

280     If mskTelDomicile.Text <> "(___) ___-____" Then
285       otlContact.HomeTelephoneNumber = mskTelDomicile.Text
290     End If

295     If txtcommentaire.Text <> "" Then
300       otlContact.Body = txtcommentaire.Text
305     End If

310     Call otlContact.Save

315     ModifierContactExchange = otlContact.EntryID

320     If bDejaOuvert = False Then
325       Call otlApp.Quit
330     End If

335     Set otlApp = Nothing

340     fraEtatOutlook.Visible = False

345     DoEvents

350     Exit Function

AfficherErreur:

355     woups "frmContact", "ModifierContactExchange", Err, Erl, "iContactID = " & iContactID)

360     fraEtatOutlook.Visible = False
End Function

Private Sub SupprimerContactExchange(ByVal iContactID As Integer)

5       On Error GoTo AfficherErreur

10      Dim otlApp      As Outlook.Application
15      Dim otlContact  As Outlook.ContactItem
20      Dim folContact  As MAPIFolder
25      Dim bDejaOuvert As Boolean

30      lblEtatOutlook.Caption = "Suppression du contact dans Outlook ..."
35      fraEtatOutlook.Visible = True

40      Set otlApp = OuvrirOutlook(bDejaOuvert)

45      Set folContact = GetFolder(otlApp, "Contacts GRB")

50      Set otlContact = folContact.Items.Find("[User1] = " & iContactID)

55      If Not otlContact Is Nothing Then
60        Call otlContact.Delete
65      End If

70      If bDejaOuvert = False Then
75        Call otlApp.Quit
80      End If

85      Set otlApp = Nothing

90      fraEtatOutlook.Visible = False

95      DoEvents

100     Exit Sub

AfficherErreur:

105     woups "frmContact", "SupprimerContactExchange", Err, Erl

110     fraEtatOutlook.Visible = False
End Sub

Private Sub cmdCopier_Click()

5       On Error GoTo AfficherErreur

        'Copie un contact
10      Dim sName    As String
15      Dim bAjouter As Boolean
  
        'On procede a la saisie du nom et du contact
20      sName = InputBox("Veuillez entrer le nom du contact", "SAISIE DU NOM", "Nom du contact")
    
25      If sName <> vbNullString Then
30        If ExisteDansBD(sName) = True Then
35          If MsgBox("Il y a déjà un contact portant ce nom!" & vbNewLine & "Voulez-vous l'ajouter quand même?", vbYesNo) = vbYes Then
40            bAjouter = True
45          Else
50            bAjouter = False
55          End If
60        Else
65          If ContientCaracteresIncorrects(sName) = True Then
70            Call MsgBox("Il y a des caractères non permis!", vbOKOnly, "Erreur")

75            bAjouter = False
80          Else
85            bAjouter = True
90          End If
95        End If
100     Else
105       bAjouter = False
110     End If

115     If bAjouter = True Then
120       Screen.MousePointer = vbHourglass
      
          'On montre seulement les boutton pour enregistrer
125       Call AfficherControles(MODE_AJOUT_MODIF)
        
          'On montre les maskEdBox
130       Call HideEdMask(False)
      
135       m_bModeAjout = True
                
          'On affiche le nom du nouveau client dans le textbox
          'pour éviter le ScrollDown durant l'ajout
        
140       txtNomContact.Text = sName
        
145       Call ViderBarrerChamps(False, False)
        
150       Call txtCompagnie.SetFocus
        
155       Screen.MousePointer = vbDefault
160     End If

165     Exit Sub

AfficherErreur:

170     woups "frmContact", "cmdCopier_Click", Err, Erl
End Sub

Private Function ExisteDansBD(ByVal sName As String) As Boolean
  
5       On Error GoTo AfficherErreur

10      Dim rstContact As ADODB.Recordset

15      Set rstContact = New ADODB.Recordset

20      Call rstContact.Open("SELECT NomContact FROM GRB_Contact WHERE NomContact = '" & Replace(sName, "'", "''") & "' AND Supprimé = False", g_connData, adOpenForwardOnly, adLockReadOnly)

25      If rstContact.EOF Then
30        ExisteDansBD = False
35      Else
40        ExisteDansBD = True
45      End If

50      Call rstContact.Close
55      Set rstContact = Nothing

60      Exit Function

AfficherErreur:

65      woups "frmContact", "ExisteDansBD", Err, Erl
End Function

Private Function ContientCaracteresIncorrects(ByVal sName As String) As Boolean

5       On Error GoTo AfficherErreur

10      If InStr(1, sName, ",") > 0 Or InStr(1, sName, ";") > 0 Or InStr(1, sName, ":") > 0 Or InStr(1, sName, "(") > 0 Or InStr(1, sName, ")") > 0 Then
15        ContientCaracteresIncorrects = True
20      Else
25        ContientCaracteresIncorrects = False
30      End If

35      Exit Function

AfficherErreur:

40      woups "frmContact", "ContientCaracteresIncorrects", Err, Erl
End Function

Private Sub cmdFax_Click()

5       On Error GoTo AfficherErreur

10      If cmbContact.ListCount > 0 Then
15        Call frmreport.Afficher(0, m_iNoContact, FRM_CONTACTS)
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmContact", "cmdFax_Click", Err, Erl
End Sub

Private Sub cmdMailList_Click()

5       On Error GoTo AfficherErreur

10      Dim otlApp       As Outlook.Application
15      Dim folContact   As Outlook.MAPIFolder
20      Dim itmContact   As Outlook.ContactItem
25      Dim otlRecipient As Outlook.Recipient
30      Dim bDejaOuvert  As Boolean
35      Dim rstContact   As ADODB.Recordset
40      Dim sIDContact   As String
45      Dim sContact     As String

50      If cmbContact.ListIndex <> -1 Then
55        If Trim$(txtEmail.Text) <> "" Then
60          sIDContact = cmbContact.ItemData(cmbContact.ListIndex)
65          sContact = cmbContact.Text
70        End If
            
75        If sIDContact <> "" Then
80          Set otlApp = OuvrirOutlook(bDejaOuvert)

85          lblEtatOutlook.Caption = "Recherche des listes de distribution..."

90          fraEtatOutlook.Visible = True

95          Call frmChoixMailList.Afficher(Me, otlApp)

100         If m_bAnnulerDistList = False Then
105           lblEtatOutlook.Caption = "Ajout du contact dans la liste de distribution..."
 
110           fraEtatOutlook.Visible = True

115           Set folContact = GetFolder(otlApp, "Contacts GRB")

120           Set itmContact = folContact.Items.Find("[User1] = " & sIDContact)

125           If Not itmContact Is Nothing Then
130             Set otlRecipient = otlApp.Session.CreateRecipient(itmContact.Email1DisplayName)

135             If otlRecipient.Resolve = True Then
140               Call m_otlDistList.AddMember(otlRecipient)
      
145               Call m_otlDistList.Save
150             Else
155               Call MsgBox("Impossible de trouver le contact '" & sContact & "' !", vbOKOnly, "Erreur")
160             End If
165           Else
170             Call MsgBox("Contact '" & sContact & "' introuvable!", vbOKOnly, "Erreur")
175           End If
180         End If

185         If bDejaOuvert = False Then
190           Call otlApp.Quit
195         End If

200         Set otlApp = Nothing

205         fraEtatOutlook.Visible = False
210       Else
215         Call MsgBox("Le ou les contacts n'ont pas d'email!", vbOKOnly, "Erreur")
220       End If
225     Else
230       Call MsgBox("Aucun contact sélectionné!", vbOKOnly, "Erreur")
235     End If

240     Exit Sub

AfficherErreur:

245     If Err.number = 287 And Erl = 145 Then
250       Call MsgBox("La liste de distribution est pleine!", vbOKOnly, "Erreur")
255     Else
260       woups "frmContact", "cmdMailList_Click", Err, Erl
265     End If

270     fraEtatOutlook.Visible = False
End Sub

Private Sub cmdRafraichir_Click()

5       On Error GoTo AfficherErreur

        'Rempli la liste avec tous les contacts après avoir fait une recherche
10      Screen.MousePointer = vbHourglass

15      Call RemplirComboContact

20      Screen.MousePointer = vbDefault

25      Exit Sub

AfficherErreur:

30      woups "frmContact", "cmdRafraichir_Click", Err, Erl
End Sub

Private Sub cmdreport_Click()

5       On Error GoTo AfficherErreur

        'Impression de la liste des contacts
10      Dim rstContact As ADODB.Recordset

15      Set rstContact = New ADODB.Recordset

20      If MsgBox("Voulez-vous imprimer ce contact uniquement?", vbYesNo) = vbYes Then
25        Call rstContact.Open("SELECT * FROM GRB_Contact WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)
30      Else
35        If MsgBox("Voulez-vous filtrer par la compagnie '" & txtCompagnie.Text & "'?", vbYesNo) = vbYes Then
40          Call rstContact.Open("SELECT * FROM GRB_Contact WHERE Compagnie = '" & Replace(txtCompagnie.Text, "'", "''") & "' AND Supprimé = False ORDER BY NomContact", g_connData, adOpenDynamic, adLockOptimistic)
45        Else
50          Call rstContact.Open("SELECT * FROM GRB_Contact WHERE Supprimé = False ORDER BY NomContact", g_connData, adOpenDynamic, adLockOptimistic)
55        End If
60      End If

65      Screen.MousePointer = vbHourglass
  
        'set le rapport
70      Set DR_ListeContact.DataSource = rstContact
  
75      DR_ListeContact.Orientation = rptOrientPortrait

80      Call DR_ListeContact.Show(vbModal)
  
85      Call rstContact.Close
90      Set rstContact = Nothing
    
95      Screen.MousePointer = vbDefault

100     Exit Sub

AfficherErreur:

105     woups "frmContact", "cmdreport_Click", Err, Erl
End Sub

Private Sub AfficherControles(ByVal eMode As enumMode)

5       On Error GoTo AfficherErreur

        'Proc qui fait le switch bouton visible/invible et enabled/disabled
10      Dim bCmbContact    As Boolean
15      Dim bTxtNomContact As Boolean
20      Dim bTxtRechercher As Boolean
25      Dim bCmdAdd        As Boolean
30      Dim bCmdEnr        As Boolean
35      Dim bCmdModif      As Boolean
40      Dim bCmdSupp       As Boolean
45      Dim bCmdAnul       As Boolean
50      Dim bCmdQuit       As Boolean
55      Dim bCmdRenommer   As Boolean
60      Dim bCmdRafraichir As Boolean
65      Dim bCmdRechercher As Boolean
70      Dim bCmdImprimer   As Boolean
75      Dim bCmdCopier     As Boolean
80      Dim bFax           As Boolean
85      Dim bCmdMailList   As Boolean
  
90      Select Case eMode
          Case MODE_AJOUT_MODIF:
95          bCmdEnr = True
100         bCmdAnul = True
105         bTxtNomContact = True
      
          Case MODE_INACTIF:
110         bCmbContact = True
115         bTxtRechercher = True
120         bCmdAdd = True
125         bCmdModif = True
130         bCmdSupp = True
135         bCmdQuit = True
140         bCmdRenommer = True
145         bCmdRafraichir = True
150         bCmdImprimer = True
155         bCmdCopier = True
160         bFax = True
165         bCmdMailList = True
      
170         If Len(Trim$(txtRechercher.Text)) > 0 Then
175           bCmdRechercher = True
180         End If
185     End Select
  
190     cmbContact.Visible = bCmbContact
195     txtNomContact.Visible = bTxtNomContact
200     txtRechercher.Enabled = bTxtRechercher
205     CmdAdd.Visible = bCmdAdd
210     CmdEnr.Visible = bCmdEnr
215     CmdModif.Visible = bCmdModif
220     CmdSupp.Visible = bCmdSupp
225     CmdAnul.Visible = bCmdAnul
230     CmdQuit.Visible = bCmdQuit
235     cmdRechercher.Enabled = bCmdRechercher
240     cmdReport.Visible = bCmdImprimer
245     cmdCopier.Visible = bCmdCopier
250     cmdFax.Visible = bFax
255     cmdMailList.Visible = bCmdMailList

260     Exit Sub

AfficherErreur:

265     woups "frmContact", "AfficherControles", Err, Erl
End Sub

Private Sub CmdAdd_Click()

5       On Error GoTo AfficherErreur

        'proc qui permet d'ajouter un contact à la BD
10      Dim sName As String
15      Dim bAjouter As Boolean
  
        'On procede a la saisie du nom et du contact
20      sName = InputBox("Veuillez entrer le nom du contact" & vbNewLine & _
                         vbNewLine & _
                         "SVP, respectez le bon orthographe!", "SAISIE DU NOM", "Nom du contact")
    
25      If sName <> vbNullString Then
30        If ExisteDansBD(sName) = True Then
35          If MsgBox("Il y a déjà un contact portant ce nom!" & vbNewLine & "Voulez-vous l'ajouter quand même?", vbYesNo) = vbYes Then
40            bAjouter = True
45          Else
50            bAjouter = False
55          End If
60        Else
65          If ContientCaracteresIncorrects(sName) = True Then
70            Call MsgBox("Il y a des caractères non permis!", vbOKOnly, "Erreur")

75            bAjouter = False
80          Else
85            bAjouter = True
90          End If
95        End If
100     Else
105       bAjouter = False
110     End If

115     If bAjouter = True Then
120       Screen.MousePointer = vbHourglass
        
125       m_bModeAjout = True
        
          'On montre seulement les boutton pour enregistrer
130       Call AfficherControles(MODE_AJOUT_MODIF)
        
          'On montre les maskEdBox
135       Call HideEdMask(False)
                
          'On affiche le nom du nouveau client dans le textbox
          'pour éviter le ScrollDown durant l'ajout
        
140       txtNomContact.Text = sName
        
145       Call ViderBarrerChamps(False, True)
        
150       Call txtCompagnie.SetFocus
        
155       Screen.MousePointer = vbDefault
160     End If

165     Exit Sub

AfficherErreur:

170     woups "frmContact", "CmdAdd_Click", Err, Erl
End Sub

Private Sub ViderBarrerChamps(ByVal bLocked As Boolean, ByVal bVider As Boolean)
        
5       On Error GoTo AfficherErreur
  
10      If bVider = True Then
15        txtCompagnie.Text = vbNullString
20        txtTelephone.Text = vbNullString
25        txtTitre.Text = vbNullString
30        txtPoste.Text = vbNullString
35        txtFax.Text = vbNullString
40        txtPagette.Text = vbNullString
45        txtCellulaire.Text = vbNullString
50        txtEmail.Text = vbNullString
55        txtTelDomicile.Text = vbNullString
60        txtcommentaire.Text = vbNullString
65        lblDateCreation.Caption = vbNullString
70        lblUserCreation.Caption = vbNullString
75        lblDateModification.Caption = vbNullString
80        lblUserModification.Caption = vbNullString
85      End If
  
90      txtNomContact.Locked = bLocked
95      txtCompagnie.Locked = bLocked
100     txtTelephone.Locked = bLocked
105     txtTitre.Locked = bLocked
110     txtPoste.Locked = bLocked
115     txtFax.Locked = bLocked
120     txtPagette.Locked = bLocked
125     txtCellulaire.Locked = bLocked
130     txtEmail.Locked = bLocked
135     txtTelDomicile.Locked = bLocked
140     txtcommentaire.Locked = bLocked

145     Exit Sub

AfficherErreur:

150     woups "frmContact", "ViderBarrerChamps", Err, Erl
End Sub

Private Sub AfficherContact()

5       On Error GoTo AfficherErreur
        
        'Affiche le contact sélectionné
10      Dim rstContact As ADODB.Recordset
  
15      Set rstContact = New ADODB.Recordset
  
20      Call rstContact.Open("SELECT * FROM GRB_Contact WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)
  
        'Compagnie
25      If Not IsNull(rstContact.Fields("Compagnie")) Then
30        txtCompagnie.Text = rstContact.Fields("Compagnie")
35      Else
40        txtCompagnie.Text = vbNullString
45      End If

        'Titre
50      If Not IsNull(rstContact.Fields("Titre")) Then
55        txtTitre.Text = rstContact.Fields("Titre")
60      Else
65        txtTitre.Text = vbNullString
70      End If

        'Telephonne
75      If Not IsNull(rstContact.Fields("Telephonne")) Then
80        txtTelephone.Text = rstContact.Fields("Telephonne")
85      Else
90        txtTelephone.Text = vbNullString
95      End If

        'noPoste
100     If Not IsNull(rstContact.Fields("noPoste")) Then
105       txtPoste.Text = rstContact.Fields("noPoste")
110     Else
115       txtPoste.Text = vbNullString
120     End If
  
        'Fax
125     If Not IsNull(rstContact.Fields("Fax")) Then
130       txtFax.Text = rstContact.Fields("Fax")
135     Else
140       txtFax.Text = vbNullString
145     End If
  
        'Pagette
150     If Not IsNull(rstContact.Fields("Pagette")) Then
155       txtPagette.Text = rstContact.Fields("Pagette")
160     Else
165       txtPagette.Text = vbNullString
170     End If

        'Cellulaire
175     If Not IsNull(rstContact.Fields("Cellulaire")) Then
180       txtCellulaire.Text = rstContact.Fields("Cellulaire")
185     Else
190       txtCellulaire.Text = vbNullString
195     End If

        'teldomicile
200     If Not IsNull(rstContact.Fields("teldomicile")) Then
205       txtTelDomicile.Text = rstContact.Fields("teldomicile")
210     Else
215       txtTelDomicile.Text = vbNullString
220     End If

        'E-mail
225     If Not IsNull(rstContact.Fields("E-mail")) Then
230       txtEmail.Text = rstContact.Fields("E-mail")
235     Else
240       txtEmail.Text = vbNullString
245     End If

        'Commentaire
250     If Not IsNull(rstContact.Fields("Commentaire")) Then
255       txtcommentaire.Text = rstContact.Fields("Commentaire")
260     Else
265       txtcommentaire.Text = vbNullString
270     End If

        'Création
275     If Not IsNull(rstContact.Fields("DateCréation")) Then
280       lblDateCreation.Caption = rstContact.Fields("DateCréation")
285     Else
290       lblDateCreation.Caption = vbNullString
295     End If

        'User Création
300     If Not IsNull(rstContact.Fields("UserCréation")) Then
305       lblUserCreation.Caption = "Par : " & rstContact.Fields("UserCréation")
310     Else
315       lblUserCreation.Caption = vbNullString
320     End If

        'Modification
325     If Not IsNull(rstContact.Fields("DateModification")) Then
330       lblDateModification.Caption = rstContact.Fields("DateModification")
335     Else
340       lblDateModification.Caption = vbNullString
345     End If

        'User Modification
350     If Not IsNull(rstContact.Fields("UserModification")) Then
355       lblUserModification.Caption = "Par : " & rstContact.Fields("UserModification")
360     Else
365       lblUserModification.Caption = vbNullString
370     End If

375     Call rstContact.Close
380     Set rstContact = Nothing

385     Exit Sub

AfficherErreur:

390     woups "frmContact", "AfficherContact", Err, Erl
End Sub

Private Sub HideEdMask(ByVal bVisible As Boolean)

5       On Error GoTo AfficherErreur
        
        'proc qui rend visible/ou non les maskEdBox
        'On en profite pour les nettoyer du dernier Enregistrement
        'et on fait l'inverse avec les textBox
10      If m_bModeAjout = True Then
15        txtTelephone.Text = vbNullString
20        txtCellulaire.Text = vbNullString
25        txtPagette.Text = vbNullString
30        txtFax.Text = vbNullString
35        txtTelDomicile = vbNullString
    
40        mskTelephone.Text = vbNullString
45        mskCellulaire.Text = vbNullString
50        mskPagette.Text = vbNullString
55        mskFax.Text = vbNullString
60        mskTelDomicile.Text = vbNullString
65      Else
70        mskTelephone.Text = txtTelephone.Text
75        mskCellulaire.Text = txtCellulaire.Text
80        mskPagette.Text = txtPagette.Text
85        mskFax.Text = txtFax.Text
90        mskTelDomicile.Text = txtTelDomicile.Text
95      End If
  
100     mskTelephone.Visible = Not bVisible
105     mskCellulaire.Visible = Not bVisible
110     mskPagette.Visible = Not bVisible
115     mskFax.Visible = Not bVisible
120     mskTelDomicile.Visible = Not bVisible
 
125     txtTelephone.Visible = bVisible
130     txtCellulaire.Visible = bVisible
135     txtPagette.Visible = bVisible
140     txtFax.Visible = bVisible
145     txtTelDomicile.Visible = bVisible

150     Exit Sub

AfficherErreur:

155     woups "frmContact", "HideEdMask", Err, Erl
End Sub

Private Sub CmdAnul_Click()

5       On Error GoTo AfficherErreur
        
        'Annule l'ajout ou la modif
10      Screen.MousePointer = vbHourglass
    
        'On cache le maskEdBox
15      Call HideEdMask(True)
  
        'commentaire unlock
        'txtNomClient.Visible = False
20      m_bModeAjout = False

        'on retablis les bouttons
25      Call AfficherControles(MODE_INACTIF)

        'jfc 15oct
        'on affiche les donnée du premier enreg
30      Call cmbContact_Click
    
35      Call ViderBarrerChamps(True, True)
  
40      Call cmbContact_Click
  
45      Screen.MousePointer = vbDefault

50      Exit Sub

AfficherErreur:

55      woups "frmContact", "CmdAnul_Click", Err, Erl
End Sub

Private Sub CmdEnr_Click()

5       On Error GoTo AfficherErreur
        
        'Enregistrement d'un contact dnas la BD
10      Dim iCompteur  As Integer
15      Dim sContact   As String
20      Dim bSave      As Boolean
  
        'Nom du contact
25      sContact = txtNomContact.Text
  
30      If Trim$(txtNomContact.Text) = "" Or Trim$(txtCompagnie.Text) = "" Then
35        Call MsgBox("Le nom du contact et la compagnie sont obligatoires!", vbOKOnly, "Erreur")

40        bSave = False
45      Else
50        If m_bModeAjout = True Then
55          bSave = True
60        Else
65          If Trim$(Left$(cmbContact.Text, InStr(1, cmbContact.Text, " - ") - 1)) = txtNomContact.Text Then
70            bSave = True
75          Else
80            If ExisteDansBD(txtNomContact.Text) = True Then
85              If MsgBox("Il y a déjà un contact portant ce nom!" & vbNewLine & "Voulez-vous l'ajouter quand même?", vbYesNo) = vbYes Then
90                bSave = True
95              Else
100               bSave = False
105             End If
110           Else
115             bSave = True
120           End If
125         End If
130       End If
135     End If
  
140     If bSave = True Then
145       Screen.MousePointer = vbHourglass
   
150       Call EnregistrerContact
     
          'On cache les MaskEdBox
155       Call HideEdMask(True)
 
          'On met a jour le combo
160       Call RemplirComboContact
        
          'Retablir les boutons
165       Call AfficherControles(MODE_INACTIF)
  
170       For iCompteur = 0 To cmbContact.ListCount - 1
175         If Trim$(Left$(cmbContact.LIST(iCompteur), InStr(1, cmbContact.LIST(iCompteur), "-") - 1)) = sContact Then
180           cmbContact.ListIndex = iCompteur
      
185           Exit For
190         End If
195       Next
  
200       Call cmbContact.SetFocus

205       Call ViderBarrerChamps(True, False)
  
210       Screen.MousePointer = vbDefault
215     End If

220     Exit Sub

AfficherErreur:

225     woups "frmContact", "CmdEnr_Click", Err, Erl
End Sub

Private Sub CmdModif_Click()

5       On Error GoTo AfficherErreur
        
        'Pour modifier l'enregistrement courant
10      If cmbContact.ListCount > 0 Then
15        Screen.MousePointer = vbHourglass
      
20        Call HideEdMask(False)
25        Call AfficherControles(MODE_AJOUT_MODIF)
30        Call ViderBarrerChamps(False, False)
35        Call txtCompagnie.SetFocus
      
40        Screen.MousePointer = vbDefault
45      Else
50        Call MsgBox("Aucun enregistrement sélectionné!")
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmContact", "CmdModif_Click", Err, Erl
End Sub

Private Sub cmdquit_Click()

5       On Error GoTo AfficherErreur
              'Fermeture de la fenêtre
10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmContact", "cmdquit_Click", Err, Erl
End Sub

Private Sub CmdSupp_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum  As ADODB.Recordset
15      Dim rstContact   As ADODB.Recordset
20      Dim rstLiaison   As ADODB.Recordset
25      Dim bPeutEffacer As Boolean
  
        'fonction qui supprime lenregistrement courant
30      If cmbContact.ListCount > 0 Then
35        If MsgBox("Etes-vous sur de supprimer cette enregistrement?", vbYesNo, "Supprimer") = vbYes Then
40          Screen.MousePointer = vbHourglass
                           
            'open table
45          Set rstProjSoum = New ADODB.Recordset
            
50          Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)

            'Si existe pas dans soumission, on peut le deleté
55          If rstProjSoum.EOF Then
60            Call rstProjSoum.Close
        
65            Call rstProjSoum.Open("SELECT * FROM GRB_ProjetMec WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)
      
70            If rstProjSoum.EOF Then
75              Call rstProjSoum.Close
          
80              Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)
          
85              If rstProjSoum.EOF Then
90                Call rstProjSoum.Close
            
95                Call rstProjSoum.Open("SELECT * FROM GRB_ProjetElec WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)
            
100               If rstProjSoum.EOF Then
105                 bPeutEffacer = True

110                 Call rstProjSoum.Close
115                 Set rstProjSoum = Nothing
120               Else
125                 bPeutEffacer = False
                  
130                 Call rstProjSoum.Close
135                 Set rstProjSoum = Nothing
140               End If
145             Else
150               bPeutEffacer = False
         
155               Call rstProjSoum.Close
160               Set rstProjSoum = Nothing
165             End If
170           Else
175             bPeutEffacer = False
           
180             Call rstProjSoum.Close
185             Set rstProjSoum = Nothing
190           End If
195         Else
200           bPeutEffacer = False
          
205           Call rstProjSoum.Close
210           Set rstProjSoum = Nothing
215         End If
220       End If
  
225       Call SupprimerContactExchange(m_iNoContact)

230       If bPeutEffacer = True Then
235         Call g_connData.Execute("DELETE * FROM GRB_Contact WHERE IDContact = " & m_iNoContact)
240       Else
245         Set rstContact = New ADODB.Recordset

250         Call rstContact.Open("SELECT * FROM GRB_Contact WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)

255         rstContact.Fields("Supprimé") = True

260         Call rstContact.Update

265         Call rstContact.Close
270         Set rstContact = Nothing
275       End If

280       Set rstLiaison = New ADODB.Recordset

285       Call rstLiaison.Open("SELECT * FROM GRB_ContactClient WHERE NoContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)

290       If Not rstLiaison.EOF Then
295         Do While Not rstLiaison.EOF
300           Call rstLiaison.Delete

305           Call rstLiaison.MoveNext
310         Loop
315       End If

320       Call rstLiaison.Close

325       Call rstLiaison.Open("SELECT * FROM GRB_ContactFRS WHERE NoContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)

330       If Not rstLiaison.EOF Then
335         Do While Not rstLiaison.EOF
340           Call rstLiaison.Delete

345           Call rstLiaison.MoveNext
350         Loop
355       End If

360       Call rstLiaison.Close
365       Set rstLiaison = Nothing
                  
370       Call RemplirComboContact
     
375       Screen.MousePointer = vbDefault
380     Else
385       Call MsgBox("Aucun enregistrement sélectionné!")
390     End If

395     Exit Sub

AfficherErreur:

400     woups "frmContact", "CmdSupp_Click", Err, Erl
End Sub

Private Sub LierContactClient(ByVal iContactID As Integer, ByVal iClientID As Integer)

5       On Error GoTo AfficherErreur

10      Dim otlApp           As Outlook.Application
15      Dim itmContact       As Outlook.ContactItem
20      Dim itmClient        As Outlook.ContactItem
25      Dim folClient        As MAPIFolder
30      Dim folContact       As MAPIFolder
35      Dim rstContactClient As ADODB.Recordset
40      Dim rstClient        As ADODB.Recordset
45      Dim bDejaOuvert      As Boolean
50      Dim iCompteur        As Integer

55      Set otlApp = OuvrirOutlook(bDejaOuvert)

65      Set folClient = GetFolder(otlApp, "Clients GRB")
70      Set folContact = GetFolder(otlApp, "Contacts GRB")

75      Set rstClient = New ADODB.Recordset

80      Call rstClient.Open("SELECT EntryIDOutlook FROM GRB_Client WHERE IDClient = " & iClientID, g_connData, adOpenForwardOnly, adLockReadOnly)

85      Set itmClient = folClient.Items.Find("[User1] = " & iClientID)

90      If Not itmClient Is Nothing Then
95        Do While itmClient.Links.count > 0
100          Set itmContact = folContact.Items.Find("[User1] = " & itmClient.Links.Item(1).Item.User1)

105          For iCompteur = 1 To itmContact.Links.count
110           If itmContact.Links.Item(1).Item.User1 = itmClient.User1 Then
115             Call itmContact.Links.Remove(iCompteur)

120             Call itmContact.Save

125             Exit For
130           End If
135         Next

140         Call itmClient.Links.Remove(1)
145       Loop

150       Call itmClient.Save

155       Call rstClient.Close
160       Set rstClient = Nothing

165       Set rstContactClient = New ADODB.Recordset

170       Call rstContactClient.Open("SELECT * FROM GRB_ContactClient WHERE NoClient = " & iClientID, g_connData, adOpenForwardOnly, adLockReadOnly)

175       Do While Not rstContactClient.EOF
180         If rstContactClient.Fields("NoContact") <> iContactID Then
185           Set itmContact = folContact.Items.Find("[User1] = " & rstContactClient.Fields("NoContact"))

190           If Not itmContact Is Nothing Then
195             Call itmClient.Links.Add(itmContact)

200             Call itmClient.Save

205             Call itmContact.Links.Add(itmClient)

210             Call itmContact.Save
215           End If
220         End If

225         Call rstContactClient.MoveNext
230       Loop

235       Call rstContactClient.Close
240       Set rstContactClient = Nothing
245     Else
250       Call MsgBox("Client introuvable!", vbOKOnly, "Erreur")

255       Call rstClient.Close
260       Set rstClient = Nothing
265     End If

270     If bDejaOuvert = False Then
275       Call otlApp.Quit
280     End If

285     Set otlApp = Nothing

290     DoEvents

295     Exit Sub

AfficherErreur:

300     woups "frmClient", "LierContactClient", Err, Erl

305     fraEtatOutlook.Visible = False
End Sub

Private Sub LierContactFournisseur(ByVal iContactID As Integer, ByVal iFournisseurID As Integer)

5       On Error GoTo AfficherErreur

10      Dim otlApp         As Outlook.Application
15      Dim itmContact     As Outlook.ContactItem
20      Dim itmFRS         As Outlook.ContactItem
25      Dim folFRS         As MAPIFolder
30      Dim folContact     As MAPIFolder
35      Dim rstContactFRS  As ADODB.Recordset
40      Dim rstFRS         As ADODB.Recordset
45      Dim bDejaOuvert    As Boolean
50      Dim iCompteur      As Integer

55      fraEtatOutlook.Visible = True

60      Set otlApp = OuvrirOutlook(bDejaOuvert)

65      Set folFRS = GetFolder(otlApp, "Fournisseurs GRB")
70      Set folContact = GetFolder(otlApp, "Contacts GRB")

75      Set rstFRS = New ADODB.Recordset

80      Call rstFRS.Open("SELECT EntryIDOutlook FROM GRB_Fournisseur WHERE IDFRS = " & iFournisseurID, g_connData, adOpenForwardOnly, adLockReadOnly)

85      Set itmFRS = folFRS.Items.Find("[User1] = " & iFournisseurID)

90      If Not itmFRS Is Nothing Then
95        Do While itmFRS.Links.count > 0
100         Set itmContact = folContact.Items.Find("[User1] = " & itmFRS.Links.Item(1).Item.User1)

105         For iCompteur = 1 To itmContact.Links.count
110           If itmContact.Links.Item(1).Item.User1 = itmFRS.User1 Then
115             Call itmContact.Links.Remove(iCompteur)

120             Call itmContact.Save

125             Exit For
130           End If
135         Next

140         Call itmFRS.Links.Remove(1)
145       Loop

150       Call itmFRS.Save

155       Call rstFRS.Close
160       Set rstFRS = Nothing

165       Set rstContactFRS = New ADODB.Recordset

170       Call rstContactFRS.Open("SELECT * FROM GRB_ContactFRS WHERE NoFRS = " & iFournisseurID, g_connData, adOpenForwardOnly, adLockReadOnly)

175       Do While Not rstContactFRS.EOF
180         If rstContactFRS.Fields("NoContact") <> iContactID Then
185           Set itmContact = folContact.Items.Find("[User1] = " & rstContactFRS.Fields("NoContact"))

190           If Not itmContact Is Nothing Then
195             Call itmFRS.Links.Add(itmContact)

200             Call itmFRS.Save

205             Call itmContact.Links.Add(itmFRS)

210             Call itmContact.Save
215           End If
220         End If

225         Call rstContactFRS.MoveNext
230       Loop

235       Call rstContactFRS.Close
240       Set rstContactFRS = Nothing
245     Else
250       Call MsgBox("Fournisseur introuvable!", vbOKOnly, "Erreur")

255       Call rstFRS.Close
260       Set rstFRS = Nothing
265     End If

270     If bDejaOuvert = False Then
275       Call otlApp.Quit
280     End If

285     Set otlApp = Nothing

290     fraEtatOutlook.Visible = False

295     DoEvents

300     Exit Sub

AfficherErreur:

305     woups "frmFRS", "LierContactFournisseur", Err, Erl

310     fraEtatOutlook.Visible = False
End Sub


Public Sub cmbContact_Click()

5       On Error GoTo AfficherErreur
        
        'Quand le user selectionne un enregistrement on se posotionne dessus
10      If cmbContact.Text <> vbNullString Then
15        txtNomContact.Text = Trim$(Left$(cmbContact.Text, InStr(1, cmbContact.Text, " - ") - 1))
20      Else
25        cmbContact.Text = txtNomContact.Text
30      End If
  
35      If cmbContact.ListIndex > -1 Then
40        If m_bRenommer = False And m_bModeAjout = False Then
45          m_iNoContact = cmbContact.ItemData(cmbContact.ListIndex)
50        End If
55      End If
  
        'remplis le combo dépendant le contact sélectionné
60      Call AfficherContact

65      Exit Sub

AfficherErreur:

70      woups "frmContact", "cmbContact_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Call RemplirComboContact

15      Call HideEdMask(True)

20      Call AfficherControles(MODE_INACTIF)
  
25      Call ActiverBoutonsGroupe

30      Screen.MousePointer = vbDefault

35      Exit Sub

AfficherErreur:

40      woups "frmContact", "Form_Load", Err, Erl
End Sub

Private Sub ActiverBoutonsGroupe()

5       On Error GoTo AfficherErreur

10      CmdAdd.Enabled = g_bModificationContacts
15      CmdModif.Enabled = g_bModificationContacts
20      CmdSupp.Enabled = g_bModificationContacts
25      cmdCopier.Enabled = g_bModificationContacts
30      cmdMailList.Enabled = g_bModificationListeDistribution

35      Exit Sub

AfficherErreur:

40      woups "frmContact", "ActiverBoutonsGroupe", Err, Erl
End Sub

Private Sub Form_Unload(Cancel As Integer)

5       On Error GoTo AfficherErreur

10      Set FrmContact = Nothing

15      Exit Sub

AfficherErreur:

20      woups "frmContact", "Form_Unload", Err, Erl
End Sub
Private Sub mskTelephone_GotFocus()

5       On Error GoTo AfficherErreur

10      mskTelephone.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmContact", "mskTelephone_GotFocus", Err, Erl
End Sub

Private Sub mskTelephone_LostFocus()

5       On Error GoTo AfficherErreur

10      mskTelephone.mask = vbNullString

15      If mskTelephone.Text = "(___) ___-____" Then
20        mskTelephone.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmContact", "mskTelephone_LostFocus", Err, Erl
End Sub

Private Sub mskCellulaire_GotFocus()

5       On Error GoTo AfficherErreur

10      mskCellulaire.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmContact", "mskCellulaire_GotFocus", Err, Erl
End Sub

Private Sub mskCellulaire_LostFocus()

5       On Error GoTo AfficherErreur

10      mskCellulaire.mask = vbNullString

15      If mskCellulaire.Text = "(___) ___-____" Then
20        mskCellulaire.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmContact", "mskCellulaire_LostFocus", Err, Erl
End Sub

Private Sub mskFax_GotFocus()

5       On Error GoTo AfficherErreur

10      mskFax.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmContact", "mskFax_GotFocus", Err, Erl
End Sub

Private Sub mskFax_LostFocus()

5       On Error GoTo AfficherErreur

10      mskFax.mask = vbNullString

15      If mskFax.Text = "(___) ___-____" Then
20        mskFax.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmContact", "mskFax_LostFocus", Err, Erl
End Sub

Private Sub mskPagette_GotFocus()

5       On Error GoTo AfficherErreur

10      mskPagette.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmContact", "mskPagette_GotFocus", Err, Erl
End Sub

Private Sub mskPagette_LostFocus()

5       On Error GoTo AfficherErreur

10      mskPagette.mask = vbNullString

15      If mskPagette.Text = "(___) ___-____" Then
20        mskPagette.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmContact", "mskPagette_LostFocus", Err, Erl
End Sub

Private Sub mskTelDomicile_GotFocus()

5       On Error GoTo AfficherErreur

10      mskTelDomicile.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmContact", "mskTelDomicile_GotFocus", Err, Erl
End Sub

Private Sub mskTelDomicile_LostFocus()

5       On Error GoTo AfficherErreur

10      mskTelDomicile.mask = vbNullString

15      If mskTelDomicile.Text = "(___) ___-____" Then
20        mskTelDomicile.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmContact", "mskTelDomicile_LostFocus", Err, Erl
End Sub

Private Sub cmdRechercher_Click()

5       On Error GoTo AfficherErreur

10      Dim rstContact As ADODB.Recordset
15      Dim sSearch    As String
  
20      sSearch = txtRechercher.Text
        
25      Screen.MousePointer = vbHourglass
        
        'vide les champs
30      Call ViderBarrerChamps(True, True)
  
        'Filtre pour selection des Nomcontact
35      Set rstContact = New ADODB.Recordset
        
40      Call rstContact.Open("SELECT NomContact, Compagnie, IDContact FROM GRB_Contact WHERE Instr(1, NomContact,'" & Replace(sSearch, "'", "''") & "') > 0 Or Instr(1, Compagnie, '" & Replace(sSearch, "'", "''") & "') > 0 AND Supprimé = False ORDER BY NomContact", g_connData, adOpenDynamic, adLockOptimistic)
              
        'vide combo
45      Call cmbContact.Clear
   
50      Do While Not rstContact.EOF
55        Call cmbContact.AddItem(rstContact.Fields("NomContact") & " - " & rstContact.Fields("Compagnie"))
60        cmbContact.ItemData(cmbContact.newIndex) = rstContact.Fields("IDContact")
                  
65        Call rstContact.MoveNext
70      Loop
   
75      Call rstContact.Close
80      Set rstContact = Nothing
    
85      Screen.MousePointer = vbDefault
  
90      If cmbContact.ListCount > 0 Then
95        cmbContact.ListIndex = 0
100     End If

105     Exit Sub

AfficherErreur:

110     woups "frmContact", "cmdRechercher_Click", Err, Erl
End Sub

Private Sub txtRechercher_Change()

5       On Error GoTo AfficherErreur

10      If Len(Trim$(txtRechercher.Text)) > 0 Then
15        cmdRechercher.Enabled = True
20      Else
25        cmdRechercher.Enabled = False
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmContact", "txtRechercher_Change", Err, Erl
End Sub
