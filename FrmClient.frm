VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmClient 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clients"
   ClientHeight    =   8385
   ClientLeft      =   2160
   ClientTop       =   1020
   ClientWidth     =   9240
   Icon            =   "FrmClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "FrmClient.frx":0442
   ScaleHeight     =   8385
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEtatOutlook 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   720
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   7575
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
         TabIndex        =   17
         Top             =   840
         Width           =   6495
      End
   End
   Begin VB.TextBox txtCategorie 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datClient"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   80
      Top             =   2280
      Width           =   3855
   End
   Begin VB.CommandButton cmdMailListClient 
      Caption         =   "Ajouter au mailing list"
      Height          =   495
      Left            =   6240
      TabIndex        =   77
      Top             =   7800
      Width           =   1695
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datClient"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   6480
      Width           =   1932
   End
   Begin VB.CommandButton cmdFax 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Envoyer Fax"
      Height          =   495
      Left            =   5040
      TabIndex        =   76
      Top             =   7800
      Width           =   1095
   End
   Begin VB.TextBox txtSiteWeb 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datClient"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   6060
      Width           =   1935
   End
   Begin VB.TextBox txtNomClient 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton cmdRechercher 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rechercher"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdRafraichir 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rafraîchir"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox TxtRechercher 
      BackColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   5040
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Frame fraContact 
      BackColor       =   &H00000000&
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   4680
      TabIndex        =   37
      Top             =   2760
      Width           =   4455
      Begin VB.CommandButton cmdMailListContact 
         Caption         =   "Ajouter au mailing list"
         Height          =   495
         Left            =   2520
         TabIndex        =   66
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox txtContactTitre 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Telephonne"
         DataSource      =   "datContact"
         Height          =   288
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton CmdAddCont 
         Caption         =   "Ajouter"
         Height          =   495
         Left            =   120
         TabIndex        =   62
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtcontact_dom 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Telephonne"
         DataSource      =   "datContact"
         Height          =   288
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton cmdsupcontact 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Supprimer"
         Height          =   495
         Left            =   1320
         TabIndex        =   65
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtcontact_cell 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Fax"
         DataSource      =   "datContact"
         Height          =   288
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtcontact_fax 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Fax"
         DataSource      =   "datContact"
         Height          =   288
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton CmdRefCont 
         Caption         =   "Enregistrer"
         Height          =   495
         Left            =   120
         TabIndex        =   63
         Top             =   3960
         Width           =   1095
      End
      Begin VB.ComboBox cmbcontact 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "FrmClient.frx":334F
         Left            =   1080
         List            =   "FrmClient.frx":3351
         TabIndex        =   40
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtcontact_tel 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Telephonne"
         DataSource      =   "datContact"
         Height          =   288
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtcontact_poste 
         BackColor       =   &H00FFFFFF&
         DataField       =   "noposte"
         DataSource      =   "datContact"
         Height          =   288
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtcontact_email 
         BackColor       =   &H00FFFFFF&
         DataField       =   "E-mail"
         DataSource      =   "datContact"
         Height          =   288
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   3480
         Width           =   3255
      End
      Begin VB.TextBox txtcontact_page 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Pagette"
         DataSource      =   "datContact"
         Height          =   288
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton cmdanulcontact 
         BackColor       =   &H00C0C0C0&
         Caption         =   "A&nnuler"
         Height          =   495
         Left            =   1320
         TabIndex        =   64
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtContact 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   360
         Visible         =   0   'False
         Width           =   3255
      End
      Begin MSMask.MaskEdBox mskContactPage 
         Height          =   285
         Left            =   1080
         TabIndex        =   58
         Top             =   3120
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskContactFax 
         Height          =   285
         Left            =   1080
         TabIndex        =   54
         Top             =   2760
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskContactCell 
         Height          =   285
         Left            =   1080
         TabIndex        =   53
         Top             =   2400
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskContactDom 
         Height          =   285
         Left            =   1080
         TabIndex        =   48
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskContactTel 
         Height          =   285
         Left            =   1080
         TabIndex        =   47
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Titre "
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
         Left            =   120
         TabIndex        =   44
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Domicile"
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
         TabIndex        =   42
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
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
         Left            =   120
         TabIndex        =   41
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         Index           =   2
         Left            =   120
         TabIndex        =   52
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
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
         TabIndex        =   56
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label1 
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
         Index           =   4
         Left            =   120
         TabIndex        =   59
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail"
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
         Left            =   120
         TabIndex        =   61
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Poste"
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
         Left            =   120
         TabIndex        =   43
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   855
      End
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
      TabIndex        =   70
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdRenommer 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Renommer"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton CmdAdd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Ajouter"
      Height          =   495
      Left            =   1440
      TabIndex        =   71
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton CmdSupp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Supprimer"
      Height          =   495
      Left            =   2640
      TabIndex        =   74
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Fermer"
      Height          =   495
      Left            =   8040
      TabIndex        =   78
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton CmdEnr 
      Caption         =   "Enregistrer"
      Height          =   495
      Left            =   1440
      TabIndex        =   72
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton CmdAnul 
      Caption         =   "A&nnuler"
      Height          =   495
      Left            =   2640
      TabIndex        =   73
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton CmdModif 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Modifier"
      Height          =   495
      Left            =   3840
      TabIndex        =   75
      Top             =   7800
      Width           =   1095
   End
   Begin VB.TextBox txtPays 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datClient"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   5220
      Width           =   1935
   End
   Begin VB.TextBox txtCP 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datClient"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   5640
      Width           =   1935
   End
   Begin VB.TextBox txtProvEtat 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datClient"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox txtVille 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datClient"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   4380
      Width           =   1932
   End
   Begin VB.TextBox txtAdresse 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datClient"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   3960
      Width           =   2535
   End
   Begin VB.TextBox txtCommentaire 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datClient"
      Height          =   735
      Left            =   6480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "FrmClient.frx":3353
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtContactGRB 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datClient"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   3540
      Width           =   1935
   End
   Begin VB.TextBox txtFax 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datClient"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtTelephone 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datClient"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2700
      Width           =   1932
   End
   Begin VB.ComboBox cmbClient 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "FrmClient.frx":3359
      Left            =   1440
      List            =   "FrmClient.frx":335B
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1560
      Width           =   4335
   End
   Begin MSMask.MaskEdBox mskTelephone 
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Top             =   2700
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskFax 
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdCategorie 
      Caption         =   "..."
      Height          =   288
      Left            =   5400
      TabIndex        =   79
      Top             =   2280
      Width           =   375
   End
   Begin VB.Frame fraPotentiel 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   6240
      TabIndex        =   82
      Top             =   2280
      Width           =   2415
      Begin VB.CheckBox chkClientPotentiel 
         BackColor       =   &H00000000&
         Caption         =   "Client Potentiel      "
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
         Left            =   240
         TabIndex        =   83
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Categorie"
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
      Left            =   240
      TabIndex        =   81
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label12 
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
      TabIndex        =   36
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label12 
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
      Left            =   240
      TabIndex        =   69
      Top             =   7380
      Width           =   1215
   End
   Begin VB.Label lblDateCreation 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2004-09-14"
      Height          =   285
      Left            =   1440
      TabIndex        =   34
      Top             =   6900
      Width           =   975
   End
   Begin VB.Label lblDateModification 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2004-09-14"
      Height          =   285
      Left            =   1440
      TabIndex        =   67
      Top             =   7320
      Width           =   975
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
      TabIndex        =   35
      Top             =   6900
      Width           =   1335
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
      TabIndex        =   68
      Top             =   7335
      Width           =   1335
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
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
      Left            =   240
      TabIndex        =   32
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Site Web"
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
      Left            =   240
      TabIndex        =   30
      Top             =   6060
      Width           =   855
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Rechercher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "ContactGRB"
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
      Left            =   240
      TabIndex        =   18
      Top             =   3540
      Width           =   1095
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Pays"
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
      Left            =   240
      TabIndex        =   26
      Top             =   5220
      Width           =   975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Code Postal"
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
      Left            =   240
      TabIndex        =   28
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Prov/Etat"
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
      TabIndex        =   24
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Ville"
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
      Left            =   240
      TabIndex        =   22
      Top             =   4380
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Adresse"
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
      Left            =   240
      TabIndex        =   20
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Commentaires"
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
      Left            =   6480
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
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
      Left            =   240
      TabIndex        =   13
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone"
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
      Left            =   240
      TabIndex        =   10
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "FrmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enumMode
  MODE_CLIENT = 0
  MODE_CONTACT = 1
  MODE_INACTIF = 2
End Enum

Public m_bCategorieBeton              As Boolean
Public m_bCategoriePave               As Boolean
Public m_bCategoriePharmaceutique     As Boolean
Public m_bCategorieAgroalimentaire    As Boolean
Public m_bCategorieMeuble             As Boolean
Public m_bCategorieMeunerie           As Boolean
Public m_bCategorieManufacturier      As Boolean
Public m_bCategorieConsultant         As Boolean
Public m_bCategorieAsphalte           As Boolean
Public m_bCategorieICPI               As Boolean
Public m_bCategorieProduitsChimiques  As Boolean
Public m_bCategorieAutre              As Boolean

'Choix d'impression
Public m_bImpressionAnnuler           As Boolean
Public m_bImpressionVille             As Boolean
Public m_bImpressionCategorie         As Boolean
Public m_bImpressionPotentiel         As Boolean
Public m_bImpressionFacturer          As Boolean

'Choix d'impression de categorie
Public m_bImpressionBeton             As Boolean
Public m_bImpressionPave              As Boolean
Public m_bImpressionPharmaceutique    As Boolean
Public m_bImpressionAgroAlimentaire   As Boolean
Public m_bImpressionMeuble            As Boolean
Public m_bImpressionMeunerie          As Boolean
Public m_bImpressionManufacturier     As Boolean
Public m_bImpressionConsultant        As Boolean
Public m_bImpressionAsphalte          As Boolean
Public m_bImpressionICPI              As Boolean
Public m_bImpressionProduitsChimiques As Boolean
Public m_bImpressionAutre             As Boolean
  
Private m_bModeAjoutContact           As Boolean
Private m_bModeAjoutClient            As Boolean
Private m_iNoContact                  As Integer
Private m_iNoClient                   As Integer
Private m_bRenommer                   As Boolean
Private m_bNewContact                 As Boolean

Public m_bAnnulerDistList             As Boolean
Public m_otlDistList                  As Outlook.DistListItem

Public m_bAnnulerVille                As Boolean
Public m_sVille                       As String
 
Private m_eMode                       As enumMode

Private Sub RemplirComboClient()

5       On Error GoTo AfficherErreur
        
        'Rempli le combo des clients
10      Dim rstClient As ADODB.Recordset
  
        'Il faut vider le combo avant de le remplir
15      Call cmbclient.Clear
 
20      Set rstClient = New ADODB.Recordset

25      Call rstClient.Open("SELECT NomClient, IDClient FROM GRB_Client WHERE Supprimé = False", g_connData, adOpenDynamic, adLockOptimistic)
     
        'Tant que ce n'est pas la fin des enregistrements
30      Do While Not rstClient.EOF
          'Ajout du nom du client dans le combo
35        Call cmbclient.AddItem(Trim(rstClient.Fields("NomClient")))
      
          'Ajout du numéro du client dans le ItemData du combo
40        cmbclient.ItemData(cmbclient.newIndex) = rstClient.Fields("IDClient")
    
45        Call rstClient.MoveNext
50      Loop
  
55      Call rstClient.Close
60      Set rstClient = Nothing
  
65      If cmbclient.ListCount > 0 Then
70        cmbclient.ListIndex = 0
75      End If

80      Exit Sub

AfficherErreur:

85      woups "frmClient", "RemplirComboClient", Err, Erl
End Sub

Private Sub AfficherClient()

5       On Error GoTo AfficherErreur
        
        'Affiche le client sélectionné
10      Dim rstClient As ADODB.Recordset
  
15      Set rstClient = New ADODB.Recordset
  
20      Call rstClient.Open("SELECT * FROM GRB_Client WHERE IDClient = " & m_iNoClient, g_connData, adOpenDynamic, adLockOptimistic)

25      Call ViderBarrerChamps(True, True)
                
        'Telephonne
30      If Not IsNull(rstClient.Fields("Telephonne")) Then
35        txtTelephone.Text = rstClient.Fields("Telephonne")
40      End If
  
        'Fax
45      If Not IsNull(rstClient.Fields("Fax")) Then
50        txtFax.Text = rstClient.Fields("Fax")
55      End If
 
        'ContactGRB
60      If Not IsNull(rstClient.Fields("ContactGRB")) Then
65        txtContactGRB.Text = rstClient.Fields("ContactGRB")
70      End If
  
        'Email
75      If Not IsNull(rstClient.Fields("Email")) Then
80        txtEmail.Text = rstClient.Fields("Email")
85      End If
  
        'AdresseLiv
90      If Not IsNull(rstClient.Fields("AdresseLiv")) Then
95        txtAdresse.Text = rstClient.Fields("AdresseLiv")
100      End If
  
        'VilleLiv
105     If Not IsNull(rstClient.Fields("VilleLiv")) Then
120       txtVille.Text = rstClient.Fields("VilleLiv")
125     End If
  
        'Prov/EtatLiv
130     If Not IsNull(rstClient.Fields("Prov/EtatLiv")) Then
135       txtProvEtat.Text = rstClient.Fields("Prov/EtatLiv")
140     End If
  
        'PaysLiv
145     If Not IsNull(rstClient.Fields("PaysLiv")) Then
150       txtPays.Text = rstClient.Fields("PaysLiv")
155     End If
  
        'CodePostalLiv
160     If Not IsNull(rstClient.Fields("CodePostalLiv")) Then
165       txtCP.Text = rstClient.Fields("CodePostalLiv")
170     End If
    
        'Commentaire
175     If Not IsNull(rstClient.Fields("Commentaire")) Then
180       txtcommentaire.Text = rstClient.Fields("Commentaire")
185     End If

        'Site Web
190     If Not IsNull(rstClient.Fields("SiteWeb")) Then
195       txtSiteWeb.Text = rstClient.Fields("SiteWeb")
200     End If

        'Création
205     If Not IsNull(rstClient.Fields("DateCréation")) Then
210       lblDateCreation.Caption = rstClient.Fields("DateCréation")
215     End If

        'User Création
220     If Not IsNull(rstClient.Fields("UserCréation")) Then
225       lblUserCreation.Caption = "Par : " & rstClient.Fields("UserCréation")
230     End If

        'Modification
235     If Not IsNull(rstClient.Fields("DateModification")) Then
240       lblDateModification.Caption = rstClient.Fields("DateModification")
245     End If

        'User Modification
250     If Not IsNull(rstClient.Fields("UserModification")) Then
255       lblUserModification.Caption = "Par : " & rstClient.Fields("UserModification")
260     End If

265     'Client Potentiel
270     If rstClient.Fields("Potentiel") = True Then
275       chkClientPotentiel.Value = vbChecked
280     End If

285     m_bCategorieBeton = rstClient.Fields("Béton")
290     m_bCategoriePave = rstClient.Fields("Pavé")
295     m_bCategoriePharmaceutique = rstClient.Fields("Pharmaceutique")
300     m_bCategorieAgroalimentaire = rstClient.Fields("Agroalimentaire")
305     m_bCategorieMeuble = rstClient.Fields("Meuble")
310     m_bCategorieMeunerie = rstClient.Fields("Meunerie")
315     m_bCategorieManufacturier = rstClient.Fields("Manufacturier")
320     m_bCategorieConsultant = rstClient.Fields("Consultant")
325     m_bCategorieAsphalte = rstClient.Fields("Asphalte")
330     m_bCategorieICPI = rstClient.Fields("ICPI")
335     m_bCategorieProduitsChimiques = rstClient.Fields("ProduitsChimiques")
340     m_bCategorieAutre = rstClient.Fields("Autre")

345     Call AfficherCategories
    
350     Call rstClient.Close
355     Set rstClient = Nothing

360     Exit Sub

AfficherErreur:

365     woups "frmClient", "AfficherClient", Err, Erl
End Sub

Private Sub AfficherCategories()

5       On Error GoTo AfficherErreur

10      txtCategorie.Text = ""

15      If m_bCategorieBeton = True Then
20        txtCategorie.Text = "Béton"
25      End If

30      If m_bCategoriePave = True Then
35        If Trim$(txtCategorie.Text) <> "" Then
40          txtCategorie.Text = txtCategorie.Text & ", Pavé"
45        Else
50          txtCategorie.Text = "Pavé"
55        End If
60      End If

65      If m_bCategoriePharmaceutique = True Then
70        If Trim$(txtCategorie.Text) <> "" Then
75          txtCategorie.Text = txtCategorie.Text & ", Pharmaceutique"
80        Else
85          txtCategorie.Text = "Pharmaceutique"
90        End If
95      End If

100     If m_bCategorieAgroalimentaire = True Then
105       If Trim$(txtCategorie.Text) <> "" Then
110         txtCategorie.Text = txtCategorie.Text & ", Agroalimentaire"
115       Else
120         txtCategorie.Text = "Agroalimentaire"
125       End If
130     End If

135     If m_bCategorieMeuble = True Then
140       If Trim$(txtCategorie.Text) <> "" Then
145         txtCategorie.Text = txtCategorie.Text & ", Meuble"
150       Else
155         txtCategorie.Text = "Meuble"
160       End If
165     End If

170     If m_bCategorieMeunerie = True Then
175       If Trim$(txtCategorie.Text) <> "" Then
180         txtCategorie.Text = txtCategorie.Text & ", Meunerie"
185       Else
190         txtCategorie.Text = "Meunerie"
195       End If
200     End If

205     If m_bCategorieManufacturier = True Then
210       If Trim$(txtCategorie.Text) <> "" Then
215         txtCategorie.Text = txtCategorie.Text & ", Manufacturier"
220       Else
225         txtCategorie.Text = "Manufacturier"
230       End If
235     End If

240     If m_bCategorieConsultant = True Then
245       If Trim$(txtCategorie.Text) <> "" Then
250         txtCategorie.Text = txtCategorie.Text & ", Consultant"
255       Else
260         txtCategorie.Text = "Consultant"
265       End If
270     End If

275     If m_bCategorieAsphalte = True Then
280       If Trim$(txtCategorie.Text) <> "" Then
285         txtCategorie.Text = txtCategorie.Text & ", Asphalte"
290       Else
295         txtCategorie.Text = "Asphalte"
300       End If
305     End If

310     If m_bCategorieICPI = True Then
315       If Trim$(txtCategorie.Text) <> "" Then
320         txtCategorie.Text = txtCategorie.Text & ", ICPI"
325       Else
330         txtCategorie.Text = "ICPI"
335       End If
340     End If

345     If m_bCategorieProduitsChimiques = True Then
350       If Trim$(txtCategorie.Text) <> "" Then
355         txtCategorie.Text = txtCategorie.Text & ", Produits chimiques"
360       Else
365         txtCategorie.Text = "Produits chimiques"
370       End If
375     End If

380     If m_bCategorieAutre = True Then
385       If Trim$(txtCategorie.Text) <> "" Then
390         txtCategorie.Text = txtCategorie.Text & ", Autre"
395       Else
400         txtCategorie.Text = "Autre"
405       End If
410     End If

415     Exit Sub

AfficherErreur:

420     woups "frmClient", "AfficherCategorie", Err, Erl
End Sub

Public Sub AfficherContact()

5       On Error GoTo AfficherErreur
        
        ''''''''''''''''''''''''''''''''''''''''
        'affiche les contacts de l'employé selectionné'
        ''''''''''''''''''''''''''''''''''''''''
10      Dim rstContact As ADODB.Recordset

        'Ouverture de la table contact
15      Set rstContact = New ADODB.Recordset
        
20      Call rstContact.Open("SELECT * FROM GRB_contact WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)
    
        'VIDE les champs
25      If m_bModeAjoutContact = True Then
30        If m_bNewContact = True Then
35          Call ViderBarrerChampsContact(False, True)
40        Else
45          Call ViderBarrerChampsContact(True, True)
50        End If
55      Else
60        Call ViderBarrerChampsContact(True, True)
65      End If

    
        'REMPLIS LES CHAMPS si il y a enregistrement
70      If Not rstContact.EOF Then
75        If Not IsNull(rstContact.Fields("Titre")) Then
80          txtContactTitre.Text = rstContact.Fields("Titre")
85        End If

90        If Not IsNull(rstContact.Fields("cellulaire")) Then
95          txtcontact_cell.Text = rstContact.Fields("cellulaire")
100       End If
      
105       If Not IsNull(rstContact.Fields("pagette")) Then
110         txtcontact_page.Text = rstContact.Fields("pagette")
115       End If
      
120       If Not IsNull(rstContact.Fields("telephonne")) Then
125         txtcontact_tel.Text = rstContact.Fields("telephonne")
130       End If
        
135       If Not IsNull(rstContact.Fields("fax")) Then
140         txtcontact_fax.Text = rstContact.Fields("Fax")
145       End If
        
150       If Not IsNull(rstContact.Fields("e-mail")) Then
155         txtcontact_email.Text = rstContact.Fields("e-mail")
160       End If
        
165       If Not IsNull(rstContact.Fields("noposte")) Then
170         txtcontact_poste.Text = rstContact.Fields("noposte")
175       End If
        
180       If Not IsNull(rstContact.Fields("teldomicile")) Then
185         txtcontact_dom.Text = rstContact.Fields("teldomicile")
190       End If
195     End If
      
        'ferme la table
200     Call rstContact.Close
205     Set rstContact = Nothing

210     Exit Sub

AfficherErreur:

215     woups "frmClient", "AfficherContact", Err, Erl
End Sub

Private Sub cmbContact_Click()

5       On Error GoTo AfficherErreur
        
        ''''''''''''''''''''''''''''''''''
        'affiche employé sélectioné
        ''''''''''''''''''''''''''''''''''
10      If cmbContact.ListIndex <> -1 Then
15        m_iNoContact = cmbContact.ItemData(cmbContact.ListIndex)
20      End If

25      Call AfficherContact

30      Exit Sub

AfficherErreur:

35      woups "frmClient", "cmbContact_Click", Err, Erl
End Sub

Private Sub cmdanulcontact_Click()

5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbHourglass

15      Call AfficherControles(MODE_INACTIF)

20      If m_bNewContact = True Then
25        Call HideEdMaskContact(True)

30        m_bNewContact = False
35      End If
              
        'n'est plus en mode ajouter
40      m_bModeAjoutContact = False
  
45      txtNomClient.Visible = False
50      txtNomClient.Locked = False

        'remplis combo contact
55      Call RemplirComboContact
  
60      Screen.MousePointer = vbDefault

65      Exit Sub

AfficherErreur:

70      woups "frmClient", "cmdanulcontact_Click", Err, Erl
End Sub

Private Sub cmdCategorie_Click()

5       On Error GoTo AfficherErreur

10      Call frmCategorieClient.AfficherClient

15      Call AfficherCategories

20      Exit Sub

AfficherErreur:

25      woups "frmClient", "cmdCategorie_Click", Err, Erl
End Sub

Private Sub cmdFax_Click()

5       On Error GoTo AfficherErreur

10      If cmbclient.ListCount > 0 Then
15        If cmbContact.ListIndex > -1 Then
20          Call frmreport.Afficher(cmbclient.ItemData(cmbclient.ListIndex), cmbContact.ItemData(cmbContact.ListIndex), FRM_CLIENTS)
25        Else
30          Call frmreport.Afficher(cmbclient.ItemData(cmbclient.ListIndex), 0, FRM_CLIENTS)
35        End If
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmClient", "cmdFax_Click", Err, Erl
End Sub

Private Sub cmdMailListContact_Click()

5       On Error GoTo AfficherErreur

10      Dim otlApp       As Outlook.Application
15      Dim folContact   As Outlook.MAPIFolder
20      Dim itmContact() As Outlook.ContactItem
25      Dim otlRecipient As Outlook.Recipient
30      Dim bDejaOuvert  As Boolean
35      Dim rstContact   As ADODB.Recordset
40      Dim sIDContact() As String
45      Dim sContact()   As String
50      Dim iCompteur    As Integer
55      Dim bArrayVide   As Boolean
60      Dim bNouveau     As Boolean
65      Dim iResult      As Integer
70      Dim bPlein       As Boolean
75      Dim sMsgPlein    As String
80      Dim iNbreRendu   As Integer

85      If cmbContact.ListIndex <> -1 Then
90        bArrayVide = True

95        iResult = MsgBox("Voulez-vous ajouter tous les contacts à la liste de distribution?" & vbNewLine & _
                    "Oui - Tous les contacts" & vbNewLine & _
                    "Non - Contact affiché seulement", vbYesNoCancel)
                    
100       If iResult = vbYes Then
105         Set rstContact = New ADODB.Recordset

110         Call rstContact.Open("SELECT * FROM GRB_ContactClient INNER JOIN GRB_Contact ON GRB_ContactClient.NoContact = GRB_Contact.IDContact WHERE GRB_ContactClient.NoClient = " & cmbclient.ItemData(cmbclient.ListIndex) & " ORDER BY NomContact", g_connData, adOpenForwardOnly, adLockOptimistic)
                       
115         iCompteur = 0
            
120         Do While Not rstContact.EOF
125           If Not IsNull(rstContact.Fields("E-mail")) Then
130             If Trim(rstContact.Fields("E-mail")) <> "" Then
135               bArrayVide = False

140               ReDim Preserve sIDContact(0 To iCompteur)
145               ReDim Preserve itmContact(0 To iCompteur)
150               ReDim Preserve sContact(0 To iCompteur)

155               sIDContact(iCompteur) = rstContact.Fields("IDContact")
160               sContact(iCompteur) = rstContact.Fields("NomContact")

165               iCompteur = iCompteur + 1
170             End If
175           End If

180           Call rstContact.MoveNext
185         Loop
190       Else
195         If iResult = vbNo Then
200           If Trim$(txtcontact_email.Text) <> "" Then
205             bArrayVide = False

210             ReDim Preserve sIDContact(0 To 0)
215             ReDim Preserve itmContact(0 To 0)
220             ReDim Preserve sContact(0 To 0)

225             sIDContact(0) = cmbContact.ItemData(cmbContact.ListIndex)
230             sContact(0) = cmbContact.Text
235           End If
240         Else
245           Exit Sub
250         End If
255       End If
            
260       If bArrayVide = False Then
265         Set otlApp = OuvrirOutlook(bDejaOuvert)

270         lblEtatOutlook.Caption = "Recherche des listes de distribution..."

275         fraEtatOutlook.Visible = True

280         Call frmChoixMailList.Afficher(Me, otlApp)

285         If m_bAnnulerDistList = False Then
290           lblEtatOutlook.Caption = "Ajout du contact dans la liste de distribution..."
 
295           fraEtatOutlook.Visible = True

300           Set folContact = GetFolder(otlApp, "Contacts GRB")

305           For iCompteur = 0 To UBound(sIDContact)
310             Set itmContact(iCompteur) = folContact.Items.Find("[User1] = " & sIDContact(iCompteur))
315           Next

320           bPlein = False

325           For iCompteur = 0 To UBound(itmContact)
330             If Not itmContact(iCompteur) Is Nothing Then
335               If m_otlDistList.MemberCount < 10 Then
340                 Set otlRecipient = otlApp.Session.CreateRecipient(itmContact(iCompteur).Email1DisplayName)

345                 If otlRecipient.Resolve = True Then
350                   Call m_otlDistList.AddMember(otlRecipient)
      
355                   Call m_otlDistList.Save
360                 Else
365                   Call MsgBox("Impossible d'ajouter le contact '" & sContact(iCompteur) & "' !", vbOKOnly, "Erreur")
370                 End If
375               Else
380                 bPlein = True

385                 Exit For
390               End If
395             Else
400               Call MsgBox("Contact '" & sContact(iCompteur) & "' introuvable!", vbOKOnly, "Erreur")
405             End If
410           Next

415           If bPlein = True Then
420             sMsgPlein = "Les contacts suivants n'ont pas pu être ajouté car la liste contient déjà 10 contacts!" & vbNewLine & _
                            vbNewLine

425             iNbreRendu = iCompteur

430             For iCompteur = iNbreRendu To UBound(sContact)
435               sMsgPlein = sMsgPlein & sContact(iCompteur)

440               If iCompteur < UBound(sContact) Then
445                 sMsgPlein = sMsgPlein & vbNewLine
450               End If
455             Next

460             Call MsgBox(sMsgPlein, vbOKOnly, "Erreur")
465           End If
470         End If

475         If bDejaOuvert = False Then
480           Call otlApp.Quit
485         End If

490         Set otlApp = Nothing

495         fraEtatOutlook.Visible = False
500       Else
505         Call MsgBox("Le ou les contacts n'ont pas d'email!", vbOKOnly, "Erreur")
510       End If
515     Else
520       Call MsgBox("Aucun contact sélectionné!", vbOKOnly, "Erreur")
525     End If

530     Exit Sub

AfficherErreur:

535     If Err.number = 287 And Erl = 305 Then
540       Call MsgBox("La liste de distribution est pleine!", vbOKOnly, "Erreur")
545     Else
550       woups "frmClient", "cmdMailListContact_Click", Err, Erl
555     End If

560     fraEtatOutlook.Visible = False
End Sub

Private Sub cmdMailListClient_Click()

5       On Error GoTo AfficherErreur

10      Dim otlApp       As Outlook.Application
15      Dim folClient    As Outlook.MAPIFolder
20      Dim itmClient    As Outlook.ContactItem
25      Dim otlRecipient As Outlook.Recipient
30      Dim bDejaOuvert  As Boolean

35      If cmbclient.ListIndex <> -1 Then
40        If Trim$(txtEmail.Text) <> "" Then
45          Set otlApp = OuvrirOutlook(bDejaOuvert)

55          lblEtatOutlook.Caption = "Recherche des listes de distribution..."

60          fraEtatOutlook.Visible = True

65          Call frmChoixMailList.Afficher(Me, otlApp)

70          If m_bAnnulerDistList = False Then
75            lblEtatOutlook.Caption = "Ajout du client dans la liste de distribution..."

80            fraEtatOutlook.Visible = True

85            If m_otlDistList.MemberCount < 10 Then
90              Set folClient = GetFolder(otlApp, "Clients GRB")

95              Set itmClient = folClient.Items.Find("[User1] = " & cmbclient.ItemData(cmbclient.ListIndex))

100             If Not itmClient Is Nothing Then
105               Set otlRecipient = otlApp.Session.CreateRecipient(itmClient.Email1DisplayName)

110               If otlRecipient.Resolve = True Then
115                 Call m_otlDistList.AddMember(otlRecipient)
      
120                 Call m_otlDistList.Save
125               Else
130                 Call MsgBox("Impossible de trouver le client!", vbOKOnly, "Erreur")
135               End If
140             Else
145               Call MsgBox("Client introuvable!", vbOKOnly, "Erreur")
150             End If
155           Else
160             Call MsgBox("Cette liste de distribution contient déjà 10 contacts!" & vbNewLine & _
                            vbNewLine & _
                            "Veuillez en choisir une autre.", vbOKOnly, "Erreur")
165           End If
170         End If

175         If bDejaOuvert = False Then
180           Call otlApp.Quit
185         End If

190         Set otlApp = Nothing

195         fraEtatOutlook.Visible = False
200       Else
205         Call MsgBox("Ce client n'a pas d'email!", vbOKOnly, "Erreur")
210       End If
215     Else
220       Call MsgBox("Aucun client sélectionné!", vbOKOnly, "Erreur")
225     End If

230     Exit Sub

AfficherErreur:

235     If Err.number = 287 And Erl = 115 Then
240       Call MsgBox("La liste de distribution est pleine!", vbOKOnly, "Erreur")
245     Else
250       woups "frmClient", "cmdMailListClient_Click", Err, Erl
255     End If

260     fraEtatOutlook.Visible = False
End Sub

Private Sub cmdreport_Click()

5       On Error GoTo AfficherErreur

        'Impression de la liste des clients
10      Dim rstClient As ADODB.Recordset
15      Dim sWhere    As String

20      Call dlgImpressionClient.Show(vbModal)

25      If m_bImpressionAnnuler = False Then
30        Set rstClient = New ADODB.Recordset

35        If m_bImpressionVille = True Then
40          Call frmChoixVille.Show(vbModal)

45          If m_bAnnulerVille = False Then
50            Call rstClient.Open("SELECT * FROM GRB_Client WHERE VilleLiv = '" & Replace(m_sVille, "'", "''") & "' AND Supprimé = False ORDER BY NomClient", g_connData, adOpenForwardOnly, adLockReadOnly)
55          Else
60            Set rstClient = Nothing

65            Exit Sub
70          End If
75        Else
80          If m_bImpressionCategorie = True Then
85            Call frmCategorieClient.AfficherImpression

90            If m_bImpressionBeton = True Then
95              sWhere = "Béton = True"
100           End If

105           If m_bImpressionPave = True Then
110             If Trim$(sWhere) <> "" Then
115               sWhere = sWhere & " OR Pavé = True"
120             Else
125               sWhere = "Pavé = True"
130             End If
135           End If

140           If m_bImpressionPharmaceutique = True Then
145             If Trim$(sWhere) <> "" Then
150               sWhere = sWhere & " OR Pharmaceutique = True"
155             Else
160               sWhere = "Pharmaceutique = True"
165             End If
170           End If
                
175           If m_bImpressionAgroAlimentaire = True Then
180             If Trim$(sWhere) <> "" Then
185               sWhere = sWhere & " OR Agroalimentaire = True"
190             Else
195               sWhere = "Agroalimentaire = True"
200             End If
205           End If

210           If m_bImpressionMeuble = True Then
215             If Trim$(sWhere) <> "" Then
220               sWhere = sWhere & " OR Meuble = True"
225             Else
230               sWhere = "Meuble = True"
235             End If
240           End If

245           If m_bImpressionMeunerie = True Then
250             If Trim$(sWhere) <> "" Then
255               sWhere = sWhere & " OR Meunerie = True"
260             Else
265               sWhere = "Meunerie = True"
270             End If
275           End If

280           If m_bImpressionManufacturier = True Then
285             If Trim$(sWhere) <> "" Then
290               sWhere = sWhere & " OR Manufacturier = True"
295             Else
300               sWhere = "Manufacturier = True"
305             End If
310           End If

315           If m_bImpressionConsultant = True Then
320             If Trim$(sWhere) <> "" Then
325               sWhere = sWhere & " OR Consultant = True"
330             Else
335               sWhere = "Consultant = True"
340             End If
345           End If

350           If m_bImpressionAsphalte = True Then
355             If Trim$(sWhere) <> "" Then
360               sWhere = sWhere & " OR Asphalte = True"
365             Else
370               sWhere = "Asphalte = True"
375             End If
380           End If

385           If m_bImpressionICPI = True Then
390             If Trim$(sWhere) <> "" Then
395               sWhere = sWhere & " OR ICPI = True"
400             Else
405               sWhere = "ICPI = True"
410             End If
415           End If

420           If m_bImpressionProduitsChimiques = True Then
425             If Trim$(sWhere) <> "" Then
430               sWhere = sWhere & " OR ProduitsChimiques = True"
435             Else
440               sWhere = "ProduitsChimiques = True"
445             End If
450           End If
                
455           If m_bImpressionAutre = True Then
460             If Trim$(sWhere) <> "" Then
465               sWhere = sWhere & " OR Autre = True"
470             Else
475               sWhere = "Autre = True"
480             End If
485           End If

490           If Trim$(sWhere) <> "" Then
495             Call rstClient.Open("SELECT * FROM GRB_Client WHERE " & sWhere & " AND Supprimé = False ORDER BY NomClient", g_connData, adOpenForwardOnly, adLockReadOnly)
500           Else
505             Call rstClient.Open("SELECT * FROM GRB_Client WHERE Supprimé = False ORDER BY NomClient", g_connData, adOpenForwardOnly, adLockReadOnly)
510           End If
515         Else
520           If m_bImpressionPotentiel = True Then
525             Call rstClient.Open("SELECT * FROM GRB_Client WHERE Potentiel = True AND Supprimé = False ORDER BY NomClient", g_connData, adOpenForwardOnly, adLockReadOnly)
530           Else
535             If m_bImpressionFacturer = True Then
540               Call rstClient.Open("SELECT * FROM GRB_Client WHERE Potentiel = False AND Supprimé = False ORDER BY NomClient", g_connData, adOpenForwardOnly, adLockReadOnly)
545             Else
550               Call rstClient.Open("SELECT * FROM GRB_Client WHERE Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)
555             End If
560           End If
565         End If
570       End If
       
575       Screen.MousePointer = vbHourglass
   
          'set le rapport
580       Set DR_ListeClient.DataSource = rstClient
    
585       DR_ListeClient.Orientation = rptOrientLandscape

590       Call DR_ListeClient.Show(vbModal)
    
595       Call rstClient.Close
600       Set rstClient = Nothing
    
605       Screen.MousePointer = vbDefault
610     End If

615     Exit Sub

AfficherErreur:

620     woups "frmClient", "cmdreport_Click", Err, Erl
End Sub

Private Sub AfficherControles(ByVal eMode As enumMode)

5       On Error GoTo AfficherErreur
        
        'Activation des controles
10      Dim bCmbClient       As Boolean
15      Dim bTxtNomClient    As Boolean
20      Dim bCmbContact      As Boolean
25      Dim bTxtContact      As Boolean
30      Dim bTxtRechercher   As Boolean
35      Dim bCmdAdd          As Boolean
40      Dim bCmdEnr          As Boolean
45      Dim bCmdModif        As Boolean
50      Dim bCmdSupp         As Boolean
55      Dim bFraContact      As Boolean
60      Dim bCmdAnul         As Boolean
65      Dim bCmdQuit         As Boolean
70      Dim bCmdAddCont      As Boolean
75      Dim bCmdSupContact   As Boolean
80      Dim bCmdAnulContact  As Boolean
85      Dim bCmdRenommer     As Boolean
90      Dim bCmdRafraichir   As Boolean
95      Dim bCmdImprimer     As Boolean
100     Dim bCmdRefCont      As Boolean
105     Dim bCmdRechercher   As Boolean
110     Dim bFax             As Boolean
115     Dim bMailListClient  As Boolean
120     Dim bMailListContact As Boolean
125     Dim bCategorie       As Boolean

130     m_eMode = eMode
  
135     Select Case eMode
          'Mode ajout et modif d'un client
          Case MODE_CLIENT:
140         bTxtNomClient = True
145         bCmdEnr = True
150         bCmdAnul = True
155         bCategorie = True
      
          'Mode ajout et modif d'un contact
          Case MODE_CONTACT:
160         bFraContact = True
165         bTxtNomClient = True
170         bCmdAnulContact = True
175         bCmdRefCont = True
    
180         If m_bNewContact = True Then
185           bTxtContact = True
190         Else
195           bCmbContact = True
200         End If
    
          Case MODE_INACTIF:
205         bFraContact = True
210         bCmbClient = True
215         bCmdImprimer = True
220         bTxtRechercher = True
225         bCmdRenommer = True
230         bCmdRafraichir = True
235         bCmdAdd = True
240         bCmdSupp = True
245         bCmdModif = True
250         bCmdQuit = True
255         bCmdAddCont = True
260         bCmdSupContact = True
265         bFax = True
270         bCmbContact = True
275         bMailListClient = True
280         bMailListContact = True
      
285         If Len(txtRechercher.Text) > 0 Then
290           bCmdRechercher = True
295         End If
300     End Select
  
305     cmbclient.Visible = bCmbClient
310     txtNomClient.Visible = bTxtNomClient
315     fracontact.Visible = bFraContact
320     txtRechercher.Enabled = bTxtRechercher
325     cmdRechercher.Enabled = bCmdRechercher
330     cmdRafraichir.Enabled = bCmdRafraichir
335     cmdrenommer.Enabled = bCmdRenommer
340     cmdReport.Visible = bCmdImprimer
345     CmdAdd.Visible = bCmdAdd
350     CmdSupp.Visible = bCmdSupp
355     CmdModif.Visible = bCmdModif
360     CmdQuit.Visible = bCmdQuit
365     CmdAnul.Visible = bCmdAnul
370     CmdEnr.Visible = bCmdEnr
375     CmdAddCont.Visible = bCmdAddCont
380     cmdsupcontact.Visible = bCmdSupContact
385     cmdanulcontact.Visible = bCmdAnulContact
390     CmdRefCont.Visible = bCmdRefCont
395     cmdFax.Visible = bFax
400     txtcontact.Visible = bTxtContact
405     cmbContact.Visible = bCmbContact
410     cmdMailListClient.Visible = bMailListClient
415     cmdMailListContact.Visible = bMailListContact
420     cmdCategorie.Visible = bCategorie

425     Exit Sub

AfficherErreur:

430     woups "frmClient", "AfficherControles", Err, Erl
End Sub

Private Sub CmdAdd_Click()

5       On Error GoTo AfficherErreur
        
        'proc qui permet dajouter un client a la BD
10      Dim sName As String
 
15      Call AfficherControles(MODE_CLIENT)
  
        'On procede a la saisie du nom du du contact
20      sName = InputBox("Veuillez entrer le nom du client" & vbNewLine & _
                         vbNewLine & _
                         "SVP, respectez le bon orthographe!", "SAISIE DU NOM", "Nom du client")
    
25      If sName <> vbNullString Then
30        If Not ComboContient(cmbclient, sName) Then
35          Screen.MousePointer = vbHourglass
          
40          m_bModeAjoutClient = True
        
            'On montre les maskEdBox
45          Call HideEdMask(False)
                
            'On affiche le nom du nouveau client dans le textbox
            'pour éviter le ScrollDown durant l'ajout
50          txtNomClient.Text = sName
                
55          Call ViderBarrerChamps(False, True)
        
60          Call mskTelephone.SetFocus
       
65          Screen.MousePointer = vbDefault
70        Else
75          Call MsgBox("Le client " & sName & " existe deja", vbCritical)
      
80          Call AfficherControles(MODE_INACTIF)
85        End If
90      Else
95        Call AfficherControles(MODE_INACTIF)
100     End If

105     Exit Sub

AfficherErreur:

110     woups "frmClient", "CmdAdd_Click", Err, Erl
End Sub

Private Sub ViderBarrerChamps(ByVal bLocked As Boolean, ByVal bVider As Boolean)

5       On Error GoTo AfficherErreur
        
        'Cette procédure vide et unlock tous les textbox pour pouvoir ajouter
10      If bVider = True Then
15        txtTelephone.Text = vbNullString
20        mskTelephone.Text = vbNullString
25        txtFax.Text = vbNullString
30        mskFax.Text = vbNullString
35        txtContactGRB.Text = vbNullString
40        txtEmail.Text = vbNullString
45        txtAdresse.Text = vbNullString
50        txtVille.Text = vbNullString
55        txtProvEtat.Text = vbNullString
60        txtPays.Text = vbNullString
65        txtCP.Text = vbNullString
70        txtcommentaire.Text = vbNullString
75        txtSiteWeb.Text = vbNullString
80        lblDateCreation.Caption = vbNullString
85        lblUserCreation.Caption = vbNullString
90        lblDateModification.Caption = vbNullString
95        lblUserModification.Caption = vbNullString
100       txtCategorie.Text = vbNullString
105       chkClientPotentiel.Value = vbUnchecked

110       m_bCategorieBeton = False
115       m_bCategoriePave = False
120       m_bCategoriePharmaceutique = False
125       m_bCategorieAgroalimentaire = False
130       m_bCategorieMeuble = False
135       m_bCategorieMeunerie = False
140       m_bCategorieManufacturier = False
145       m_bCategorieConsultant = False
150       m_bCategorieAsphalte = False
155       m_bCategorieICPI = False
160       m_bCategorieProduitsChimiques = False
165       m_bCategorieAutre = False
170     End If
  
175     txtTelephone.Locked = bLocked
180     txtFax.Locked = bLocked
185     txtContactGRB.Locked = bLocked
190     txtEmail.Locked = bLocked
195     txtAdresse.Locked = bLocked
200     txtVille.Locked = bLocked
205     txtProvEtat.Locked = bLocked
210     txtPays.Locked = bLocked
215     txtCP.Locked = bLocked
220     txtcommentaire.Locked = bLocked
225     txtSiteWeb.Locked = bLocked
230     fraPotentiel.Enabled = Not bLocked

235     Exit Sub

AfficherErreur:

240     woups "frmClient", "ViderBarrerChamps", Err, Erl
End Sub

Private Sub ViderBarrerChampsContact(ByVal bLocked As Boolean, ByVal bVider As Boolean)

5       On Error GoTo AfficherErreur

10      If bVider = True Then
15        txtContactTitre.Text = vbNullString
20        txtcontact_cell.Text = vbNullString
25        txtcontact_dom.Text = vbNullString
30        txtcontact_email.Text = vbNullString
35        txtcontact_fax.Text = vbNullString
40        txtcontact_page.Text = vbNullString
45        txtcontact_poste.Text = vbNullString
50        txtcontact_tel.Text = vbNullString
55      End If

60      txtContactTitre.Locked = bLocked
65      txtcontact_cell.Locked = bLocked
70      txtcontact_dom.Locked = bLocked
75      txtcontact_email.Locked = bLocked
80      txtcontact_fax.Locked = bLocked
85      txtcontact_page.Locked = bLocked
90      txtcontact_poste.Locked = bLocked
95      txtcontact_tel.Locked = bLocked

100     Exit Sub

AfficherErreur:

105     woups "frmClient", "ViderBarrerChampsContact", Err, Erl
End Sub


Private Sub CmdAddCont_Click()

5       On Error GoTo AfficherErreur
            
        'Pour faire l'ajout d'un contact
10      Dim sNom       As String
15      Dim rstContact As ADODB.Recordset
20      Dim bAjouter   As Boolean

25      If cmbclient.ListCount > 0 Then
30        m_bModeAjoutContact = True

35        If MsgBox("Voulez-vous ajouter un nouveau contact?" & vbNewLine & _
                    "Oui - Nouveau contact" & vbNewLine & _
                    "Non - Sélection dans la liste des contacts existant", vbYesNo) = vbYes Then
40          m_bNewContact = True

45          sNom = InputBox("Quel est le nom du nouveau contact?" & vbNewLine & _
                            vbNewLine & _
                            "SVP, respectez le bon orthographe!")

50          If sNom <> vbNullString Then
55            If ExisteDansBD(sNom) = True Then
60              If MsgBox("Il y a déjà un contact portant ce nom!" & vbNewLine & "Voulez-vous l'ajouter quand même?", vbYesNo) = vbYes Then
65                bAjouter = True
70              Else
75                bAjouter = False
80              End If
85            Else
90              If ContientCaracteresIncorrects(sNom) = True Then
95                Call MsgBox("Il y a des caractères non permis!", vbOKOnly, "Erreur")

100               bAjouter = False
105             Else
110               bAjouter = True
115             End If
120           End If
125         Else
130           bAjouter = False
135         End If

140         If bAjouter = True Then
145           txtcontact.Text = sNom

150           Call ViderBarrerChampsContact(False, True)

155           Call HideEdMaskContact(False)

160           Call mskContactTel.SetFocus

165           txtNomClient.Visible = True
170           txtNomClient.Locked = True
      
              'Remplis combo avec tout les contact existant
175           Call AfficherControles(MODE_CONTACT)

180           Call txtContactTitre.SetFocus
185         End If
190       Else
195         Screen.MousePointer = vbHourglass


200         m_bNewContact = False

205         txtNomClient.Visible = True
210         txtNomClient.Locked = True
      
            'Remplis combo avec tout les contact existant
215         Call AfficherControles(MODE_CONTACT)
            
            'Affiche client no modifiable
220         Call RemplirComboContact
225       End If
      
230       Screen.MousePointer = vbDefault
235     Else
240       Call MsgBox("Aucun enregistrement de sélectionné")
245     End If

250     Exit Sub

AfficherErreur:

255      woups "frmClient", "CmdAddCont_Click", Err, Erl
End Sub

Private Sub CmdAnul_Click()

5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbHourglass
  
        'On cache le maskEdBox
15      Call HideEdMask(True)
        
        'commentaire unlock
20      txtcommentaire.Locked = True
  
25      m_bModeAjoutClient = False

        'on retablis les bouttons
30      Call AfficherControles(MODE_INACTIF)

        'on affiche les donnée du premier enreg
35      Call ViderBarrerChamps(True, True)
  
40      Call cmbclient_Click
        
45      Screen.MousePointer = vbDefault

50      Exit Sub

AfficherErreur:

55      woups "frmClient", "CmdAnul_Click", Err, Erl
End Sub

Private Sub CmdEnr_Click()

5       On Error GoTo AfficherErreur

10      Dim sClient   As String
15      Dim iCompteur As Integer
    
        'Nom du client
20      sClient = txtNomClient.Text
  
        'Enregistrement d'un client dans le BD
25      Screen.MousePointer = vbHourglass
  
30      Call EnregistrerClient
          
35      Call ViderBarrerChamps(True, True)
    
        'On cache les MaskEdBox
40      Call HideEdMask(True)
 
        'On met à jour le combo
45      Call RemplirComboClient
  
        'Retablir les boutons
50      Call AfficherControles(MODE_INACTIF)
  
55      For iCompteur = 0 To cmbclient.ListCount - 1
60        If cmbclient.LIST(iCompteur) = sClient Then
65          cmbclient.ListIndex = iCompteur
      
70          Exit For
75        End If
80      Next
  
85      Call cmbclient.SetFocus
  
90      Screen.MousePointer = vbDefault

95      Exit Sub

AfficherErreur:

100     woups "frmClient", "CmdEnr_Click", Err, Erl
End Sub

Private Sub EnregistrerClient()

5       On Error GoTo AfficherErreur

10      Dim rstClient As ADODB.Recordset
  
15      Set rstClient = New ADODB.Recordset
 
20      If m_bModeAjoutClient = True Then
25        Call rstClient.Open("SELECT * FROM GRB_Client", g_connData, adOpenDynamic, adLockOptimistic)
  
30        Call rstClient.AddNew

35        rstClient.Fields("DateCréation") = ConvertDate(Date)
40        rstClient.Fields("UserCréation") = g_sInitiale
45      Else
50        Call rstClient.Open("SELECT * FROM GRB_Client WHERE IDClient = " & m_iNoClient, g_connData, adOpenDynamic, adLockOptimistic)

55        rstClient.Fields("DateModification") = ConvertDate(Date)
60        rstClient.Fields("UserModification") = g_sInitiale
65      End If
      
        'Enregistre le client
70      rstClient.Fields("NomClient").Value = txtNomClient.Text
75      rstClient.Fields("Telephonne").Value = mskTelephone.Text
80      rstClient.Fields("Fax").Value = mskFax.Text
85      rstClient.Fields("ContactGRB").Value = txtContactGRB.Text
90      rstClient.Fields("Email").Value = txtEmail.Text
95      rstClient.Fields("AdresseLiv").Value = txtAdresse.Text
100     rstClient.Fields("VilleLiv").Value = txtVille.Text
105     rstClient.Fields("Prov/EtatLiv").Value = txtProvEtat.Text
110     rstClient.Fields("PaysLiv").Value = txtPays.Text
115     rstClient.Fields("Commentaire").Value = txtcommentaire.Text
120     rstClient.Fields("CodePostalLiv").Value = txtCP.Text
125     rstClient.Fields("SiteWeb").Value = txtSiteWeb.Text

130     If chkClientPotentiel.Value = vbChecked Then
135       rstClient.Fields("Potentiel").Value = True
140     Else
145       rstClient.Fields("Potentiel").Value = False
150     End If

155     rstClient.Fields("Béton").Value = m_bCategorieBeton
160     rstClient.Fields("Pavé").Value = m_bCategoriePave
165     rstClient.Fields("Pharmaceutique").Value = m_bCategoriePharmaceutique
170     rstClient.Fields("Agroalimentaire").Value = m_bCategorieAgroalimentaire
175     rstClient.Fields("Meuble").Value = m_bCategorieMeuble
180     rstClient.Fields("Meunerie").Value = m_bCategorieMeunerie
185     rstClient.Fields("Manufacturier").Value = m_bCategorieManufacturier
190     rstClient.Fields("Consultant").Value = m_bCategorieConsultant
195     rstClient.Fields("Asphalte").Value = m_bCategorieAsphalte
200     rstClient.Fields("ICPI").Value = m_bCategorieICPI
205     rstClient.Fields("ProduitsChimiques").Value = m_bCategorieProduitsChimiques
210     rstClient.Fields("Autre").Value = m_bCategorieAutre

215     rstClient.Fields("EntryIDOutlook") = ModifierClientExchange(rstClient.Fields("IDClient"))
    
220     If m_bModeAjoutClient = True Then
225       m_bModeAjoutClient = False
230     End If

235     Call rstClient.Update

240     Call rstClient.Close
245     Set rstClient = Nothing

250     Exit Sub

AfficherErreur:

255     woups "frmClient", "EnregistrerClient", Err, Erl
End Sub

Private Sub ModifierNomClientExchange(ByVal sName As String, ByVal iClientID As Integer)

5       On Error GoTo AfficherErreur
  
10      Dim otlApp      As Outlook.Application
15      Dim otlClient   As Outlook.ContactItem
20      Dim folClient   As Outlook.MAPIFolder
25      Dim bDejaOuvert As Boolean

30      lblEtatOutlook.Caption = "Modification du nom du client dans Outlook ..."
35      fraEtatOutlook.Visible = True

40      Set otlApp = OuvrirOutlook(bDejaOuvert)

45      Set folClient = GetFolder(otlApp, "Clients GRB")

50      Set otlClient = folClient.Items.Find("[User1] = " & iClientID)
       
55      If otlClient Is Nothing Then
60        Call MsgBox("Le client " & txtNomClient.Text & " n'a pas été trouvé pour la mise à jour Exchange!", vbOKOnly, "Erreur")

65        fraEtatOutlook.Visible = False

70        DoEvents

75        Exit Sub
80      End If

85      otlClient.CompanyName = sName
        
90      Call otlClient.Save

95      If bDejaOuvert = False Then
100       Call otlApp.Quit
105     End If

110     Set otlApp = Nothing

115     fraEtatOutlook.Visible = False

120     DoEvents

125     Exit Sub

AfficherErreur:

130     woups "frmClient", "ModifierNomClientExchange", Err, Erl, "iClientID = " & iClientID)

135     fraEtatOutlook.Visible = False
End Sub

Private Sub LierContactClient(ByVal iClientID As Integer)

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

55      lblEtatOutlook.Caption = "Liaison du contact avec le client dans Outlook ..."
60      fraEtatOutlook.Visible = True

65      Set otlApp = OuvrirOutlook(bDejaOuvert)

70      Set folClient = GetFolder(otlApp, "Clients GRB")
75      Set folContact = GetFolder(otlApp, "Contacts GRB")

80      Set rstClient = New ADODB.Recordset

85      Call rstClient.Open("SELECT EntryIDOutlook FROM GRB_Client WHERE IDClient = " & m_iNoClient, g_connData, adOpenForwardOnly, adLockReadOnly)

90      Set itmClient = folClient.Items.Find("[User1] = " & iClientID)

95      If Not itmClient Is Nothing Then
100       Do While itmClient.Links.count > 0
105          Set itmContact = folContact.Items.Find("[User1] = " & itmClient.Links.Item(1).Item.User1)

110          For iCompteur = 1 To itmContact.Links.count
115           If itmContact.Links.Item(1).Item.User1 = itmClient.User1 Then
120             Call itmContact.Links.Remove(iCompteur)

125             Call itmContact.Save

130             Exit For
135           End If
140         Next

145         Call itmClient.Links.Remove(1)
150       Loop

155       Call itmClient.Save

160       Call rstClient.Close
165       Set rstClient = Nothing

170       Set rstContactClient = New ADODB.Recordset

175       Call rstContactClient.Open("SELECT * FROM GRB_ContactClient WHERE NoClient = " & m_iNoClient, g_connData, adOpenForwardOnly, adLockReadOnly)

180       Do While Not rstContactClient.EOF
185         Set itmContact = folContact.Items.Find("[User1] = " & rstContactClient.Fields("NoContact"))

190         If Not itmContact Is Nothing Then
195           Call itmClient.Links.Add(itmContact)

200           Call itmClient.Save

205           Call itmContact.Links.Add(itmClient)

210           Call itmContact.Save
215         End If

220         Call rstContactClient.MoveNext
225       Loop

230       Call rstContactClient.Close
235       Set rstContactClient = Nothing
240     Else
245       Call MsgBox("Client introuvable!", vbOKOnly, "Erreur")

250       Call rstClient.Close
255       Set rstClient = Nothing
260     End If

265     If bDejaOuvert = False Then
270       Call otlApp.Quit
275     End If

280     Set otlApp = Nothing

285     fraEtatOutlook.Visible = False

290     DoEvents

295     Exit Sub

AfficherErreur:

300     If InStr(1, UCase(Err.Description), "THE OPERATION FAILED") > 0 Then
305       Call MsgBox("Une erreur est survenue ! " & vbNewLine & _
                      vbNewLine & _
                      "Pour réparer cette erreur, veuillez suivre les instructions suivantes : " & vbNewLine & _
                      vbNewLine & _
                      "- Dans Outlook, ouvrez le client '" & txtNomClient.Text & "' dans Clients GRB" & vbNewLine & _
                      "- Cliquez sur tous les contacts de ce client 1 à la fois pour trouver le contact incorrect." & vbNewLine & _
                      "- Effacez ce contact de la liste des contacts de ce client." & vbNewLine & _
                      "- Dans le logiciel GRB, recommencez l'opération (effacez le contact et l'ajouter de nouveau).", vbOKOnly, "Erreur")
310     Else
315       woups "frmClient", "LierContactClient", Err, Erl, txtNomClient.Text)
320     End If

325     fraEtatOutlook.Visible = False
End Sub

Private Function ModifierClientExchange(ByVal iClientID As Integer) As String

5       On Error GoTo AfficherErreur

10      Dim otlApp      As Outlook.Application
15      Dim otlClient   As Outlook.ContactItem
20      Dim folClient   As Outlook.MAPIFolder
25      Dim bDejaOuvert As Boolean

30      If m_bModeAjoutClient = True Then
35        lblEtatOutlook.Caption = "Ajout du client dans Outlook ..."
40      Else
45        lblEtatOutlook.Caption = "Modification du client dans Outlook ..."
50      End If

55      fraEtatOutlook.Visible = True

60      Set otlApp = OuvrirOutlook(bDejaOuvert)

65      Set folClient = GetFolder(otlApp, "Clients GRB")

70      If m_bModeAjoutClient = True Then
75        Set otlClient = folClient.Items.Add(olContactItem)

80        otlClient.User1 = iClientID
85      Else
90        Set otlClient = folClient.Items.Find("[User1] = " & iClientID)
95      End If

100     If otlClient Is Nothing Then
105       Call MsgBox("Le client " & txtNomClient.Text & " n'a pas été trouvé pour la mise à jour Exchange!", vbOKOnly, "Erreur")

110       fraEtatOutlook.Visible = False

115       DoEvents

120       Exit Function
125     End If

130     otlClient.CompanyName = txtNomClient.Text
    
135     If mskTelephone.Text <> "(___) ___-____" Then
140       otlClient.BusinessTelephoneNumber = mskTelephone.Text
145     End If
  
150     If mskFax.Text <> "(___) ___-____" Then
155       otlClient.BusinessFaxNumber = mskFax.Text
160     End If
   
165     otlClient.Email1Address = txtEmail.Text
170     otlClient.BusinessAddressStreet = txtAdresse.Text
175     otlClient.BusinessAddressCity = txtVille.Text
180     otlClient.BusinessAddressState = txtProvEtat.Text
185     otlClient.BusinessAddressCountry = txtPays.Text
190     otlClient.BusinessAddressPostalCode = txtCP.Text
195     otlClient.Body = txtcommentaire.Text
200     otlClient.WebPage = txtSiteWeb.Text
   
205     Call otlClient.Save

210     ModifierClientExchange = otlClient.EntryID

215     If bDejaOuvert = False Then
220       Call otlApp.Quit
225     End If

230     Set otlApp = Nothing

235     fraEtatOutlook.Visible = False

240     DoEvents

245     Exit Function

AfficherErreur:

250     woups "frmClient", "ModifierClientExchange", Err, Erl, "iClientID = " & iClientID)

255     fraEtatOutlook.Visible = False
End Function

Private Function AjouterContactExchange(ByVal iContactID As Integer) As String
  
5       On Error GoTo AfficherErreur
  
10      Dim otlApp      As Outlook.Application
15      Dim otlContact  As Outlook.ContactItem
20      Dim folContact  As Outlook.MAPIFolder
25      Dim bDejaOuvert As Boolean
30      Dim sNom()      As String

35      lblEtatOutlook.Caption = "Ajout du contact dans Outlook ..."
40      fraEtatOutlook.Visible = True

45      Set otlApp = OuvrirOutlook(bDejaOuvert)

50      Set folContact = GetFolder(otlApp, "Contacts GRB")

55      Set otlContact = folContact.Items.Add(olContactItem)

60      otlContact.User1 = iContactID
    
65      sNom = Split(Trim$(txtcontact.Text), " ")

70      Select Case UBound(sNom)
          Case 0:
75          otlContact.FirstName = sNom(0)
  
          Case 1:
80          otlContact.FirstName = sNom(0)
85          otlContact.LastName = sNom(1)

          Case 2
90          otlContact.FirstName = sNom(0)
95          otlContact.MiddleName = sNom(1)
100         otlContact.LastName = sNom(2)
105      End Select
        
110     otlContact.Title = ""
    
115     otlContact.CompanyName = txtNomClient.Text
120     otlContact.JobTitle = txtContactTitre.Text

125     If Trim$(mskContactTel.Text) <> "" Then
130       If txtTelephone.Text <> "(___) ___-____" Then
135         If Trim$(txtcontact_poste.Text) <> "" Then
140           otlContact.BusinessTelephoneNumber = mskContactTel.Text & " Ext : " & txtcontact_poste.Text
145         Else
150           otlContact.BusinessTelephoneNumber = mskContactTel.Text
155         End If
160       End If
165     End If
    
170     If mskContactFax.Text <> "(___) ___-____" Then
175       otlContact.BusinessFaxNumber = mskContactFax.Text
180     End If

185     If mskContactDom.Text <> "(___) ___-____" Then
190       otlContact.HomeTelephoneNumber = mskContactDom.Text
195     End If
    
200     If mskContactCell.Text <> "(___) ___-____" Then
205       otlContact.MobileTelephoneNumber = mskContactCell.Text
210     End If
    
215     If mskContactPage.Text <> "(___) ___-____" Then
220       otlContact.PagerNumber = mskContactPage.Text
225     End If
    
230     otlContact.Email1Address = txtcontact_email.Text
        
235     Call otlContact.Save

240     AjouterContactExchange = otlContact.EntryID

245     If bDejaOuvert = False Then
250       Call otlApp.Quit
255     End If

260     Set otlApp = Nothing

265     fraEtatOutlook.Visible = False

270     DoEvents

275     Exit Function

AfficherErreur:

280     woups "frmClient", "AjouterContactExchange", Err, Erl, "iContactID = " & iContactID)

285     fraEtatOutlook.Visible = False
End Function

Private Sub SupprimerClientExchange(ByVal iClientID As Integer)

5       On Error GoTo AfficherErreur

10      Dim otlApp      As Outlook.Application
15      Dim otlClient   As Outlook.ContactItem
20      Dim folClient   As MAPIFolder
25      Dim bDejaOuvert As Boolean

30      lblEtatOutlook.Caption = "Suppression du client dans Outlook ..."
35      fraEtatOutlook.Visible = True

40      Set otlApp = OuvrirOutlook(bDejaOuvert)

45      Set folClient = GetFolder(otlApp, "Clients GRB")

50      Set otlClient = folClient.Items.Find("[User1] = " & iClientID)

55      If Not otlClient Is Nothing Then
60        Call otlClient.Delete
65      End If

70      If bDejaOuvert = False Then
75        Call otlApp.Quit
80      End If

85      Set otlApp = Nothing

90      fraEtatOutlook.Visible = False

95      DoEvents

100     Exit Sub

AfficherErreur:

105     woups "frmClient", "SupprimerClientExchange", Err, Erl

110     fraEtatOutlook.Visible = False
End Sub

Private Sub CmdModif_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
  
15      If cmbclient.ListCount > 0 Then
20        Screen.MousePointer = vbHourglass
    
          'proc qui permet de modifier l'enr courant
          'on montre/cache des buttons
25        Call HideEdMask(False)
                    
30        Call AfficherControles(MODE_CLIENT)
              
35        Call ViderBarrerChamps(False, False)
              
          'pour que le nom ne soit pas modifiable
40        txtNomClient.Visible = True
45        txtNomClient.Locked = True
      
50        Screen.MousePointer = vbDefault
55      Else
60        Call MsgBox("Aucun enregistrement de sélectionné!")
65      End If

70      Exit Sub

AfficherErreur:

75      woups "frmClient", "CmdModif_Click", Err, Erl
End Sub

Private Sub cmdquit_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmClient", "cmdquit_Click", Err, Erl
End Sub

Private Sub CmdRefCont_Click()

5       On Error GoTo AfficherErreur

10      Dim rstContactClient As ADODB.Recordset
15      Dim rstContact       As ADODB.Recordset
    
20      Screen.MousePointer = vbHourglass
    
25      Set rstContactClient = New ADODB.Recordset

30      If m_bNewContact = True Then
35        Set rstContact = New ADODB.Recordset

40        Call rstContact.Open("SELECT * FROM GRB_Contact", g_connData, adOpenDynamic, adLockOptimistic)

45        Call rstContact.AddNew

50        rstContact.Fields("NomContact").Value = txtcontact.Text
55        rstContact.Fields("Titre").Value = txtContactTitre.Text
60        rstContact.Fields("Compagnie").Value = txtNomClient.Text
65        rstContact.Fields("Telephonne").Value = mskContactTel.Text
70        rstContact.Fields("Fax").Value = mskContactFax.Text
75        rstContact.Fields("Pagette").Value = mskContactPage.Text
80        rstContact.Fields("Cellulaire").Value = mskContactCell.Text
85        rstContact.Fields("E-mail").Value = txtcontact_email.Text
90        rstContact.Fields("NoPoste").Value = txtcontact_poste.Text
95        rstContact.Fields("TelDomicile").Value = mskContactDom.Text
100       rstContact.Fields("UserCréation").Value = g_sInitiale
105       rstContact.Fields("DateCréation").Value = ConvertDate(Date)
  
110       rstContact.Fields("EntryIDOutlook") = AjouterContactExchange(rstContact.Fields("IDContact"))

115       Call rstContact.Update

          'Set la table
120       Call rstContactClient.Open("SELECT * FROM GRB_ContactClient WHERE NoClient = " & m_iNoClient & " And NoContact = " & rstContact.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)
    
          'Si pas déjà existant, on ajoute dans la table
125       If rstContactClient.EOF Then
            'Ajoute dans la table
130         Call rstContactClient.AddNew
      
135         rstContactClient.Fields("NoClient") = m_iNoClient
140         rstContactClient.Fields("NoContact") = rstContact.Fields("IDContact")
      
145         Call rstContactClient.Update
150       End If
              
155       Call rstContact.Close
160       Set rstContact = Nothing

165       m_bNewContact = False

170       Call HideEdMaskContact(True)
175     Else
          'Set la table
180       Call rstContactClient.Open("SELECT * FROM GRB_ContactClient WHERE NoClient = " & m_iNoClient & " AND NoContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)
    
          'Si pas déjà existant, on ajoute dans la table
185       If rstContactClient.EOF Then
            'Ajoute dans la table
190         Call rstContactClient.AddNew
      
195         rstContactClient.Fields("NoClient") = m_iNoClient
200         rstContactClient.Fields("NoContact") = m_iNoContact
      
205         Call rstContactClient.Update
210       End If
    
          'Ferme tables et connexion
215       Call rstContactClient.Close
220     End If

225     Call LierContactClient(m_iNoClient)

        'Ferme tables et connection
230     Set rstContactClient = Nothing
 
        'Bouton comme avant(apparait)
235     Call AfficherControles(MODE_INACTIF)
    
        'n'est plus en mode ajouter
240     m_bModeAjoutContact = False

        'remplis combo contact
245     Call RemplirComboContact

250     Call ViderBarrerChampsContact(True, False)

255     Screen.MousePointer = vbDefault

260     Exit Sub

AfficherErreur:

265     woups "frmClient", "CmdRefCont_Click", Err, Erl
End Sub

Private Sub cmdRenommer_Click()

5       On Error GoTo AfficherErreur
        
        '''''''''''''''''''''''''''''''''''''''
        'on renomme le nom du CLIENT
        ''''''''''''''''''''''''''''''''''''''''
10      Dim rstClient As ADODB.Recordset
15      Dim sName     As String

20      If cmbclient.ListCount > 0 Then
          'Proc qui permet de modifié un CLIENT a la BD
          'On procède à la saisie du nom du CLIENT
25        sName = InputBox("Veuillez entrer le nom du client", "Renommer client", txtNomClient.Text)
        
30        If sName <> vbNullString Then
35          If sName <> txtNomClient.Text Then
40            Set rstClient = New ADODB.Recordset
  
45            Call rstClient.Open("SELECT * FROM GRB_Client WHERE NomClient = '" & Replace(sName, "'", "''") & "' AND Supprimé = False", g_connData, adOpenDynamic, adLockOptimistic)
      
50            If rstClient.EOF Then
                'Transfert du nom au premier textBox
60              Screen.MousePointer = vbHourglass
          
65              Call rstClient.Close
              
70              Call rstClient.Open("SELECT NomClient, IDClient, EntryIDOutlook FROM GRB_Client WHERE IDClient = " & m_iNoClient, g_connData, adOpenDynamic, adLockOptimistic)
                
75              Call ModifierNomClientExchange(sName, m_iNoClient)
                
80              txtNomClient.Text = sName
                
                'Transfert des données
85              rstClient.Fields("NomClient").Value = txtNomClient.Text
                  
                'Mise à jour de la base de données
90              Call rstClient.Update
                     
95              Call rstClient.Close
                          
100             Call RemplirComboClient
            
105             cmbclient.Text = sName
            
110             m_bRenommer = True
            
115             Call cmbclient_Click
            
120             m_bRenommer = False
            
125             Screen.MousePointer = vbDefault
130           Else
135             Call MsgBox("Le client " & sName & " existe déjà!", vbCritical)
  
140             Call rstClient.Close
145           End If
  
150           Set rstClient = Nothing
155         End If
160       End If
165     Else
170       Call MsgBox("Aucun enregistrement de sélectionné!")
175     End If

180     Exit Sub

AfficherErreur:

185     woups "frmClient", "cmdRenommer_Click", Err, Erl
End Sub

Private Sub cmdsupcontact_Click()

5       On Error GoTo AfficherErreur
        
        'fonction qui supprime l'enregistrement courant
10      If cmbContact.ListCount > 0 Then
15        If MsgBox("Etes-vous sur de vouloir supprimer cette enregistrement?", vbYesNo, "Supprimer") = vbYes Then
20          Screen.MousePointer = vbHourglass
      
25          Call g_connData.Execute("DELETE * FROM GRB_ContactClient WHERE NoClient = " & m_iNoClient & " AND NoContact = " & m_iNoContact)
             
30          Call LierContactClient(m_iNoClient)
             
            'remplis le combo employé
35          Call RemplirComboContact
      
40          Screen.MousePointer = vbDefault
45        End If
50      End If

55      Exit Sub

AfficherErreur:

60      woups "frmClient", "cmdsupcontact_Click", Err, Erl
End Sub

Private Sub CmdSupp_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum  As ADODB.Recordset
15      Dim rstClient    As ADODB.Recordset
20      Dim bPeutEffacer As Boolean
    
25      If cmbclient.ListCount > 0 Then
          'fonction qui supprime l'enregistrement courant
30        If MsgBox("Êtes-vous sûr de supprimer cet enregistrement?", vbYesNo, "Supprimer") = vbYes Then
35          Screen.MousePointer = vbHourglass
        
40          Set rstProjSoum = New ADODB.Recordset

            'open table
45          Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDClient = " & m_iNoClient, g_connData, adOpenDynamic, adLockOptimistic)

            'si existe pas dans soumission, on peut le deleter
50          If rstProjSoum.EOF Then
55            Call rstProjSoum.Close
            
              'S'il existe pas dans projet, on peut le deleter
60            Call rstProjSoum.Open("SELECT * FROM GRB_ProjetMec WHERE IDClient = " & m_iNoClient, g_connData, adOpenDynamic, adLockOptimistic)
          
65            If rstProjSoum.EOF Then
70              Call rstProjSoum.Close
              
75              Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDClient = " & m_iNoClient, g_connData, adOpenDynamic, adLockOptimistic)
            
80              If rstProjSoum.EOF Then
85                Call rstProjSoum.Close
              
90                Call rstProjSoum.Open("SELECT * FROM GRB_ProjetElec WHERE IDClient = " & m_iNoClient, g_connData, adOpenDynamic, adLockOptimistic)
              
95                If rstProjSoum.EOF Then
100                 bPeutEffacer = True

105                 Call rstProjSoum.Close
110                 Set rstProjSoum = Nothing
115               Else
120                 bPeutEffacer = False
                
125                 Call rstProjSoum.Close
130                 Set rstProjSoum = Nothing
135               End If
140             Else
145               bPeutEffacer = False
         
150               Call rstProjSoum.Close
155               Set rstProjSoum = Nothing
160             End If
165           Else
170             bPeutEffacer = False
                    
175             Call rstProjSoum.Close
180             Set rstProjSoum = Nothing
185           End If
190         Else
195           bPeutEffacer = False
          
200           Call rstProjSoum.Close
205           Set rstProjSoum = Nothing
210         End If

215         If cmbContact.ListCount > 0 Then
              'Delete les contact pour ce client
220           Call g_connData.Execute("DELETE * FROM GRB_ContactClient WHERE NoClient = " & m_iNoClient)
225         End If

230         Call SupprimerClientExchange(m_iNoClient)

235         If bPeutEffacer = True Then
              'Delete le client
240           Call g_connData.Execute("DELETE * FROM GRB_Client WHERE IDClient = " & m_iNoClient)
245         Else
250           Set rstClient = New ADODB.Recordset

255           Call rstClient.Open("SELECT * FROM GRB_Client WHERE IDClient = " & m_iNoClient, g_connData, adOpenDynamic, adLockOptimistic)

260           rstClient.Fields("Supprimé") = True

265           Call rstClient.Update

270           Call rstClient.Close
275           Set rstClient = Nothing
280         End If

285         Call RemplirComboClient
      
290         Screen.MousePointer = vbDefault
295       End If
300     Else
305       Call MsgBox("Aucun enregistrement de sélectionné!")
310     End If

315     Exit Sub

AfficherErreur:

320     woups "frmClient", "CmdSupp_Click", Err, Erl
End Sub

Private Sub cmbclient_Click()

5       On Error GoTo AfficherErreur
        
        'Quand le user selectionne un enregistrement on se posotionne dessus
10      If cmbclient.Text <> vbNullString Then
15        txtNomClient.Text = cmbclient.Text
20      Else
25        If ComboContient(cmbclient, txtNomClient.Text) = False Then
30          Call RemplirComboClient
35        End If

40        cmbclient.Text = txtNomClient.Text
45      End If
  
50      If cmbclient.ListIndex > -1 Then
55        If m_bRenommer = False And m_bModeAjoutClient = False Then
60          m_iNoClient = cmbclient.ItemData(cmbclient.ListIndex)
65        End If
70      End If
  
        'remplis le combo dépendant le client sélectionné
75      Call AfficherClient

80      Call RemplirComboContact

85      Exit Sub

AfficherErreur:

90      woups "frmClient", "cmbclient_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Call RemplirComboClient
  
15      Call AfficherControles(MODE_INACTIF)

20      Call ActiverBoutonsGroupe

25      Screen.MousePointer = vbDefault

30      Exit Sub

AfficherErreur:

35      woups "frmClient", "Form_Load", Err, Erl
End Sub

Private Sub ActiverBoutonsGroupe()

5       On Error GoTo AfficherErreur

10      CmdAdd.Enabled = g_bModificationClients
15      CmdAddCont.Enabled = g_bModificationClients
20      CmdModif.Enabled = g_bModificationClients
25      cmdrenommer.Enabled = g_bModificationClients
30      cmdsupcontact.Enabled = g_bModificationClients
35      CmdSupp.Enabled = g_bModificationClients
40      cmdMailListClient.Enabled = g_bModificationListeDistribution
45      cmdMailListContact.Enabled = g_bModificationListeDistribution

50      Exit Sub

AfficherErreur:

55      woups "frmClient", "ActiverBoutonsGroupe", Err, Erl
End Sub

Private Sub HideEdMask(ByVal bVisible As Boolean)

5       On Error GoTo AfficherErreur
        
        'proc qui rend visible/ou non les maskEdBox
        'On en profite pour les nettoyer du dernier Enregistrement
        'et on fait l'inverse avec les textBox
10      If m_bModeAjoutClient = True Then
15        txtTelephone.Text = vbNullString
20        txtFax.Text = vbNullString
25      Else
30        mskTelephone.Text = txtTelephone.Text
35        mskFax.Text = txtFax.Text
40      End If
  
45      mskTelephone.Visible = Not bVisible
50      mskFax.Visible = Not bVisible
 
55      txtTelephone.Visible = bVisible
60      txtFax.Visible = bVisible

65      Exit Sub

AfficherErreur:

70      woups "frmClient", "HideEdMask", Err, Erl
End Sub

Private Sub HideEdMaskContact(ByVal bVisible As Boolean)

5       On Error GoTo AfficherErreur
        
        'proc qui rend visible/ou non les maskEdBox
        'On en profite pour les nettoyer du dernier Enregistrement
        'et on fait l'inverse avec les textBox
10      If m_bModeAjoutContact = True Then
15        txtcontact_tel.Text = vbNullString
20        txtcontact_fax.Text = vbNullString
25        txtcontact_page.Text = vbNullString
30        txtcontact_cell.Text = vbNullString
35        txtcontact_dom.Text = vbNullString
       
40        mskContactTel.Text = vbNullString
45        mskContactFax.Text = vbNullString
50        mskContactPage.Text = vbNullString
55        mskContactCell.Text = vbNullString
60        mskContactDom.Text = vbNullString
65      End If
  
70      mskContactTel.Visible = Not bVisible
75      txtcontact_tel.Visible = bVisible

80      mskContactFax.Visible = Not bVisible
85      txtcontact_fax.Visible = bVisible

90      mskContactPage.Visible = Not bVisible
95      txtcontact_page.Visible = bVisible

100     mskContactCell.Visible = Not bVisible
105     txtcontact_cell.Visible = bVisible

110     mskContactDom.Visible = Not bVisible
115     txtcontact_dom.Visible = bVisible

120     Exit Sub

AfficherErreur:

125     woups "frmClient", "HideEdMaskContact", Err, Erl
End Sub

Private Sub Form_Unload(Cancel As Integer)

5       On Error GoTo AfficherErreur

10      Set FrmClient = Nothing

15      Exit Sub

AfficherErreur:

20      woups "frmClient", "Form_Unload", Err, Erl
End Sub
Private Sub mskTelephone_GotFocus()

5       On Error GoTo AfficherErreur

10      mskTelephone.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmClient", "mskTelephone_GotFocus", Err, Erl
End Sub

Private Sub mskTelephone_LostFocus()

5       On Error GoTo AfficherErreur

10      mskTelephone.mask = vbNullString
15      If mskTelephone.Text = "(___) ___-____" Then
20        mskTelephone.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmClient", "mskTelephone_LostFocus", Err, Erl
End Sub

Private Sub mskFax_GotFocus()

5       On Error GoTo AfficherErreur

10      mskFax.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmClient", "mskFax_GotFocus", Err, Erl
End Sub

Private Sub mskFax_LostFocus()

5       On Error GoTo AfficherErreur

10      mskFax.mask = vbNullString
15      If mskFax.Text = "(___) ___-____" Then
20        mskFax.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmClient", "mskFax_LostFocus", Err, Erl
End Sub

Public Sub RemplirComboContact()

5       On Error GoTo AfficherErreur
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'remplis le combo contact dépendant le client sélectionné
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
10      Dim rstClient  As ADODB.Recordset
15      Dim rstContact As ADODB.Recordset
    
20      Set rstContact = New ADODB.Recordset
    
25      If m_bModeAjoutContact = True Then
30        Call rstContact.Open("SELECT NomContact, IDContact FROM GRB_contact WHERE Supprimé = False ORDER BY NomContact", g_connData, adOpenDynamic, adLockOptimistic)
35      Else
40        Call rstContact.Open("SELECT GRB_Contact.NomContact, GRB_Contact.IDContact FROM GRB_Contact INNER JOIN GRB_ContactClient ON GRB_Contact.IDContact = GRB_ContactClient.NoContact WHERE GRB_ContactClient.NoClient = " & m_iNoClient & " ORDER BY NomContact", g_connData, adOpenDynamic, adLockOptimistic)
45      End If
    
50      Call cmbContact.Clear
    
55      Do While Not rstContact.EOF
60        Call cmbContact.AddItem(Trim(rstContact.Fields("NomContact")))
65        cmbContact.ItemData(cmbContact.newIndex) = rstContact.Fields("IDContact")
        
70       Call rstContact.MoveNext
75      Loop
    
        'ferme la table "GRB_Contact"
80      Call rstContact.Close
85      Set rstContact = Nothing
        
        'affiche le contact de la table client
        'si combo est pas vide, affiche le premier contact, sinon le contact inscrit dans table client
90      If cmbContact.ListCount > 0 Then
95        cmbContact.ListIndex = 0
100     Else
          'VIDE les champs
105       txtContactTitre.Text = vbNullString
110       txtcontact_cell.Text = vbNullString
115       txtcontact_email = vbNullString
120       txtcontact_fax = vbNullString
125       txtcontact_page = vbNullString
130       txtcontact_poste = vbNullString
135       txtcontact_tel = vbNullString
140       txtcontact_dom.Text = vbNullString
145     End If

150     Exit Sub

AfficherErreur:

155     woups "frmClient", "RemplirComboContact", Err, Erl
End Sub

Private Sub cmdRechercher_Click()

5       On Error GoTo AfficherErreur

10      Dim rstClient As ADODB.Recordset
15      Dim sSearch   As String
  
20      Screen.MousePointer = vbHourglass
  
25      sSearch = txtRechercher.Text
  
        'vide les champs
30      Call ViderBarrerChamps(True, True)
      
        'Filtre pour sélection des Nomcontact
        'goSQL = "SELECT * FROM GRB_contact order by NomContact"
35      Set rstClient = New ADODB.Recordset
        
40      Call rstClient.Open("SELECT NomClient, IDClient FROM GRB_Client WHERE Instr(1, NomClient, '" & Replace(sSearch, "'", "''") & "') > 0 AND Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)
                
        'vide combo
45      Call cmbclient.Clear
 
50      Do While Not rstClient.EOF
55        Call cmbclient.AddItem(rstClient.Fields("NomClient"))
60        cmbclient.ItemData(cmbclient.newIndex) = rstClient.Fields("IDClient")
                    
65        Call rstClient.MoveNext
70      Loop
      
75      Call rstClient.Close
80      Set rstClient = Nothing
    
85      Screen.MousePointer = vbDefault

90      If cmbclient.ListCount > 0 Then
95        cmbclient.ListIndex = 0
100     Else
105       Call cmbContact.Clear
    
          'VIDE les champs
110       txtContactTitre.Text = vbNullString
115       txtcontact_cell.Text = vbNullString
120       txtcontact_email.Text = vbNullString
125       txtcontact_fax.Text = vbNullString
130       txtcontact_page.Text = vbNullString
135       txtcontact_poste.Text = vbNullString
140       txtcontact_tel.Text = vbNullString
145       txtcontact_dom.Text = vbNullString
150     End If

155     Exit Sub

AfficherErreur:

160     woups "frmClient", "cmdRechercher_Click", Err, Erl
End Sub

Private Sub cmdRafraichir_Click()

5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbHourglass
  
15      Call RemplirComboClient
  
20      Screen.MousePointer = vbDefault

25      Exit Sub

AfficherErreur:

30      woups "frmClient", "cmdRafraichir_Click", Err, Erl
End Sub

Private Sub txtRechercher_Change()

5       On Error GoTo AfficherErreur

10      If Len(txtRechercher.Text) > 0 Then
15        cmdRechercher.Enabled = True
20      Else
25        cmdRechercher.Enabled = False
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmClient", "txtRechercher_Change", Err, Erl
End Sub

Private Sub mskContactTel_GotFocus()

5       On Error GoTo AfficherErreur

10      mskContactTel.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmClient", "mskContactTel_GotFocus", Err, Erl
End Sub

Private Sub mskContactTel_LostFocus()

5       On Error GoTo AfficherErreur

10      mskContactTel.mask = vbNullString

15      If mskContactTel.Text = "(___) ___-____" Then
20        mskContactTel.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmClient", "mskContactTel_LostFocus", Err, Erl
End Sub

Private Sub mskContactFax_GotFocus()

5       On Error GoTo AfficherErreur

10      mskContactFax.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmClient", "mskContactFax_GotFocus", Err, Erl
End Sub

Private Sub mskContactFax_LostFocus()

5       On Error GoTo AfficherErreur

10      mskContactFax.mask = vbNullString

15      If mskContactFax.Text = "(___) ___-____" Then
20        mskContactFax.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmClient", "mskContactFax_LostFocus", Err, Erl
End Sub

Private Sub mskContactCell_GotFocus()

5       On Error GoTo AfficherErreur

10      mskContactCell.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmClient", "mskContactCell_GotFocus", Err, Erl
End Sub

Private Sub mskContactCell_LostFocus()

5       On Error GoTo AfficherErreur

10      mskContactCell.mask = vbNullString

15      If mskContactCell.Text = "(___) ___-____" Then
20        mskContactCell.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmClient", "mskContactCell_LostFocus", Err, Erl
End Sub

Private Sub mskContactDom_GotFocus()

5       On Error GoTo AfficherErreur

10      mskContactDom.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmClient", "mskContactDom_GotFocus", Err, Erl
End Sub

Private Sub mskContactDom_LostFocus()

5       On Error GoTo AfficherErreur

10      mskContactDom.mask = vbNullString

15      If mskContactDom.Text = "(___) ___-____" Then
20        mskContactDom.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmClient", "mskContactDom_LostFocus", Err, Erl
End Sub

Private Sub mskContactPage_GotFocus()

5       On Error GoTo AfficherErreur

10      mskContactPage.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmClient", "mskContactPage_GotFocus", Err, Erl
End Sub

Private Sub mskContactPage_LostFocus()

5       On Error GoTo AfficherErreur

10      mskContactPage.mask = vbNullString

15      If mskContactPage.Text = "(___) ___-____" Then
20        mskContactPage.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmClient", "mskContactPage_LostFocus", Err, Erl
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

65      woups "frmClient", "ExisteDansBD", Err, Erl
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

40      woups "frmClient", "ContientCaracteresIncorrects", Err, Erl
End Function
