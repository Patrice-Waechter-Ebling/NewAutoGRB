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
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   9240
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
         ItemData        =   "FrmClient.frx":0442
         Left            =   1080
         List            =   "FrmClient.frx":0444
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
      Text            =   "FrmClient.frx":0446
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
      ItemData        =   "FrmClient.frx":044C
      Left            =   1440
      List            =   "FrmClient.frx":044E
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

Public m_bCategorieBeton As Boolean
Public m_bCategoriePave As Boolean
Public m_bCategoriePharmaceutique As Boolean
Public m_bCategorieAgroalimentaire As Boolean
Public m_bCategorieMeuble As Boolean
Public m_bCategorieMeunerie As Boolean
Public m_bCategorieManufacturier As Boolean
Public m_bCategorieConsultant As Boolean
Public m_bCategorieAsphalte As Boolean
Public m_bCategorieICPI As Boolean
Public m_bCategorieProduitsChimiques As Boolean
Public m_bCategorieAutre As Boolean

'Choix d'impression
Public m_bImpressionAnnuler As Boolean
Public m_bImpressionVille As Boolean
Public m_bImpressionCategorie As Boolean
Public m_bImpressionPotentiel As Boolean
Public m_bImpressionFacturer As Boolean

'Choix d'impression de categorie
Public m_bImpressionBeton As Boolean
Public m_bImpressionPave As Boolean
Public m_bImpressionPharmaceutique As Boolean
Public m_bImpressionAgroAlimentaire As Boolean
Public m_bImpressionMeuble As Boolean
Public m_bImpressionMeunerie As Boolean
Public m_bImpressionManufacturier As Boolean
Public m_bImpressionConsultant As Boolean
Public m_bImpressionAsphalte As Boolean
Public m_bImpressionICPI As Boolean
Public m_bImpressionProduitsChimiques As Boolean
Public m_bImpressionAutre As Boolean
 
Private m_bModeAjoutContact As Boolean
Private m_bModeAjoutClient As Boolean
Private m_iNoContact As Integer
Private m_iNoClient As Integer
Private m_bRenommer As Boolean
Private m_bNewContact As Boolean

Public m_bAnnulerDistList As Boolean
'Public m_otlDistList As Outlook.DistListItem

Public m_bAnnulerVille As Boolean
Public m_sVille As String
 
Private m_eMode As enumMode

Private Sub RemplirComboClient()

 On Error GoTo Oups
 
 'Rempli le combo des clients
 Dim rstClient As ADODB.Recordset
 
 'Il faut vider le combo avant de le remplir
 Call cmbclient.Clear
 
 Set rstClient = New ADODB.Recordset

 Call rstClient.Open("SELECT NomClient, IDClient FROM GrbClient WHERE Supprimé = False", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstClient.EOF
 'Ajout du nom du client dans le combo
 Call cmbclient.AddItem(Trim(rstClient.Fields("NomClient")))
 
 'Ajout du numéro du client dans le ItemData du combo
 cmbclient.ItemData(cmbclient.newIndex) = rstClient.Fields("IDClient")
 
 Call rstClient.MoveNext
 Loop
 
 Call rstClient.Close
  Set rstClient = Nothing
 
  If cmbclient.ListCount > 0 Then
  cmbclient.ListIndex = 0
  End If

  Exit Sub

Oups:

  wOups "frmClient", "RemplirComboClient", Err, Err.number, Err.Description
End Sub

Private Sub AfficherClient()

 On Error GoTo Oups
 
 'Affiche le client sélectionné
 Dim rstClient As ADODB.Recordset
 
 Set rstClient = New ADODB.Recordset
 
 Call rstClient.Open("SELECT * FROM GrbClient WHERE IDClient = " & m_iNoClient, g_connData, adOpenDynamic, adLockOptimistic)

 Call ViderBarrerChamps(True, True)
 
 'Telephonne
 If Not IsNull(rstClient.Fields("Telephonne")) Then
 txtTelephone.Text = rstClient.Fields("Telephonne")
 End If
 
 'Fax
 If Not IsNull(rstClient.Fields("Fax")) Then
 txtFax.Text = rstClient.Fields("Fax")
 End If
 
 'ContactGRB
  If Not IsNull(rstClient.Fields("ContactGRB")) Then
  txtContactGRB.Text = rstClient.Fields("ContactGRB")
  End If
 
 'Email
  If Not IsNull(rstClient.Fields("Email")) Then
  txtEmail.Text = rstClient.Fields("Email")
  End If
 
 'AdresseLiv
  If Not IsNull(rstClient.Fields("AdresseLiv")) Then
  txtAdresse.Text = rstClient.Fields("AdresseLiv")
10 End If
 
 'VilleLiv
If Not IsNull(rstClient.Fields("VilleLiv")) Then
 txtVille.Text = rstClient.Fields("VilleLiv")
End If
 
 'Prov/EtatLiv
If Not IsNull(rstClient.Fields("Prov/EtatLiv")) Then
 txtProvEtat.Text = rstClient.Fields("Prov/EtatLiv")
End If
 
 'PaysLiv
If Not IsNull(rstClient.Fields("PaysLiv")) Then
 txtPays.Text = rstClient.Fields("PaysLiv")
End If
 
 'CodePostalLiv
1  If Not IsNull(rstClient.Fields("CodePostalLiv")) Then
 txtCP.Text = rstClient.Fields("CodePostalLiv")
 End If
 
 'Commentaire
If Not IsNull(rstClient.Fields("Commentaire")) Then
 txtcommentaire.Text = rstClient.Fields("Commentaire")
End If

 'Site Web
 If Not IsNull(rstClient.Fields("SiteWeb")) Then
1  txtSiteWeb.Text = rstClient.Fields("SiteWeb")
 End If

 'Création
 If Not IsNull(rstClient.Fields("DateCréation")) Then
 lblDateCreation.Caption = rstClient.Fields("DateCréation")
End If

 'User Création
If Not IsNull(rstClient.Fields("UserCréation")) Then
 lblUserCreation.Caption = "Par : " & rstClient.Fields("UserCréation")
End If

 'Modification
If Not IsNull(rstClient.Fields("DateModification")) Then
 lblDateModification.Caption = rstClient.Fields("DateModification")
End If

 'User Modification
If Not IsNull(rstClient.Fields("UserModification")) Then
 lblUserModification.Caption = "Par : " & rstClient.Fields("UserModification")
2  End If

'Client Potentiel
2  If rstClient.Fields("Potentiel") = True Then
 chkClientPotentiel.Value = vbChecked
2  End If

m_bCategorieBeton = rstClient.Fields("Béton")
2  m_bCategoriePave = rstClient.Fields("Pavé")
m_bCategoriePharmaceutique = rstClient.Fields("Pharmaceutique")
30 m_bCategorieAgroalimentaire = rstClient.Fields("Agroalimentaire")
m_bCategorieMeuble = rstClient.Fields("Meuble")
m_bCategorieMeunerie = rstClient.Fields("Meunerie")
m_bCategorieManufacturier = rstClient.Fields("Manufacturier")
m_bCategorieConsultant = rstClient.Fields("Consultant")
m_bCategorieAsphalte = rstClient.Fields("Asphalte")
m_bCategorieICPI = rstClient.Fields("ICPI")
m_bCategorieProduitsChimiques = rstClient.Fields("ProduitsChimiques")
m_bCategorieAutre = rstClient.Fields("Autre")

Call AfficherCategories
 
Call rstClient.Close
Set rstClient = Nothing

3  Exit Sub

Oups:

wOups "frmClient", "AfficherClient", Err, Err.number, Err.Description
End Sub

Private Sub AfficherCategories()

 On Error GoTo Oups

 txtCategorie.Text = ""

 If m_bCategorieBeton = True Then
 txtCategorie.Text = "Béton"
 End If

 If m_bCategoriePave = True Then
 If Trim$(txtCategorie.Text) <> "" Then
 txtCategorie.Text = txtCategorie.Text & ", Pavé"
 Else
 txtCategorie.Text = "Pavé"
 End If
  End If

  If m_bCategoriePharmaceutique = True Then
  If Trim$(txtCategorie.Text) <> "" Then
  txtCategorie.Text = txtCategorie.Text & ", Pharmaceutique"
  Else
  txtCategorie.Text = "Pharmaceutique"
  End If
  End If

10 If m_bCategorieAgroalimentaire = True Then
1 If Trim$(txtCategorie.Text) <> "" Then
 txtCategorie.Text = txtCategorie.Text & ", Agroalimentaire"
 Else
 txtCategorie.Text = "Agroalimentaire"
 End If
End If

If m_bCategorieMeuble = True Then
 If Trim$(txtCategorie.Text) <> "" Then
 txtCategorie.Text = txtCategorie.Text & ", Meuble"
 Else
 txtCategorie.Text = "Meuble"
End If
End If

 If m_bCategorieMeunerie = True Then
 If Trim$(txtCategorie.Text) <> "" Then
 txtCategorie.Text = txtCategorie.Text & ", Meunerie"
 Else
 txtCategorie.Text = "Meunerie"
1  End If
 End If

 If m_bCategorieManufacturier = True Then
 If Trim$(txtCategorie.Text) <> "" Then
 txtCategorie.Text = txtCategorie.Text & ", Manufacturier"
 Else
 txtCategorie.Text = "Manufacturier"
 End If
End If

If m_bCategorieConsultant = True Then
 If Trim$(txtCategorie.Text) <> "" Then
 txtCategorie.Text = txtCategorie.Text & ", Consultant"
 Else
 txtCategorie.Text = "Consultant"
 End If
2  End If

If m_bCategorieAsphalte = True Then
If Trim$(txtCategorie.Text) <> "" Then
 txtCategorie.Text = txtCategorie.Text & ", Asphalte"
Else
 txtCategorie.Text = "Asphalte"
End If
End If

If m_bCategorieICPI = True Then
 If Trim$(txtCategorie.Text) <> "" Then
 txtCategorie.Text = txtCategorie.Text & ", ICPI"
 Else
 txtCategorie.Text = "ICPI"
 End If
End If

If m_bCategorieProduitsChimiques = True Then
 If Trim$(txtCategorie.Text) <> "" Then
 txtCategorie.Text = txtCategorie.Text & ", Produits chimiques"
Else
 txtCategorie.Text = "Produits chimiques"
End If
End If

3  If m_bCategorieAutre = True Then
 If Trim$(txtCategorie.Text) <> "" Then
 txtCategorie.Text = txtCategorie.Text & ", Autre"
 Else
 txtCategorie.Text = "Autre"
4 End If
4 End If

4 Exit Sub

Oups:

4 wOups "frmClient", "AfficherCategorie", Err, Err.number, Err.Description
End Sub

Public Sub AfficherContact()

 On Error GoTo Oups
 
 ''''''''''''''''''''''''''''''''''''''''
 'affiche les contacts de l'employé selectionné'
 ''''''''''''''''''''''''''''''''''''''''
 Dim rstContact As ADODB.Recordset

 'Ouverture de la table contact
 Set rstContact = New ADODB.Recordset
 
 Call rstContact.Open("SELECT * FROM Grbcontact WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)
 
 'VIDE les champs
 If m_bModeAjoutContact = True Then
 If m_bNewContact = True Then
 Call ViderBarrerChampsContact(False, True)
 Else
 Call ViderBarrerChampsContact(True, True)
 End If
 Else
  Call ViderBarrerChampsContact(True, True)
  End If

 
 'REMPLIS LES CHAMPS si il y a enregistrement
  If Not rstContact.EOF Then
  If Not IsNull(rstContact.Fields("Titre")) Then
  txtContactTitre.Text = rstContact.Fields("Titre")
  End If

  If Not IsNull(rstContact.Fields("cellulaire")) Then
  txtcontact_cell.Text = rstContact.Fields("cellulaire")
End If
 
1 If Not IsNull(rstContact.Fields("pagette")) Then
 txtcontact_page.Text = rstContact.Fields("pagette")
 End If
 
 If Not IsNull(rstContact.Fields("telephonne")) Then
 txtcontact_tel.Text = rstContact.Fields("telephonne")
 End If
 
 If Not IsNull(rstContact.Fields("fax")) Then
 txtcontact_fax.Text = rstContact.Fields("Fax")
 End If
 
 If Not IsNull(rstContact.Fields("e-mail")) Then
 txtcontact_email.Text = rstContact.Fields("e-mail")
End If
 
 If Not IsNull(rstContact.Fields("noposte")) Then
 txtcontact_poste.Text = rstContact.Fields("noposte")
 End If
 
 If Not IsNull(rstContact.Fields("teldomicile")) Then
 txtcontact_dom.Text = rstContact.Fields("teldomicile")
 End If
1  End If
 
 'ferme la table
 Call rstContact.Close
 Set rstContact = Nothing

Exit Sub

Oups:

wOups "frmClient", "AfficherContact", Err, Err.number, Err.Description
End Sub

Private Sub cmbContact_Click()

 On Error GoTo Oups
 
 ''''''''''''''''''''''''''''''''''
 'affiche employé sélectioné
 ''''''''''''''''''''''''''''''''''
 If cmbContact.ListIndex <> -1 Then
 m_iNoContact = cmbContact.ItemData(cmbContact.ListIndex)
 End If

 Call AfficherContact

 Exit Sub

Oups:

 wOups "frmClient", "cmbContact_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdanulcontact_Click()

 On Error GoTo Oups

 Screen.MousePointer = vbHourglass

 Call AfficherControles(MODE_INACTIF)

 If m_bNewContact = True Then
 Call HideEdMaskContact(True)

 m_bNewContact = False
 End If
 
 'n'est plus en mode ajouter
 m_bModeAjoutContact = False
 
 txtNomClient.Visible = False
 txtNomClient.Locked = False

 'remplis combo contact
 Call RemplirComboContact
 
  Screen.MousePointer = vbDefault

  Exit Sub

Oups:

  wOups "frmClient", "cmdanulcontact_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdCategorie_Click()

 On Error GoTo Oups

 Call frmCategorieClient.AfficherClient

 Call AfficherCategories

 Exit Sub

Oups:

 wOups "frmClient", "cmdCategorie_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdFax_Click()

 On Error GoTo Oups

 If cmbclient.ListCount > 0 Then
 If cmbContact.ListIndex > -1 Then
 Call frmreport.Afficher(cmbclient.ItemData(cmbclient.ListIndex), cmbContact.ItemData(cmbContact.ListIndex), FRM_CLIENTS)
 Else
 Call frmreport.Afficher(cmbclient.ItemData(cmbclient.ListIndex), 0, FRM_CLIENTS)
 End If
 End If

 Exit Sub

Oups:

 wOups "frmClient", "cmdFax_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdMailListContact_Click()

 On Error GoTo Oups

 Dim otlApp As Outlook.Application
 Dim folContact As Outlook.MAPIFolder
 Dim itmContact() As Outlook.ContactItem
 Dim otlRecipient As Outlook.Recipient
 Dim bDejaOuvert As Boolean
 Dim rstContact As ADODB.Recordset
 Dim sIDContact() As String
 Dim sContact() As String
 Dim iCompteur As Integer
 Dim bArrayVide As Boolean
  Dim bNouveau As Boolean
  Dim iResult As Integer
  Dim bPlein As Boolean
  Dim sMsgPlein As String
  Dim iNbreRendu As Integer

  If cmbContact.ListIndex <> -1 Then
  bArrayVide = True

  iResult = MsgBox("Voulez-vous ajouter tous les contacts à la liste de distribution?" & vbNewLine & _
 "Oui - Tous les contacts" & vbNewLine & _
 "Non - Contact affiché seulement", vbYesNoCancel)
 
If iResult = vbYes Then
Set rstContact = New ADODB.Recordset

 Call rstContact.Open("SELECT * FROM GrbContactClient INNER JOIN GrbContact ON GrbContactClient.NoContact = GrbContact.IDContact WHERE GrbContactClient.NoClient = " & cmbclient.ItemData(cmbclient.ListIndex) & " ORDER BY NomContact", g_connData, adOpenForwardOnly, adLockOptimistic)
 
 iCompteur = 0
 
 Do While Not rstContact.EOF
 If Not IsNull(rstContact.Fields("E-mail")) Then
 If Trim(rstContact.Fields("E-mail")) <> "" Then
 bArrayVide = False

 ReDim Preserve sIDContact(0 To iCompteur)
 ReDim Preserve itmContact(0 To iCompteur)
 ReDim Preserve sContact(0 To iCompteur)

 sIDContact(iCompteur) = rstContact.Fields("IDContact")
 sContact(iCompteur) = rstContact.Fields("NomContact")

 iCompteur = iCompteur + 1
 End If
 End If

 Call rstContact.MoveNext
 Loop
 Else
1  If iResult = vbNo Then
 If Trim$(txtcontact_email.Text) <> "" Then
 bArrayVide = False

 ReDim Preserve sIDContact(0 To 0)
 ReDim Preserve itmContact(0 To 0)
 ReDim Preserve sContact(0 To 0)

 sIDContact(0) = cmbContact.ItemData(cmbContact.ListIndex)
 sContact(0) = cmbContact.Text
 End If
 Else
 Exit Sub
 End If
 End If
 
If bArrayVide = False Then
 Set otlApp = OuvrirOutlook(bDejaOuvert)

 lblEtatOutlook.Caption = "Recherche des listes de distribution..."

 fraEtatOutlook.Visible = True

 Call frmChoixMailList.Afficher(Me, otlApp)

 If m_bAnnulerDistList = False Then
 lblEtatOutlook.Caption = "Ajout du contact dans la liste de distribution..."
 
 fraEtatOutlook.Visible = True

 Set folContact = GetFolder(otlApp, "Contacts GRB")

 For iCompteur = 0 To UBound(sIDContact)
 Set itmContact(iCompteur) = folContact.Items.Find("[User1] = " & sIDContact(iCompteur))
 Next

 bPlein = False

 For iCompteur = 0 To UBound(itmContact)
 If Not itmContact(iCompteur) Is Nothing Then
 If m_otlDistList.MemberCount < 10 Then
 Set otlRecipient = otlApp.Session.CreateRecipient(itmContact(iCompteur).Email1DisplayName)

 If otlRecipient.Resolve = True Then
 Call m_otlDistList.AddMember(otlRecipient)
 
 Call m_otlDistList.Save
 Else
 Call MsgBox("Impossible d'ajouter le contact '" & sContact(iCompteur) & "' !", vbOKOnly, "Erreur")
 End If
 Else
 bPlein = True

 Exit For
 End If
 Else
 Call MsgBox("Contact '" & sContact(iCompteur) & "' introuvable!", vbOKOnly, "Erreur")
4 End If
4 Next

4 If bPlein = True Then
4 sMsgPlein = "Les contacts suivants n'ont pas pu être ajouté car la liste contient déjà 10 contacts!" & vbNewLine & _
 vbNewLine

4 iNbreRendu = iCompteur

4 For iCompteur = iNbreRendu To UBound(sContact)
4 sMsgPlein = sMsgPlein & sContact(iCompteur)

4 If iCompteur < UBound(sContact) Then
4 sMsgPlein = sMsgPlein & vbNewLine
4 End If
4 Next

4  Call MsgBox(sMsgPlein, vbOKOnly, "Erreur")
4  End If
4  End If

4  If bDejaOuvert = False Then
4  Call otlApp.Quit
4  End If

4  Set otlApp = Nothing

4  fraEtatOutlook.Visible = False
50 Else
Call MsgBox("Le ou les contacts n'ont pas d'email!", vbOKOnly, "Erreur")
 End If
 Else
 Call MsgBox("Aucun contact sélectionné!", vbOKOnly, "Erreur")
 End If

 Exit Sub

Oups:

 If Err.number = 2 And Erl = 305 Then
 Call MsgBox("La liste de distribution est pleine!", vbOKOnly, "Erreur")
 Else
 wOups "frmClient", "cmdMailListContact_Click", Err, Err.number, Err.Description
 End If

5  fraEtatOutlook.Visible = False
End Sub

Private Sub cmdMailListClient_Click()

 On Error GoTo Oups

 Dim otlApp As Outlook.Application
 Dim folClient As Outlook.MAPIFolder
 Dim itmClient As Outlook.ContactItem
 Dim otlRecipient As Outlook.Recipient
 Dim bDejaOuvert As Boolean

 If cmbclient.ListIndex <> -1 Then
 If Trim$(txtEmail.Text) <> "" Then
 Set otlApp = OuvrirOutlook(bDejaOuvert)

 lblEtatOutlook.Caption = "Recherche des listes de distribution..."

  fraEtatOutlook.Visible = True

  Call frmChoixMailList.Afficher(Me, otlApp)

  If m_bAnnulerDistList = False Then
  lblEtatOutlook.Caption = "Ajout du client dans la liste de distribution..."

  fraEtatOutlook.Visible = True

  If m_otlDistList.MemberCount < 10 Then
  Set folClient = GetFolder(otlApp, "Clients GRB")

  Set itmClient = folClient.Items.Find("[User1] = " & cmbclient.ItemData(cmbclient.ListIndex))

 If Not itmClient Is Nothing Then
 Set otlRecipient = otlApp.Session.CreateRecipient(itmClient.Email1DisplayName)

 If otlRecipient.Resolve = True Then
 Call m_otlDistList.AddMember(otlRecipient)
 
 Call m_otlDistList.Save
 Else
 Call MsgBox("Impossible de trouver le client!", vbOKOnly, "Erreur")
 End If
 Else
 Call MsgBox("Client introuvable!", vbOKOnly, "Erreur")
 End If
 Else
 Call MsgBox("Cette liste de distribution contient déjà 10 contacts!" & vbNewLine & _
 vbNewLine & _
 "Veuillez en choisir une autre.", vbOKOnly, "Erreur")
 End If
 End If

 If bDejaOuvert = False Then
 Call otlApp.Quit
 End If

 Set otlApp = Nothing

1  fraEtatOutlook.Visible = False
 Else
 Call MsgBox("Ce client n'a pas d'email!", vbOKOnly, "Erreur")
 End If
Else
 Call MsgBox("Aucun client sélectionné!", vbOKOnly, "Erreur")
End If

Exit Sub

Oups:

If Err.number = 2 And Erl = 115 Then
 Call MsgBox("La liste de distribution est pleine!", vbOKOnly, "Erreur")
Else
 wOups "frmClient", "cmdMailListClient_Click", Err, Err.number, Err.Description
End If

2  fraEtatOutlook.Visible = False
End Sub

Private Sub cmdreport_Click()

 On Error GoTo Oups

 'Impression de la liste des clients
 Dim rstClient As ADODB.Recordset
 Dim sWhere As String

 Call dlgImpressionClient.Show(vbModal)

 If m_bImpressionAnnuler = False Then
 Set rstClient = New ADODB.Recordset

 If m_bImpressionVille = True Then
 Call frmChoixVille.Show(vbModal)

 If m_bAnnulerVille = False Then
 Call rstClient.Open("SELECT * FROM GrbClient WHERE VilleLiv = '" & Replace(m_sVille, "'", "''") & "' AND Supprimé = False ORDER BY NomClient", g_connData, adOpenForwardOnly, adLockReadOnly)
 Else
  Set rstClient = Nothing

  Exit Sub
  End If
  Else
  If m_bImpressionCategorie = True Then
  Call frmCategorieClient.AfficherImpression

  If m_bImpressionBeton = True Then
  sWhere = "Béton = True"
 End If

 If m_bImpressionPave = True Then
 If Trim$(sWhere) <> "" Then
 sWhere = sWhere & " OR Pavé = True"
 Else
 sWhere = "Pavé = True"
 End If
 End If

 If m_bImpressionPharmaceutique = True Then
 If Trim$(sWhere) <> "" Then
 sWhere = sWhere & " OR Pharmaceutique = True"
 Else
 sWhere = "Pharmaceutique = True"
 End If
 End If
 
 If m_bImpressionAgroAlimentaire = True Then
 If Trim$(sWhere) <> "" Then
 sWhere = sWhere & " OR Agroalimentaire = True"
 Else
1  sWhere = "Agroalimentaire = True"
 End If
 End If

 If m_bImpressionMeuble = True Then
 If Trim$(sWhere) <> "" Then
 sWhere = sWhere & " OR Meuble = True"
 Else
 sWhere = "Meuble = True"
 End If
 End If

 If m_bImpressionMeunerie = True Then
 If Trim$(sWhere) <> "" Then
 sWhere = sWhere & " OR Meunerie = True"
 Else
 sWhere = "Meunerie = True"
 End If
 End If

 If m_bImpressionManufacturier = True Then
 If Trim$(sWhere) <> "" Then
 sWhere = sWhere & " OR Manufacturier = True"
 Else
 sWhere = "Manufacturier = True"
 End If
 End If

 If m_bImpressionConsultant = True Then
 If Trim$(sWhere) <> "" Then
 sWhere = sWhere & " OR Consultant = True"
 Else
 sWhere = "Consultant = True"
 End If
 End If

 If m_bImpressionAsphalte = True Then
 If Trim$(sWhere) <> "" Then
 sWhere = sWhere & " OR Asphalte = True"
 Else
 sWhere = "Asphalte = True"
 End If
 End If

 If m_bImpressionICPI = True Then
 If Trim$(sWhere) <> "" Then
 sWhere = sWhere & " OR ICPI = True"
 Else
4 sWhere = "ICPI = True"
4 End If
4 End If

4 If m_bImpressionProduitsChimiques = True Then
4 If Trim$(sWhere) <> "" Then
4 sWhere = sWhere & " OR ProduitsChimiques = True"
4 Else
4 sWhere = "ProduitsChimiques = True"
4 End If
4 End If
 
4 If m_bImpressionAutre = True Then
4  If Trim$(sWhere) <> "" Then
4  sWhere = sWhere & " OR Autre = True"
4  Else
4  sWhere = "Autre = True"
4  End If
4  End If

4  If Trim$(sWhere) <> "" Then
4  Call rstClient.Open("SELECT * FROM GrbClient WHERE " & sWhere & " AND Supprimé = False ORDER BY NomClient", g_connData, adOpenForwardOnly, adLockReadOnly)
50 Else
 Call rstClient.Open("SELECT * FROM GrbClient WHERE Supprimé = False ORDER BY NomClient", g_connData, adOpenForwardOnly, adLockReadOnly)
 End If
 Else
 If m_bImpressionPotentiel = True Then
 Call rstClient.Open("SELECT * FROM GrbClient WHERE Potentiel = True AND Supprimé = False ORDER BY NomClient", g_connData, adOpenForwardOnly, adLockReadOnly)
 Else
 If m_bImpressionFacturer = True Then
 Call rstClient.Open("SELECT * FROM GrbClient WHERE Potentiel = False AND Supprimé = False ORDER BY NomClient", g_connData, adOpenForwardOnly, adLockReadOnly)
 Else
 Call rstClient.Open("SELECT * FROM GrbClient WHERE Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)
 End If
5  End If
5  End If
5  End If
 
5  Screen.MousePointer = vbHourglass
 
 'set le rapport
5  Set DR_ListeClient.DataSource = rstClient
 
5  DR_ListeClient.Orientation = rptOrientLandscape

5  Call DR_ListeClient.Show(vbModal)
 
5  Call rstClient.Close
60 Set rstClient = Nothing
 
  Screen.MousePointer = vbDefault
  End If

  Exit Sub

Oups:

  wOups "frmClient", "cmdreport_Click", Err, Err.number, Err.Description
End Sub

Private Sub AfficherControles(ByVal eMode As enumMode)

 On Error GoTo Oups
 
 'Activation des controles
 Dim bCmbClient As Boolean
 Dim bTxtNomClient As Boolean
 Dim bCmbContact As Boolean
 Dim bTxtContact As Boolean
 Dim bTxtRechercher As Boolean
 Dim bCmdAdd As Boolean
 Dim bCmdEnr As Boolean
 Dim bCmdModif As Boolean
 Dim bCmdSupp As Boolean
 Dim bFraContact As Boolean
  Dim bCmdAnul As Boolean
  Dim bCmdQuit As Boolean
  Dim bCmdAddCont As Boolean
  Dim bCmdSupContact As Boolean
  Dim bCmdAnulContact As Boolean
  Dim bCmdRenommer As Boolean
  Dim bCmdRafraichir As Boolean
  Dim bCmdImprimer As Boolean
10 Dim bCmdRefCont As Boolean
Dim bCmdRechercher As Boolean
Dim bFax As Boolean
Dim bMailListClient As Boolean
Dim bMailListContact As Boolean
Dim bCategorie As Boolean

m_eMode = eMode
 
Select Case eMode
 'Mode ajout et modif d'un client
 Case MODE_CLIENT:
 bTxtNomClient = True
 bCmdEnr = True
 bCmdAnul = True
 bCategorie = True
 
 'Mode ajout et modif d'un contact
 Case MODE_CONTACT:
 bFraContact = True
 bTxtNomClient = True
 bCmdAnulContact = True
 bCmdRefCont = True
 
 If m_bNewContact = True Then
 bTxtContact = True
 Else
1  bCmbContact = True
 End If
 
 Case MODE_INACTIF:
 bFraContact = True
 bCmbClient = True
 bCmdImprimer = True
 bTxtRechercher = True
 bCmdRenommer = True
 bCmdRafraichir = True
 bCmdAdd = True
 bCmdSupp = True
 bCmdModif = True
 bCmdQuit = True
 bCmdAddCont = True
 bCmdSupContact = True
 bFax = True
 bCmbContact = True
 bMailListClient = True
 bMailListContact = True
 
 If Len(txtRechercher.Text) > 0 Then
 bCmdRechercher = True
 End If
30 End Select
 
cmbclient.Visible = bCmbClient
txtNomClient.Visible = bTxtNomClient
fracontact.Visible = bFraContact
txtRechercher.Enabled = bTxtRechercher
cmdRechercher.Enabled = bCmdRechercher
cmdRafraichir.Enabled = bCmdRafraichir
cmdrenommer.Enabled = bCmdRenommer
cmdReport.Visible = bCmdImprimer
CmdAdd.Visible = bCmdAdd
CmdSupp.Visible = bCmdSupp
CmdModif.Visible = bCmdModif
3  CmdQuit.Visible = bCmdQuit
CmdAnul.Visible = bCmdAnul
3  CmdEnr.Visible = bCmdEnr
CmdAddCont.Visible = bCmdAddCont
3  cmdsupcontact.Visible = bCmdSupContact
cmdanulcontact.Visible = bCmdAnulContact
3  CmdRefCont.Visible = bCmdRefCont
 cmdFax.Visible = bFax
40 txtcontact.Visible = bTxtContact
cmbContact.Visible = bCmbContact
4 cmdMailListClient.Visible = bMailListClient
4 cmdMailListContact.Visible = bMailListContact
4 cmdCategorie.Visible = bCategorie

4 Exit Sub

Oups:

4 wOups "frmClient", "AfficherControles", Err, Err.number, Err.Description
End Sub

Private Sub CmdAdd_Click()

 On Error GoTo Oups
 
 'proc qui permet dajouter un client a la BD
 Dim sName As String
 
 Call AfficherControles(MODE_CLIENT)
 
 'On procede a la saisie du nom du du contact
 sName = InputBox("Veuillez entrer le nom du client" & vbNewLine & _
 vbNewLine & _
 "SVP, respectez le bon orthographe!", "SAISIE DU NOM", "Nom du client")
 
 If sName <> vbNullString Then
 If Not ComboContient(cmbclient, sName) Then
 Screen.MousePointer = vbHourglass
 
 m_bModeAjoutClient = True
 
 'On montre les maskEdBox
 Call HideEdMask(False)
 
 'On affiche le nom du nouveau client dans le textbox
 'pour éviter le ScrollDown durant l'ajout
 txtNomClient.Text = sName
 
 Call ViderBarrerChamps(False, True)
 
  Call mskTelephone.SetFocus
 
  Screen.MousePointer = vbDefault
  Else
  Call MsgBox("Le client " & sName & " existe deja", vbCritical)
 
  Call AfficherControles(MODE_INACTIF)
  End If
  Else
  Call AfficherControles(MODE_INACTIF)
10 End If

Exit Sub

Oups:

wOups "frmClient", "CmdAdd_Click", Err, Err.number, Err.Description
End Sub

Private Sub ViderBarrerChamps(ByVal bLocked As Boolean, ByVal bVider As Boolean)

 On Error GoTo Oups
 
 'Cette procédure vide et unlock tous les textbox pour pouvoir ajouter
 If bVider = True Then
 txtTelephone.Text = vbNullString
 mskTelephone.Text = vbNullString
 txtFax.Text = vbNullString
 mskFax.Text = vbNullString
 txtContactGRB.Text = vbNullString
 txtEmail.Text = vbNullString
 txtAdresse.Text = vbNullString
 txtVille.Text = vbNullString
 txtProvEtat.Text = vbNullString
  txtPays.Text = vbNullString
  txtCP.Text = vbNullString
  txtcommentaire.Text = vbNullString
  txtSiteWeb.Text = vbNullString
  lblDateCreation.Caption = vbNullString
  lblUserCreation.Caption = vbNullString
  lblDateModification.Caption = vbNullString
  lblUserModification.Caption = vbNullString
txtCategorie.Text = vbNullString
1 chkClientPotentiel.Value = vbUnchecked

 m_bCategorieBeton = False
 m_bCategoriePave = False
 m_bCategoriePharmaceutique = False
 m_bCategorieAgroalimentaire = False
 m_bCategorieMeuble = False
 m_bCategorieMeunerie = False
 m_bCategorieManufacturier = False
 m_bCategorieConsultant = False
 m_bCategorieAsphalte = False
 m_bCategorieICPI = False
m_bCategorieProduitsChimiques = False
 m_bCategorieAutre = False
 End If
 
txtTelephone.Locked = bLocked
 txtFax.Locked = bLocked
txtContactGRB.Locked = bLocked
 txtEmail.Locked = bLocked
1  txtAdresse.Locked = bLocked
 txtVille.Locked = bLocked
 txtProvEtat.Locked = bLocked
txtPays.Locked = bLocked
txtCP.Locked = bLocked
txtcommentaire.Locked = bLocked
txtSiteWeb.Locked = bLocked
fraPotentiel.Enabled = Not bLocked

Exit Sub

Oups:

wOups "frmClient", "ViderBarrerChamps", Err, Err.number, Err.Description
End Sub

Private Sub ViderBarrerChampsContact(ByVal bLocked As Boolean, ByVal bVider As Boolean)

 On Error GoTo Oups

 If bVider = True Then
 txtContactTitre.Text = vbNullString
 txtcontact_cell.Text = vbNullString
 txtcontact_dom.Text = vbNullString
 txtcontact_email.Text = vbNullString
 txtcontact_fax.Text = vbNullString
 txtcontact_page.Text = vbNullString
 txtcontact_poste.Text = vbNullString
 txtcontact_tel.Text = vbNullString
 End If

  txtContactTitre.Locked = bLocked
  txtcontact_cell.Locked = bLocked
  txtcontact_dom.Locked = bLocked
  txtcontact_email.Locked = bLocked
  txtcontact_fax.Locked = bLocked
  txtcontact_page.Locked = bLocked
  txtcontact_poste.Locked = bLocked
  txtcontact_tel.Locked = bLocked

10 Exit Sub

Oups:

wOups "frmClient", "ViderBarrerChampsContact", Err, Err.number, Err.Description
End Sub


Private Sub CmdAddCont_Click()

 On Error GoTo Oups
 
 'Pour faire l'ajout d'un contact
 Dim sNom As String
 Dim rstContact As ADODB.Recordset
 Dim bAjouter As Boolean

 If cmbclient.ListCount > 0 Then
 m_bModeAjoutContact = True

 If MsgBox("Voulez-vous ajouter un nouveau contact?" & vbNewLine & _
 "Oui - Nouveau contact" & vbNewLine & _
 "Non - Sélection dans la liste des contacts existant", vbYesNo) = vbYes Then
 m_bNewContact = True

 sNom = InputBox("Quel est le nom du nouveau contact?" & vbNewLine & _
 vbNewLine & _
 "SVP, respectez le bon orthographe!")

 If sNom <> vbNullString Then
 If ExisteDansBD(sNom) = True Then
  If MsgBox("Il y a déjà un contact portant ce nom!" & vbNewLine & "Voulez-vous l'ajouter quand même?", vbYesNo) = vbYes Then
  bAjouter = True
  Else
  bAjouter = False
  End If
  Else
  If ContientCaracteresIncorrects(sNom) = True Then
  Call MsgBox("Il y a des caractères non permis!", vbOKOnly, "Erreur")

 bAjouter = False
 Else
 bAjouter = True
 End If
 End If
 Else
 bAjouter = False
 End If

 If bAjouter = True Then
 txtcontact.Text = sNom

 Call ViderBarrerChampsContact(False, True)

 Call HideEdMaskContact(False)

 Call mskContactTel.SetFocus

 txtNomClient.Visible = True
 txtNomClient.Locked = True
 
 'Remplis combo avec tout les contact existant
 Call AfficherControles(MODE_CONTACT)

 Call txtContactTitre.SetFocus
 End If
 Else
1  Screen.MousePointer = vbHourglass


 m_bNewContact = False

 txtNomClient.Visible = True
 txtNomClient.Locked = True
 
 'Remplis combo avec tout les contact existant
 Call AfficherControles(MODE_CONTACT)
 
 'Affiche client no modifiable
 Call RemplirComboContact
 End If
 
 Screen.MousePointer = vbDefault
Else
 Call MsgBox("Aucun enregistrement de sélectionné")
End If

Exit Sub

Oups:

 wOups "frmClient", "CmdAddCont_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdAnul_Click()

 On Error GoTo Oups

 Screen.MousePointer = vbHourglass
 
 'On cache le maskEdBox
 Call HideEdMask(True)
 
 'commentaire unlock
 txtcommentaire.Locked = True
 
 m_bModeAjoutClient = False

 'on retablis les bouttons
 Call AfficherControles(MODE_INACTIF)

 'on affiche les donnée du premier enreg
 Call ViderBarrerChamps(True, True)
 
 Call cmbclient_Click
 
 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmClient", "CmdAnul_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdEnr_Click()

 On Error GoTo Oups

 Dim sClient As String
 Dim iCompteur As Integer
 
 'Nom du client
 sClient = txtNomClient.Text
 
 'Enregistrement d'un client dans le BD
 Screen.MousePointer = vbHourglass
 
 Call EnregistrerClient
 
 Call ViderBarrerChamps(True, True)
 
 'On cache les MaskEdBox
 Call HideEdMask(True)
 
 'On met à jour le combo
 Call RemplirComboClient
 
 'Retablir les boutons
 Call AfficherControles(MODE_INACTIF)
 
 For iCompteur = 0 To cmbclient.ListCount - 1
  If cmbclient.LIST(iCompteur) = sClient Then
  cmbclient.ListIndex = iCompteur
 
  Exit For
  End If
  Next
 
  Call cmbclient.SetFocus
 
  Screen.MousePointer = vbDefault

  Exit Sub

Oups:

10 wOups "frmClient", "CmdEnr_Click", Err, Err.number, Err.Description
End Sub

Private Sub EnregistrerClient()

 On Error GoTo Oups

 Dim rstClient As ADODB.Recordset
 
 Set rstClient = New ADODB.Recordset
 
 If m_bModeAjoutClient = True Then
 Call rstClient.Open("SELECT * FROM GrbClient", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call rstClient.AddNew

 rstClient.Fields("DateCréation") = ConvertDate(Date)
 rstClient.Fields("UserCréation") = g_sInitiale
 Else
 Call rstClient.Open("SELECT * FROM GrbClient WHERE IDClient = " & m_iNoClient, g_connData, adOpenDynamic, adLockOptimistic)

 rstClient.Fields("DateModification") = ConvertDate(Date)
  rstClient.Fields("UserModification") = g_sInitiale
  End If
 
 'Enregistre le client
  rstClient.Fields("NomClient").Value = txtNomClient.Text
  rstClient.Fields("Telephonne").Value = mskTelephone.Text
  rstClient.Fields("Fax").Value = mskFax.Text
  rstClient.Fields("ContactGRB").Value = txtContactGRB.Text
  rstClient.Fields("Email").Value = txtEmail.Text
  rstClient.Fields("AdresseLiv").Value = txtAdresse.Text
10 rstClient.Fields("VilleLiv").Value = txtVille.Text
rstClient.Fields("Prov/EtatLiv").Value = txtProvEtat.Text
rstClient.Fields("PaysLiv").Value = txtPays.Text
rstClient.Fields("Commentaire").Value = txtcommentaire.Text
rstClient.Fields("CodePostalLiv").Value = txtCP.Text
rstClient.Fields("SiteWeb").Value = txtSiteWeb.Text

If chkClientPotentiel.Value = vbChecked Then
 rstClient.Fields("Potentiel").Value = True
Else
 rstClient.Fields("Potentiel").Value = False
End If

rstClient.Fields("Béton").Value = m_bCategorieBeton
1  rstClient.Fields("Pavé").Value = m_bCategoriePave
rstClient.Fields("Pharmaceutique").Value = m_bCategoriePharmaceutique
 rstClient.Fields("Agroalimentaire").Value = m_bCategorieAgroalimentaire
rstClient.Fields("Meuble").Value = m_bCategorieMeuble
 rstClient.Fields("Meunerie").Value = m_bCategorieMeunerie
rstClient.Fields("Manufacturier").Value = m_bCategorieManufacturier
 rstClient.Fields("Consultant").Value = m_bCategorieConsultant
1  rstClient.Fields("Asphalte").Value = m_bCategorieAsphalte
 rstClient.Fields("ICPI").Value = m_bCategorieICPI
 rstClient.Fields("ProduitsChimiques").Value = m_bCategorieProduitsChimiques
rstClient.Fields("Autre").Value = m_bCategorieAutre

rstClient.Fields("EntryIDOutlook") = ModifierClientExchange(rstClient.Fields("IDClient"))
 
If m_bModeAjoutClient = True Then
 m_bModeAjoutClient = False
End If

Call rstClient.Update

Call rstClient.Close
Set rstClient = Nothing

Exit Sub

Oups:

wOups "frmClient", "EnregistrerClient", Err, Err.number, Err.Description
End Sub

Private Sub ModifierNomClientExchange(ByVal sName As String, ByVal iClientID As Integer)

 On Error GoTo Oups
 
 Dim otlApp As Outlook.Application
 Dim otlClient As Outlook.ContactItem
 Dim folClient As Outlook.MAPIFolder
 Dim bDejaOuvert As Boolean

 lblEtatOutlook.Caption = "Modification du nom du client dans Outlook ..."
 fraEtatOutlook.Visible = True

 Set otlApp = OuvrirOutlook(bDejaOuvert)

 Set folClient = GetFolder(otlApp, "Clients GRB")

 Set otlClient = folClient.Items.Find("[User1] = " & iClientID)
 
 If otlClient Is Nothing Then
  Call MsgBox("Le client " & txtNomClient.Text & " n'a pas été trouvé pour la mise à jour Exchange!", vbOKOnly, "Erreur")

  fraEtatOutlook.Visible = False

  DoEvents

  Exit Sub
  End If

  otlClient.CompanyName = sName
 
  Call otlClient.Save

  If bDejaOuvert = False Then
Call otlApp.Quit
End If

Set otlApp = Nothing

fraEtatOutlook.Visible = False

DoEvents

Exit Sub

Oups:

woups"frmClient", "ModifierNomClientExchange", Err, Erl, "iClientID = " & iClientID)

fraEtatOutlook.Visible = False
End Sub

Private Sub LierContactClient(ByVal iClientID As Integer)

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

 lblEtatOutlook.Caption = "Liaison du contact avec le client dans Outlook ..."
  fraEtatOutlook.Visible = True

  Set otlApp = OuvrirOutlook(bDejaOuvert)

  Set folClient = GetFolder(otlApp, "Clients GRB")
  Set folContact = GetFolder(otlApp, "Contacts GRB")

  Set rstClient = New ADODB.Recordset

  Call rstClient.Open("SELECT EntryIDOutlook FROM GrbClient WHERE IDClient = " & m_iNoClient, g_connData, adOpenForwardOnly, adLockReadOnly)

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

 Call rstContactClient.Open("SELECT * FROM GrbContactClient WHERE NoClient = " & m_iNoClient, g_connData, adOpenForwardOnly, adLockReadOnly)

 Do While Not rstContactClient.EOF
 Set itmContact = folContact.Items.Find("[User1] = " & rstContactClient.Fields("NoContact"))

 If Not itmContact Is Nothing Then
1  Call itmClient.Links.Add(itmContact)

 Call itmClient.Save

 Call itmContact.Links.Add(itmClient)

 Call itmContact.Save
 End If

 Call rstContactClient.MoveNext
 Loop

 Call rstContactClient.Close
 Set rstContactClient = Nothing
Else
 Call MsgBox("Client introuvable!", vbOKOnly, "Erreur")

 Call rstClient.Close
 Set rstClient = Nothing
2  End If

If bDejaOuvert = False Then
Call otlApp.Quit
End If

2  Set otlApp = Nothing

fraEtatOutlook.Visible = False

2  DoEvents

Exit Sub

Oups:

30 If InStr(1, UCase(Err.Description), "THE OPERATION FAILED") > 0 Then
3 Call MsgBox("Une erreur est survenue ! " & vbNewLine & _
 vbNewLine & _
 "Pour réparer cette erreur, veuillez suivre les instructions suivantes : " & vbNewLine & _
 vbNewLine & _
 "- Dans Outlook, ouvrez le client '" & txtNomClient.Text & "' dans Clients GRB" & vbNewLine & _
 "- Cliquez sur tous les contacts de ce client 1 à la fois pour trouver le contact incorrect." & vbNewLine & _
 "- Effacez ce contact de la liste des contacts de ce client." & vbNewLine & _
 "- Dans le logiciel GRB, recommencez l'opération (effacez le contact et l'ajouter de nouveau).", vbOKOnly, "Erreur")
Else
 woups"frmClient", "LierContactClient", Err, Erl, txtNomClient.Text)
End If

fraEtatOutlook.Visible = False
End Sub

Private Function ModifierClientExchange(ByVal iClientID As Integer) As String

 On Error GoTo Oups

 Dim otlApp As Outlook.Application
 Dim otlClient As Outlook.ContactItem
 Dim folClient As Outlook.MAPIFolder
 Dim bDejaOuvert As Boolean

 If m_bModeAjoutClient = True Then
 lblEtatOutlook.Caption = "Ajout du client dans Outlook ..."
 Else
 lblEtatOutlook.Caption = "Modification du client dans Outlook ..."
 End If

 fraEtatOutlook.Visible = True

  Set otlApp = OuvrirOutlook(bDejaOuvert)

  Set folClient = GetFolder(otlApp, "Clients GRB")

  If m_bModeAjoutClient = True Then
  Set otlClient = folClient.Items.Add(olContactItem)

  otlClient.User1 = iClientID
  Else
  Set otlClient = folClient.Items.Find("[User1] = " & iClientID)
  End If

10 If otlClient Is Nothing Then
1 Call MsgBox("Le client " & txtNomClient.Text & " n'a pas été trouvé pour la mise à jour Exchange!", vbOKOnly, "Erreur")

 fraEtatOutlook.Visible = False

 DoEvents

 Exit Function
End If

otlClient.CompanyName = txtNomClient.Text
 
If mskTelephone.Text <> "(___) ___-____" Then
 otlClient.BusinessTelephoneNumber = mskTelephone.Text
End If
 
If mskFax.Text <> "(___) ___-____" Then
 otlClient.BusinessFaxNumber = mskFax.Text
1  End If
 
otlClient.Email1Address = txtEmail.Text
 otlClient.BusinessAddressStreet = txtAdresse.Text
otlClient.BusinessAddressCity = txtVille.Text
 otlClient.BusinessAddressState = txtProvEtat.Text
otlClient.BusinessAddressCountry = txtPays.Text
 otlClient.BusinessAddressPostalCode = txtCP.Text
1  otlClient.Body = txtcommentaire.Text
 otlClient.WebPage = txtSiteWeb.Text
 
 Call otlClient.Save

ModifierClientExchange = otlClient.EntryID

If bDejaOuvert = False Then
 Call otlApp.Quit
End If

Set otlApp = Nothing

fraEtatOutlook.Visible = False

DoEvents

Exit Function

Oups:

woups"frmClient", "ModifierClientExchange", Err, Erl, "iClientID = " & iClientID)

fraEtatOutlook.Visible = False
End Function

Private Function AjouterContactExchange(ByVal iContactID As Integer) As String
 
 On Error GoTo Oups
 
 Dim otlApp As Outlook.Application
 Dim otlContact As Outlook.ContactItem
 Dim folContact As Outlook.MAPIFolder
 Dim bDejaOuvert As Boolean
 Dim sNom() As String

 lblEtatOutlook.Caption = "Ajout du contact dans Outlook ..."
 fraEtatOutlook.Visible = True

 Set otlApp = OuvrirOutlook(bDejaOuvert)

 Set folContact = GetFolder(otlApp, "Contacts GRB")

 Set otlContact = folContact.Items.Add(olContactItem)

  otlContact.User1 = iContactID
 
  sNom = Split(Trim$(txtcontact.Text), " ")

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
1End Select
 
otlContact.Title = ""
 
otlContact.CompanyName = txtNomClient.Text
otlContact.JobTitle = txtContactTitre.Text

If Trim$(mskContactTel.Text) <> "" Then
 If txtTelephone.Text <> "(___) ___-____" Then
 If Trim$(txtcontact_poste.Text) <> "" Then
 otlContact.BusinessTelephoneNumber = mskContactTel.Text & " Ext : " & txtcontact_poste.Text
 Else
 otlContact.BusinessTelephoneNumber = mskContactTel.Text
 End If
End If
End If
 
 If mskContactFax.Text <> "(___) ___-____" Then
 otlContact.BusinessFaxNumber = mskContactFax.Text
 End If

If mskContactDom.Text <> "(___) ___-____" Then
 otlContact.HomeTelephoneNumber = mskContactDom.Text
1  End If
 
 If mskContactCell.Text <> "(___) ___-____" Then
 otlContact.MobileTelephoneNumber = mskContactCell.Text
End If
 
If mskContactPage.Text <> "(___) ___-____" Then
 otlContact.PagerNumber = mskContactPage.Text
End If
 
otlContact.Email1Address = txtcontact_email.Text
 
Call otlContact.Save

AjouterContactExchange = otlContact.EntryID

If bDejaOuvert = False Then
 Call otlApp.Quit
End If

2  Set otlApp = Nothing

fraEtatOutlook.Visible = False

2  DoEvents

Exit Function

Oups:

2  woups"frmClient", "AjouterContactExchange", Err, Erl, "iContactID = " & iContactID)

fraEtatOutlook.Visible = False
End Function

Private Sub SupprimerClientExchange(ByVal iClientID As Integer)

 On Error GoTo Oups

 Dim otlApp As Outlook.Application
 Dim otlClient As Outlook.ContactItem
 Dim folClient As MAPIFolder
 Dim bDejaOuvert As Boolean

 lblEtatOutlook.Caption = "Suppression du client dans Outlook ..."
 fraEtatOutlook.Visible = True

 Set otlApp = OuvrirOutlook(bDejaOuvert)

 Set folClient = GetFolder(otlApp, "Clients GRB")

 Set otlClient = folClient.Items.Find("[User1] = " & iClientID)

 If Not otlClient Is Nothing Then
  Call otlClient.Delete
  End If

  If bDejaOuvert = False Then
  Call otlApp.Quit
  End If

  Set otlApp = Nothing

  fraEtatOutlook.Visible = False

  DoEvents

10 Exit Sub

Oups:

wOups "frmClient", "SupprimerClientExchange", Err, Err.number, Err.Description

fraEtatOutlook.Visible = False
End Sub

Private Sub CmdModif_Click()

 On Error GoTo Oups

 Dim iCompteur As Integer
 
 If cmbclient.ListCount > 0 Then
 Screen.MousePointer = vbHourglass
 
 'proc qui permet de modifier l'enr courant
 'on montre/cache des buttons
 Call HideEdMask(False)
 
 Call AfficherControles(MODE_CLIENT)
 
 Call ViderBarrerChamps(False, False)
 
 'pour que le nom ne soit pas modifiable
 txtNomClient.Visible = True
 txtNomClient.Locked = True
 
 Screen.MousePointer = vbDefault
 Else
  Call MsgBox("Aucun enregistrement de sélectionné!")
  End If

  Exit Sub

Oups:

  wOups "frmClient", "CmdModif_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdquit_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmClient", "cmdquit_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdRefCont_Click()

 On Error GoTo Oups

 Dim rstContactClient As ADODB.Recordset
 Dim rstContact As ADODB.Recordset
 
 Screen.MousePointer = vbHourglass
 
 Set rstContactClient = New ADODB.Recordset

 If m_bNewContact = True Then
 Set rstContact = New ADODB.Recordset

 Call rstContact.Open("SELECT * FROM GrbContact", g_connData, adOpenDynamic, adLockOptimistic)

 Call rstContact.AddNew

 rstContact.Fields("NomContact").Value = txtcontact.Text
 rstContact.Fields("Titre").Value = txtContactTitre.Text
  rstContact.Fields("Compagnie").Value = txtNomClient.Text
  rstContact.Fields("Telephonne").Value = mskContactTel.Text
  rstContact.Fields("Fax").Value = mskContactFax.Text
  rstContact.Fields("Pagette").Value = mskContactPage.Text
  rstContact.Fields("Cellulaire").Value = mskContactCell.Text
  rstContact.Fields("E-mail").Value = txtcontact_email.Text
  rstContact.Fields("NoPoste").Value = txtcontact_poste.Text
  rstContact.Fields("TelDomicile").Value = mskContactDom.Text
rstContact.Fields("UserCréation").Value = g_sInitiale
1 rstContact.Fields("DateCréation").Value = ConvertDate(Date)
 
 rstContact.Fields("EntryIDOutlook") = AjouterContactExchange(rstContact.Fields("IDContact"))

 Call rstContact.Update

 'Set la table
 Call rstContactClient.Open("SELECT * FROM GrbContactClient WHERE NoClient = " & m_iNoClient & " And NoContact = " & rstContact.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'Si pas déjà existant, on ajoute dans la table
 If rstContactClient.EOF Then
 'Ajoute dans la table
 Call rstContactClient.AddNew
 
 rstContactClient.Fields("NoClient") = m_iNoClient
 rstContactClient.Fields("NoContact") = rstContact.Fields("IDContact")
 
 Call rstContactClient.Update
 End If
 
 Call rstContact.Close
Set rstContact = Nothing

 m_bNewContact = False

 Call HideEdMaskContact(True)
Else
 'Set la table
 Call rstContactClient.Open("SELECT * FROM GrbContactClient WHERE NoClient = " & m_iNoClient & " AND NoContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)
 
 'Si pas déjà existant, on ajoute dans la table
 If rstContactClient.EOF Then
 'Ajoute dans la table
 Call rstContactClient.AddNew
 
1  rstContactClient.Fields("NoClient") = m_iNoClient
 rstContactClient.Fields("NoContact") = m_iNoContact
 
 Call rstContactClient.Update
 End If
 
 'Ferme tables et connexion
 Call rstContactClient.Close
End If

Call LierContactClient(m_iNoClient)

 'Ferme tables et connection
Set rstContactClient = Nothing
 
 'Bouton comme avant(apparait)
Call AfficherControles(MODE_INACTIF)
 
 'n'est plus en mode ajouter
m_bModeAjoutContact = False

 'remplis combo contact
Call RemplirComboContact

Call ViderBarrerChampsContact(True, False)

Screen.MousePointer = vbDefault

2  Exit Sub

Oups:

wOups "frmClient", "CmdRefCont_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRenommer_Click()

 On Error GoTo Oups
 
 '''''''''''''''''''''''''''''''''''''''
 'on renomme le nom du CLIENT
 ''''''''''''''''''''''''''''''''''''''''
 Dim rstClient As ADODB.Recordset
 Dim sName As String

 If cmbclient.ListCount > 0 Then
 'Proc qui permet de modifié un CLIENT a la BD
 'On procède à la saisie du nom du CLIENT
 sName = InputBox("Veuillez entrer le nom du client", "Renommer client", txtNomClient.Text)
 
 If sName <> vbNullString Then
 If sName <> txtNomClient.Text Then
 Set rstClient = New ADODB.Recordset
 
 Call rstClient.Open("SELECT * FROM GrbClient WHERE NomClient = '" & Replace(sName, "'", "''") & "' AND Supprimé = False", g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstClient.EOF Then
 'Transfert du nom au premier textBox
  Screen.MousePointer = vbHourglass
 
  Call rstClient.Close
 
  Call rstClient.Open("SELECT NomClient, IDClient, EntryIDOutlook FROM GrbClient WHERE IDClient = " & m_iNoClient, g_connData, adOpenDynamic, adLockOptimistic)
 
  Call ModifierNomClientExchange(sName, m_iNoClient)
 
  txtNomClient.Text = sName
 
 'Transfert des données
  rstClient.Fields("NomClient").Value = txtNomClient.Text
 
 'Mise à jour de la base de données
  Call rstClient.Update
 
  Call rstClient.Close
 
 Call RemplirComboClient
 
 cmbclient.Text = sName
 
 m_bRenommer = True
 
 Call cmbclient_Click
 
 m_bRenommer = False
 
 Screen.MousePointer = vbDefault
 Else
 Call MsgBox("Le client " & sName & " existe déjà!", vbCritical)
 
 Call rstClient.Close
 End If
 
 Set rstClient = Nothing
 End If
End If
Else
 Call MsgBox("Aucun enregistrement de sélectionné!")
End If

 Exit Sub

Oups:

wOups "frmClient", "cmdRenommer_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdsupcontact_Click()

 On Error GoTo Oups
 
 'fonction qui supprime l'enregistrement courant
 If cmbContact.ListCount > 0 Then
 If MsgBox("Etes-vous sur de vouloir supprimer cette enregistrement?", vbYesNo, "Supprimer") = vbYes Then
 Screen.MousePointer = vbHourglass
 
 Call g_connData.Execute("DELETE * FROM GrbContactClient WHERE NoClient = " & m_iNoClient & " AND NoContact = " & m_iNoContact)
 
 Call LierContactClient(m_iNoClient)
 
 'remplis le combo employé
 Call RemplirComboContact
 
 Screen.MousePointer = vbDefault
 End If
 End If

 Exit Sub

Oups:

  wOups "frmClient", "cmdsupcontact_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdSupp_Click()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim rstClient As ADODB.Recordset
 Dim bPeutEffacer As Boolean
 
 If cmbclient.ListCount > 0 Then
 'fonction qui supprime l'enregistrement courant
 If MsgBox("Êtes-vous sûr de supprimer cet enregistrement?", vbYesNo, "Supprimer") = vbYes Then
 Screen.MousePointer = vbHourglass
 
 Set rstProjSoum = New ADODB.Recordset

 'open table
 Call rstProjSoum.Open("SELECT * FROM GrbSoumissionMec WHERE IDClient = " & m_iNoClient, g_connData, adOpenDynamic, adLockOptimistic)

 'si existe pas dans soumission, on peut le deleter
 If rstProjSoum.EOF Then
 Call rstProjSoum.Close
 
 'S'il existe pas dans projet, on peut le deleter
  Call rstProjSoum.Open("SELECT * FROM GrbProjetMec WHERE IDClient = " & m_iNoClient, g_connData, adOpenDynamic, adLockOptimistic)
 
  If rstProjSoum.EOF Then
  Call rstProjSoum.Close
 
  Call rstProjSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDClient = " & m_iNoClient, g_connData, adOpenDynamic, adLockOptimistic)
 
  If rstProjSoum.EOF Then
  Call rstProjSoum.Close
 
  Call rstProjSoum.Open("SELECT * FROM GrbProjetElec WHERE IDClient = " & m_iNoClient, g_connData, adOpenDynamic, adLockOptimistic)
 
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
 Else
1  bPeutEffacer = False
 
 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 End If

 If cmbContact.ListCount > 0 Then
 'Delete les contact pour ce client
 Call g_connData.Execute("DELETE * FROM GrbContactClient WHERE NoClient = " & m_iNoClient)
 End If

 Call SupprimerClientExchange(m_iNoClient)

 If bPeutEffacer = True Then
 'Delete le client
 Call g_connData.Execute("DELETE * FROM GrbClient WHERE IDClient = " & m_iNoClient)
 Else
 Set rstClient = New ADODB.Recordset

 Call rstClient.Open("SELECT * FROM GrbClient WHERE IDClient = " & m_iNoClient, g_connData, adOpenDynamic, adLockOptimistic)

 rstClient.Fields("Supprimé") = True

 Call rstClient.Update

 Call rstClient.Close
 Set rstClient = Nothing
 End If

 Call RemplirComboClient
 
 Screen.MousePointer = vbDefault
 End If
30 Else
3 Call MsgBox("Aucun enregistrement de sélectionné!")
End If

Exit Sub

Oups:

wOups "frmClient", "CmdSupp_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbclient_Click()

 On Error GoTo Oups
 
 'Quand le user selectionne un enregistrement on se posotionne dessus
 If cmbclient.Text <> vbNullString Then
 txtNomClient.Text = cmbclient.Text
 Else
 If ComboContient(cmbclient, txtNomClient.Text) = False Then
 Call RemplirComboClient
 End If

 cmbclient.Text = txtNomClient.Text
 End If
 
 If cmbclient.ListIndex > -1 Then
 If m_bRenommer = False And m_bModeAjoutClient = False Then
  m_iNoClient = cmbclient.ItemData(cmbclient.ListIndex)
  End If
  End If
 
 'remplis le combo dépendant le client sélectionné
  Call AfficherClient

  Call RemplirComboContact

  Exit Sub

Oups:

  wOups "frmClient", "cmbclient_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Call RemplirComboClient
 
 Call AfficherControles(MODE_INACTIF)

 Call ActiverBoutonsGroupe

 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmClient", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub ActiverBoutonsGroupe()

 On Error GoTo Oups

 CmdAdd.Enabled = g_bModificationClients
 CmdAddCont.Enabled = g_bModificationClients
 CmdModif.Enabled = g_bModificationClients
 cmdrenommer.Enabled = g_bModificationClients
 cmdsupcontact.Enabled = g_bModificationClients
 CmdSupp.Enabled = g_bModificationClients
 cmdMailListClient.Enabled = g_bModificationListeDistribution
 cmdMailListContact.Enabled = g_bModificationListeDistribution

 Exit Sub

Oups:

 wOups "frmClient", "ActiverBoutonsGroupe", Err, Err.number, Err.Description
End Sub

Private Sub HideEdMask(ByVal bVisible As Boolean)

 On Error GoTo Oups
 
 'proc qui rend visible/ou non les maskEdBox
 'On en profite pour les nettoyer du dernier Enregistrement
 'et on fait l'inverse avec les textBox
 If m_bModeAjoutClient = True Then
 txtTelephone.Text = vbNullString
 txtFax.Text = vbNullString
 Else
 mskTelephone.Text = txtTelephone.Text
 mskFax.Text = txtFax.Text
 End If
 
 mskTelephone.Visible = Not bVisible
 mskFax.Visible = Not bVisible
 
 txtTelephone.Visible = bVisible
  txtFax.Visible = bVisible

  Exit Sub

Oups:

  wOups "frmClient", "HideEdMask", Err, Err.number, Err.Description
End Sub

Private Sub HideEdMaskContact(ByVal bVisible As Boolean)

 On Error GoTo Oups
 
 'proc qui rend visible/ou non les maskEdBox
 'On en profite pour les nettoyer du dernier Enregistrement
 'et on fait l'inverse avec les textBox
 If m_bModeAjoutContact = True Then
 txtcontact_tel.Text = vbNullString
 txtcontact_fax.Text = vbNullString
 txtcontact_page.Text = vbNullString
 txtcontact_cell.Text = vbNullString
 txtcontact_dom.Text = vbNullString
 
 mskContactTel.Text = vbNullString
 mskContactFax.Text = vbNullString
 mskContactPage.Text = vbNullString
 mskContactCell.Text = vbNullString
  mskContactDom.Text = vbNullString
  End If
 
  mskContactTel.Visible = Not bVisible
  txtcontact_tel.Visible = bVisible

  mskContactFax.Visible = Not bVisible
  txtcontact_fax.Visible = bVisible

  mskContactPage.Visible = Not bVisible
  txtcontact_page.Visible = bVisible

10 mskContactCell.Visible = Not bVisible
txtcontact_cell.Visible = bVisible

mskContactDom.Visible = Not bVisible
txtcontact_dom.Visible = bVisible

Exit Sub

Oups:

wOups "frmClient", "HideEdMaskContact", Err, Err.number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)

 On Error GoTo Oups

 Set FrmClient = Nothing

 Exit Sub

Oups:

 wOups "frmClient", "Form_Unload", Err, Err.number, Err.Description
End Sub
Private Sub mskTelephone_GotFocus()

 On Error GoTo Oups

 mskTelephone.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmClient", "mskTelephone_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskTelephone_LostFocus()

 On Error GoTo Oups

 mskTelephone.mask = vbNullString
 If mskTelephone.Text = "(___) ___-____" Then
 mskTelephone.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmClient", "mskTelephone_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskFax_GotFocus()

 On Error GoTo Oups

 mskFax.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmClient", "mskFax_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskFax_LostFocus()

 On Error GoTo Oups

 mskFax.mask = vbNullString
 If mskFax.Text = "(___) ___-____" Then
 mskFax.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmClient", "mskFax_LostFocus", Err, Err.number, Err.Description
End Sub

Public Sub RemplirComboContact()

 On Error GoTo Oups
 
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 'remplis le combo contact dépendant le client sélectionné
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Dim rstClient As ADODB.Recordset
 Dim rstContact As ADODB.Recordset
 
 Set rstContact = New ADODB.Recordset
 
 If m_bModeAjoutContact = True Then
 Call rstContact.Open("SELECT NomContact, IDContact FROM Grbcontact WHERE Supprimé = False ORDER BY NomContact", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstContact.Open("SELECT GrbContact.NomContact, GrbContact.IDContact FROM GrbContact INNER JOIN GrbContactClient ON GrbContact.IDContact = GrbContactClient.NoContact WHERE GrbContactClient.NoClient = " & m_iNoClient & " ORDER BY NomContact", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 
 Call cmbContact.Clear
 
 Do While Not rstContact.EOF
  Call cmbContact.AddItem(Trim(rstContact.Fields("NomContact")))
  cmbContact.ItemData(cmbContact.newIndex) = rstContact.Fields("IDContact")
 
  Call rstContact.MoveNext
  Loop
 
 'ferme la table "GrbContact"
  Call rstContact.Close
  Set rstContact = Nothing
 
 'affiche le contact de la table client
 'si combo est pas vide, affiche le premier contact, sinon le contact inscrit dans table client
  If cmbContact.ListCount > 0 Then
  cmbContact.ListIndex = 0
10 Else
 'VIDE les champs
1 txtContactTitre.Text = vbNullString
 txtcontact_cell.Text = vbNullString
 txtcontact_email = vbNullString
 txtcontact_fax = vbNullString
 txtcontact_page = vbNullString
 txtcontact_poste = vbNullString
 txtcontact_tel = vbNullString
 txtcontact_dom.Text = vbNullString
End If

Exit Sub

Oups:

wOups "frmClient", "RemplirComboContact", Err, Err.number, Err.Description
End Sub

Private Sub cmdRechercher_Click()

 On Error GoTo Oups

 Dim rstClient As ADODB.Recordset
 Dim sSearch As String
 
 Screen.MousePointer = vbHourglass
 
 sSearch = txtRechercher.Text
 
 'vide les champs
 Call ViderBarrerChamps(True, True)
 
 'Filtre pour sélection des Nomcontact
 'goSQL = "SELECT * FROM Grbcontact order by NomContact"
 Set rstClient = New ADODB.Recordset
 
 Call rstClient.Open("SELECT NomClient, IDClient FROM GrbClient WHERE Instr(1, NomClient, '" & Replace(sSearch, "'", "''") & "') > 0 AND Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)
 
 'vide combo
 Call cmbclient.Clear
 
 Do While Not rstClient.EOF
 Call cmbclient.AddItem(rstClient.Fields("NomClient"))
  cmbclient.ItemData(cmbclient.newIndex) = rstClient.Fields("IDClient")
 
  Call rstClient.MoveNext
  Loop
 
  Call rstClient.Close
  Set rstClient = Nothing
 
  Screen.MousePointer = vbDefault

  If cmbclient.ListCount > 0 Then
  cmbclient.ListIndex = 0
10 Else
1 Call cmbContact.Clear
 
 'VIDE les champs
 txtContactTitre.Text = vbNullString
 txtcontact_cell.Text = vbNullString
 txtcontact_email.Text = vbNullString
 txtcontact_fax.Text = vbNullString
 txtcontact_page.Text = vbNullString
 txtcontact_poste.Text = vbNullString
 txtcontact_tel.Text = vbNullString
 txtcontact_dom.Text = vbNullString
End If

Exit Sub

Oups:

1  wOups "frmClient", "cmdRechercher_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRafraichir_Click()

 On Error GoTo Oups

 Screen.MousePointer = vbHourglass
 
 Call RemplirComboClient
 
 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmClient", "cmdRafraichir_Click", Err, Err.number, Err.Description
End Sub

Private Sub txtRechercher_Change()

 On Error GoTo Oups

 If Len(txtRechercher.Text) > 0 Then
 cmdRechercher.Enabled = True
 Else
 cmdRechercher.Enabled = False
 End If

 Exit Sub

Oups:

 wOups "frmClient", "txtRechercher_Change", Err, Err.number, Err.Description
End Sub

Private Sub mskContactTel_GotFocus()

 On Error GoTo Oups

 mskContactTel.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmClient", "mskContactTel_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskContactTel_LostFocus()

 On Error GoTo Oups

 mskContactTel.mask = vbNullString

 If mskContactTel.Text = "(___) ___-____" Then
 mskContactTel.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmClient", "mskContactTel_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskContactFax_GotFocus()

 On Error GoTo Oups

 mskContactFax.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmClient", "mskContactFax_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskContactFax_LostFocus()

 On Error GoTo Oups

 mskContactFax.mask = vbNullString

 If mskContactFax.Text = "(___) ___-____" Then
 mskContactFax.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmClient", "mskContactFax_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskContactCell_GotFocus()

 On Error GoTo Oups

 mskContactCell.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmClient", "mskContactCell_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskContactCell_LostFocus()

 On Error GoTo Oups

 mskContactCell.mask = vbNullString

 If mskContactCell.Text = "(___) ___-____" Then
 mskContactCell.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmClient", "mskContactCell_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskContactDom_GotFocus()

 On Error GoTo Oups

 mskContactDom.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmClient", "mskContactDom_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskContactDom_LostFocus()

 On Error GoTo Oups

 mskContactDom.mask = vbNullString

 If mskContactDom.Text = "(___) ___-____" Then
 mskContactDom.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmClient", "mskContactDom_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskContactPage_GotFocus()

 On Error GoTo Oups

 mskContactPage.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmClient", "mskContactPage_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskContactPage_LostFocus()

 On Error GoTo Oups

 mskContactPage.mask = vbNullString

 If mskContactPage.Text = "(___) ___-____" Then
 mskContactPage.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmClient", "mskContactPage_LostFocus", Err, Err.number, Err.Description
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

  wOups "frmClient", "ExisteDansBD", Err, Err.number, Err.Description
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

 wOups "frmClient", "ContientCaracteresIncorrects", Err, Err.number, Err.Description
End Function
