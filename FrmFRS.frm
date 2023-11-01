VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmFRS 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fournisseurs"
   ClientHeight    =   7440
   ClientLeft      =   2760
   ClientTop       =   1950
   ClientWidth     =   9225
   Icon            =   "FrmFRS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "FrmFRS.frx":0442
   ScaleHeight     =   7440
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm_Categorie 
      Caption         =   "Frame1"
      Height          =   6615
      Left            =   0
      TabIndex        =   79
      Top             =   840
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton cmdcatmod 
         Caption         =   "Modifier"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5160
         TabIndex        =   83
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmb_modAnu 
         Caption         =   "Annuller"
         Height          =   375
         Left            =   5160
         TabIndex        =   89
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdCatAdd 
         Caption         =   "Ajouter"
         Height          =   375
         Left            =   5160
         TabIndex        =   82
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdcatval 
         Caption         =   "Accepter"
         Height          =   375
         Left            =   5160
         TabIndex        =   88
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame FrmCatMod 
         Caption         =   "Frame1"
         Height          =   735
         Left            =   2160
         TabIndex        =   86
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
         Begin VB.TextBox txtmodcat 
            Height          =   285
            Left            =   360
            TabIndex        =   87
            Text            =   "Text1"
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmdAnnuller 
         Caption         =   "Annuller"
         Height          =   375
         Left            =   5160
         TabIndex        =   85
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdcatenr 
         Caption         =   "Enregistrer"
         Height          =   375
         Left            =   5160
         TabIndex        =   84
         Top             =   1320
         Width           =   1455
      End
      Begin MSComctlLib.ListView Lst_Cat 
         Height          =   6015
         Left            =   0
         TabIndex        =   80
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   10610
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Actif"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton cmb_modCat 
      Caption         =   "Modifier"
      Height          =   375
      Left            =   1440
      TabIndex        =   81
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox cmbcatégorie 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   77
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdMailListFournisseur 
      Caption         =   "Ajouter au mailing list"
      Height          =   495
      Left            =   6240
      TabIndex        =   75
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Frame fraEtatOutlook 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   720
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   7575
      Begin VB.Label lblEtatOutlook 
         Alignment       =   2  'Center
         Caption         =   "Liaison du contact avec le fournisseur dans Outlook ..."
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
         Left            =   0
         TabIndex        =   17
         Top             =   840
         Width           =   7575
      End
   End
   Begin VB.CommandButton cmdFax 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Envoyer Fax"
      Height          =   495
      Left            =   5040
      TabIndex        =   73
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox txtFax 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Telephonne"
      DataSource      =   "DatFRS"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   2640
      Width           =   2055
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
      Top             =   2160
      Width           =   4455
      Begin VB.CommandButton cmdMailListContact 
         Caption         =   "Ajouter au mailing list"
         Height          =   495
         Left            =   2520
         TabIndex        =   76
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
         TabIndex        =   42
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtContactPage 
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
      Begin VB.TextBox txtContactEmail 
         BackColor       =   &H00FFFFFF&
         DataField       =   "E-mail"
         DataSource      =   "datContact"
         Height          =   288
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   3480
         Width           =   3255
      End
      Begin VB.TextBox txtContactPoste 
         BackColor       =   &H00FFFFFF&
         DataField       =   "noposte"
         DataSource      =   "datContact"
         Height          =   288
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtContactTel 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Telephonne"
         DataSource      =   "datContact"
         Height          =   288
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox cmbcontact 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "FrmFRS.frx":334F
         Left            =   1080
         List            =   "FrmFRS.frx":3351
         TabIndex        =   39
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtContactFax 
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
      Begin VB.TextBox txtContactCell 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Fax"
         DataSource      =   "datContact"
         Height          =   288
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmdsupcontact 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Supprimer"
         Height          =   495
         Left            =   1320
         TabIndex        =   64
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtContactDom 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Telephonne"
         DataSource      =   "datContact"
         Height          =   288
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton CmdAddCont 
         Caption         =   "Ajouter"
         Height          =   495
         Left            =   120
         TabIndex        =   63
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdEnregistrerContact 
         Caption         =   "Enregistrer"
         Height          =   495
         Left            =   120
         TabIndex        =   62
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdAnnulerContact 
         BackColor       =   &H00C0C0C0&
         Caption         =   "A&nnuler"
         Height          =   495
         Left            =   1320
         TabIndex        =   65
         Top             =   3960
         Width           =   1095
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
         TabIndex        =   56
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
         TabIndex        =   44
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtContact 
         BackColor       =   &H00FFFFFF&
         DataField       =   "NomFournisseur"
         DataSource      =   "DatFRS"
         Height          =   288
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   360
         Visible         =   0   'False
         Width           =   3255
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
         TabIndex        =   41
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         Width           =   1095
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
         TabIndex        =   50
         Top             =   2040
         Width           =   1095
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
         TabIndex        =   60
         Top             =   3480
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
         TabIndex        =   54
         Top             =   2760
         Width           =   975
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
         TabIndex        =   49
         Top             =   2400
         Width           =   975
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
         TabIndex        =   45
         Top             =   1320
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
         TabIndex        =   46
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdRechercher 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rechercher"
      Default         =   -1  'True
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   600
      Width           =   1215
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
      TabIndex        =   67
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox txtSiteWeb 
      BackColor       =   &H00FFFFFF&
      DataField       =   "siteweb"
      DataSource      =   "DatFRS"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   5580
      Width           =   2055
   End
   Begin VB.CommandButton cmdrenommer 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Renommer"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtCommentaire 
      BackColor       =   &H00FFFFFF&
      DataField       =   "commentaire"
      DataSource      =   "DatFRS"
      Height          =   645
      Left            =   6480
      Locked          =   -1  'True
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "FrmFRS.frx":3353
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtRechercher 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton CmdModif 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Modifier"
      Height          =   495
      Left            =   3840
      TabIndex        =   72
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H00FFFFFF&
      DataField       =   "E-mail"
      DataSource      =   "DatFRS"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   3060
      Width           =   2055
   End
   Begin VB.TextBox txtTelephone 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Telephonne"
      DataSource      =   "DatFRS"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2220
      Width           =   2055
   End
   Begin VB.TextBox txtCP 
      BackColor       =   &H00FFFFFF&
      DataField       =   "CodePostal"
      DataSource      =   "DatFRS"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   5160
      Width           =   2052
   End
   Begin VB.TextBox txtPays 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Pays"
      DataSource      =   "DatFRS"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   4740
      Width           =   2052
   End
   Begin VB.TextBox txtProvEtat 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Prov/Etat"
      DataSource      =   "DatFRS"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   4320
      Width           =   2052
   End
   Begin VB.TextBox txtVille 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Ville"
      DataSource      =   "DatFRS"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   3900
      Width           =   2055
   End
   Begin VB.TextBox txtAdresse 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Adresse"
      DataSource      =   "DatFRS"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton CmdFerme 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Fermer"
      Height          =   495
      Left            =   8040
      TabIndex        =   74
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton CmdSupp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Supprimer"
      Height          =   495
      Left            =   2640
      TabIndex        =   71
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton CmdAdd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Ajouter"
      Height          =   495
      Left            =   1440
      TabIndex        =   68
      Top             =   6840
      Width           =   1095
   End
   Begin MSMask.MaskEdBox mskTelephone 
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Top             =   2220
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton CmdEnr 
      Caption         =   "&Enregistrer"
      Height          =   495
      Left            =   1440
      TabIndex        =   69
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton CmdAnul 
      Caption         =   "A&nnuler"
      Height          =   495
      Left            =   2640
      TabIndex        =   70
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox txtNomFournisseur 
      BackColor       =   &H00FFFFFF&
      DataField       =   "NomFournisseur"
      DataSource      =   "DatFRS"
      Height          =   288
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.ComboBox cmbFournisseur 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1320
      Width           =   4335
   End
   Begin MSMask.MaskEdBox mskFax 
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Top             =   2640
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Catégorie"
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
      TabIndex        =   78
      Top             =   1800
      Width           =   1095
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
      TabIndex        =   36
      Top             =   6420
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
      TabIndex        =   33
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
      TabIndex        =   35
      Top             =   6420
      Width           =   975
   End
   Begin VB.Label lblDateCreation 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2004-09-14"
      Height          =   285
      Left            =   1440
      TabIndex        =   32
      Top             =   6000
      Width           =   975
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
      Left            =   120
      TabIndex        =   66
      Top             =   6480
      Width           =   1215
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
      Left            =   120
      TabIndex        =   34
      Top             =   6060
      Width           =   1095
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
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Site web"
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
      TabIndex        =   31
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label16 
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
   Begin VB.Label lblRechercher 
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
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
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
      Left            =   120
      TabIndex        =   18
      Top             =   3060
      Width           =   1095
   End
   Begin VB.Label Label10 
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
      Left            =   120
      TabIndex        =   10
      Top             =   2220
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "CodePostal"
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
      TabIndex        =   28
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label8 
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
      Left            =   120
      TabIndex        =   26
      Top             =   4740
      Width           =   735
   End
   Begin VB.Label Label7 
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
      Left            =   120
      TabIndex        =   24
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label6 
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
      Left            =   120
      TabIndex        =   22
      Top             =   3900
      Width           =   615
   End
   Begin VB.Label Label5 
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
      Left            =   120
      TabIndex        =   20
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "FrmFRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enumMode
  MODE_FRS = 0
  MODE_CONTACT = 1
  MODE_INACTIF = 2
End Enum

Private m_bModeAjoutContact As Boolean
Private m_bModeAjoutFRS     As Boolean
Private m_bRenommer         As Boolean
Private m_bNewContact       As Boolean
Private m_bCategorie        As Boolean
Private m_iNoContact        As Integer
Private m_iNoFournisseur    As Integer
Private m_iNoCategorie      As Integer
Private m_Tag               As String 'V1.44 GLL

Public m_bAnnulerDistList   As Boolean
Public m_otlDistList        As Outlook.DistListItem

Private Sub AfficherCatFour() 'V1.44
'Afficher les fournisseur dans le combobox selon la catégorie choisis
5       On Error GoTo AfficherErreur
10      Dim i As Integer
15      Dim sString As String
20      Dim rstlist As ADODB.Recordset

25      Set rstlist = New ADODB.Recordset
30      sString = "Select * from GRB_Fournisseur "
35      Call rstlist.Open(sString, g_connData, adOpenDynamic, adLockOptimistic)

40      i = 0

45      If Not rstlist.EOF Then
50          Do While Not rstlist.EOF
55              If rstlist.Fields("NomFournisseur") = cmbFournisseur.Text Then Exit Do
60              Call rstlist.MoveNext
65          Loop
    
70          For i = 1 To Lst_Cat.ListItems.count
75              If rstlist.Fields("Categorie") = Null Then Exit For 'si aucune catégorie sélectionner on fait rien

80              If InStr(1, rstlist.Fields("categorie"), Lst_Cat.ListItems(i).Tag, vbTextCompare) > 0 Then
85                  Lst_Cat.ListItems(i).Checked = True
                End If
            Next
        End If
Exit Sub
AfficherErreur:
90      woups "FrmFRS", "AfficherCatFour", Err, Erl
End Sub

Private Sub RemplirComboFournisseur()

5       On Error GoTo AfficherErreur

        'Rempli le combo des fournisseurs
10      Dim rstFournisseur As ADODB.Recordset

15      Set rstFournisseur = New ADODB.Recordset

20      Call rstFournisseur.Open("SELECT NomFournisseur, IDFRS FROM GRB_Fournisseur WHERE Supprimé = False", g_connData, adOpenDynamic, adLockOptimistic)

        'Il faut vider le combo avant de le remplir
25      Call cmbFournisseur.Clear

        'Tant que ce n'est pas la fin des enregistrements
30      Do While Not rstFournisseur.EOF
          'Ajout du nom du fournisseur dans le combo
35        Call cmbFournisseur.AddItem(rstFournisseur.Fields("NomFournisseur"))

          'Ajout du numéro du fournisseur dans le ItemData du combo
40        cmbFournisseur.ItemData(cmbFournisseur.newIndex) = rstFournisseur.Fields("IDFRS")

45        Call rstFournisseur.MoveNext
50      Loop

55      Call rstFournisseur.Close
60      Set rstFournisseur = Nothing

65      If cmbFournisseur.ListCount > 0 Then
70        cmbFournisseur.ListIndex = 0
75      End If

80      Exit Sub

AfficherErreur:

85      woups "frmFRS", "RemplirComboFournisseur", Err, Erl
End Sub

Private Sub AfficherFournisseur()

5       On Error GoTo AfficherErreur
              
        'Affiche le fournisseur sélectionné dans le combo
10      Dim rstFournisseur As ADODB.Recordset
        Dim i As Integer

15      Set rstFournisseur = New ADODB.Recordset
  
20      Call rstFournisseur.Open("SELECT * FROM GRB_Fournisseur WHERE IDFRS = " & m_iNoFournisseur, g_connData, adOpenDynamic, adLockOptimistic)
 
25      Call ViderBarrerChamps(True, True)
 
        'Adresse
30      If Not IsNull(rstFournisseur.Fields("Adresse")) Then
35        txtAdresse.Text = rstFournisseur.Fields("Adresse")
40      End If
  
        'Ville
45      If Not IsNull(rstFournisseur.Fields("Ville")) Then
50        txtVille.Text = rstFournisseur.Fields("Ville")
55      End If
    
        'Prov/Etat
60      If Not IsNull(rstFournisseur.Fields("Prov/Etat")) Then
65        txtProvEtat.Text = rstFournisseur.Fields("Prov/Etat")
70      End If
    
        'Pays
75      If Not IsNull(rstFournisseur.Fields("Pays")) Then
80        txtPays.Text = rstFournisseur.Fields("Pays")
85      End If
    
        'CodePostal
90      If Not IsNull(rstFournisseur.Fields("CodePostal")) Then
95        txtCP.Text = rstFournisseur.Fields("CodePostal")
100     End If
    
        'Telephonne
105     If Not IsNull(rstFournisseur.Fields("Telephonne")) Then
110       txtTelephone.Text = rstFournisseur.Fields("Telephonne")
115     End If

        'Fax
120     If Not IsNull(rstFournisseur.Fields("Fax")) Then
125       txtFax.Text = rstFournisseur.Fields("Fax")
130     End If
    
        'E-mail
135     If Not IsNull(rstFournisseur.Fields("E-mail")) Then
140       txtEmail.Text = rstFournisseur.Fields("E-mail")
145     End If
    
        'Site Web
150     If Not IsNull(rstFournisseur.Fields("SiteWeb")) Then
155       txtSiteWeb.Text = rstFournisseur.Fields("SiteWeb")
160     End If
  
        'commentaire
165     If Not IsNull(rstFournisseur.Fields("Commentaire")) Then
170       txtcommentaire.Text = rstFournisseur.Fields("Commentaire")
175     End If

        'Création
180     If Not IsNull(rstFournisseur.Fields("DateCréation")) Then
185       lblDateCreation.Caption = rstFournisseur.Fields("DateCréation")
190     End If

        'User Création
195     If Not IsNull(rstFournisseur.Fields("UserCréation")) Then
200       lblUserCreation.Caption = "Par : " & rstFournisseur.Fields("UserCréation")
205     End If

        'Modification
210     If Not IsNull(rstFournisseur.Fields("DateModification")) Then
215       lblDateModification.Caption = rstFournisseur.Fields("DateModification")
220     End If

        'User Modification
225     If Not IsNull(rstFournisseur.Fields("UserModification")) Then
230       lblUserModification.Caption = "Par : " & rstFournisseur.Fields("UserModification")
235     End If
        'Catégorie
        If Not IsNull(rstFournisseur.Fields("Categorie")) Then
          For i = 0 To cmbcatégorie.ListCount
           
                If cmbcatégorie.LIST(i) = rstFournisseur.Fields("Categorie") Then
                  cmbcatégorie.ListIndex = (i)
                  Exit For
                End If
            Next
        End If
240     Call rstFournisseur.Close
245     Set rstFournisseur = Nothing

250     Exit Sub

AfficherErreur:

255     woups "frmFRS", "AfficherFournisseur", Err, Erl
End Sub

Public Sub AfficherContact()

5       On Error GoTo AfficherErreur
        
        ''''''''''''''''''''''''''''''''''''''''
        'affiche les contacts de l'employé selectionné'
        ''''''''''''''''''''''''''''''''''''''''
10      Dim rstContact As ADODB.Recordset

        'Ouverture de la table contact
15      Set rstContact = New ADODB.Recordset
        
20      Call rstContact.Open("SELECT * FROM GRB_Contact WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)
    
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
        
        'REMPLIS LES CHAMPS s'il y a enregistrement
70      If Not rstContact.EOF Then
75        If Not IsNull(rstContact.Fields("Titre")) Then
80          txtContactTitre.Text = rstContact.Fields("Titre")
85        End If

90        If Not IsNull(rstContact.Fields("cellulaire")) Then
95          txtContactCell.Text = rstContact.Fields("cellulaire")
100       End If
      
105       If Not IsNull(rstContact.Fields("pagette")) Then
110         txtContactPage.Text = rstContact.Fields("pagette")
115       End If
      
120       If Not IsNull(rstContact.Fields("telephonne")) Then
125         txtContactTel.Text = rstContact.Fields("telephonne")
130       End If
      
135       If Not IsNull(rstContact.Fields("fax")) Then
140         txtContactFax.Text = rstContact.Fields("fax")
145       End If
        
150       If Not IsNull(rstContact.Fields("e-mail")) Then
155         txtContactEmail.Text = rstContact.Fields("e-mail")
160       End If
        
165       If Not IsNull(rstContact.Fields("noposte")) Then
170         txtContactPoste.Text = rstContact.Fields("noposte")
175       End If

180       If Not IsNull(rstContact.Fields("teldomicile")) Then
185         txtContactDom.Text = rstContact.Fields("teldomicile")
190       End If
195     End If
      
        'Ferme la table
200     Call rstContact.Close
205     Set rstContact = Nothing

210     Exit Sub

AfficherErreur:

215     woups "frmFRS", "AfficherContact", Err, Erl
End Sub



Private Sub cmb_modAnu_Click() 'V1.44 GLL
        'Bouton d'annulation des changement apporter a la liste des catégorie
5       On Error GoTo AfficherErreur
10       m_Tag = ""
15      FrmCatMod.Visible = False
20      cmdcatval.Visible = False
25      cmb_modAnu.Visible = False
30      cmdCatAdd.Visible = True
35      cmdcatmod.Visible = True
40      cmdcatenr.Visible = True
45      cmdAnnuller.Visible = True
Exit Sub
AfficherErreur:
50      woups "frmFRS", "cmb_modAnu_Click", Err, Erl
End Sub

Private Sub cmb_modCat_Click() 'V1.44 GLL
        'Bouton pour modifier le nom d'une catégorie
5        On Error GoTo AfficherErreur

10       If m_bCategorie = True Then
15          frm_Categorie.Visible = True
20          frm_Categorie.Caption = "Catégorie pour :" & cmbFournisseur.Text
25          Call AfficherCatList
30          Call AfficherCatFour
35          If Lst_Cat.ListItems.count >= 34 Then cmdCatAdd.Enabled = False
          End If
Exit Sub
AfficherErreur:
40       woups "frmFRS", "cmb_modCat_Click", Err, Erl
End Sub

Private Sub cmbcatégorie_click() 'GLL 2017-11-28 V1.44
        'Active la réduction du nombre de fournisseur par catégorie
5 On Error GoTo AfficherErreur
10        If m_bCategorie = False Then
15           If cmbcatégorie.ListIndex <> -1 Then
20              m_iNoCategorie = cmbcatégorie.ItemData(cmbcatégorie.ListIndex)
25              Call AfficherCategorie
30           End If
35        End If
Exit Sub
AfficherErreur:
40       woups "frmFRS", "cmbCatégorie_Click", Err, Erl
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

35      woups "frmFRS", "cmbContact_Click", Err, Erl
End Sub

Private Sub cmdAnnulerContact_Click()

5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbHourglass

15      Call AfficherControles(MODE_INACTIF)

20      If m_bNewContact = True Then
25        Call HideEdMaskContact(True)

30        m_bNewContact = False
35      End If
        
        'n'est plus en mode ajouter
40      m_bModeAjoutContact = False
  
45      txtNomFournisseur.Visible = False
50      txtNomFournisseur.Locked = False

        'remplis combo contact
55      Call RemplirComboContact
  
60      Screen.MousePointer = vbDefault

65      Exit Sub

AfficherErreur:

70      woups "frmFRS", "cmdanulcontact_Click", Err, Erl
End Sub

Public Sub RemplirComboContact()

5       On Error GoTo AfficherErreur
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'remplis le combo contact dépendant le client sélectionné
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
10      Dim rstContact As ADODB.Recordset
  
15      Set rstContact = New ADODB.Recordset
  
20      If m_bModeAjoutContact = True Then
25        Call rstContact.Open("SELECT NomContact, IDContact FROM GRB_Contact WHERE Supprimé = False ORDER BY NomContact", g_connData, adOpenDynamic, adLockOptimistic)
30      Else
35        Call rstContact.Open("SELECT GRB_Contact.NomContact, GRB_Contact.IDContact FROM GRB_Contact INNER JOIN GRB_ContactFRS ON GRB_Contact.IDContact = GRB_ContactFRS.NoContact WHERE GRB_ContactFRS.NoFRS = " & m_iNoFournisseur & " ORDER BY NomContact", g_connData, adOpenDynamic, adLockOptimistic)
40      End If
    
45      Call cmbContact.Clear
    
50      Do While Not rstContact.EOF
55        Call cmbContact.AddItem(rstContact.Fields("NomContact"))
60        cmbContact.ItemData(cmbContact.newIndex) = rstContact.Fields("IDContact")
        
65        Call rstContact.MoveNext
70      Loop
    
        'Ferme la table "GRB_Contact"
75      Call rstContact.Close
80      Set rstContact = Nothing
        
        'Affiche le contact de la table client
        'si combo est pas vide, affiche le premier contact, sinon le contact inscrit dans table client
85      If cmbContact.ListCount > 0 Then
90        cmbContact.ListIndex = 0
95      Else
100       txtContactTitre.Text = vbNullString
105       txtContactCell.Text = vbNullString
110       txtContactDom.Text = vbNullString
115       txtContactEmail.Text = vbNullString
120       txtContactFax.Text = vbNullString
125       txtContactPage.Text = vbNullString
130       txtContactPoste.Text = vbNullString
135       txtContactTel.Text = vbNullString
140     End If

145     Exit Sub

AfficherErreur:

150     woups "frmFRS", "RemplirComboContact", Err, Erl
End Sub

Private Sub EnregistrerFournisseur()

5       On Error GoTo AfficherErreur

10      Dim rstFournisseur As ADODB.Recordset

15      Set rstFournisseur = New ADODB.Recordset

20      If m_bModeAjoutFRS = True Then
25        Call rstFournisseur.Open("SELECT * FROM GRB_Fournisseur", g_connData, adOpenDynamic, adLockOptimistic)
    
30        Call rstFournisseur.AddNew

35        rstFournisseur.Fields("DateCréation") = ConvertDate(Date)
40        rstFournisseur.Fields("UserCréation") = g_sInitiale
45      Else
50        Call rstFournisseur.Open("SELECT * FROM GRB_Fournisseur WHERE IDFRS = " & m_iNoFournisseur, g_connData, adOpenDynamic, adLockOptimistic)

55        rstFournisseur.Fields("DateModification") = ConvertDate(Date)
60        rstFournisseur.Fields("UserModification") = g_sInitiale
65      End If

        'Enregistrement du fournisseur
70      rstFournisseur.Fields("NomFournisseur").Value = txtNomFournisseur.Text
75      rstFournisseur.Fields("Adresse").Value = txtAdresse.Text
80      rstFournisseur.Fields("Ville").Value = txtVille.Text
85      rstFournisseur.Fields("Prov/Etat").Value = txtProvEtat.Text
90      rstFournisseur.Fields("Pays").Value = txtPays.Text
95      rstFournisseur.Fields("CodePostal").Value = txtCP.Text
100     rstFournisseur.Fields("Telephonne").Value = mskTelephone.Text
105     rstFournisseur.Fields("Fax").Value = mskFax.Text
110     rstFournisseur.Fields("E-mail").Value = txtEmail.Text
115     rstFournisseur.Fields("siteweb").Value = txtSiteWeb.Text
120     rstFournisseur.Fields("Commentaire").Value = txtcommentaire.Text

125     rstFournisseur.Fields("EntryIDOutlook") = ModifierFRSExchange(rstFournisseur.Fields("IDFRS"))

130     If m_bModeAjoutFRS = True Then
135       m_bModeAjoutFRS = False
140     End If

145     Call rstFournisseur.Update
      
150     Call rstFournisseur.Close
155     Set rstFournisseur = Nothing

160     Exit Sub

AfficherErreur:

165     woups "frmFRS", "EnregistrerFournisseur", Err, Erl
End Sub

Private Sub ModifierNomFRSExchange(ByVal sName As String, ByVal iFournisseurID As Integer)

5       On Error GoTo AfficherErreur
  
10      Dim otlApp      As Outlook.Application
15      Dim otlFRS      As Outlook.ContactItem
20      Dim folFRS      As Outlook.MAPIFolder
25      Dim bDejaOuvert As Boolean

30      lblEtatOutlook.Caption = "Modification du nom du fournisseur dans Outlook ..."
35      fraEtatOutlook.Visible = True

40      Set otlApp = OuvrirOutlook(bDejaOuvert)

45      Set folFRS = GetFolder(otlApp, "Fournisseurs GRB")

50      Set otlFRS = folFRS.Items.Find("[User1] = " & iFournisseurID)

55      If otlFRS Is Nothing Then
60        Call MsgBox("Le fournisseur " & txtNomFournisseur.Text & " n'a pas été trouvé pour la mise à jour Exchange!", vbOKOnly, "Erreur")

65        fraEtatOutlook.Visible = False

70        DoEvents

75        Exit Sub
80      End If

85      otlFRS.CompanyName = sName
        
90      Call otlFRS.Save

95      If bDejaOuvert = False Then
100       Call otlApp.Quit
105     End If

110     Set otlApp = Nothing

115     fraEtatOutlook.Visible = False

120     DoEvents

125     Exit Sub

AfficherErreur:

130     woups "frmFRS", "ModifierNomFRSExchange", Err, Erl, "iFournisseurID = " & iFournisseurID)

135     fraEtatOutlook.Visible = False
End Sub

Private Sub LierContactFournisseur(ByVal iFournisseurID As Integer)

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

55      lblEtatOutlook.Caption = "Liaison du contact avec le fournisseur dans Outlook ..."
60      fraEtatOutlook.Visible = True

65      Set otlApp = OuvrirOutlook(bDejaOuvert)

70      Set folFRS = GetFolder(otlApp, "Fournisseurs GRB")
75      Set folContact = GetFolder(otlApp, "Contacts GRB")

80      Set rstFRS = New ADODB.Recordset

85      Call rstFRS.Open("SELECT EntryIDOutlook FROM GRB_Fournisseur WHERE IDFRS = " & m_iNoFournisseur, g_connData, adOpenForwardOnly, adLockReadOnly)

90      Set itmFRS = folFRS.Items.Find("[User1] = " & iFournisseurID)

95      If Not itmFRS Is Nothing Then
100       Do While itmFRS.Links.count > 0
105         Set itmContact = folContact.Items.Find("[User1] = " & itmFRS.Links.Item(1).Item.User1)

110         For iCompteur = 1 To itmContact.Links.count
115           If itmContact.Links.Item(1).Item.User1 = itmFRS.User1 Then
120             Call itmContact.Links.Remove(iCompteur)

125             Call itmContact.Save

130             Exit For
135           End If
140         Next

145         Call itmFRS.Links.Remove(1)
150       Loop

155       Call itmFRS.Save

160       Call rstFRS.Close
165       Set rstFRS = Nothing

170       Set rstContactFRS = New ADODB.Recordset

175       Call rstContactFRS.Open("SELECT * FROM GRB_ContactFRS WHERE NoFRS = " & m_iNoFournisseur, g_connData, adOpenForwardOnly, adLockReadOnly)

180       Do While Not rstContactFRS.EOF
185         Set itmContact = folContact.Items.Find("[User1] = " & rstContactFRS.Fields("NoContact"))

190         If Not itmContact Is Nothing Then
195           Call itmFRS.Links.Add(itmContact)

200           Call itmFRS.Save

205           Call itmContact.Links.Add(itmFRS)

210           Call itmContact.Save
215         End If

220         Call rstContactFRS.MoveNext
225       Loop

230       Call rstContactFRS.Close
235       Set rstContactFRS = Nothing
240     Else
245       Call MsgBox("Fournisseur introuvable!", vbOKOnly, "Erreur")

250       Call rstFRS.Close
255       Set rstFRS = Nothing
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
                      "- Dans Outlook, ouvrez le fournisseur '" & txtNomFournisseur.Text & "' dans Fournisseurs GRB" & vbNewLine & _
                      "- Cliquez sur tous les contacts de ce fournisseur 1 à la fois pour trouver le contact incorrect." & vbNewLine & _
                      "- Effacez ce contact de la liste des contacts de ce fournisseur." & vbNewLine & _
                      "- Dans le logiciel GRB, recommencez l'opération (effacez le contact et l'ajouter de nouveau).", vbOKOnly, "Erreur")
310     Else
315       woups "frmFRS", "LierContactFournisseur", Err, Erl, txtNomFournisseur.Text)
320     End If

325     fraEtatOutlook.Visible = False
End Sub

Private Function ModifierFRSExchange(ByVal iFournisseurID As Integer) As String
  
5       On Error GoTo AfficherErreur
  
10      Dim otlApp      As Outlook.Application
15      Dim otlFRS      As Outlook.ContactItem
20      Dim folFRS      As Outlook.MAPIFolder
25      Dim bDejaOuvert As Boolean

30      If m_bModeAjoutFRS = True Then
35        lblEtatOutlook.Caption = "Ajout du fournisseur dans Outlook ..."
40      Else
45        lblEtatOutlook.Caption = "Modification du fournisseur dans Outlook ..."
50      End If

55      fraEtatOutlook.Visible = True

60      Set otlApp = OuvrirOutlook(bDejaOuvert)

65      Set folFRS = GetFolder(otlApp, "Fournisseurs GRB")

70      If m_bModeAjoutFRS = True Then
75        Set otlFRS = folFRS.Items.Add(olContactItem)

80        otlFRS.User1 = iFournisseurID
85      Else
90        Set otlFRS = folFRS.Items.Find("[User1] = " & iFournisseurID)
95      End If

100     If otlFRS Is Nothing Then
105       Call MsgBox("Le fournisseur " & txtNomFournisseur.Text & " n'a pas été trouvé pour la mise à jour Exchange!", vbOKOnly, "Erreur")

110       fraEtatOutlook.Visible = False

115       DoEvents

120       Exit Function
125     End If

130     otlFRS.CompanyName = txtNomFournisseur.Text
    
135     If mskTelephone.Text <> "(___) ___-____" Then
140       otlFRS.BusinessTelephoneNumber = mskTelephone.Text
145     End If
   
150     If mskFax.Text <> "(___) ___-____" Then
155       otlFRS.BusinessFaxNumber = mskFax.Text
160     End If
   
165     otlFRS.Email1Address = txtEmail.Text
170     otlFRS.BusinessAddressStreet = txtAdresse.Text
175     otlFRS.BusinessAddressCity = txtVille.Text
180     otlFRS.BusinessAddressState = txtProvEtat.Text
185     otlFRS.BusinessAddressCountry = txtPays.Text
190     otlFRS.BusinessAddressPostalCode = txtCP.Text
195     otlFRS.Body = txtcommentaire.Text
200     otlFRS.WebPage = txtSiteWeb.Text
        
205     Call otlFRS.Save

210     ModifierFRSExchange = otlFRS.EntryID

215     If bDejaOuvert = False Then
220       Call otlApp.Quit
225     End If

230     Set otlApp = Nothing

235     fraEtatOutlook.Visible = False

240     DoEvents

245     Exit Function

AfficherErreur:

250     woups "frmFRS", "ModifierFRSExchange", Err, Erl, "iFournisseurID = " & iFournisseurID)

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
105     End Select
        
110     otlContact.Title = ""

115     otlContact.CompanyName = txtNomFournisseur.Text
120     otlContact.JobTitle = txtContactTitre.Text

125     If Trim$(mskContactTel.Text) <> "" Then
130       If mskContactTel.Text <> "(___) ___-____" Then
135         If Trim$(txtContactPoste.Text) <> "" Then
140           otlContact.BusinessTelephoneNumber = mskContactTel.Text & " Ext : " & txtContactPoste.Text
145         Else
150           otlContact.BusinessTelephoneNumber = mskContactTel.Text
155         End If
160       End If
165     End If
    
170     If mskContactFax.Text <> "(___) ___-____" Then
175       otlContact.BusinessFaxNumber = mskContactFax.Text
180     End If
    
185     If mskContactCell.Text <> "(___) ___-____" Then
190       otlContact.MobileTelephoneNumber = mskContactCell.Text
195     End If

200     If mskContactDom.Text <> "(___) ___-____" Then
205       otlContact.HomeTelephoneNumber = mskContactDom.Text
210     End If
    
215     If mskContactPage.Text <> "(___) ___-____" Then
220       otlContact.PagerNumber = mskContactPage.Text
225     End If
    
230     otlContact.Email1Address = txtContactEmail.Text
        
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

280     woups "frmFRS", "AjouterContactExchange", Err, Erl, "iContactID = " & iContactID)

285     fraEtatOutlook.Visible = False
End Function

Private Sub SupprimerFRSExchange(ByVal iFournisseurID As Integer)

5       On Error GoTo AfficherErreur

10      Dim otlApp      As Outlook.Application
15      Dim otlFRS      As Outlook.ContactItem
20      Dim folFRS      As MAPIFolder
25      Dim bDejaOuvert As Boolean

30      lblEtatOutlook.Caption = "Suppression du fournisseur dans Outlook ..."
35      fraEtatOutlook.Visible = True

40      Set otlApp = OuvrirOutlook(bDejaOuvert)

45      Set folFRS = GetFolder(otlApp, "Fournisseurs GRB")

50      Set otlFRS = folFRS.Items.Find("[User1] = " & iFournisseurID)

55      If Not otlFRS Is Nothing Then
60        Call otlFRS.Delete
65      End If

70      If bDejaOuvert = False Then
75        Call otlApp.Quit
80      End If

85      Set otlApp = Nothing

90      fraEtatOutlook.Visible = False

95      DoEvents

100     Exit Sub

AfficherErreur:

105     woups "frmFRS", "SupprimerFRSExchange", Err, Erl

110     fraEtatOutlook.Visible = False
End Sub

Private Sub cmdAnnuller_Click()
frm_Categorie.Visible = False
Call RemplirComboCatégorie
End Sub

Private Sub cmdCatAdd_Click() 'V1.44 GLL
    'Bouton pour ajouter une catégorie a la base de donné
5    On Error GoTo AfficherErreur

10      If Lst_Cat.ListItems.count >= 34 Then 'Méthode utilisé pour géré les catégorie on une limite de 34 alors je bloque les futur addition pour ne pas avoir de problème
15          MsgBox "Vous Avez attent le maximum de catégorie"
20          cmdCatAdd.Enabled = False
            Exit Sub
        End If
        
25      m_Tag = ""

30      FrmCatMod.Visible = True
35      cmdcatval.Visible = True
40      cmb_modAnu.Visible = True
45      cmdCatAdd.Visible = False
50      cmdcatmod.Visible = False
55      cmdcatenr.Visible = False
60      cmdAnnuller.Visible = False
65      cmdcatval.Default = True

70      txtmodcat.SetFocus
75      FrmCatMod.Caption = "Ajouter"
80      txtmodcat.Text = ""

        Exit Sub
        
AfficherErreur:
85      woups "FrmFRS", "cmdCatAdd_Click", Err, Erl

End Sub

Private Sub cmdcatenr_Click() '1.44 GLL Enregistre les catégorie pour le founisseur
On Error GoTo AfficherErreur

5       Dim rstcat As ADODB.Recordset
10      Dim i As Integer
15      Dim sCat As String

20      Set rstcat = New ADODB.Recordset
25      sCat = ""
30      Call rstcat.Open("Select * from Grb_Fournisseur Where NomFournisseur ='" & cmbFournisseur.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

35      If rstcat.EOF Then 'Vérifie si un Fournisseur est sélectionner
40          MsgBox "Erreur aucun fournisseur sélectionner"
            Exit Sub
        End If
        
45      For i = 1 To Lst_Cat.ListItems.count 'Fabric le nouveau code pour la catégorie
50          If Lst_Cat.ListItems(i).Checked Then sCat = sCat & Lst_Cat.ListItems(i).Tag
        Next
        
55      If sCat = "" Then rstcat.Fields("Categorie").Value = Null 'si aucune catégorie est selectionner on rend null la case categorie
60      If sCat <> "" Then rstcat.Fields("Categorie").Value = sCat 'On envoie le code dans la catégorie du fournisseur

65      Call rstcat.Update
70      Call rstcat.Close
75      Set rstcat = Nothing


80      frm_Categorie.Visible = False
85      Call RemplirComboCatégorie
        Exit Sub
AfficherErreur:
90      woups "frmFrs", "cmdcatenr_click", Err, Erl
End Sub


Private Sub cmdcatmod_Click() 'V1.44 GLL
        'bouton pour modifier le nom d'une catégorie
5       On Error GoTo AfficherErreur

10      m_Tag = Lst_Cat.SelectedItem.Tag
15      FrmCatMod.Visible = True
20      cmdcatval.Visible = True
25      cmb_modAnu.Visible = True
30      cmdCatAdd.Visible = False
35      cmdcatmod.Visible = False
40      cmdcatenr.Visible = False
45      cmdAnnuller.Visible = False
50      txtmodcat.SetFocus
55      FrmCatMod.Caption = "Modifier"
60      txtmodcat.Text = Lst_Cat.SelectedItem.Text
        Exit Sub
AfficherErreur:
woups "FrmFrs", "cmdcatmod_Click", Err, Erl
End Sub
Private Sub cmdcatval_Click() 'V1.44 GLL
    'Bouton pour valider l'Additon/modification d'une catégorie
5       On Error GoTo AfficherErreur
10      Dim rstag As ADODB.Recordset
15      Dim bDelete As Boolean
20      Set rstag = New ADODB.Recordset
25      bDelete = False

30      If m_Tag <> "" Then 'pour faire une modification
35          Call rstag.Open("SELECT * FROM TBL_Categorie where Correspondance ='" & m_Tag & "'", g_connData, adOpenDynamic, adLockOptimistic)
40              rstag.Fields("Nom").Value = txtmodcat.Text
45          Else 'pour faire une addition d'une catégorie
50              Call rstag.Open("SELECT * FROM TBL_Categorie", g_connData, adOpenDynamic, adLockOptimistic)

55              Do While Not rstag.EOF
60                  If UCase(rstag.Fields("nom")) = UCase(txtmodcat.Text) Then 'Vérifie si ce nom de catégorie existe déja
65                      MsgBox "vous avez déja cette Categorie"
70                      GoTo Fin
                    End If
75                  Call rstag.MoveNext
                Loop
80              rstag.MoveFirst
                
85              Do While Not rstag.EOF

90                  If IsNull(rstag.Fields("nom")) Then
95                      rstag.Fields("Nom").Value = txtmodcat.Text
100                     Exit Do
105                 End If

110                 Call rstag.MoveNext
                Loop
        End If
        
115     If txtmodcat.Text = "" Then 'Si on a pas miss de text on efface le nom de la catégorie
120         rstag.Fields("Nom").Value = Null
125         bDelete = True
130     End If

135     Call rstag.Update

Fin:
140     Call rstag.Close
145     Set rstag = Nothing
        
150     If bDelete Then Call DeleteCategorie
        
155     Call AfficherCatList
160     Call AfficherCatFour

165     FrmCatMod.Visible = False
170     cmdcatval.Visible = False
175     cmb_modAnu.Visible = False
180     cmdCatAdd.Visible = True
190     cmdcatmod.Visible = True
195     cmdcatenr.Visible = True
200     cmdAnnuller.Visible = True
205     cmdcatmod.Enabled = False
210     m_Tag = ""

Exit Sub
AfficherErreur:
215     woups "Frm_FRS", "Cmdcatval_Click", Err, Erl
End Sub

Private Sub cmdEnregistrerContact_Click()

5       On Error GoTo AfficherErreur
        
10      Dim rstContactFRS As ADODB.Recordset
15      Dim rstContact    As ADODB.Recordset
    
20      Screen.MousePointer = vbHourglass
  
25      Set rstContactFRS = New ADODB.Recordset
  
30      If m_bNewContact = True Then
35        Set rstContact = New ADODB.Recordset

40        Call rstContact.Open("SELECT * FROM GRB_Contact", g_connData, adOpenDynamic, adLockOptimistic)

45        Call rstContact.AddNew

50        rstContact.Fields("NomContact").Value = txtcontact.Text
55        rstContact.Fields("Titre").Value = txtContactTitre.Text
60        rstContact.Fields("Compagnie").Value = txtNomFournisseur.Text
65        rstContact.Fields("Telephonne").Value = mskContactTel.Text
70        rstContact.Fields("Fax").Value = mskContactFax.Text
75        rstContact.Fields("Pagette").Value = mskContactPage.Text
80        rstContact.Fields("Cellulaire").Value = mskContactCell.Text
85        rstContact.Fields("E-mail").Value = txtContactEmail.Text
90        rstContact.Fields("noposte").Value = txtContactPoste.Text
95        rstContact.Fields("teldomicile").Value = mskContactDom.Text
100       rstContact.Fields("UserCréation").Value = g_sInitiale
105       rstContact.Fields("DateCréation").Value = ConvertDate(Date)

110       rstContact.Fields("EntryIDOutlook") = AjouterContactExchange(rstContact.Fields("IDContact"))
 
115       Call rstContact.Update

          'set la table
120       Call rstContactFRS.Open("SELECT * FROM GRB_ContactFRS WHERE NoFRS = " & m_iNoFournisseur & " And NoContact = " & rstContact.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)
    
          'si pas deja existant, on ajoute dans la table
125       If rstContactFRS.EOF Then
            'ajoute dans la table
130         Call rstContactFRS.AddNew
      
135         rstContactFRS.Fields("NoFRS") = m_iNoFournisseur
140         rstContactFRS.Fields("NoContact") = rstContact.Fields("IDContact")
      
145         Call rstContactFRS.Update
150       End If
   
155       Call rstContact.Close
160       Set rstContact = Nothing

165       m_bNewContact = False

170       Call HideEdMaskContact(True)
175     Else
          'set la table
180       Call rstContactFRS.Open("SELECT * FROM GRB_ContactFRS WHERE NoFRS = " & m_iNoFournisseur & " And NoContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)
    
          'Si pas deja existant, on ajoute dans la table
185       If rstContactFRS.EOF Then
            'ajoute dans la table
190         Call rstContactFRS.AddNew
     
195         rstContactFRS.Fields("NoFRS") = m_iNoFournisseur
200         rstContactFRS.Fields("NoContact") = m_iNoContact
      
205         Call rstContactFRS.Update
210       End If

          'Ferme tables et connection
215       Call rstContactFRS.Close
220     End If
    
225     Call LierContactFournisseur(m_iNoFournisseur)
    
230     Set rstContactFRS = Nothing
    
        'Bouton comme avant(apparait)
235     Call AfficherControles(MODE_INACTIF)
    
        'N'est plus en mode ajouter
240     m_bModeAjoutContact = False

        'Remplis combo contact
245     Call RemplirComboContact

250     Call ViderBarrerChampsContact(True, False)

255     Screen.MousePointer = vbDefault

260     Exit Sub

AfficherErreur:

265     woups "frmFRS", "cmdEnregistrerContact_Click", Err, Erl
End Sub

Private Sub cmdFax_Click()

5       On Error GoTo AfficherErreur

10      If cmbFournisseur.ListCount > 0 Then
15        If cmbContact.ListIndex > -1 Then
20          Call frmreport.Afficher(cmbFournisseur.ItemData(cmbFournisseur.ListIndex), cmbContact.ItemData(cmbContact.ListIndex), FRM_FRS)
25        Else
30          Call frmreport.Afficher(cmbFournisseur.ItemData(cmbFournisseur.ListIndex), 0, FRM_FRS)
35        End If
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmFRS", "cmdFax_Click", Err, Erl
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
75      Dim sMsgPlein    As Boolean
80      Dim iNbreRendu   As Integer

85      If cmbContact.ListIndex <> -1 Then
90        bArrayVide = True

95        iResult = MsgBox("Voulez-vous ajouter tous les contacts à la liste de distribution?" & vbNewLine & _
                    "Oui - Tous les contacts" & vbNewLine & _
                    "Non - Contact affiché seulement", vbYesNoCancel)

100       If iResult = vbYes Then
105         Set rstContact = New ADODB.Recordset

110         Call rstContact.Open("SELECT * FROM GRB_ContactFRS INNER JOIN GRB_Contact ON GRB_ContactFRS.NoContact = GRB_Contact.IDContact WHERE GRB_ContactFRS.NoFRS = " & cmbFournisseur.ItemData(cmbFournisseur.ListIndex) & " ORDER BY NomContact", g_connData, adOpenForwardOnly, adLockOptimistic)
                       
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
200           If Trim$(txtContactEmail.Text) <> "" Then
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
550       woups "frmFRS", "cmdMailListContact_Click", Err, Erl
555     End If

560     fraEtatOutlook.Visible = False
End Sub

Private Sub cmdMailListFournisseur_Click()

5       On Error GoTo AfficherErreur

10      Dim otlApp       As Outlook.Application
15      Dim folFRS       As Outlook.MAPIFolder
20      Dim itmFRS       As Outlook.ContactItem
25      Dim otlRecipient As Outlook.Recipient
30      Dim bDejaOuvert  As Boolean

35      If cmbFournisseur.ListIndex <> -1 Then
40        If Trim$(txtEmail.Text) <> "" Then
45          Set otlApp = OuvrirOutlook(bDejaOuvert)

55          lblEtatOutlook.Caption = "Recherche des listes de distribution..."

60          fraEtatOutlook.Visible = True

65          Call frmChoixMailList.Afficher(Me, otlApp)

70          If m_bAnnulerDistList = False Then
75            lblEtatOutlook.Caption = "Ajout du fournisseur dans la liste de distribution..."

80            fraEtatOutlook.Visible = True

85            If m_otlDistList.MemberCount < 10 Then
90              Set folFRS = GetFolder(otlApp, "Fournisseurs GRB")

95              Set itmFRS = folFRS.Items.Find("[User1] = " & cmbFournisseur.ItemData(cmbFournisseur.ListIndex))

100             If Not itmFRS Is Nothing Then
105               Set otlRecipient = otlApp.Session.CreateRecipient(itmFRS.Email1DisplayName)

110               If otlRecipient.Resolve = True Then
115                 Call m_otlDistList.AddMember(otlRecipient)

120                 Call m_otlDistList.Save
125               Else
130                 Call MsgBox("Impossible de trouver le fournisseur!", vbOKOnly, "Erreur")
135               End If
140             Else
145               Call MsgBox("Fournisseur introuvable!", vbOKOnly, "Erreur")
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
205         Call MsgBox("Ce fournisseur n'a pas d'email!", vbOKOnly, "Erreur")
210       End If
215     Else
220       Call MsgBox("Aucun fournisseur sélectionné!", vbOKOnly, "Erreur")
225     End If

230     Exit Sub

AfficherErreur:

235     If Err.number = 287 And Erl = 115 Then
240       Call MsgBox("La liste de distribution est pleine!", vbOKOnly, "Erreur")
245     Else
250       woups "frmFRS", "cmdMailListFournisseur_Click", Err, Erl
255     End If

260     fraEtatOutlook.Visible = False
End Sub

Private Sub cmdRafraichir_Click()

5       On Error GoTo AfficherErreur
              
        'Rafraichir la liste après avoir fait une recherche
10      Screen.MousePointer = vbHourglass
  
        'Remplir le combo
15      Call RemplirComboFournisseur
16      cmbcatégorie.ListIndex = -1
20      Screen.MousePointer = vbDefault

25      Exit Sub

AfficherErreur:

30      woups "frmFRS", "cmdRafraichir_Click", Err, Erl
End Sub

Private Sub ViderBarrerChamps(ByVal bLocked As Boolean, ByVal bVider As Boolean)

5       On Error GoTo AfficherErreur
        'Cette procédure vide et unlock tous les textbox pour pouvoir ajouter
  
10      If bVider = True Then
20        txtAdresse.Text = vbNullString
25        txtVille.Text = vbNullString
30        txtProvEtat.Text = vbNullString
35        txtPays.Text = vbNullString
40        txtCP.Text = vbNullString
45        txtEmail.Text = vbNullString
50        txtSiteWeb.Text = vbNullString
55        txtcommentaire.Text = vbNullString
60        txtTelephone.Text = vbNullString
65        txtFax.Text = vbNullString
70        lblDateCreation.Caption = vbNullString
75        lblUserCreation.Caption = vbNullString
80        lblDateModification.Caption = vbNullString
85        lblUserModification.Caption = vbNullString
90      End If
  
95      txtAdresse.Locked = bLocked
100     txtVille.Locked = bLocked
105     txtProvEtat.Locked = bLocked
110     txtPays.Locked = bLocked
115     txtCP.Locked = bLocked
120     txtEmail.Locked = bLocked
125     txtSiteWeb.Locked = bLocked
130     txtTelephone.Locked = bLocked
135     txtFax.Locked = bLocked
140     txtcommentaire.Locked = bLocked

145     Exit Sub

AfficherErreur:

150     woups "frmFRS", "ViderBarrerChamps", Err, Erl
End Sub

Private Sub ViderBarrerChampsContact(ByVal bLocked As Boolean, ByVal bVider As Boolean)

5       On Error GoTo AfficherErreur
        'Cette procédure vide et unlock tous les textbox pour pouvoir ajouter
  
10      If bVider = True Then
15        txtContactTitre.Text = vbNullString
20        txtContactCell.Text = vbNullString
25        txtContactDom.Text = vbNullString
30        txtContactEmail.Text = vbNullString
35        txtContactFax.Text = vbNullString
40        txtContactPage.Text = vbNullString
45        txtContactPoste.Text = vbNullString
50        txtContactTel.Text = vbNullString
55      End If
  
60      txtContactTitre.Locked = bLocked
65      txtContactCell.Locked = bLocked
70      txtContactDom.Locked = bLocked
75      txtContactEmail.Locked = bLocked
80      txtContactFax.Locked = bLocked
85      txtContactPage.Locked = bLocked
90      txtContactPoste.Locked = bLocked
95      txtContactTel.Locked = bLocked

100     Exit Sub

AfficherErreur:

105     woups "frmFRS", "ViderBarrerChampsContact", Err, Erl
End Sub
Private Sub CmdAddCont_Click()

5       On Error GoTo AfficherErreur
        
        'Pour faire l'ajout d'un contact
10      Dim sNom       As String
15      Dim rstContact As ADODB.Recordset
20      Dim bAjouter   As Boolean

25      If cmbFournisseur.ListCount > 0 Then
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

165           txtNomFournisseur.Visible = True
170           txtNomFournisseur.Locked = True

              'Remplis le combo avec tous les contacts
175           Call AfficherControles(MODE_CONTACT)

180           Call txtContactTitre.SetFocus
185         End If
190       Else
195         Screen.MousePointer = vbHourglass

200         m_bNewContact = False

205         txtNomFournisseur.Visible = True
210         txtNomFournisseur.Locked = True

            'Remplis le combo avec tous les contacts
215         Call AfficherControles(MODE_CONTACT)

220         Call RemplirComboContact
225       End If

230       Screen.MousePointer = vbDefault
235     Else
240       Call MsgBox("Aucun enregistrement de sélectionné")
245     End If

250     Exit Sub

AfficherErreur:

255     woups "frmFRS", "CmdAddCont_Click", Err, Erl
End Sub

Private Sub HideEdMask(ByVal bVisible As Boolean)

5       On Error GoTo AfficherErreur
        
        'proc qui rend visible/ou non les maskEdBox
        'On en profite pour les nettoyer du dernier Enregistrement
        'et on fait l'inverse avec les textBox
10      If m_bModeAjoutFRS = True Then
15        txtTelephone.Text = vbNullString
20        txtFax.Text = vbNullString
       
25        mskTelephone.Text = vbNullString
30        mskFax.Text = vbNullString
35      Else
40        mskTelephone.Text = txtTelephone.Text
45        mskFax.Text = txtFax.Text
50      End If
  
55      mskTelephone.Visible = Not bVisible
60      txtTelephone.Visible = bVisible

65      mskFax.Visible = Not bVisible
70      txtFax.Visible = bVisible

75      Exit Sub

AfficherErreur:

80      woups "frmFRS", "HideEdMask", Err, Erl
End Sub

Private Sub HideEdMaskContact(ByVal bVisible As Boolean)

5       On Error GoTo AfficherErreur
        
        'proc qui rend visible/ou non les maskEdBox
        'On en profite pour les nettoyer du dernier Enregistrement
        'et on fait l'inverse avec les textBox
10      If m_bModeAjoutContact = True Then
15        txtContactTel.Text = vbNullString
20        txtContactFax.Text = vbNullString
25        txtContactPage.Text = vbNullString
30        txtContactCell.Text = vbNullString
35        txtContactDom.Text = vbNullString
       
40        mskContactTel.Text = vbNullString
45        mskContactFax.Text = vbNullString
50        mskContactPage.Text = vbNullString
55        mskContactCell.Text = vbNullString
60        mskContactDom.Text = vbNullString
65      End If
  
70      mskContactTel.Visible = Not bVisible
75      txtContactTel.Visible = bVisible

80      mskContactFax.Visible = Not bVisible
85      txtContactFax.Visible = bVisible

90      mskContactPage.Visible = Not bVisible
95      txtContactPage.Visible = bVisible

100     mskContactCell.Visible = Not bVisible
105     txtContactCell.Visible = bVisible

110     mskContactDom.Visible = Not bVisible
115     txtContactDom.Visible = bVisible

120     Exit Sub

AfficherErreur:

125     woups "frmFRS", "HideEdMaskContact", Err, Erl
End Sub

Private Sub cmdreport_Click()

5       On Error GoTo AfficherErreur
        
        'Impression de la liste des fournisseurs
10      Dim rstFournisseur As ADODB.Recordset

15      Set rstFournisseur = New ADODB.Recordset

20      If MsgBox("Voulez-vous imprimer ce fournisseur uniquement?", vbYesNo) = vbYes Then
25        Call rstFournisseur.Open("SELECT * FROM GRB_Fournisseur WHERE IDFRS = " & m_iNoFournisseur, g_connData, adOpenDynamic, adLockOptimistic)
30      Else
35        If MsgBox("Voulez-vous filtrer par la ville '" & txtVille.Text & "'?", vbYesNo) = vbYes Then
40          Call rstFournisseur.Open("SELECT * FROM GRB_Fournisseur WHERE ville = '" & Replace(txtVille.Text, "'", "''") & "' AND Supprimé = False ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
45        Else
50          Call rstFournisseur.Open("SELECT * FROM GRB_Fournisseur WHERE Supprimé = False ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
55        End If
60      End If
    
65      Screen.MousePointer = vbHourglass
    
        'Set le rapport
70      Set DR_ListeFournisseur.DataSource = rstFournisseur
    
75      DR_ListeFournisseur.Orientation = rptOrientLandscape

80      Call DR_ListeFournisseur.Show(vbModal)
    
85      Call rstFournisseur.Close
90      Set rstFournisseur = Nothing
        
95      Screen.MousePointer = vbDefault

100     Exit Sub

AfficherErreur:

105     woups "frmFRS", "cmdreport_Click", Err, Erl
End Sub

Private Sub AfficherControles(ByVal eMode As enumMode)

5       On Error GoTo AfficherErreur
        
        'Proc qui fait le switch boutton visible/invible
        'on utilise le textBox dummy pour montrer contact
10      Dim bCmbFournisseur  As Boolean
15      Dim bTxtFournisseur  As Boolean
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
115     Dim bMailListFRS     As Boolean
120     Dim bMailListContact As Boolean
  
125     Select Case eMode
          Case MODE_FRS:
130         bTxtFournisseur = True
135         bCmdEnr = True
140         bCmdAnul = True
141         m_bCategorie = True 'GLL 1.44
142         cmb_modCat.Visible = True 'GLL 1.44

          Case MODE_CONTACT:
145         bFraContact = True
150         bTxtFournisseur = True
155         bCmdAnulContact = True
160         bCmdRefCont = True

165         If m_bNewContact = True Then
170           bTxtContact = True
175         Else
180           bCmbContact = True
185         End If

          Case MODE_INACTIF:
190         bFraContact = True
195         bCmbFournisseur = True
200         bCmdImprimer = True
205         bTxtRechercher = True
210         bCmdRenommer = True
215         bCmdRafraichir = True
220         bCmdAdd = True
225         bCmdSupp = True
230         bCmdModif = True
235         bCmdQuit = True
240         bCmdAddCont = True
245         bCmdSupContact = True
250         bFax = True
255         bCmbContact = True
260         bMailListContact = True
265         bMailListFRS = True
270         m_bCategorie = False 'GLL V1.44
271         cmb_modCat.Visible = False 'GLL V1.44
275         If Len(txtRechercher.Text) > 0 Then
280           bCmdRechercher = True
285         End If
290     End Select
     
295     cmbFournisseur.Visible = bCmbFournisseur
300     txtNomFournisseur.Visible = bTxtFournisseur
305     fracontact.Visible = bFraContact
310     txtRechercher.Enabled = bTxtRechercher
315     cmdRechercher.Enabled = bCmdRechercher
320     cmdRafraichir.Enabled = bCmdRafraichir
325     cmdrenommer.Enabled = bCmdRenommer
330     cmdReport.Visible = bCmdImprimer
335     CmdAdd.Visible = bCmdAdd
340     CmdSupp.Visible = bCmdSupp
345     CmdModif.Visible = bCmdModif
350     CmdFerme.Visible = bCmdQuit
355     CmdAnul.Visible = bCmdAnul
360     CmdEnr.Visible = bCmdEnr
365     CmdAddCont.Visible = bCmdAddCont
370     cmdsupcontact.Visible = bCmdSupContact
375     cmdAnnulerContact.Visible = bCmdAnulContact
380     cmdEnregistrerContact.Visible = bCmdRefCont
385     cmdFax.Visible = bFax
390     cmbContact.Visible = bCmbContact
395     txtcontact.Visible = bTxtContact
400     cmdMailListFournisseur.Visible = bMailListFRS
405     cmdMailListContact.Visible = bMailListContact

410     Exit Sub

AfficherErreur:

415     woups "frmFRS", "AfficherControles", Err, Erl
End Sub

Private Sub CmdAdd_Click()

5       On Error GoTo AfficherErreur
              'proc qui permet d'ajouter un contact à la BD
10      Dim sName As String
 
15      Call AfficherControles(MODE_FRS)

20      sName = InputBox("Veuillez entrer le nom du fournisseur" & vbNewLine & _
                         vbNewLine & _
                         "SVP, respectez le bon orthographe!", "SAISIE DU NOM", "Nom du fournisseur")
    
25      If sName <> vbNullString Then
30        If Not ComboContient(cmbFournisseur, sName) Then
35          Screen.MousePointer = vbHourglass
      
40          m_bModeAjoutFRS = True
        
            'On montre les maskEdBox
45          Call HideEdMask(False)
                
            'On affiche le nom du nouveau client dans le textbox
            'pour éviter le ScrollDown durant l'ajout
50          txtNomFournisseur.Text = sName

55          Call ViderBarrerChamps(False, True)

60          Call mskTelephone.SetFocus
              
65          Screen.MousePointer = vbDefault
70        Else
75          Call MsgBox("Ce fournisseur existe déjà!")
     
80          Call AfficherControles(MODE_INACTIF)
85        End If
90      Else
95        Call AfficherControles(MODE_INACTIF)
100     End If

105     Exit Sub

AfficherErreur:

110     woups "frmFRS", "CmdAdd_Click", Err, Erl
End Sub

Private Sub CmdAnul_Click()

5       On Error GoTo AfficherErreur
              
        'Annule l'ajout ou la modif
10      Screen.MousePointer = vbHourglass

        'On cache le maskEdBox
15      Call HideEdMask(True)
                
        'commentaire unlock
        'txtNomClient.Visible = False
  
        'on retablis les bouttons
20      Call AfficherControles(MODE_INACTIF)

25      m_bModeAjoutFRS = False
    
30      Call ViderBarrerChamps(True, True)
  
35      Call cmbFournisseur_Click
  
40      Screen.MousePointer = vbDefault

45      Exit Sub

AfficherErreur:

50      woups "frmFRS", "CmdAnul_Click", Err, Erl
End Sub

Private Sub CmdEnr_Click()

5       On Error GoTo AfficherErreur

        'Enregistrement d'une modif ou d'un ajout
10      Dim sFournisseur   As String
15      Dim iCompteur      As Integer
     
        'Nom du fournisseur
20      sFournisseur = txtNomFournisseur.Text
     
        'Enregistrement d'un frs dans la BD
25      Screen.MousePointer = vbHourglass
        
30      Call EnregistrerFournisseur
    
        'On cache les MaskEdBox
35      Call HideEdMask(True)
 
        'On met a jour le combo
40      Call RemplirComboFournisseur
        
        'Retablir les boutons
45      Call AfficherControles(MODE_INACTIF)
  
50      For iCompteur = 0 To cmbFournisseur.ListCount - 1
55        If cmbFournisseur.LIST(iCompteur) = sFournisseur Then
60          cmbFournisseur.ListIndex = iCompteur
      
65          Exit For
70        End If
75      Next
  
80      Call cmbFournisseur.SetFocus
  
85      Screen.MousePointer = vbDefault

90      Exit Sub

AfficherErreur:

95      woups "frmFRS", "CmdEnr_Click", Err, Erl
End Sub

Private Sub CmdFerme_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmFRS", "CmdFerme_Click", Err, Erl
End Sub

Private Sub CmdModif_Click()

5       On Error GoTo AfficherErreur

10      If cmbFournisseur.ListCount > 0 Then
15        Screen.MousePointer = vbHourglass
    
                'proc qui permet de modifier l'enr courant
20        Call HideEdMask(False)
        
25        Call AfficherControles(MODE_FRS)
        
30        Call ViderBarrerChamps(False, False)
        
35        Screen.MousePointer = vbDefault
40      Else
45        Call MsgBox("Aucun enregistrement sélectionné!")
50      End If

55      Exit Sub

AfficherErreur:

60      woups "frmFRS", "CmdModif_Click", Err, Erl
End Sub

Private Sub cmdRenommer_Click()

5       On Error GoTo AfficherErreur
        
        ''''''''''''''''''''''''''''''''''''''
        'on renomme le nom du FOURNISSEUR
        ''''''''''''''''''''''''''''''''''''''''
10      Dim rstFournisseur As ADODB.Recordset
15      Dim sName          As String

20      If cmbFournisseur.ListCount > 0 Then
          'Proc qui permet de modifié un CLIENT a la BD
          'On procede a la saisie du nom du CLIENT
25        sName = InputBox("Veuillez entrer le nom du Fournisseur", "Renommer fournisseur", txtNomFournisseur.Text)
        
30        If sName <> vbNullString Then
35          If sName <> txtNomFournisseur.Text Then
40            If Not ComboContient(cmbFournisseur, sName) Then
45              Screen.MousePointer = vbHourglass
              
50              Set rstFournisseur = New ADODB.Recordset
              
55              Call rstFournisseur.Open("SELECT IDFrs, NomFournisseur, EntryIDOutlook FROM GRB_Fournisseur WHERE IDFRS = " & m_iNoFournisseur, g_connData, adOpenDynamic, adLockOptimistic)
              
60              Call ModifierNomFRSExchange(sName, m_iNoFournisseur)
              
65              txtNomFournisseur = sName
              
                'transfert des donnes
70              rstFournisseur.Fields("NomFournisseur").Value = txtNomFournisseur.Text
                  
                'mise a jour de la base de donne
75              Call rstFournisseur.Update
            
80              Call rstFournisseur.Close
85              Set rstFournisseur = Nothing
            
90              Call RemplirComboFournisseur
            
95              cmbFournisseur.Text = sName
            
100             m_bRenommer = True
            
105              Call cmbFournisseur_Click
            
110             m_bRenommer = False
            
115             Screen.MousePointer = vbDefault
120           Else
125             Call MsgBox("Le fournisseur " & sName & " existe déjà!", vbCritical)
130           End If
135         End If
140       End If
145     Else
150       Call MsgBox("Aucun enregistrement de sélectionné!", vbOKOnly, "Erreur")
155     End If

160     Exit Sub

AfficherErreur:

165     woups "frmFRS", "cmdRenommer_Click", Err, Erl
End Sub

Private Sub cmdsupcontact_Click()

5       On Error GoTo AfficherErreur
        
        'Fonction qui supprime l'enregistrement courant
10      If cmbContact.ListCount > 0 Then
15        If MsgBox("Etes-vous sur de vouloir supprimer cette enregistrement?", vbYesNo, "Supprimer") = vbYes Then
20          Screen.MousePointer = vbHourglass
      
25          Call g_connData.Execute("DELETE * FROM GRB_ContactFRS WHERE NoFRS = " & m_iNoFournisseur & " AND NoContact = " & m_iNoContact)

30          Call LierContactFournisseur(m_iNoFournisseur)

            'Femplis le combo employé
35          Call RemplirComboContact

40          Screen.MousePointer = vbDefault
45        End If
50      End If

55      Exit Sub

AfficherErreur:

60      woups "frmFRS", "cmdsupcontact_Click", Err, Erl
End Sub

Private Sub CmdSupp_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum  As ADODB.Recordset
15      Dim rstCatalogue As ADODB.Recordset
20      Dim rstFRS       As ADODB.Recordset
25      Dim bPeutEffacer As Boolean
        
        'Fonction qui supprime lenregistrement courant
30      If cmbFournisseur.ListCount > 0 Then
35        If MsgBox("Etes-vous sur de supprimer cette enregistrement?", vbYesNo, "Supprimer") = vbYes Then
40          Screen.MousePointer = vbHourglass
               
            'Open table
45          Set rstProjSoum = New ADODB.Recordset
50          Set rstCatalogue = New ADODB.Recordset
            
55          Call rstProjSoum.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDFRS = " & m_iNoFournisseur, g_connData, adOpenDynamic, adLockOptimistic)
60          Call rstCatalogue.Open("SELECT * FROM GRB_PiecesFRS WHERE IDFRS = " & m_iNoFournisseur, g_connData, adOpenDynamic, adLockOptimistic)

            'si existe pas dans soumission
65          If rstProjSoum.EOF Then
70            Call rstProjSoum.Close
              'si existe pas dans projet
75            Call rstProjSoum.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDFRS = " & m_iNoFournisseur, g_connData, adOpenDynamic, adLockOptimistic)
          
80            If rstProjSoum.EOF Then
                'si existe pas dans la table fournisseurs pour une piece
85              If rstCatalogue.EOF Then
90                bPeutEffacer = True
95              Else
100               bPeutEffacer = False
105             End If
110           Else
115             bPeutEffacer = False
120           End If
125         Else
130           bPeutEffacer = False
135         End If
        
140         Call rstCatalogue.Close
145         Set rstCatalogue = Nothing
                  
150         Call rstProjSoum.Close
155         Set rstProjSoum = Nothing
             
160         If cmbContact.ListCount > 0 Then
              'Delete les contact» pour ce fournisseur
165           Call g_connData.Execute("DELETE * FROM GRB_ContactFRS WHERE NoFRS = " & m_iNoFournisseur)
170         End If

175         Call SupprimerFRSExchange(m_iNoFournisseur)

180         If bPeutEffacer = True Then
185           Call g_connData.Execute("DELETE * FROM GRB_Fournisseur WHERE IDFRS = " & m_iNoFournisseur)
190         Else
195           Set rstFRS = New ADODB.Recordset

200           Call rstFRS.Open("SELECT * FROM GRB_Fournisseur WHERE IDFRS = " & m_iNoFournisseur, g_connData, adOpenDynamic, adLockOptimistic)

205           rstFRS.Fields("Supprimé") = True

210           Call rstFRS.Update

215           Call rstFRS.Close
220           Set rstFRS = Nothing
225         End If

            'Remplir le combo des fournisseurs
230         Call RemplirComboFournisseur

235         Screen.MousePointer = vbDefault
240       End If
245     Else
250       Call MsgBox("Aucun enregistrement sélectionné!")
255     End If

260     Exit Sub

AfficherErreur:

265     woups "frmFRS", "CmdSupp_Click", Err, Erl
End Sub

Private Sub cmbFournisseur_Click()

5        On Error GoTo AfficherErreur
        
        'Quand le user selectionne un enregistrement on se posotionne dessus
10      If cmbFournisseur.Text <> vbNullString Then
15        txtNomFournisseur.Text = cmbFournisseur.Text
20      Else
25        If ComboContient(cmbFournisseur, txtNomFournisseur.Text) = False Then
30          Call RemplirComboFournisseur
35        End If

40        cmbFournisseur.Text = txtNomFournisseur.Text
45      End If
  
50      If cmbFournisseur.ListIndex > -1 Then
55        If m_bRenommer = False And m_bModeAjoutFRS = False Then
60          m_iNoFournisseur = cmbFournisseur.ItemData(cmbFournisseur.ListIndex)
65        End If
70      End If
  
        'Affiche le fournisseur sélectionné dans le combo
75      Call AfficherFournisseur
80      Call RemplirComboContact

85      Exit Sub

AfficherErreur:

90      woups "frmFRS", "cmbFournisseur_Click", Err, Erl
End Sub






Private Sub Form_Load()

5       On Error GoTo AfficherErreur
        
10

15       Call tbl_cat_exist 'GLL 2017-11-28 V1.44
20      Call FindFieldsExist 'GLL 2017-11-28 V1.44

25      Call RemplirComboFournisseur
30      Call RemplirComboCatégorie 'GLL 2017-11-28 V1.44
    
35      Call HideEdMask(True)
  
40      Call AfficherControles(MODE_INACTIF)
  
45      Call ActiverBoutonsGroupe

50      Screen.MousePointer = vbDefault


55      Exit Sub

AfficherErreur:

60      woups "frmFRS", "Form_Load", Err, Erl
End Sub

Private Sub ActiverBoutonsGroupe()

5       On Error GoTo AfficherErreur
        
        'Activation des boutons selon le groupe
10      CmdAdd.Enabled = g_bModificationFournisseurs
15      CmdModif.Enabled = g_bModificationFournisseurs
20      cmdrenommer.Enabled = g_bModificationFournisseurs
25      CmdSupp.Enabled = g_bModificationFournisseurs
30      CmdAddCont.Enabled = g_bModificationFournisseurs
35      cmdsupcontact.Enabled = g_bModificationFournisseurs
40      cmdMailListContact.Enabled = g_bModificationListeDistribution
45      cmdMailListFournisseur.Enabled = g_bModificationListeDistribution
  
50      Exit Sub

AfficherErreur:

55      woups "frmFRS", "ActiverBoutonsGroupe", Err, Erl
End Sub

Private Sub Form_Unload(Cancel As Integer)

5       On Error GoTo AfficherErreur

10      Set FrmFRS = Nothing

15      Exit Sub

AfficherErreur:

20      woups "frmFRS", "Form_Unload", Err, Erl
End Sub

Private Sub Lst_Cat_ItemClick(ByVal Item As MSComctlLib.ListItem)
cmdcatmod.Enabled = True
End Sub



Private Sub mskTelephone_GotFocus()

5       On Error GoTo AfficherErreur

10      mskTelephone.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmFRS", "mskTelephone_GotFocus", Err, Erl
End Sub

Private Sub mskTelephone_LostFocus()

5       On Error GoTo AfficherErreur

10      mskTelephone.mask = vbNullString

15      If mskTelephone.Text = "(___) ___-____" Then
20        mskTelephone.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmFRS", "mskTelephone_LostFocus", Err, Erl
End Sub

Private Sub mskFax_GotFocus()

5       On Error GoTo AfficherErreur

10      mskFax.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmFRS", "mskFax_GotFocus", Err, Erl
End Sub

Private Sub mskFax_LostFocus()

5       On Error GoTo AfficherErreur

10      mskFax.mask = vbNullString

15      If mskFax.Text = "(___) ___-____" Then
20        mskFax.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmFRS", "mskFax_LostFocus", Err, Erl
End Sub

Private Sub cmdRechercher_Click()

5       On Error GoTo AfficherErreur
        
        'Rempli le combo des fournisseurs selon le texte à rechercher
10      Dim rstFournisseur As ADODB.Recordset
15      Dim sSearch    As String
  
20      Screen.MousePointer = vbHourglass
  
25      sSearch = txtRechercher.Text
  
        'vide les champs
30      Call ViderBarrerChamps(True, True)
    
        'Filtre pour selection des Nomcontact
35      Set rstFournisseur = New ADODB.Recordset
        
40      Call rstFournisseur.Open("SELECT NomFournisseur, IDFRS FROM GRB_Fournisseur WHERE Instr(1, NomFournisseur, '" & Replace(sSearch, "'", "''") & "') > 0 AND Supprimé = False ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
                
        'vide combo
45      Call cmbFournisseur.Clear
                
        'Tant que ce n'est pas la fin des enregistrements
50      Do While Not rstFournisseur.EOF
55        Call cmbFournisseur.AddItem(rstFournisseur.Fields("NomFournisseur"))
60        cmbFournisseur.ItemData(cmbFournisseur.newIndex) = rstFournisseur.Fields("IDFRS")
                
65        Call rstFournisseur.MoveNext
70      Loop
      
75      Call rstFournisseur.Close
80      Set rstFournisseur = Nothing
  
85      If cmbFournisseur.ListCount > 0 Then
90        cmbFournisseur.ListIndex = 0
95      Else
100       Call cmbContact.Clear

105       txtContactCell.Text = vbNullString
110       txtContactDom.Text = vbNullString
115       txtContactEmail.Text = vbNullString
120       txtContactFax.Text = vbNullString
125       txtContactPage.Text = vbNullString
130       txtContactPoste.Text = vbNullString
135       txtContactTel.Text = vbNullString
140     End If
    
145     Screen.MousePointer = vbDefault

150     Exit Sub

AfficherErreur:

155     woups "frmFRS", "cmdRechercher_Click", Err, Erl
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

40      woups "frmFRS", "txtRechercher_Change", Err, Erl
End Sub

Private Sub mskContactTel_GotFocus()

5       On Error GoTo AfficherErreur

10      mskContactTel.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmFRS", "mskContactTel_GotFocus", Err, Erl
End Sub

Private Sub mskContactTel_LostFocus()

5       On Error GoTo AfficherErreur

10      mskContactTel.mask = vbNullString

15      If mskContactTel.Text = "(___) ___-____" Then
20        mskContactTel.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmFRS", "mskContactTel_LostFocus", Err, Erl
End Sub

Private Sub mskContactFax_GotFocus()

5       On Error GoTo AfficherErreur

10      mskContactFax.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmFRS", "mskContactFax_GotFocus", Err, Erl
End Sub

Private Sub mskContactFax_LostFocus()

5       On Error GoTo AfficherErreur

10      mskContactFax.mask = vbNullString

15      If mskContactFax.Text = "(___) ___-____" Then
20        mskContactFax.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmFRS", "mskContactFax_LostFocus", Err, Erl
End Sub

Private Sub mskContactCell_GotFocus()

5       On Error GoTo AfficherErreur

10      mskContactCell.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmFRS", "mskContactCell_GotFocus", Err, Erl
End Sub

Private Sub mskContactCell_LostFocus()

5       On Error GoTo AfficherErreur

10      mskContactCell.mask = vbNullString

15      If mskContactCell.Text = "(___) ___-____" Then
20        mskContactCell.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmFRS", "mskContactCell_LostFocus", Err, Erl
End Sub

Private Sub mskContactDom_GotFocus()

5       On Error GoTo AfficherErreur

10      mskContactDom.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmFRS", "mskContactDom_GotFocus", Err, Erl
End Sub

Private Sub mskContactDom_LostFocus()

5       On Error GoTo AfficherErreur

10      mskContactDom.mask = vbNullString

15      If mskContactDom.Text = "(___) ___-____" Then
20        mskContactDom.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmFRS", "mskContactDom_LostFocus", Err, Erl
End Sub

Private Sub mskContactPage_GotFocus()

5       On Error GoTo AfficherErreur

10      mskContactPage.mask = "(###) ###-####"

15      Exit Sub

AfficherErreur:

20      woups "frmFRS", "mskContactPage_GotFocus", Err, Erl
End Sub

Private Sub mskContactPage_LostFocus()

5       On Error GoTo AfficherErreur

10      mskContactPage.mask = vbNullString

15      If mskContactPage.Text = "(___) ___-____" Then
20        mskContactPage.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmFRS", "mskContactPage_LostFocus", Err, Erl
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

65      woups "frmFRS", "ExisteDansBD", Err, Erl
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

40      woups "frmFRS", "ContientCaracteresIncorrects", Err, Erl
End Function
Private Sub AfficherCategorie() 'GLL 2017-11-28 V1.44

5       On Error GoTo AfficherErreur
        
        ''''''''''''''''''''''''''''''''''''''''
        'affiche les contacts selon leur catégorie'
        ''''''''''''''''''''''''''''''''''''''''
10      Dim rstCategorie As ADODB.Recordset
15      Dim rstFournisseur As ADODB.Recordset
20      Dim i As Integer
25      Dim j As Integer
30      Dim id As Integer
35      Dim sString As String
40      Dim cString As String

45      'Ouverture de la table contact
50      Set rstCategorie = New ADODB.Recordset
55      Set rstFournisseur = New ADODB.Recordset
60      sString = "Select * From Tbl_Categorie where nom <> Null"
65      cmbFournisseur.Clear
70      Call rstCategorie.Open(sString, g_connData, adOpenDynamic, adLockOptimistic)

75          Do While Not rstCategorie.EOF
80              If rstCategorie.Fields("Nom") = cmbcatégorie.Text Then sString = rstCategorie.Fields("Correspondance")
85              Call rstCategorie.MoveNext
90          Loop
95          Call rstCategorie.Close
100         Set rstCategorie = Nothing
            
105     Call rstFournisseur.Open("Select * From GRB_Fournisseur where categorie <> Null", g_connData, adOpenDynamic, adLockOptimistic)
        
110         Do While Not rstFournisseur.EOF
115             If InStr(1, rstFournisseur.Fields("Categorie"), sString, vbTextCompare) > 0 Then
120                 cmbFournisseur.AddItem (rstFournisseur.Fields("NomFournisseur"))
125                 cmbFournisseur.ItemData(cmbFournisseur.newIndex) = rstFournisseur.Fields("IDFRS")
                End If
130         Call rstFournisseur.MoveNext
            Loop
135         Call rstFournisseur.Close
140         Set rstFournisseur = Nothing

145        If cmbFournisseur.ListCount > 0 Then 'Afficher le premier Fournisseur qui est dans cette catégorie
150             cmbFournisseur.ListIndex = 0
155             Call cmbFournisseur_Click
160        End If

165     Exit Sub

AfficherErreur:
170     woups "FrmFrs", "AfficherCategorie", Err, Erl
End Sub
Private Sub RemplirComboCatégorie()
5       On Error GoTo AfficherErreur 'GLL 2017-11-28 V1.44
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'remplis le combo contact dépendant le client sélectionné
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
10      Dim rstCategorie As ADODB.Recordset
  
15      Set rstCategorie = New ADODB.Recordset
  

35        Call rstCategorie.Open("SELECT  Nom FROM TBL_Categorie where nom <> Null order by Nom", g_connData, adOpenDynamic, adLockOptimistic)

    
45      Call cmbcatégorie.Clear
    
50      Do While Not rstCategorie.EOF
55        Call cmbcatégorie.AddItem(rstCategorie.Fields("nom"))
65        Call rstCategorie.MoveNext
70      Loop
    
        'Ferme la table "GRB_Contact"
75      Call rstCategorie.Close
80      Set rstCategorie = Nothing
        
85     Exit Sub

AfficherErreur:

90     woups "frmFRS", "RemplirComboCatégorie", Err, Erl
End Sub
Private Sub AfficherCatList() 'V1.44 GLL
On Error GoTo AfficherErreur
        'Affiche dans Rstlist tout les catégorie enregistrer
5       Dim rstlist As ADODB.Recordset
10      Dim itemlist As ListItem

15      Set rstlist = New ADODB.Recordset
20      Call rstlist.Open("Select * from tbl_categorie where nom <> Null order by nom", g_connData, adOpenDynamic, adLockOptimistic)
25      Call Lst_Cat.ListItems.Clear

    
30      Do While Not rstlist.EOF 'Ajoute dans la liste tout le catégorie trouver
35          Set itemlist = Lst_Cat.ListItems.Add
40          itemlist.Text = rstlist.Fields("Nom")
45          itemlist.Tag = rstlist.Fields("Correspondance")
50          Call rstlist.MoveNext
55      Loop
Exit Sub
AfficherErreur:
60      woups "FrmFRS", "AfficherCatList", Err, Erl
End Sub
Private Sub DeleteCategorie() 'V1.44 GLL
        'efface une catégorie de tout les fournisseur qui la possêde
5       On Error GoTo AfficherErreur

10      Dim rstCategorie As ADODB.Recordset
15      Dim sString As String

20      Set rstCategorie = New ADODB.Recordset
25      Call rstCategorie.Open("Select categorie from GRB_Fournisseur where categorie <> Null or categorie ='""'", g_connData, adOpenStatic, adLockPessimistic)

30      Do While Not rstCategorie.EOF

35          If InStr(1, rstCategorie.Fields("categorie"), m_Tag, vbTextCompare) > 0 Then
40              sString = rstCategorie.Fields("categorie")
45              sString = Replace(sString, m_Tag, "", 1)
                
50              If sString = "" Then
55                  rstCategorie.Fields("categorie").Value = Null
                Else
60                  rstCategorie.Fields("categorie").Value = sString
                End If
            End If
65      Call rstCategorie.MoveNext
        Loop

70      Call rstCategorie.Close
75      Set rstCategorie = Nothing
    Exit Sub
AfficherErreur:
80      woups "FrmFRS", "DeleteCategorie", Err, Erl
End Sub

Private Sub tbl_cat_exist() 'V1.44
'Vérifie si la tbl catégorie exist dans la basse de donné si non elle la crée

5       On Error GoTo AfficherErreur
10      Dim adoxconnection As adox.Catalog
15      Dim i As Integer

20      Set adoxconnection = New adox.Catalog
25      adoxconnection.ActiveConnection = g_connData
30      For i = 0 To adoxconnection.Tables.count - 1

35          If LCase(adoxconnection.Tables(i).Name) = LCase("TBL_Categorie") Then 'Si elle exist on sort de la sous routine
40              Set adoxconnection = Nothing
                Exit Sub
45          End If
        Next

50      Call g_connData.Execute("Create TABLE TBL_Categorie (Correspondance text(1), Nom Text (100))")
55      Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('A');")
60      Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('B');")
65      Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('C');")
70      Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('D');")
75      Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('E');")
80      Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('F');")
85      Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('G');")
90      Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('H');")
95      Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('I');")
100     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('J');")
105     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('K');")
110     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('M');")
115     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('N');")
120     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('O');")
125     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('P');")
130     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('Q');")
135     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('R');")
140     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('S');")
145     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('T');")
150     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('U');")
155     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('V');")
160     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('W');")
165     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('X');")
170     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('Y');")
175     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('Z');")
180     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('1');")
185     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('2');")
190     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('3');")
195     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('4');")
200     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('5');")
205     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('6');")
210     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('7');")
215     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('8');")
220     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('9');")
225     Call g_connData.Execute("Insert into TBL_Categorie  (Correspondance) Values ('0');")
230     Set adoxconnection = Nothing
    Exit Sub
AfficherErreur:
235     woups "FrmFrs", "tbl_Cat_exist", Err, Erl
End Sub

Private Sub FindFieldsExist() 'V1.44
5    On Error GoTo AfficherErreur

10    Dim strName As String

15    Dim Findfield As ADODB.Recordset

20    Dim i As Integer
    
25    Set Findfield = New ADODB.Recordset

30    Call Findfield.Open("Select * from Grb_Fournisseur", g_connData, adOpenDynamic, adLockOptimistic)

35    For i = 0 To Findfield.Fields.count - 1

40    strName = Findfield.Fields(i).Name

45    If strName = "Categorie" Then

50        Call Findfield.Close

55        Set Findfield = Nothing

60        Exit Sub

65    End If

70    Next

75    Call Findfield.Close

80    Set Findfield = Nothing

85    Call g_connData.Execute("ALTER TABLE GRB_Fournisseur Add Categorie Text(40);")
    Exit Sub
    
AfficherErreur:
90    woups "frmFRS", "FindFieldsExist()", Err, Erl
    End Sub
