VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmFRS 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fournisseurs"
   ClientHeight    =   7440
   ClientLeft      =   2760
   ClientTop       =   1950
   ClientWidth     =   9225
   Icon            =   "FrmFRS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   9225
   Begin VB.Frame frm_Categorie 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   0
      TabIndex        =   79
      Top             =   720
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
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
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
      Left            =   0
      TabIndex        =   16
      Top             =   2640
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
         Left            =   360
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
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
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
         ItemData        =   "FrmFRS.frx":0442
         Left            =   1080
         List            =   "FrmFRS.frx":0444
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
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   41
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   50
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   60
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   59
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   54
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   49
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   45
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
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
      Appearance      =   0  'Flat
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
      Text            =   "FrmFRS.frx":0446
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
      Appearance      =   0  'Flat
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
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Supprimer"
      Height          =   495
      Left            =   2640
      TabIndex        =   71
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton CmdAdd 
      Appearance      =   0  'Flat
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
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   78
      Top             =   1800
      Width           =   1095
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
      TabIndex        =   36
      Top             =   6420
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
      TabIndex        =   33
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label lblDateModification 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2004-09-14"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1440
      TabIndex        =   35
      Top             =   6420
      Width           =   975
   End
   Begin VB.Label lblDateCreation 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2004-09-14"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1440
      TabIndex        =   32
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label12 
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
      TabIndex        =   66
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label12 
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
      Left            =   120
      TabIndex        =   34
      Top             =   6060
      Width           =   1095
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
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
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   3060
      Width           =   1095
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2220
      Width           =   1095
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   4740
      Width           =   735
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3900
      Width           =   615
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
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
Private m_bModeAjoutFRS As Boolean
Private m_bRenommer As Boolean
Private m_bNewContact As Boolean
Private m_bCategorie As Boolean
Private m_iNoContact As Integer
Private m_iNoFournisseur As Integer
Private m_iNoCategorie As Integer
Private m_Tag As String 'V1.44 GLL

Public m_bAnnulerDistList As Boolean
'Public m_otlDistList As Outlook.DistListItem

Private Sub AfficherCatFour() 'V1.44
'Afficher les fournisseur dans le combobox selon la catégorie choisis
 On Error GoTo Oups
 Dim i As Integer
 Dim sString As String
 Dim rstlist As ADODB.Recordset

 Set rstlist = New ADODB.Recordset
 sString = "Select * from GrbFournisseur "
 Call rstlist.Open(sString, g_connData, adOpenDynamic, adLockOptimistic)

 i = 0

 If Not rstlist.EOF Then
 Do While Not rstlist.EOF
 If rstlist.Fields("NomFournisseur") = cmbFournisseur.Text Then Exit Do
  Call rstlist.MoveNext
  Loop
 
  For i = 1 To Lst_Cat.ListItems.count
  If rstlist.Fields("Categorie") = Null Then Exit For 'si aucune catégorie sélectionner on fait rien

  If InStr(1, rstlist.Fields("categorie"), Lst_Cat.ListItems(i).Tag, vbTextCompare) > 0 Then
  Lst_Cat.ListItems(i).Checked = True
 End If
 Next
 End If
Exit Sub
Oups:
  wOups "FrmFRS", "AfficherCatFour", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboFournisseur()

 On Error GoTo Oups

 'Rempli le combo des fournisseurs
 Dim rstFournisseur As ADODB.Recordset

 Set rstFournisseur = New ADODB.Recordset

 Call rstFournisseur.Open("SELECT NomFournisseur, IDFRS FROM GrbFournisseur WHERE Supprimé = False", g_connData, adOpenDynamic, adLockOptimistic)

 'Il faut vider le combo avant de le remplir
 Call cmbFournisseur.Clear

 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstFournisseur.EOF
 'Ajout du nom du fournisseur dans le combo
 Call cmbFournisseur.AddItem(rstFournisseur.Fields("NomFournisseur"))

 'Ajout du numéro du fournisseur dans le ItemData du combo
 cmbFournisseur.ItemData(cmbFournisseur.newIndex) = rstFournisseur.Fields("IDFRS")

 Call rstFournisseur.MoveNext
 Loop

 Call rstFournisseur.Close
  Set rstFournisseur = Nothing

  If cmbFournisseur.ListCount > 0 Then
  cmbFournisseur.ListIndex = 0
  End If

  Exit Sub

Oups:

  wOups "frmFRS", "RemplirComboFournisseur", Err, Err.number, Err.Description
End Sub

Private Sub AfficherFournisseur()

 On Error GoTo Oups
 
 'Affiche le fournisseur sélectionné dans le combo
 Dim rstFournisseur As ADODB.Recordset
 Dim i As Integer

 Set rstFournisseur = New ADODB.Recordset
 
 Call rstFournisseur.Open("SELECT * FROM GrbFournisseur WHERE IDFRS = " & m_iNoFournisseur, g_connData, adOpenDynamic, adLockOptimistic)
 
 Call ViderBarrerChamps(True, True)
 
 'Adresse
 If Not IsNull(rstFournisseur.Fields("Adresse")) Then
 txtAdresse.Text = rstFournisseur.Fields("Adresse")
 End If
 
 'Ville
 If Not IsNull(rstFournisseur.Fields("Ville")) Then
 txtVille.Text = rstFournisseur.Fields("Ville")
 End If
 
 'Prov/Etat
  If Not IsNull(rstFournisseur.Fields("Prov/Etat")) Then
  txtProvEtat.Text = rstFournisseur.Fields("Prov/Etat")
  End If
 
 'Pays
  If Not IsNull(rstFournisseur.Fields("Pays")) Then
  txtPays.Text = rstFournisseur.Fields("Pays")
  End If
 
 'CodePostal
  If Not IsNull(rstFournisseur.Fields("CodePostal")) Then
  txtCP.Text = rstFournisseur.Fields("CodePostal")
10 End If
 
 'Telephonne
If Not IsNull(rstFournisseur.Fields("Telephonne")) Then
 txtTelephone.Text = rstFournisseur.Fields("Telephonne")
End If

 'Fax
If Not IsNull(rstFournisseur.Fields("Fax")) Then
 txtFax.Text = rstFournisseur.Fields("Fax")
End If
 
 'E-mail
If Not IsNull(rstFournisseur.Fields("E-mail")) Then
 txtEmail.Text = rstFournisseur.Fields("E-mail")
End If
 
 'Site Web
If Not IsNull(rstFournisseur.Fields("SiteWeb")) Then
 txtSiteWeb.Text = rstFournisseur.Fields("SiteWeb")
1  End If
 
 'commentaire
If Not IsNull(rstFournisseur.Fields("Commentaire")) Then
 txtcommentaire.Text = rstFournisseur.Fields("Commentaire")
End If

 'Création
 If Not IsNull(rstFournisseur.Fields("DateCréation")) Then
 lblDateCreation.Caption = rstFournisseur.Fields("DateCréation")
 End If

 'User Création
1  If Not IsNull(rstFournisseur.Fields("UserCréation")) Then
 lblUserCreation.Caption = "Par : " & rstFournisseur.Fields("UserCréation")
 End If

 'Modification
If Not IsNull(rstFournisseur.Fields("DateModification")) Then
 lblDateModification.Caption = rstFournisseur.Fields("DateModification")
End If

 'User Modification
If Not IsNull(rstFournisseur.Fields("UserModification")) Then
 lblUserModification.Caption = "Par : " & rstFournisseur.Fields("UserModification")
End If
 'Catégorie
 If Not IsNull(rstFournisseur.Fields("Categorie")) Then
 For i = 0 To cmbcatégorie.ListCount
 
 If cmbcatégorie.LIST(i) = rstFournisseur.Fields("Categorie") Then
 cmbcatégorie.ListIndex = (i)
 Exit For
 End If
 Next
 End If
Call rstFournisseur.Close
Set rstFournisseur = Nothing

Exit Sub

Oups:

wOups "frmFRS", "AfficherFournisseur", Err, Err.number, Err.Description
End Sub

Public Sub AfficherContact()

 On Error GoTo Oups
 
 ''''''''''''''''''''''''''''''''''''''''
 'affiche les contacts de l'employé selectionné'
 ''''''''''''''''''''''''''''''''''''''''
 Dim rstContact As ADODB.Recordset

 'Ouverture de la table contact
 Set rstContact = New ADODB.Recordset
 
 Call rstContact.Open("SELECT * FROM GrbContact WHERE IDContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)
 
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
 
 'REMPLIS LES CHAMPS s'il y a enregistrement
  If Not rstContact.EOF Then
  If Not IsNull(rstContact.Fields("Titre")) Then
  txtContactTitre.Text = rstContact.Fields("Titre")
  End If

  If Not IsNull(rstContact.Fields("cellulaire")) Then
  txtContactCell.Text = rstContact.Fields("cellulaire")
End If
 
1 If Not IsNull(rstContact.Fields("pagette")) Then
 txtContactPage.Text = rstContact.Fields("pagette")
 End If
 
 If Not IsNull(rstContact.Fields("telephonne")) Then
 txtContactTel.Text = rstContact.Fields("telephonne")
 End If
 
 If Not IsNull(rstContact.Fields("fax")) Then
 txtContactFax.Text = rstContact.Fields("fax")
 End If
 
 If Not IsNull(rstContact.Fields("e-mail")) Then
 txtContactEmail.Text = rstContact.Fields("e-mail")
End If
 
 If Not IsNull(rstContact.Fields("noposte")) Then
 txtContactPoste.Text = rstContact.Fields("noposte")
 End If

 If Not IsNull(rstContact.Fields("teldomicile")) Then
 txtContactDom.Text = rstContact.Fields("teldomicile")
 End If
1  End If
 
 'Ferme la table
 Call rstContact.Close
 Set rstContact = Nothing

Exit Sub

Oups:

wOups "frmFRS", "AfficherContact", Err, Err.number, Err.Description
End Sub



Private Sub cmb_modAnu_Click() 'V1.44 GLL
 'Bouton d'annulation des changement apporter a la liste des catégorie
 On Error GoTo Oups
 m_Tag = ""
 FrmCatMod.Visible = False
 cmdcatval.Visible = False
 cmb_modAnu.Visible = False
 cmdCatAdd.Visible = True
 cmdcatmod.Visible = True
 cmdcatenr.Visible = True
 cmdAnnuller.Visible = True
Exit Sub
Oups:
 wOups "frmFRS", "cmb_modAnu_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmb_modCat_Click() 'V1.44 GLL
 'Bouton pour modifier le nom d'une catégorie
 On Error GoTo Oups

 If m_bCategorie = True Then
 frm_Categorie.Visible = True
 frm_Categorie.Caption = "Catégorie pour :" & cmbFournisseur.Text
 Call AfficherCatList
 Call AfficherCatFour
 If Lst_Cat.ListItems.count >= 34 Then cmdCatAdd.Enabled = False
 End If
Exit Sub
Oups:
 wOups "frmFRS", "cmb_modCat_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbcatégorie_click() 'GLL 2017-11-2  V1.44
 'Active la réduction du nombre de fournisseur par catégorie
5 On Error GoTo Oups
 If m_bCategorie = False Then
 If cmbcatégorie.ListIndex <> -1 Then
 m_iNoCategorie = cmbcatégorie.ItemData(cmbcatégorie.ListIndex)
 Call AfficherCategorie
 End If
 End If
Exit Sub
Oups:
 wOups "frmFRS", "cmbCatégorie_Click", Err, Err.number, Err.Description
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

 wOups "frmFRS", "cmbContact_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulerContact_Click()

 On Error GoTo Oups

 Screen.MousePointer = vbHourglass

 Call AfficherControles(MODE_INACTIF)

 If m_bNewContact = True Then
 Call HideEdMaskContact(True)

 m_bNewContact = False
 End If
 
 'n'est plus en mode ajouter
 m_bModeAjoutContact = False
 
 txtNomFournisseur.Visible = False
 txtNomFournisseur.Locked = False

 'remplis combo contact
 Call RemplirComboContact
 
  Screen.MousePointer = vbDefault

  Exit Sub

Oups:

  wOups "frmFRS", "cmdanulcontact_Click", Err, Err.number, Err.Description
End Sub

Public Sub RemplirComboContact()

 On Error GoTo Oups
 
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 'remplis le combo contact dépendant le client sélectionné
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Dim rstContact As ADODB.Recordset
 
 Set rstContact = New ADODB.Recordset
 
 If m_bModeAjoutContact = True Then
 Call rstContact.Open("SELECT NomContact, IDContact FROM GrbContact WHERE Supprimé = False ORDER BY NomContact", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstContact.Open("SELECT GrbContact.NomContact, GrbContact.IDContact FROM GrbContact INNER JOIN GrbContactFRS ON GrbContact.IDContact = GrbContactFRS.NoContact WHERE GrbContactFRS.NoFRS = " & m_iNoFournisseur & " ORDER BY NomContact", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 
 Call cmbContact.Clear
 
 Do While Not rstContact.EOF
 Call cmbContact.AddItem(rstContact.Fields("NomContact"))
  cmbContact.ItemData(cmbContact.newIndex) = rstContact.Fields("IDContact")
 
  Call rstContact.MoveNext
  Loop
 
 'Ferme la table "GrbContact"
  Call rstContact.Close
  Set rstContact = Nothing
 
 'Affiche le contact de la table client
 'si combo est pas vide, affiche le premier contact, sinon le contact inscrit dans table client
  If cmbContact.ListCount > 0 Then
  cmbContact.ListIndex = 0
  Else
txtContactTitre.Text = vbNullString
1 txtContactCell.Text = vbNullString
 txtContactDom.Text = vbNullString
 txtContactEmail.Text = vbNullString
 txtContactFax.Text = vbNullString
 txtContactPage.Text = vbNullString
 txtContactPoste.Text = vbNullString
 txtContactTel.Text = vbNullString
End If

Exit Sub

Oups:

wOups "frmFRS", "RemplirComboContact", Err, Err.number, Err.Description
End Sub

Private Sub EnregistrerFournisseur()

 On Error GoTo Oups

 Dim rstFournisseur As ADODB.Recordset

 Set rstFournisseur = New ADODB.Recordset

 If m_bModeAjoutFRS = True Then
 Call rstFournisseur.Open("SELECT * FROM GrbFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call rstFournisseur.AddNew

 rstFournisseur.Fields("DateCréation") = ConvertDate(Date)
 rstFournisseur.Fields("UserCréation") = g_sInitiale
 Else
 Call rstFournisseur.Open("SELECT * FROM GrbFournisseur WHERE IDFRS = " & m_iNoFournisseur, g_connData, adOpenDynamic, adLockOptimistic)

 rstFournisseur.Fields("DateModification") = ConvertDate(Date)
  rstFournisseur.Fields("UserModification") = g_sInitiale
  End If

 'Enregistrement du fournisseur
  rstFournisseur.Fields("NomFournisseur").Value = txtNomFournisseur.Text
  rstFournisseur.Fields("Adresse").Value = txtAdresse.Text
  rstFournisseur.Fields("Ville").Value = txtVille.Text
  rstFournisseur.Fields("Prov/Etat").Value = txtProvEtat.Text
  rstFournisseur.Fields("Pays").Value = txtPays.Text
  rstFournisseur.Fields("CodePostal").Value = txtCP.Text
10 rstFournisseur.Fields("Telephonne").Value = mskTelephone.Text
rstFournisseur.Fields("Fax").Value = mskFax.Text
rstFournisseur.Fields("E-mail").Value = txtEmail.Text
rstFournisseur.Fields("siteweb").Value = txtSiteWeb.Text
rstFournisseur.Fields("Commentaire").Value = txtcommentaire.Text

rstFournisseur.Fields("EntryIDOutlook") = ModifierFRSExchange(rstFournisseur.Fields("IDFRS"))

If m_bModeAjoutFRS = True Then
 m_bModeAjoutFRS = False
End If

Call rstFournisseur.Update
 
Call rstFournisseur.Close
Set rstFournisseur = Nothing

1  Exit Sub

Oups:

wOups "frmFRS", "EnregistrerFournisseur", Err, Err.number, Err.Description
End Sub

Private Sub ModifierNomFRSExchange(ByVal sName As String, ByVal iFournisseurID As Integer)

 On Error GoTo Oups
 
 Dim otlApp As Outlook.Application
 Dim otlFRS As Outlook.ContactItem
 Dim folFRS As Outlook.MAPIFolder
 Dim bDejaOuvert As Boolean

 lblEtatOutlook.Caption = "Modification du nom du fournisseur dans Outlook ..."
 fraEtatOutlook.Visible = True

 Set otlApp = OuvrirOutlook(bDejaOuvert)

 Set folFRS = GetFolder(otlApp, "Fournisseurs GRB")

 Set otlFRS = folFRS.Items.Find("[User1] = " & iFournisseurID)

 If otlFRS Is Nothing Then
  Call MsgBox("Le fournisseur " & txtNomFournisseur.Text & " n'a pas été trouvé pour la mise à jour Exchange!", vbOKOnly, "Erreur")

  fraEtatOutlook.Visible = False

  DoEvents

  Exit Sub
  End If

  otlFRS.CompanyName = sName
 
  Call otlFRS.Save

  If bDejaOuvert = False Then
Call otlApp.Quit
End If

Set otlApp = Nothing

fraEtatOutlook.Visible = False

DoEvents

Exit Sub

Oups:

woups"frmFRS", "ModifierNomFRSExchange", Err, Erl, "iFournisseurID = " & iFournisseurID)

fraEtatOutlook.Visible = False
End Sub

Private Sub LierContactFournisseur(ByVal iFournisseurID As Integer)

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

 lblEtatOutlook.Caption = "Liaison du contact avec le fournisseur dans Outlook ..."
  fraEtatOutlook.Visible = True

  Set otlApp = OuvrirOutlook(bDejaOuvert)

  Set folFRS = GetFolder(otlApp, "Fournisseurs GRB")
  Set folContact = GetFolder(otlApp, "Contacts GRB")

  Set rstFRS = New ADODB.Recordset

  Call rstFRS.Open("SELECT EntryIDOutlook FROM GrbFournisseur WHERE IDFRS = " & m_iNoFournisseur, g_connData, adOpenForwardOnly, adLockReadOnly)

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

 Call rstContactFRS.Open("SELECT * FROM GrbContactFRS WHERE NoFRS = " & m_iNoFournisseur, g_connData, adOpenForwardOnly, adLockReadOnly)

 Do While Not rstContactFRS.EOF
 Set itmContact = folContact.Items.Find("[User1] = " & rstContactFRS.Fields("NoContact"))

 If Not itmContact Is Nothing Then
1  Call itmFRS.Links.Add(itmContact)

 Call itmFRS.Save

 Call itmContact.Links.Add(itmFRS)

 Call itmContact.Save
 End If

 Call rstContactFRS.MoveNext
 Loop

 Call rstContactFRS.Close
 Set rstContactFRS = Nothing
Else
 Call MsgBox("Fournisseur introuvable!", vbOKOnly, "Erreur")

 Call rstFRS.Close
 Set rstFRS = Nothing
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
 "- Dans Outlook, ouvrez le fournisseur '" & txtNomFournisseur.Text & "' dans Fournisseurs GRB" & vbNewLine & _
 "- Cliquez sur tous les contacts de ce fournisseur 1 à la fois pour trouver le contact incorrect." & vbNewLine & _
 "- Effacez ce contact de la liste des contacts de ce fournisseur." & vbNewLine & _
 "- Dans le logiciel GRB, recommencez l'opération (effacez le contact et l'ajouter de nouveau).", vbOKOnly, "Erreur")
Else
 woups"frmFRS", "LierContactFournisseur", Err, Erl, txtNomFournisseur.Text)
End If

fraEtatOutlook.Visible = False
End Sub

Private Function ModifierFRSExchange(ByVal iFournisseurID As Integer) As String
 
 On Error GoTo Oups
 
 Dim otlApp As Outlook.Application
 Dim otlFRS As Outlook.ContactItem
 Dim folFRS As Outlook.MAPIFolder
 Dim bDejaOuvert As Boolean

 If m_bModeAjoutFRS = True Then
 lblEtatOutlook.Caption = "Ajout du fournisseur dans Outlook ..."
 Else
 lblEtatOutlook.Caption = "Modification du fournisseur dans Outlook ..."
 End If

 fraEtatOutlook.Visible = True

  Set otlApp = OuvrirOutlook(bDejaOuvert)

  Set folFRS = GetFolder(otlApp, "Fournisseurs GRB")

  If m_bModeAjoutFRS = True Then
  Set otlFRS = folFRS.Items.Add(olContactItem)

  otlFRS.User1 = iFournisseurID
  Else
  Set otlFRS = folFRS.Items.Find("[User1] = " & iFournisseurID)
  End If

10 If otlFRS Is Nothing Then
1 Call MsgBox("Le fournisseur " & txtNomFournisseur.Text & " n'a pas été trouvé pour la mise à jour Exchange!", vbOKOnly, "Erreur")

 fraEtatOutlook.Visible = False

 DoEvents

 Exit Function
End If

otlFRS.CompanyName = txtNomFournisseur.Text
 
If mskTelephone.Text <> "(___) ___-____" Then
 otlFRS.BusinessTelephoneNumber = mskTelephone.Text
End If
 
If mskFax.Text <> "(___) ___-____" Then
 otlFRS.BusinessFaxNumber = mskFax.Text
1  End If
 
otlFRS.Email1Address = txtEmail.Text
 otlFRS.BusinessAddressStreet = txtAdresse.Text
otlFRS.BusinessAddressCity = txtVille.Text
 otlFRS.BusinessAddressState = txtProvEtat.Text
otlFRS.BusinessAddressCountry = txtPays.Text
 otlFRS.BusinessAddressPostalCode = txtCP.Text
1  otlFRS.Body = txtcommentaire.Text
 otlFRS.WebPage = txtSiteWeb.Text
 
 Call otlFRS.Save

ModifierFRSExchange = otlFRS.EntryID

If bDejaOuvert = False Then
 Call otlApp.Quit
End If

Set otlApp = Nothing

fraEtatOutlook.Visible = False

DoEvents

Exit Function

Oups:

woups"frmFRS", "ModifierFRSExchange", Err, Erl, "iFournisseurID = " & iFournisseurID)

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
End Select
 
otlContact.Title = ""

otlContact.CompanyName = txtNomFournisseur.Text
otlContact.JobTitle = txtContactTitre.Text

If Trim$(mskContactTel.Text) <> "" Then
 If mskContactTel.Text <> "(___) ___-____" Then
 If Trim$(txtContactPoste.Text) <> "" Then
 otlContact.BusinessTelephoneNumber = mskContactTel.Text & " Ext : " & txtContactPoste.Text
 Else
 otlContact.BusinessTelephoneNumber = mskContactTel.Text
 End If
End If
End If
 
 If mskContactFax.Text <> "(___) ___-____" Then
 otlContact.BusinessFaxNumber = mskContactFax.Text
 End If
 
If mskContactCell.Text <> "(___) ___-____" Then
 otlContact.MobileTelephoneNumber = mskContactCell.Text
1  End If

 If mskContactDom.Text <> "(___) ___-____" Then
 otlContact.HomeTelephoneNumber = mskContactDom.Text
End If
 
If mskContactPage.Text <> "(___) ___-____" Then
 otlContact.PagerNumber = mskContactPage.Text
End If
 
otlContact.Email1Address = txtContactEmail.Text
 
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

2  woups"frmFRS", "AjouterContactExchange", Err, Erl, "iContactID = " & iContactID)

fraEtatOutlook.Visible = False
End Function

Private Sub SupprimerFRSExchange(ByVal iFournisseurID As Integer)

 On Error GoTo Oups

 Dim otlApp As Outlook.Application
 Dim otlFRS As Outlook.ContactItem
 Dim folFRS As MAPIFolder
 Dim bDejaOuvert As Boolean

 lblEtatOutlook.Caption = "Suppression du fournisseur dans Outlook ..."
 fraEtatOutlook.Visible = True

 Set otlApp = OuvrirOutlook(bDejaOuvert)

 Set folFRS = GetFolder(otlApp, "Fournisseurs GRB")

 Set otlFRS = folFRS.Items.Find("[User1] = " & iFournisseurID)

 If Not otlFRS Is Nothing Then
  Call otlFRS.Delete
  End If

  If bDejaOuvert = False Then
  Call otlApp.Quit
  End If

  Set otlApp = Nothing

  fraEtatOutlook.Visible = False

  DoEvents

10 Exit Sub

Oups:

wOups "frmFRS", "SupprimerFRSExchange", Err, Err.number, Err.Description

fraEtatOutlook.Visible = False
End Sub

Private Sub cmdAnnuller_Click()
frm_Categorie.Visible = False
Call RemplirComboCatégorie
End Sub

Private Sub cmdCatAdd_Click() 'V1.44 GLL
 'Bouton pour ajouter une catégorie a la base de donné
 On Error GoTo Oups

 If Lst_Cat.ListItems.count >= 34 Then 'Méthode utilisé pour géré les catégorie on une limite de 34 alors je bloque les futur addition pour ne pas avoir de problème
 MsgBox "Vous Avez attent le maximum de catégorie"
 cmdCatAdd.Enabled = False
 Exit Sub
 End If
 
 m_Tag = ""

 FrmCatMod.Visible = True
 cmdcatval.Visible = True
 cmb_modAnu.Visible = True
 cmdCatAdd.Visible = False
 cmdcatmod.Visible = False
 cmdcatenr.Visible = False
  cmdAnnuller.Visible = False
  cmdcatval.Default = True

  txtmodcat.SetFocus
  FrmCatMod.Caption = "Ajouter"
  txtmodcat.Text = ""

 Exit Sub
 
Oups:
  wOups "FrmFRS", "cmdCatAdd_Click", Err, Err.number, Err.Description

End Sub

Private Sub cmdcatenr_Click() '1.44 GLL Enregistre les catégorie pour le founisseur
On Error GoTo Oups

 Dim rstcat As ADODB.Recordset
 Dim i As Integer
 Dim sCat As String

 Set rstcat = New ADODB.Recordset
 sCat = ""
 Call rstcat.Open("Select * from GrbFournisseur Where NomFournisseur ='" & cmbFournisseur.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstcat.EOF Then 'Vérifie si un Fournisseur est sélectionner
 MsgBox "Erreur aucun fournisseur sélectionner"
 Exit Sub
 End If
 
 For i = 1 To Lst_Cat.ListItems.count 'Fabric le nouveau code pour la catégorie
 If Lst_Cat.ListItems(i).Checked Then sCat = sCat & Lst_Cat.ListItems(i).Tag
 Next
 
 If sCat = "" Then rstcat.Fields("Categorie").Value = Null 'si aucune catégorie est selectionner on rend null la case categorie
  If sCat <> "" Then rstcat.Fields("Categorie").Value = sCat 'On envoie le code dans la catégorie du fournisseur

  Call rstcat.Update
  Call rstcat.Close
  Set rstcat = Nothing


  frm_Categorie.Visible = False
  Call RemplirComboCatégorie
 Exit Sub
Oups:
  wOups "frmFrs", "cmdcatenr_click", Err, Err.number, Err.Description
End Sub


Private Sub cmdcatmod_Click() 'V1.44 GLL
 'bouton pour modifier le nom d'une catégorie
 On Error GoTo Oups

 m_Tag = Lst_Cat.SelectedItem.Tag
 FrmCatMod.Visible = True
 cmdcatval.Visible = True
 cmb_modAnu.Visible = True
 cmdCatAdd.Visible = False
 cmdcatmod.Visible = False
 cmdcatenr.Visible = False
 cmdAnnuller.Visible = False
 txtmodcat.SetFocus
 FrmCatMod.Caption = "Modifier"
  txtmodcat.Text = Lst_Cat.SelectedItem.Text
 Exit Sub
Oups:
wOups "FrmFrs", "cmdcatmod_Click", Err, Err.number, Err.Description
End Sub
Private Sub cmdcatval_Click() 'V1.44 GLL
 'Bouton pour valider l'Additon/modification d'une catégorie
 On Error GoTo Oups
 Dim rstag As ADODB.Recordset
 Dim bDelete As Boolean
 Set rstag = New ADODB.Recordset
 bDelete = False

 If m_Tag <> "" Then 'pour faire une modification
 Call rstag.Open("SELECT * FROM TBL_Categorie where Correspondance ='" & m_Tag & "'", g_connData, adOpenDynamic, adLockOptimistic)
 rstag.Fields("Nom").Value = txtmodcat.Text
 Else 'pour faire une addition d'une catégorie
 Call rstag.Open("SELECT * FROM TBL_Categorie", g_connData, adOpenDynamic, adLockOptimistic)

 Do While Not rstag.EOF
  If UCase(rstag.Fields("nom")) = UCase(txtmodcat.Text) Then 'Vérifie si ce nom de catégorie existe déja
  MsgBox "vous avez déja cette Categorie"
  GoTo Fin
 End If
  Call rstag.MoveNext
 Loop
  rstag.MoveFirst
 
  Do While Not rstag.EOF

  If IsNull(rstag.Fields("nom")) Then
  rstag.Fields("Nom").Value = txtmodcat.Text
 Exit Do
 End If

 Call rstag.MoveNext
 Loop
 End If
 
If txtmodcat.Text = "" Then 'Si on a pas miss de text on efface le nom de la catégorie
 rstag.Fields("Nom").Value = Null
 bDelete = True
End If

Call rstag.Update

Fin:
Call rstag.Close
Set rstag = Nothing
 
If bDelete Then Call DeleteCategorie
 
Call AfficherCatList
1  Call AfficherCatFour

FrmCatMod.Visible = False
 cmdcatval.Visible = False
cmb_modAnu.Visible = False
 cmdCatAdd.Visible = True
 cmdcatmod.Visible = True
1  cmdcatenr.Visible = True
 cmdAnnuller.Visible = True
 cmdcatmod.Enabled = False
m_Tag = ""

Exit Sub
Oups:
wOups "Frm_FRS", "Cmdcatval_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdEnregistrerContact_Click()

 On Error GoTo Oups
 
 Dim rstContactFRS As ADODB.Recordset
 Dim rstContact As ADODB.Recordset
 
 Screen.MousePointer = vbHourglass
 
 Set rstContactFRS = New ADODB.Recordset
 
 If m_bNewContact = True Then
 Set rstContact = New ADODB.Recordset

 Call rstContact.Open("SELECT * FROM GrbContact", g_connData, adOpenDynamic, adLockOptimistic)

 Call rstContact.AddNew

 rstContact.Fields("NomContact").Value = txtcontact.Text
 rstContact.Fields("Titre").Value = txtContactTitre.Text
  rstContact.Fields("Compagnie").Value = txtNomFournisseur.Text
  rstContact.Fields("Telephonne").Value = mskContactTel.Text
  rstContact.Fields("Fax").Value = mskContactFax.Text
  rstContact.Fields("Pagette").Value = mskContactPage.Text
  rstContact.Fields("Cellulaire").Value = mskContactCell.Text
  rstContact.Fields("E-mail").Value = txtContactEmail.Text
  rstContact.Fields("noposte").Value = txtContactPoste.Text
  rstContact.Fields("teldomicile").Value = mskContactDom.Text
rstContact.Fields("UserCréation").Value = g_sInitiale
1 rstContact.Fields("DateCréation").Value = ConvertDate(Date)

 rstContact.Fields("EntryIDOutlook") = AjouterContactExchange(rstContact.Fields("IDContact"))
 
 Call rstContact.Update

 'set la table
 Call rstContactFRS.Open("SELECT * FROM GrbContactFRS WHERE NoFRS = " & m_iNoFournisseur & " And NoContact = " & rstContact.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'si pas deja existant, on ajoute dans la table
 If rstContactFRS.EOF Then
 'ajoute dans la table
 Call rstContactFRS.AddNew
 
 rstContactFRS.Fields("NoFRS") = m_iNoFournisseur
 rstContactFRS.Fields("NoContact") = rstContact.Fields("IDContact")
 
 Call rstContactFRS.Update
 End If
 
 Call rstContact.Close
Set rstContact = Nothing

 m_bNewContact = False

 Call HideEdMaskContact(True)
Else
 'set la table
 Call rstContactFRS.Open("SELECT * FROM GrbContactFRS WHERE NoFRS = " & m_iNoFournisseur & " And NoContact = " & m_iNoContact, g_connData, adOpenDynamic, adLockOptimistic)
 
 'Si pas deja existant, on ajoute dans la table
 If rstContactFRS.EOF Then
 'ajoute dans la table
 Call rstContactFRS.AddNew
 
1  rstContactFRS.Fields("NoFRS") = m_iNoFournisseur
 rstContactFRS.Fields("NoContact") = m_iNoContact
 
 Call rstContactFRS.Update
 End If

 'Ferme tables et connection
 Call rstContactFRS.Close
End If
 
Call LierContactFournisseur(m_iNoFournisseur)
 
Set rstContactFRS = Nothing
 
 'Bouton comme avant(apparait)
Call AfficherControles(MODE_INACTIF)
 
 'N'est plus en mode ajouter
m_bModeAjoutContact = False

 'Remplis combo contact
Call RemplirComboContact

Call ViderBarrerChampsContact(True, False)

Screen.MousePointer = vbDefault

2  Exit Sub

Oups:

wOups "frmFRS", "cmdEnregistrerContact_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdFax_Click()

 On Error GoTo Oups

 If cmbFournisseur.ListCount > 0 Then
 If cmbContact.ListIndex > -1 Then
 Call frmreport.Afficher(cmbFournisseur.ItemData(cmbFournisseur.ListIndex), cmbContact.ItemData(cmbContact.ListIndex), FRM_FRS)
 Else
 Call frmreport.Afficher(cmbFournisseur.ItemData(cmbFournisseur.ListIndex), 0, FRM_FRS)
 End If
 End If

 Exit Sub

Oups:

 wOups "frmFRS", "cmdFax_Click", Err, Err.number, Err.Description
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
  Dim sMsgPlein As Boolean
  Dim iNbreRendu As Integer

  If cmbContact.ListIndex <> -1 Then
  bArrayVide = True

  iResult = MsgBox("Voulez-vous ajouter tous les contacts à la liste de distribution?" & vbNewLine & _
 "Oui - Tous les contacts" & vbNewLine & _
 "Non - Contact affiché seulement", vbYesNoCancel)

If iResult = vbYes Then
Set rstContact = New ADODB.Recordset

 Call rstContact.Open("SELECT * FROM GrbContactFRS INNER JOIN GrbContact ON GrbContactFRS.NoContact = GrbContact.IDContact WHERE GrbContactFRS.NoFRS = " & cmbFournisseur.ItemData(cmbFournisseur.ListIndex) & " ORDER BY NomContact", g_connData, adOpenForwardOnly, adLockOptimistic)
 
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
 If Trim$(txtContactEmail.Text) <> "" Then
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
 wOups "frmFRS", "cmdMailListContact_Click", Err, Err.number, Err.Description
 End If

5  fraEtatOutlook.Visible = False
End Sub

Private Sub cmdMailListFournisseur_Click()

 On Error GoTo Oups

 Dim otlApp As Outlook.Application
 Dim folFRS As Outlook.MAPIFolder
 Dim itmFRS As Outlook.ContactItem
 Dim otlRecipient As Outlook.Recipient
 Dim bDejaOuvert As Boolean

 If cmbFournisseur.ListIndex <> -1 Then
 If Trim$(txtEmail.Text) <> "" Then
 Set otlApp = OuvrirOutlook(bDejaOuvert)

 lblEtatOutlook.Caption = "Recherche des listes de distribution..."

  fraEtatOutlook.Visible = True

  Call frmChoixMailList.Afficher(Me, otlApp)

  If m_bAnnulerDistList = False Then
  lblEtatOutlook.Caption = "Ajout du fournisseur dans la liste de distribution..."

  fraEtatOutlook.Visible = True

  If m_otlDistList.MemberCount < 10 Then
  Set folFRS = GetFolder(otlApp, "Fournisseurs GRB")

  Set itmFRS = folFRS.Items.Find("[User1] = " & cmbFournisseur.ItemData(cmbFournisseur.ListIndex))

 If Not itmFRS Is Nothing Then
 Set otlRecipient = otlApp.Session.CreateRecipient(itmFRS.Email1DisplayName)

 If otlRecipient.Resolve = True Then
 Call m_otlDistList.AddMember(otlRecipient)

 Call m_otlDistList.Save
 Else
 Call MsgBox("Impossible de trouver le fournisseur!", vbOKOnly, "Erreur")
 End If
 Else
 Call MsgBox("Fournisseur introuvable!", vbOKOnly, "Erreur")
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
 Call MsgBox("Ce fournisseur n'a pas d'email!", vbOKOnly, "Erreur")
 End If
Else
 Call MsgBox("Aucun fournisseur sélectionné!", vbOKOnly, "Erreur")
End If

Exit Sub

Oups:

If Err.number = 2 And Erl = 115 Then
 Call MsgBox("La liste de distribution est pleine!", vbOKOnly, "Erreur")
Else
 wOups "frmFRS", "cmdMailListFournisseur_Click", Err, Err.number, Err.Description
End If

2  fraEtatOutlook.Visible = False
End Sub

Private Sub cmdRafraichir_Click()

 On Error GoTo Oups
 
 'Rafraichir la liste après avoir fait une recherche
 Screen.MousePointer = vbHourglass
 
 'Remplir le combo
 Call RemplirComboFournisseur
 cmbcatégorie.ListIndex = -1
 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmFRS", "cmdRafraichir_Click", Err, Err.number, Err.Description
End Sub

Private Sub ViderBarrerChamps(ByVal bLocked As Boolean, ByVal bVider As Boolean)

 On Error GoTo Oups
 'Cette procédure vide et unlock tous les textbox pour pouvoir ajouter
 
 If bVider = True Then
 txtAdresse.Text = vbNullString
 txtVille.Text = vbNullString
 txtProvEtat.Text = vbNullString
 txtPays.Text = vbNullString
 txtCP.Text = vbNullString
 txtEmail.Text = vbNullString
 txtSiteWeb.Text = vbNullString
 txtcommentaire.Text = vbNullString
  txtTelephone.Text = vbNullString
  txtFax.Text = vbNullString
  lblDateCreation.Caption = vbNullString
  lblUserCreation.Caption = vbNullString
  lblDateModification.Caption = vbNullString
  lblUserModification.Caption = vbNullString
  End If
 
  txtAdresse.Locked = bLocked
10 txtVille.Locked = bLocked
txtProvEtat.Locked = bLocked
txtPays.Locked = bLocked
txtCP.Locked = bLocked
txtEmail.Locked = bLocked
txtSiteWeb.Locked = bLocked
txtTelephone.Locked = bLocked
txtFax.Locked = bLocked
txtcommentaire.Locked = bLocked

Exit Sub

Oups:

wOups "frmFRS", "ViderBarrerChamps", Err, Err.number, Err.Description
End Sub

Private Sub ViderBarrerChampsContact(ByVal bLocked As Boolean, ByVal bVider As Boolean)

 On Error GoTo Oups
 'Cette procédure vide et unlock tous les textbox pour pouvoir ajouter
 
 If bVider = True Then
 txtContactTitre.Text = vbNullString
 txtContactCell.Text = vbNullString
 txtContactDom.Text = vbNullString
 txtContactEmail.Text = vbNullString
 txtContactFax.Text = vbNullString
 txtContactPage.Text = vbNullString
 txtContactPoste.Text = vbNullString
 txtContactTel.Text = vbNullString
 End If
 
  txtContactTitre.Locked = bLocked
  txtContactCell.Locked = bLocked
  txtContactDom.Locked = bLocked
  txtContactEmail.Locked = bLocked
  txtContactFax.Locked = bLocked
  txtContactPage.Locked = bLocked
  txtContactPoste.Locked = bLocked
  txtContactTel.Locked = bLocked

10 Exit Sub

Oups:

wOups "frmFRS", "ViderBarrerChampsContact", Err, Err.number, Err.Description
End Sub
Private Sub CmdAddCont_Click()

 On Error GoTo Oups
 
 'Pour faire l'ajout d'un contact
 Dim sNom As String
 Dim rstContact As ADODB.Recordset
 Dim bAjouter As Boolean

 If cmbFournisseur.ListCount > 0 Then
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

 txtNomFournisseur.Visible = True
 txtNomFournisseur.Locked = True

 'Remplis le combo avec tous les contacts
 Call AfficherControles(MODE_CONTACT)

 Call txtContactTitre.SetFocus
 End If
 Else
1  Screen.MousePointer = vbHourglass

 m_bNewContact = False

 txtNomFournisseur.Visible = True
 txtNomFournisseur.Locked = True

 'Remplis le combo avec tous les contacts
 Call AfficherControles(MODE_CONTACT)

 Call RemplirComboContact
 End If

 Screen.MousePointer = vbDefault
Else
 Call MsgBox("Aucun enregistrement de sélectionné")
End If

Exit Sub

Oups:

wOups "frmFRS", "CmdAddCont_Click", Err, Err.number, Err.Description
End Sub

Private Sub HideEdMask(ByVal bVisible As Boolean)

 On Error GoTo Oups
 
 'proc qui rend visible/ou non les maskEdBox
 'On en profite pour les nettoyer du dernier Enregistrement
 'et on fait l'inverse avec les textBox
 If m_bModeAjoutFRS = True Then
 txtTelephone.Text = vbNullString
 txtFax.Text = vbNullString
 
 mskTelephone.Text = vbNullString
 mskFax.Text = vbNullString
 Else
 mskTelephone.Text = txtTelephone.Text
 mskFax.Text = txtFax.Text
 End If
 
 mskTelephone.Visible = Not bVisible
  txtTelephone.Visible = bVisible

  mskFax.Visible = Not bVisible
  txtFax.Visible = bVisible

  Exit Sub

Oups:

  wOups "frmFRS", "HideEdMask", Err, Err.number, Err.Description
End Sub

Private Sub HideEdMaskContact(ByVal bVisible As Boolean)

 On Error GoTo Oups
 
 'proc qui rend visible/ou non les maskEdBox
 'On en profite pour les nettoyer du dernier Enregistrement
 'et on fait l'inverse avec les textBox
 If m_bModeAjoutContact = True Then
 txtContactTel.Text = vbNullString
 txtContactFax.Text = vbNullString
 txtContactPage.Text = vbNullString
 txtContactCell.Text = vbNullString
 txtContactDom.Text = vbNullString
 
 mskContactTel.Text = vbNullString
 mskContactFax.Text = vbNullString
 mskContactPage.Text = vbNullString
 mskContactCell.Text = vbNullString
  mskContactDom.Text = vbNullString
  End If
 
  mskContactTel.Visible = Not bVisible
  txtContactTel.Visible = bVisible

  mskContactFax.Visible = Not bVisible
  txtContactFax.Visible = bVisible

  mskContactPage.Visible = Not bVisible
  txtContactPage.Visible = bVisible

10 mskContactCell.Visible = Not bVisible
txtContactCell.Visible = bVisible

mskContactDom.Visible = Not bVisible
txtContactDom.Visible = bVisible

Exit Sub

Oups:

wOups "frmFRS", "HideEdMaskContact", Err, Err.number, Err.Description
End Sub

Private Sub cmdreport_Click()

 On Error GoTo Oups
 
 'Impression de la liste des fournisseurs
 Dim rstFournisseur As ADODB.Recordset

 Set rstFournisseur = New ADODB.Recordset

 If MsgBox("Voulez-vous imprimer ce fournisseur uniquement?", vbYesNo) = vbYes Then
 Call rstFournisseur.Open("SELECT * FROM GrbFournisseur WHERE IDFRS = " & m_iNoFournisseur, g_connData, adOpenDynamic, adLockOptimistic)
 Else
 If MsgBox("Voulez-vous filtrer par la ville '" & txtVille.Text & "'?", vbYesNo) = vbYes Then
 Call rstFournisseur.Open("SELECT * FROM GrbFournisseur WHERE ville = '" & Replace(txtVille.Text, "'", "''") & "' AND Supprimé = False ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstFournisseur.Open("SELECT * FROM GrbFournisseur WHERE Supprimé = False ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
 End If
  End If
 
  Screen.MousePointer = vbHourglass
 
 'Set le rapport
  Set DR_ListeFournisseur.DataSource = rstFournisseur
 
  DR_ListeFournisseur.Orientation = rptOrientLandscape

  Call DR_ListeFournisseur.Show(vbModal)
 
  Call rstFournisseur.Close
  Set rstFournisseur = Nothing
 
  Screen.MousePointer = vbDefault

10 Exit Sub

Oups:

wOups "frmFRS", "cmdreport_Click", Err, Err.number, Err.Description
End Sub

Private Sub AfficherControles(ByVal eMode As enumMode)

 On Error GoTo Oups
 
 'Proc qui fait le switch boutton visible/invible
 'on utilise le textBox dummy pour montrer contact
 Dim bCmbFournisseur As Boolean
 Dim bTxtFournisseur As Boolean
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
Dim bMailListFRS As Boolean
Dim bMailListContact As Boolean
 
Select Case eMode
 Case MODE_FRS:
 bTxtFournisseur = True
 bCmdEnr = True
 bCmdAnul = True
14 m_bCategorie = True 'GLL 1.44
14 cmb_modCat.Visible = True 'GLL 1.44

 Case MODE_CONTACT:
 bFraContact = True
 bTxtFournisseur = True
 bCmdAnulContact = True
 bCmdRefCont = True

 If m_bNewContact = True Then
 bTxtContact = True
 Else
 bCmbContact = True
 End If

 Case MODE_INACTIF:
 bFraContact = True
1  bCmbFournisseur = True
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
 bMailListContact = True
 bMailListFRS = True
 m_bCategorie = False 'GLL V1.44
 cmb_modCat.Visible = False 'GLL V1.44
 If Len(txtRechercher.Text) > 0 Then
 bCmdRechercher = True
 End If
2  End Select
 
cmbFournisseur.Visible = bCmbFournisseur
30 txtNomFournisseur.Visible = bTxtFournisseur
fracontact.Visible = bFraContact
txtRechercher.Enabled = bTxtRechercher
cmdRechercher.Enabled = bCmdRechercher
cmdRafraichir.Enabled = bCmdRafraichir
cmdrenommer.Enabled = bCmdRenommer
cmdReport.Visible = bCmdImprimer
CmdAdd.Visible = bCmdAdd
CmdSupp.Visible = bCmdSupp
CmdModif.Visible = bCmdModif
CmdFerme.Visible = bCmdQuit
CmdAnul.Visible = bCmdAnul
3  CmdEnr.Visible = bCmdEnr
CmdAddCont.Visible = bCmdAddCont
3  cmdsupcontact.Visible = bCmdSupContact
cmdAnnulerContact.Visible = bCmdAnulContact
3  cmdEnregistrerContact.Visible = bCmdRefCont
cmdFax.Visible = bFax
3  cmbContact.Visible = bCmbContact
 txtcontact.Visible = bTxtContact
40 cmdMailListFournisseur.Visible = bMailListFRS
cmdMailListContact.Visible = bMailListContact

4 Exit Sub

Oups:

4 wOups "frmFRS", "AfficherControles", Err, Err.number, Err.Description
End Sub

Private Sub CmdAdd_Click()

 On Error GoTo Oups
 'proc qui permet d'ajouter un contact à la BD
 Dim sName As String
 
 Call AfficherControles(MODE_FRS)

 sName = InputBox("Veuillez entrer le nom du fournisseur" & vbNewLine & _
 vbNewLine & _
 "SVP, respectez le bon orthographe!", "SAISIE DU NOM", "Nom du fournisseur")
 
 If sName <> vbNullString Then
 If Not ComboContient(cmbFournisseur, sName) Then
 Screen.MousePointer = vbHourglass
 
 m_bModeAjoutFRS = True
 
 'On montre les maskEdBox
 Call HideEdMask(False)
 
 'On affiche le nom du nouveau client dans le textbox
 'pour éviter le ScrollDown durant l'ajout
 txtNomFournisseur.Text = sName

 Call ViderBarrerChamps(False, True)

  Call mskTelephone.SetFocus
 
  Screen.MousePointer = vbDefault
  Else
  Call MsgBox("Ce fournisseur existe déjà!")
 
  Call AfficherControles(MODE_INACTIF)
  End If
  Else
  Call AfficherControles(MODE_INACTIF)
10 End If

Exit Sub

Oups:

wOups "frmFRS", "CmdAdd_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdAnul_Click()

 On Error GoTo Oups
 
 'Annule l'ajout ou la modif
 Screen.MousePointer = vbHourglass

 'On cache le maskEdBox
 Call HideEdMask(True)
 
 'commentaire unlock
 'txtNomClient.Visible = False
 
 'on retablis les bouttons
 Call AfficherControles(MODE_INACTIF)

 m_bModeAjoutFRS = False
 
 Call ViderBarrerChamps(True, True)
 
 Call cmbFournisseur_Click
 
 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmFRS", "CmdAnul_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdEnr_Click()

 On Error GoTo Oups

 'Enregistrement d'une modif ou d'un ajout
 Dim sFournisseur As String
 Dim iCompteur As Integer
 
 'Nom du fournisseur
 sFournisseur = txtNomFournisseur.Text
 
 'Enregistrement d'un frs dans la BD
 Screen.MousePointer = vbHourglass
 
 Call EnregistrerFournisseur
 
 'On cache les MaskEdBox
 Call HideEdMask(True)
 
 'On met a jour le combo
 Call RemplirComboFournisseur
 
 'Retablir les boutons
 Call AfficherControles(MODE_INACTIF)
 
 For iCompteur = 0 To cmbFournisseur.ListCount - 1
 If cmbFournisseur.LIST(iCompteur) = sFournisseur Then
  cmbFournisseur.ListIndex = iCompteur
 
  Exit For
  End If
  Next
 
  Call cmbFournisseur.SetFocus
 
  Screen.MousePointer = vbDefault

  Exit Sub

Oups:

  wOups "frmFRS", "CmdEnr_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdFerme_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmFRS", "CmdFerme_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdModif_Click()

 On Error GoTo Oups

 If cmbFournisseur.ListCount > 0 Then
 Screen.MousePointer = vbHourglass
 
 'proc qui permet de modifier l'enr courant
 Call HideEdMask(False)
 
 Call AfficherControles(MODE_FRS)
 
 Call ViderBarrerChamps(False, False)
 
 Screen.MousePointer = vbDefault
 Else
 Call MsgBox("Aucun enregistrement sélectionné!")
 End If

 Exit Sub

Oups:

  wOups "frmFRS", "CmdModif_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRenommer_Click()

 On Error GoTo Oups
 
 ''''''''''''''''''''''''''''''''''''''
 'on renomme le nom du FOURNISSEUR
 ''''''''''''''''''''''''''''''''''''''''
 Dim rstFournisseur As ADODB.Recordset
 Dim sName As String

 If cmbFournisseur.ListCount > 0 Then
 'Proc qui permet de modifié un CLIENT a la BD
 'On procede a la saisie du nom du CLIENT
 sName = InputBox("Veuillez entrer le nom du Fournisseur", "Renommer fournisseur", txtNomFournisseur.Text)
 
 If sName <> vbNullString Then
 If sName <> txtNomFournisseur.Text Then
 If Not ComboContient(cmbFournisseur, sName) Then
 Screen.MousePointer = vbHourglass
 
 Set rstFournisseur = New ADODB.Recordset
 
 Call rstFournisseur.Open("SELECT IDFrs, NomFournisseur, EntryIDOutlook FROM GrbFournisseur WHERE IDFRS = " & m_iNoFournisseur, g_connData, adOpenDynamic, adLockOptimistic)
 
  Call ModifierNomFRSExchange(sName, m_iNoFournisseur)
 
  txtNomFournisseur = sName
 
 'transfert des donnes
  rstFournisseur.Fields("NomFournisseur").Value = txtNomFournisseur.Text
 
 'mise a jour de la base de donne
  Call rstFournisseur.Update
 
  Call rstFournisseur.Close
  Set rstFournisseur = Nothing
 
  Call RemplirComboFournisseur
 
  cmbFournisseur.Text = sName
 
 m_bRenommer = True
 
 Call cmbFournisseur_Click
 
 m_bRenommer = False
 
 Screen.MousePointer = vbDefault
 Else
 Call MsgBox("Le fournisseur " & sName & " existe déjà!", vbCritical)
 End If
 End If
 End If
Else
 Call MsgBox("Aucun enregistrement de sélectionné!", vbOKOnly, "Erreur")
End If

1  Exit Sub

Oups:

wOups "frmFRS", "cmdRenommer_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdsupcontact_Click()

 On Error GoTo Oups
 
 'Fonction qui supprime l'enregistrement courant
 If cmbContact.ListCount > 0 Then
 If MsgBox("Etes-vous sur de vouloir supprimer cette enregistrement?", vbYesNo, "Supprimer") = vbYes Then
 Screen.MousePointer = vbHourglass
 
 Call g_connData.Execute("DELETE * FROM GrbContactFRS WHERE NoFRS = " & m_iNoFournisseur & " AND NoContact = " & m_iNoContact)

 Call LierContactFournisseur(m_iNoFournisseur)

 'Femplis le combo employé
 Call RemplirComboContact

 Screen.MousePointer = vbDefault
 End If
 End If

 Exit Sub

Oups:

  wOups "frmFRS", "cmdsupcontact_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdSupp_Click()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim rstCatalogue As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim bPeutEffacer As Boolean
 
 'Fonction qui supprime lenregistrement courant
 If cmbFournisseur.ListCount > 0 Then
 If MsgBox("Etes-vous sur de supprimer cette enregistrement?", vbYesNo, "Supprimer") = vbYes Then
 Screen.MousePointer = vbHourglass
 
 'Open table
 Set rstProjSoum = New ADODB.Recordset
 Set rstCatalogue = New ADODB.Recordset
 
 Call rstProjSoum.Open("SELECT * FROM GrbSoumission_Pieces WHERE IDFRS = " & m_iNoFournisseur, g_connData, adOpenDynamic, adLockOptimistic)
  Call rstCatalogue.Open("SELECT * FROM GrbPiecesFRS WHERE IDFRS = " & m_iNoFournisseur, g_connData, adOpenDynamic, adLockOptimistic)

 'si existe pas dans soumission
  If rstProjSoum.EOF Then
  Call rstProjSoum.Close
 'si existe pas dans projet
  Call rstProjSoum.Open("SELECT * FROM GrbProjet_Pieces WHERE IDFRS = " & m_iNoFournisseur, g_connData, adOpenDynamic, adLockOptimistic)
 
  If rstProjSoum.EOF Then
 'si existe pas dans la table fournisseurs pour une piece
  If rstCatalogue.EOF Then
  bPeutEffacer = True
  Else
 bPeutEffacer = False
 End If
 Else
 bPeutEffacer = False
 End If
 Else
 bPeutEffacer = False
 End If
 
 Call rstCatalogue.Close
 Set rstCatalogue = Nothing
 
 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 
 If cmbContact.ListCount > 0 Then
 'Delete les contact» pour ce fournisseur
 Call g_connData.Execute("DELETE * FROM GrbContactFRS WHERE NoFRS = " & m_iNoFournisseur)
 End If

 Call SupprimerFRSExchange(m_iNoFournisseur)

 If bPeutEffacer = True Then
 Call g_connData.Execute("DELETE * FROM GrbFournisseur WHERE IDFRS = " & m_iNoFournisseur)
 Else
1  Set rstFRS = New ADODB.Recordset

 Call rstFRS.Open("SELECT * FROM GrbFournisseur WHERE IDFRS = " & m_iNoFournisseur, g_connData, adOpenDynamic, adLockOptimistic)

 rstFRS.Fields("Supprimé") = True

 Call rstFRS.Update

 Call rstFRS.Close
 Set rstFRS = Nothing
 End If

 'Remplir le combo des fournisseurs
 Call RemplirComboFournisseur

 Screen.MousePointer = vbDefault
 End If
Else
 Call MsgBox("Aucun enregistrement sélectionné!")
End If

2  Exit Sub

Oups:

wOups "frmFRS", "CmdSupp_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbFournisseur_Click()

 On Error GoTo Oups
 
 'Quand le user selectionne un enregistrement on se posotionne dessus
 If cmbFournisseur.Text <> vbNullString Then
 txtNomFournisseur.Text = cmbFournisseur.Text
 Else
 If ComboContient(cmbFournisseur, txtNomFournisseur.Text) = False Then
 Call RemplirComboFournisseur
 End If

 cmbFournisseur.Text = txtNomFournisseur.Text
 End If
 
 If cmbFournisseur.ListIndex > -1 Then
 If m_bRenommer = False And m_bModeAjoutFRS = False Then
  m_iNoFournisseur = cmbFournisseur.ItemData(cmbFournisseur.ListIndex)
  End If
  End If
 
 'Affiche le fournisseur sélectionné dans le combo
  Call AfficherFournisseur
  Call RemplirComboContact

  Exit Sub

Oups:

  wOups "frmFRS", "cmbFournisseur_Click", Err, Err.number, Err.Description
End Sub






Private Sub Form_Load()

 On Error GoTo Oups
 
10

 Call tbl_cat_exist 'GLL 2017-11-2  V1.44
 Call FindFieldsExist 'GLL 2017-11-2  V1.44

 Call RemplirComboFournisseur
 Call RemplirComboCatégorie 'GLL 2017-11-2  V1.44
 
 Call HideEdMask(True)
 
 Call AfficherControles(MODE_INACTIF)
 
 Call ActiverBoutonsGroupe

 Screen.MousePointer = vbDefault


 Exit Sub

Oups:

  wOups "frmFRS", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub ActiverBoutonsGroupe()

 On Error GoTo Oups
 
 'Activation des boutons selon le groupe
 CmdAdd.Enabled = g_bModificationFournisseurs
 CmdModif.Enabled = g_bModificationFournisseurs
 cmdrenommer.Enabled = g_bModificationFournisseurs
 CmdSupp.Enabled = g_bModificationFournisseurs
 CmdAddCont.Enabled = g_bModificationFournisseurs
 cmdsupcontact.Enabled = g_bModificationFournisseurs
 cmdMailListContact.Enabled = g_bModificationListeDistribution
 cmdMailListFournisseur.Enabled = g_bModificationListeDistribution
 
 Exit Sub

Oups:

 wOups "frmFRS", "ActiverBoutonsGroupe", Err, Err.number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)

 On Error GoTo Oups

 Set FrmFRS = Nothing

 Exit Sub

Oups:

 wOups "frmFRS", "Form_Unload", Err, Err.number, Err.Description
End Sub

Private Sub Lst_Cat_ItemClick(ByVal Item As MSComctlLib.ListItem)
cmdcatmod.Enabled = True
End Sub



Private Sub mskTelephone_GotFocus()

 On Error GoTo Oups

 mskTelephone.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmFRS", "mskTelephone_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskTelephone_LostFocus()

 On Error GoTo Oups

 mskTelephone.mask = vbNullString

 If mskTelephone.Text = "(___) ___-____" Then
 mskTelephone.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmFRS", "mskTelephone_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskFax_GotFocus()

 On Error GoTo Oups

 mskFax.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmFRS", "mskFax_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskFax_LostFocus()

 On Error GoTo Oups

 mskFax.mask = vbNullString

 If mskFax.Text = "(___) ___-____" Then
 mskFax.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmFRS", "mskFax_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub cmdRechercher_Click()

 On Error GoTo Oups
 
 'Rempli le combo des fournisseurs selon le texte à rechercher
 Dim rstFournisseur As ADODB.Recordset
 Dim sSearch As String
 
 Screen.MousePointer = vbHourglass
 
 sSearch = txtRechercher.Text
 
 'vide les champs
 Call ViderBarrerChamps(True, True)
 
 'Filtre pour selection des Nomcontact
 Set rstFournisseur = New ADODB.Recordset
 
 Call rstFournisseur.Open("SELECT NomFournisseur, IDFRS FROM GrbFournisseur WHERE Instr(1, NomFournisseur, '" & Replace(sSearch, "'", "''") & "') > 0 AND Supprimé = False ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
 
 'vide combo
 Call cmbFournisseur.Clear
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstFournisseur.EOF
 Call cmbFournisseur.AddItem(rstFournisseur.Fields("NomFournisseur"))
  cmbFournisseur.ItemData(cmbFournisseur.newIndex) = rstFournisseur.Fields("IDFRS")
 
  Call rstFournisseur.MoveNext
  Loop
 
  Call rstFournisseur.Close
  Set rstFournisseur = Nothing
 
  If cmbFournisseur.ListCount > 0 Then
  cmbFournisseur.ListIndex = 0
  Else
Call cmbContact.Clear

1 txtContactCell.Text = vbNullString
 txtContactDom.Text = vbNullString
 txtContactEmail.Text = vbNullString
 txtContactFax.Text = vbNullString
 txtContactPage.Text = vbNullString
 txtContactPoste.Text = vbNullString
 txtContactTel.Text = vbNullString
End If
 
Screen.MousePointer = vbDefault

Exit Sub

Oups:

wOups "frmFRS", "cmdRechercher_Click", Err, Err.number, Err.Description
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

 wOups "frmFRS", "txtRechercher_Change", Err, Err.number, Err.Description
End Sub

Private Sub mskContactTel_GotFocus()

 On Error GoTo Oups

 mskContactTel.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmFRS", "mskContactTel_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskContactTel_LostFocus()

 On Error GoTo Oups

 mskContactTel.mask = vbNullString

 If mskContactTel.Text = "(___) ___-____" Then
 mskContactTel.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmFRS", "mskContactTel_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskContactFax_GotFocus()

 On Error GoTo Oups

 mskContactFax.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmFRS", "mskContactFax_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskContactFax_LostFocus()

 On Error GoTo Oups

 mskContactFax.mask = vbNullString

 If mskContactFax.Text = "(___) ___-____" Then
 mskContactFax.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmFRS", "mskContactFax_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskContactCell_GotFocus()

 On Error GoTo Oups

 mskContactCell.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmFRS", "mskContactCell_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskContactCell_LostFocus()

 On Error GoTo Oups

 mskContactCell.mask = vbNullString

 If mskContactCell.Text = "(___) ___-____" Then
 mskContactCell.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmFRS", "mskContactCell_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskContactDom_GotFocus()

 On Error GoTo Oups

 mskContactDom.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmFRS", "mskContactDom_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskContactDom_LostFocus()

 On Error GoTo Oups

 mskContactDom.mask = vbNullString

 If mskContactDom.Text = "(___) ___-____" Then
 mskContactDom.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmFRS", "mskContactDom_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskContactPage_GotFocus()

 On Error GoTo Oups

 mskContactPage.mask = "(###) ###-####"

 Exit Sub

Oups:

 wOups "frmFRS", "mskContactPage_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskContactPage_LostFocus()

 On Error GoTo Oups

 mskContactPage.mask = vbNullString

 If mskContactPage.Text = "(___) ___-____" Then
 mskContactPage.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmFRS", "mskContactPage_LostFocus", Err, Err.number, Err.Description
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

  wOups "frmFRS", "ExisteDansBD", Err, Err.number, Err.Description
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

 wOups "frmFRS", "ContientCaracteresIncorrects", Err, Err.number, Err.Description
End Function
Private Sub AfficherCategorie() 'GLL 2017-11-2  V1.44

 On Error GoTo Oups
 
 ''''''''''''''''''''''''''''''''''''''''
 'affiche les contacts selon leur catégorie'
 ''''''''''''''''''''''''''''''''''''''''
 Dim rstCategorie As ADODB.Recordset
 Dim rstFournisseur As ADODB.Recordset
 Dim i As Integer
 Dim j As Integer
 Dim id As Integer
 Dim sString As String
 Dim cString As String

 'Ouverture de la table contact
 Set rstCategorie = New ADODB.Recordset
 Set rstFournisseur = New ADODB.Recordset
  sString = "Select * From Tbl_Categorie where nom <> Null"
  cmbFournisseur.Clear
  Call rstCategorie.Open(sString, g_connData, adOpenDynamic, adLockOptimistic)

  Do While Not rstCategorie.EOF
  If rstCategorie.Fields("Nom") = cmbcatégorie.Text Then sString = rstCategorie.Fields("Correspondance")
  Call rstCategorie.MoveNext
  Loop
  Call rstCategorie.Close
 Set rstCategorie = Nothing
 
Call rstFournisseur.Open("Select * From GrbFournisseur where categorie <> Null", g_connData, adOpenDynamic, adLockOptimistic)
 
 Do While Not rstFournisseur.EOF
 If InStr(1, rstFournisseur.Fields("Categorie"), sString, vbTextCompare) > 0 Then
 cmbFournisseur.AddItem (rstFournisseur.Fields("NomFournisseur"))
 cmbFournisseur.ItemData(cmbFournisseur.newIndex) = rstFournisseur.Fields("IDFRS")
 End If
 Call rstFournisseur.MoveNext
 Loop
 Call rstFournisseur.Close
 Set rstFournisseur = Nothing

 If cmbFournisseur.ListCount > 0 Then 'Afficher le premier Fournisseur qui est dans cette catégorie
 cmbFournisseur.ListIndex = 0
 Call cmbFournisseur_Click
 End If

Exit Sub

Oups:
 wOups "FrmFrs", "AfficherCategorie", Err, Err.number, Err.Description
End Sub
Private Sub RemplirComboCatégorie()
 On Error GoTo Oups 'GLL 2017-11-2  V1.44
 
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 'remplis le combo contact dépendant le client sélectionné
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Dim rstCategorie As ADODB.Recordset
 
 Set rstCategorie = New ADODB.Recordset
 

 Call rstCategorie.Open("SELECT Nom FROM TBL_Categorie where nom <> Null order by Nom", g_connData, adOpenDynamic, adLockOptimistic)

 
 Call cmbcatégorie.Clear
 
 Do While Not rstCategorie.EOF
 Call cmbcatégorie.AddItem(rstCategorie.Fields("nom"))
  Call rstCategorie.MoveNext
  Loop
 
 'Ferme la table "GrbContact"
  Call rstCategorie.Close
  Set rstCategorie = Nothing
 
  Exit Sub

Oups:

  wOups "frmFRS", "RemplirComboCatégorie", Err, Err.number, Err.Description
End Sub
Private Sub AfficherCatList() 'V1.44 GLL
On Error GoTo Oups
 'Affiche dans Rstlist tout les catégorie enregistrer
 Dim rstlist As ADODB.Recordset
 Dim itemlist As ListItem

 Set rstlist = New ADODB.Recordset
 Call rstlist.Open("Select * from tbl_categorie where nom <> Null order by nom", g_connData, adOpenDynamic, adLockOptimistic)
 Call Lst_Cat.ListItems.Clear

 
 Do While Not rstlist.EOF 'Ajoute dans la liste tout le catégorie trouver
 Set itemlist = Lst_Cat.ListItems.Add
 itemlist.Text = rstlist.Fields("Nom")
 itemlist.Tag = rstlist.Fields("Correspondance")
 Call rstlist.MoveNext
 Loop
Exit Sub
Oups:
  wOups "FrmFRS", "AfficherCatList", Err, Err.number, Err.Description
End Sub
Private Sub DeleteCategorie() 'V1.44 GLL
 'efface une catégorie de tout les fournisseur qui la possêde
 On Error GoTo Oups

 Dim rstCategorie As ADODB.Recordset
 Dim sString As String

 Set rstCategorie = New ADODB.Recordset
 Call rstCategorie.Open("Select categorie from GrbFournisseur where categorie <> Null or categorie ='""'", g_connData, adOpenStatic, adLockPessimistic)

 Do While Not rstCategorie.EOF

 If InStr(1, rstCategorie.Fields("categorie"), m_Tag, vbTextCompare) > 0 Then
 sString = rstCategorie.Fields("categorie")
 sString = Replace(sString, m_Tag, "", 1)
 
 If sString = "" Then
 rstCategorie.Fields("categorie").Value = Null
 Else
  rstCategorie.Fields("categorie").Value = sString
 End If
 End If
  Call rstCategorie.MoveNext
 Loop

  Call rstCategorie.Close
  Set rstCategorie = Nothing
 Exit Sub
Oups:
  wOups "FrmFRS", "DeleteCategorie", Err, Err.number, Err.Description
End Sub

Private Sub tbl_cat_exist() 'V1.44
'Vérifie si la tbl catégorie exist dans la basse de donné si non elle la crée

 On Error GoTo Oups
 Dim adoxconnection As adox.Catalog
 Dim i As Integer

 Set adoxconnection = New adox.Catalog
 adoxconnection.ActiveConnection = g_connData
 For i = 0 To adoxconnection.Tables.count - 1

 If LCase(adoxconnection.Tables(i).Name) = LCase("TBL_Categorie") Then 'Si elle exist on sort de la sous routine
 Set adoxconnection = Nothing
 Exit Sub
 End If
 Next

 Call g_connData.Execute("Create TABLE TBL_Categorie (Correspondance text(1), Nom Text (100))")
 Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('A');")
  Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('B');")
  Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('C');")
  Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('D');")
  Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('E');")
  Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('F');")
  Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('G');")
  Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('H');")
  Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('I');")
10 Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('J');")
Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('K');")
Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('M');")
Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('N');")
Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('O');")
Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('P');")
Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('Q');")
Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('R');")
Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('S');")
Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('T');")
Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('U');")
Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('V');")
1  Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('W');")
Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('X');")
 Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('Y');")
Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('Z');")
 Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('1');")
Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('2');")
 Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('3');")
1  Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('4');")
 Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('5');")
 Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('6');")
Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('7');")
Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('8');")
Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('9');")
Call g_connData.Execute("Insert into TBL_Categorie (Correspondance) Values ('0');")
Set adoxconnection = Nothing
 Exit Sub
Oups:
wOups "FrmFrs", "tbl_Cat_exist", Err, Err.number, Err.Description
End Sub

Private Sub FindFieldsExist() 'V1.44
 On Error GoTo Oups

 Dim strName As String

 Dim Findfield As ADODB.Recordset

 Dim i As Integer
 
 Set Findfield = New ADODB.Recordset

 Call Findfield.Open("Select * from GrbFournisseur", g_connData, adOpenDynamic, adLockOptimistic)

 For i = 0 To Findfield.Fields.count - 1

 strName = Findfield.Fields(i).Name

 If strName = "Categorie" Then

 Call Findfield.Close

 Set Findfield = Nothing

  Exit Sub

  End If

  Next

  Call Findfield.Close

  Set Findfield = Nothing

  Call g_connData.Execute("ALTER TABLE GrbFournisseur Add Categorie Text(40);")
 Exit Sub
 
Oups:
  wOups "frmFRS", "FindFieldsExist()", Err, Err.number, Err.Description
 End Sub
